"""排程掃描核心邏輯（不依賴 Qt）：
1. 由本機橋梁總表（.xlsx）或 Google 試算表讀取各鄉鎮工作表的橋梁清單（含二代系統ID、是否已完成）。
2. 重複橋名自動分流（disambiguate_bridge）：
   優先序：
   (1) 總表二代系統ID 精準比對（若命中名稱含「已拆除」，標記警告且不自動掃描）；
   (2) 排除名稱含「已拆除」的候選；
   (3) 以總表工作表名稱（鄉鎮市區）比對網站 bridge.place；
   (4) 備用方案（仍有未標註已拆除的多座同名橋時）：依 2 年一檢週期判斷活躍橋梁，
       但若疑似「改建新橋（無前期或僅本年度）與未標註舊橋並存」，保守判定為需人工確認，絕不誤判。
3. 監看時段輪詢佇列（WatchQueue）：
   支援設定監看時段（如 13:00～18:00）、檢查間隔、以及上傳穩定緩衝（偵測到本期後等 5～10 分鐘且缺失筆數未變動才掃描），
   並提供「立即掃描剩餘排程」模式供提早收工時馬上執行。
"""
import re
from dataclasses import dataclass, field
from datetime import datetime, time
from pathlib import Path

from openpyxl import load_workbook

from .bridge_list import DONE_HEAD, key as header_key

DEMOLISHED_RE = re.compile(r'已拆除|拆除|已廢除|廢棄|停用')
TOWNSHIP_SUFFIX_RE = re.compile(r'(?:鄉|鎮|市|區)$')
SHEET_NOISE_RE = re.compile(r'[\(（][^()（）]*[\)）]|\d+年?|\s+')
SKIP_SHEETS = frozenset({'範本', '說明', '總計', '統計', '摘要', '下拉選單', '參數'})


# ---------------------------------------------------------------- 總表橋梁目錄讀取

def _clean_id(value):
    text = str(value or '').strip()
    if re.fullmatch(r'\d+\.0+', text):
        text = text.split('.', 1)[0]
    return text if text.isdigit() else ''


def load_list_bridges(local_path=None, google_client=None):
    """從本機 Excel 總表或 Google 試算表讀取全部工作表的橋梁目錄。

    回傳 {'sheets': [sheet_name, ...],
          'bridges': [{'sheet': str, 'name': str, 'inventory_id': str, 'done': bool, 'row': int}, ...]}
    """
    if google_client is not None:
        raw = google_client.list_bridges()
        sheets, bridges = [], []
        if isinstance(raw, dict) and 'sheets' in raw and 'bridges' in raw:
            sheets = list(raw.get('sheets') or [])
            items = raw.get('bridges') or []
        elif isinstance(raw, dict):
            # 容錯：Code.gs list_bridges 可能回傳 {'sheets': [{'name': ..., 'bridges': [...]}]}
            items = []
            for entry in raw.get('sheets') or []:
                if isinstance(entry, dict):
                    sname = str(entry.get('name') or entry.get('sheet') or '').strip()
                    if sname:
                        sheets.append(sname)
                    for b in entry.get('bridges') or []:
                        items.append({**b, 'sheet': b.get('sheet') or sname})
        else:
            items = list(raw or [])

        for b in items:
            sheet = str(b.get('sheet') or '').strip()
            name = str(b.get('name') or b.get('橋梁名稱') or '').strip()
            if not name:
                continue
            if sheet and sheet not in sheets:
                sheets.append(sheet)
            bridges.append({
                'sheet': sheet,
                'name': name,
                'inventory_id': _clean_id(b.get('inventory_id') or b.get('id') or b.get('二代系統ID')),
                'done': bool(b.get('done') or str(b.get(DONE_HEAD) or '').strip()),
                'row': int(b.get('row') or 0),
            })
        return {'sheets': sheets, 'bridges': bridges}

    if not local_path:
        return {'sheets': [], 'bridges': []}

    path = Path(local_path)
    if not path.is_file():
        return {'sheets': [], 'bridges': []}

    wb = load_workbook(path, read_only=True, data_only=True)
    sheets, bridges = [], []
    try:
        name_k = header_key('橋梁名稱')
        id_k = header_key('二代系統ID')
        done_k = header_key(DONE_HEAD)
        for ws in wb.worksheets:
            if ws.title.strip() in SKIP_SHEETS:
                continue
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            header_idx, col_map = None, {}
            for ridx, row in enumerate(rows[:5]):
                mapped = {header_key(c): i for i, c in enumerate(row) if c not in (None, '')}
                if name_k in mapped:
                    header_idx, col_map = ridx, mapped
                    break
            if header_idx is None:
                continue
            ni = col_map[name_k]
            ii = col_map.get(id_k)
            di = col_map.get(done_k)
            sheet_entries = []
            for offset, row in enumerate(rows[header_idx + 1:], header_idx + 2):
                bname = str(row[ni] or '').strip() if ni < len(row) else ''
                if not bname or bname in ('橋梁名稱', '合計', '總計'):
                    continue
                bid = _clean_id(row[ii]) if ii is not None and ii < len(row) else ''
                done = bool(str(row[di] or '').strip()) if di is not None and di < len(row) else False
                sheet_entries.append({
                    'sheet': ws.title,
                    'name': bname,
                    'inventory_id': bid,
                    'done': done,
                    'row': offset,
                })
            if sheet_entries:
                sheets.append(ws.title)
                bridges.extend(sheet_entries)
    finally:
        wb.close()
    return {'sheets': sheets, 'bridges': bridges}


# ---------------------------------------------------------------- 重複橋名智慧分流

def is_demolished(name):
    return bool(DEMOLISHED_RE.search(str(name or '')))


def normalize_township(text):
    """擷取鄉鎮市區核心詞：'北港鎮(115)' → '北港'、'雲林縣斗六市' → '斗六'。"""
    cleaned = SHEET_NOISE_RE.sub('', str(text or '').strip())
    if not cleaned:
        return ''
    # 去掉前綴縣市名（例：雲林縣北港鎮 → 北港鎮）
    cleaned = re.sub(r'^.{2,3}[縣市]', '', cleaned) or cleaned
    return TOWNSHIP_SUFFIX_RE.sub('', cleaned)


def matches_township(place, township):
    """網站 bridge.place（例：雲林縣北港鎮）是否符合總表工作表鄉鎮名（例：北港鎮／北港）。"""
    t = normalize_township(township)
    p = str(place or '').strip()
    if not t or not p:
        return False
    return t in p


def _record_years(bridge):
    years = []
    for rec in getattr(bridge, 'records', None) or ():
        m = re.match(r'(\d{4})', str(getattr(rec, 'date', '') or ''))
        if m:
            years.append(int(m[1]))
    return sorted(set(years), reverse=True)


@dataclass
class DisambiguationResult:
    bridge: object | None = None
    method: str = ''               # exact / id / exclude_demolished / township / cycle_date / ambiguous / not_found
    ambiguous: bool = False
    warning: str = ''
    reason: str = ''


def disambiguate_bridge(candidates, wanted_name='', inventory_id='', township='', current_year=None):
    """重複橋名防誤判分流。

    優先順序（依使用者確認之規則）：
    1. 若有提供二代系統ID（inventory_id），精準比對 ID；若命中名稱含「已拆除」，回傳警告與 ambiguous=True。
    2. 過濾名稱含「已拆除」的候選；若只剩一座（或恰好一座完全同名），直接採用。
    3. 若提供鄉鎮（township），以 bridge.place 過濾；若剩一座直接採用。
    4. 備用方案（未標註已拆除且仍有多座）：檢視 2 年一檢週期檢測年度：
       - 若存在疑似「改建新橋（無歷史紀錄或僅本年度）」與「活躍舊橋（近 3 年內有紀錄）」並存，
         保守回傳 ambiguous=True，避免誤判。
       - 若恰有一座為近期活躍橋（最新檢測 >= current_year - 2）且其餘全為已停檢舊橋（最新檢測 <= current_year - 4），
         則選定該活躍橋。
       - 其餘無法確定的情況一律回傳 ambiguous=True，交人工確認。
    """
    candidates = list(candidates or [])
    wanted = str(wanted_name or '').strip()
    target_id = _clean_id(inventory_id)
    year = int(current_year or datetime.now().year)

    if not candidates:
        return DisambiguationResult(None, 'not_found', False, '', f'查無「{wanted or target_id}」的檢測資料')

    # 1. 總表二代系統ID 精準比對
    if target_id:
        by_id = [b for b in candidates if _clean_id(getattr(b, 'inventory_id', '')) == target_id]
        if len(by_id) == 1:
            hit = by_id[0]
            if is_demolished(getattr(hit, 'name', '')):
                return DisambiguationResult(
                    None, 'ambiguous', True,
                    f'總表 ID {target_id} 對應橋梁「{hit.name}」標註已拆除',
                    f'ID {target_id} 對應已拆除橋梁「{hit.name}」，請人工確認總表 ID 是否需更新',
                )
            return DisambiguationResult(hit, 'id', False, '', f'依總表二代系統ID {target_id} 精準對應')

    # 若未指定 ID 且本來就只有一座且完全同名（且未標已拆除）
    if len(candidates) == 1 and (not wanted or getattr(candidates[0], 'name', '').strip() == wanted):
        only = candidates[0]
        if is_demolished(getattr(only, 'name', '')):
            return DisambiguationResult(
                None, 'ambiguous', True,
                f'唯一結果「{only.name}」標註已拆除',
                f'「{wanted}」查到的唯一橋梁標註已拆除，請人工確認',
            )
        return DisambiguationResult(only, 'exact', False, '', '唯一符合橋梁')

    # 2. 排除名稱含「已拆除」的候選
    active = [b for b in candidates if not is_demolished(getattr(b, 'name', ''))]
    removed_count = len(candidates) - len(active)
    if not active:
        return DisambiguationResult(
            None, 'ambiguous', True,
            f'查到的 {len(candidates)} 座橋均標註已拆除',
            f'「{wanted}」查到的橋梁均標註已拆除，請人工確認',
        )

    # 在未拆除候選中，若有精準同名優先縮小（避免「仁愛橋」混入「仁愛橋(2)」或「新仁愛橋」）
    if wanted:
        exact = [b for b in active if getattr(b, 'name', '').strip() == wanted]
        if exact:
            active = exact

    if len(active) == 1:
        method = 'exclude_demolished' if removed_count > 0 else 'exact'
        reason = f'已排除 {removed_count} 座「已拆除」橋梁後唯一符合' if removed_count > 0 else '唯一同名橋梁'
        return DisambiguationResult(active[0], method, False, '', reason)

    # 3. 以總表鄉鎮（sheet）比對 bridge.place
    if township:
        by_town = [b for b in active if matches_township(getattr(b, 'place', ''), township)]
        if len(by_town) == 1:
            return DisambiguationResult(by_town[0], 'township', False, '',
                                        f'依總表鄉鎮「{township}」對應（{by_town[0].place}）')
        if by_town:
            active = by_town

    # 4. 備用方案：2 年一檢週期日期判定（防改建新舊橋誤判）
    years_map = {id(b): _record_years(b) for b in active}
    # 檢查是否有「疑似改建新橋（無前期紀錄，只有本年度或空）」與「近 3 年仍有紀錄的舊橋」並存
    new_like = [b for b in active if not [y for y in years_map[id(b)] if y < year]]
    recent_with_prior = [b for b in active if any(year - 3 <= y < year for y in years_map[id(b)])]
    if new_like and recent_with_prior:
        return DisambiguationResult(
            None, 'ambiguous', True,
            '疑似改建新橋與未標註舊橋並存',
            f'「{wanted}」同時有新橋（無前期紀錄）與近期舊橋，無法自動判斷，需人工確認',
        )

    # 活躍橋（近 2 年有檢：>= year - 2）vs 明顯停檢舊橋（最新檢測 <= year - 4，例 2026 年對 2020 年）
    recent_active = [b for b in active if years_map[id(b)] and years_map[id(b)][0] >= year - 2]
    retired = [b for b in active if years_map[id(b)] and years_map[id(b)][0] <= year - 4]
    if len(recent_active) == 1 and len(retired) == len(active) - 1:
        chosen = recent_active[0]
        latest_y = years_map[id(chosen)][0]
        return DisambiguationResult(
            chosen, 'cycle_date', False, '',
            f'依 2 年一檢週期判定：選定活躍橋梁（最新 {latest_y} 年，其餘最新皆在 {year - 4} 年前）',
        )

    return DisambiguationResult(
        None, 'ambiguous', True, '',
        f'「{wanted or target_id}」仍有 {len(active)} 座無法自動區分的橋梁，請人工確認',
    )


# ---------------------------------------------------------------- 監看時段輪詢佇列

STATUS_WAITING = 'waiting'
STATUS_STABILIZING = 'stabilizing'
STATUS_SCANNING = 'scanning'
STATUS_DONE = 'done'
STATUS_AMBIGUOUS = 'ambiguous'
STATUS_FAILED = 'failed'

STATUS_ICON = {
    STATUS_WAITING: '🕒',
    STATUS_STABILIZING: '⏳',
    STATUS_SCANNING: '🔄',
    STATUS_DONE: '✅',
    STATUS_AMBIGUOUS: '⚠',
    STATUS_FAILED: '❌',
}

STATUS_LABEL = {
    STATUS_WAITING: '等待上傳',
    STATUS_STABILIZING: '等待上傳穩定',
    STATUS_SCANNING: '掃描中',
    STATUS_DONE: '已掃完',
    STATUS_AMBIGUOUS: '同名需人工確認',
    STATUS_FAILED: '掃描失敗',
}


@dataclass
class WatchItem:
    name: str
    inventory_id: str = ''
    sheet: str = ''
    status: str = STATUS_WAITING
    seen_count: int | None = None
    seen_at: datetime | None = None
    scanned_at: datetime | None = None
    message: str = '等待本期上傳'

    @property
    def key(self):
        return self.inventory_id or f'{self.sheet}:{self.name}' if self.sheet else self.name

    @property
    def display(self):
        icon = STATUS_ICON.get(self.status, '•')
        label = STATUS_LABEL.get(self.status, self.status)
        meta = '／'.join(x for x in (self.sheet, f'ID {self.inventory_id}' if self.inventory_id else '') if x)
        head = f'{icon} {self.name}' + (f'（{meta}）' if meta else '')
        return f'{head} — {label}' + (f'：{self.message}' if self.message and self.message != label else '')


@dataclass
class WatchQueue:
    """監看時段輪詢狀態機：解決提早收工、App 分批上傳尚未傳完、或排太早漏掃的問題。"""
    start_time: time = field(default_factory=lambda: time(13, 0))
    end_time: time = field(default_factory=lambda: time(18, 0))
    interval_minutes: int = 10
    stabilize_minutes: int = 5
    enabled: bool = False
    items: list[WatchItem] = field(default_factory=list)
    last_check: datetime | None = None

    def add_items(self, entries):
        """加入監看項目（dict 或 WatchItem），依 key 去重。回傳實際新增的清單。"""
        existing = {item.key for item in self.items}
        added = []
        for entry in entries:
            item = entry if isinstance(entry, WatchItem) else WatchItem(
                name=str(entry.get('name') or '').strip(),
                inventory_id=_clean_id(entry.get('inventory_id') or entry.get('id')),
                sheet=str(entry.get('sheet') or '').strip(),
            )
            if not item.name and not item.inventory_id:
                continue
            if item.key in existing:
                continue
            existing.add(item.key)
            self.items.append(item)
            added.append(item)
        return added

    def remove_at(self, index):
        if 0 <= index < len(self.items):
            return self.items.pop(index)
        return None

    def clear(self):
        self.items.clear()
        self.last_check = None

    def pending_items(self):
        return [item for item in self.items if item.status in (STATUS_WAITING, STATUS_STABILIZING)]

    def is_within_window(self, now=None):
        now = now or datetime.now()
        t = now.time()
        if self.start_time <= self.end_time:
            return self.start_time <= t <= self.end_time
        return t >= self.start_time or t <= self.end_time

    def is_due(self, now=None):
        now = now or datetime.now()
        if not self.enabled or not self.pending_items() or not self.is_within_window(now):
            return False
        if self.last_check is None:
            return True
        return (now - self.last_check).total_seconds() >= max(1, self.interval_minutes) * 60

    def evaluate_probe(self, item, has_current, defect_count=0, now=None, force_immediate=False):
        """依探測結果更新項目狀態，回傳是否應立刻執行完整掃描（bool）。"""
        now = now or datetime.now()
        if not has_current:
            item.status = STATUS_WAITING
            item.seen_count = None
            item.seen_at = None
            item.message = '尚未查到本年度檢測紀錄，稍後自動重試'
            return False

        if force_immediate or self.stabilize_minutes <= 0:
            item.status = STATUS_SCANNING
            item.message = '開始掃描本期'
            return True

        # 上傳穩定緩衝：首次偵測到本期，或缺失筆數仍在增加中 → 重新計時
        if item.seen_count is None or item.seen_count != defect_count:
            prev = item.seen_count
            item.seen_count = defect_count
            item.seen_at = now
            item.status = STATUS_STABILIZING
            if prev is None:
                item.message = f'已查到本期（{defect_count} 筆缺失），等待 {self.stabilize_minutes} 分鐘確認上傳完成'
            else:
                item.message = f'缺失筆數變動（{prev} → {defect_count}），重新等待 {self.stabilize_minutes} 分鐘'
            return False

        elapsed = (now - item.seen_at).total_seconds() if item.seen_at else 0
        need = self.stabilize_minutes * 60
        if elapsed >= need:
            item.status = STATUS_SCANNING
            item.message = f'缺失筆數已穩定（{defect_count} 筆），開始掃描'
            return True

        remain = max(1, int((need - elapsed + 59) // 60))
        item.status = STATUS_STABILIZING
        item.message = f'已查到本期（{defect_count} 筆），尚需等待約 {remain} 分鐘確認穩定'
        return False

    def mark_done(self, item, now=None, message=''):
        item.status = STATUS_DONE
        item.scanned_at = now or datetime.now()
        item.message = message or f'已於 {item.scanned_at:%H:%M} 掃描完成'

    def mark_ambiguous(self, item, message):
        item.status = STATUS_AMBIGUOUS
        item.message = message

    def mark_failed(self, item, message):
        item.status = STATUS_FAILED
        item.message = message

    def status_summary(self):
        if not self.items:
            return ''
        counts = {}
        for item in self.items:
            counts[item.status] = counts.get(item.status, 0) + 1
        parts = []
        for st in (STATUS_DONE, STATUS_STABILIZING, STATUS_WAITING, STATUS_AMBIGUOUS, STATUS_FAILED):
            if counts.get(st):
                parts.append(f'{STATUS_ICON[st]} {STATUS_LABEL[st]} {counts[st]}')
        return '｜'.join(parts)
