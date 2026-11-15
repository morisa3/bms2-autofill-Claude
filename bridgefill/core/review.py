"""審查紀錄（第 5 步）：保存使用者對每組配對的確認或改判，可存檔、讀檔、接續。

一組配對以 Pair.key 識別（本期的建議維修工法 ID，或只有前期時的 der_id），
重新掃描、重新配對後仍能對回原本的審查結果。本模組只記錄判斷，不修改網站。
"""
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from .matching import (KIND_KEEP, defect_id, KIND_LARGER, KIND_MISSING, KIND_NEW, KIND_REPAIRED, KIND_SMALLER, KIND_UNIT)

KIND_GONE = '已消失'                 # 例：河道掛淤被河水沖走，非維修
KIND_CONFIRMED_MISSING = '確實漏建'   # 前期缺失仍在，但本期沒有建立
# 審查時可選的判定（配對類與只有前期類分開）
PAIRED_CHOICES = (KIND_KEEP, KIND_LARGER, KIND_SMALLER, KIND_UNIT)
PRIOR_ONLY_CHOICES = (KIND_REPAIRED, KIND_GONE, KIND_CONFIRMED_MISSING, KIND_MISSING)
CURRENT_ONLY_CHOICES = (KIND_NEW,)
PENDING, ACCEPTED, CHANGED = '待審查', '已確認', '已改判'


def choices_for(pair):
    if pair.prior and pair.current:
        return PAIRED_CHOICES
    return PRIOR_ONLY_CHOICES if pair.prior else CURRENT_ONLY_CHOICES


@dataclass
class Decision:
    kind: str                 # 最終判定
    status: str = PENDING     # 待審查／已確認／已改判
    note: str = ''
    at: str = ''


@dataclass
class Review:
    """一期檢測的審查紀錄。decisions：Pair.key → Decision。
    links：人工配對（前期 ID → 本期 ID 清單）；splits：人工拆開的 [前期 ID, 本期 ID]。"""
    bridge: str = ''
    current_date: str = ''
    prior_date: str = ''
    decisions: dict = field(default_factory=dict)
    links: dict = field(default_factory=dict)
    splits: list = field(default_factory=list)
    edits: dict = field(default_factory=dict)       # 本期 Pair.key → {欄位: 修改後的值}（第 6 步填寫用）
    kept: dict = field(default_factory=dict)        # 本期 Pair.key → 確認維持網站現值的時間（填寫「需確認」已處理）
    opinion: dict = field(default_factory=dict)     # 檢測員意見草稿（文字、勾選、待選項目）

    def sync(self, pairs):
        """配對結果更新後：保留已審查的判斷；新出現的配對以程式判定、待審查；高信心者視為已確認；
        人工配對視為已改判。已存的判定若不適用於目前的配對（配對方式改變），改回程式判定重新審查。"""
        for pair in pairs:
            old = self.decisions.get(pair.key)
            if old is not None and old.kind in choices_for(pair):
                continue
            if pair.manual:
                self.decisions[pair.key] = Decision(pair.kind, CHANGED, '人工配對', datetime.now().isoformat(timespec='seconds'))
            else:
                self.decisions[pair.key] = Decision(pair.kind, ACCEPTED if pair.confidence == '高' else PENDING)
        return self

    def set_edit(self, pair, changes):
        """記錄本期欄位的修改；只保留與網站現值不同的欄位，全部相同時清除。"""
        if not pair.current:
            raise ValueError('只有前期的缺失不能修改')
        kept = {k: str(v).strip() for k, v in changes.items()
                if str(v).strip() != site_value(pair.current, k)}
        if kept:
            self.edits[pair.key] = kept
        else:
            self.edits.pop(pair.key, None)
        return kept

    def keep(self, pair, on=True):
        """填寫清單的「需確認」：人工確認維持網站現值，不再顯示為未確認。"""
        if on:
            self.kept[pair.key] = datetime.now().isoformat(timespec='seconds')
        else:
            self.kept.pop(pair.key, None)

    def edit_of(self, pair):
        return self.edits.get(pair.key, {}) if pair.current else {}

    def value(self, pair, name):
        """本期欄位修改後的值（沒改則為網站現值）。"""
        if not pair.current:
            return ''
        return self.edit_of(pair).get(name, site_value(pair.current, name))

    def forget(self, pairs, prior_ids=(), current_ids=()):
        """配對方式改變：清除牽涉到這些缺失的配對的審查結果，之後由 sync 重新給定。"""
        for pair in pairs:
            if defect_id(pair.prior) in prior_ids or defect_id(pair.current) in current_ids:
                self.decisions.pop(pair.key, None)

    def link(self, pairs, prior_id, current_ids):
        """人工配對：prior_id 對應 current_ids（取代這筆前期原本的人工配對）。"""
        current_ids = list(dict.fromkeys(current_ids))
        old = self.links.get(prior_id, [])
        for pid in list(self.links):
            self.links[pid] = [c for c in self.links[pid] if c not in current_ids]
            if not self.links[pid]:
                del self.links[pid]
        if current_ids:
            self.links[prior_id] = current_ids
        self.splits = [s for s in self.splits if s[0] != prior_id and s[1] not in current_ids]
        self.forget(pairs, {prior_id}, set(current_ids) | set(old))

    def split(self, pairs, pair):
        """拆開一組配對：兩筆各自單獨，之後也不再自動配在一起。"""
        if not (pair.prior and pair.current):
            raise ValueError('這筆沒有配對，不需拆開')
        pid, cid = defect_id(pair.prior), defect_id(pair.current)
        if cid in self.links.get(pid, []):
            self.links[pid] = [c for c in self.links[pid] if c != cid]
            if not self.links[pid]:
                del self.links[pid]
        if [pid, cid] not in self.splits:
            self.splits.append([pid, cid])
        self.forget(pairs, {pid}, {cid})

    def decide(self, pair, kind=None, note=''):
        kind = kind or pair.kind
        if kind not in choices_for(pair):
            raise ValueError(f'「{kind}」不適用於這組配對')
        status = ACCEPTED if kind == pair.kind else CHANGED
        self.decisions[pair.key] = Decision(kind, status, note, datetime.now().isoformat(timespec='seconds'))
        return self.decisions[pair.key]

    def decision(self, pair):
        return self.decisions.get(pair.key) or Decision(pair.kind)

    def counts(self, pairs):
        result = {PENDING: 0, ACCEPTED: 0, CHANGED: 0}
        for pair in pairs:
            result[self.decision(pair).status] += 1
        return result

    # ------------------------------------------------------------ 存檔
    def to_dict(self):
        return {'version': 2, 'bridge': self.bridge, 'current_date': self.current_date, 'prior_date': self.prior_date,
                'decisions': {key: asdict(value) for key, value in self.decisions.items()},
                'links': self.links, 'splits': self.splits, 'edits': self.edits,
                'kept': self.kept, 'opinion': self.opinion}

    def save(self, path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix('.tmp')
        temporary.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(path)    # 先寫暫存檔再取代，避免存到一半中斷造成檔案損毀
        return path

    @classmethod
    def load(cls, path):
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        return cls(data.get('bridge', ''), data.get('current_date', ''), data.get('prior_date', ''),
                   {key: Decision(**value) for key, value in data.get('decisions', {}).items()},
                   data.get('links', {}), [list(s) for s in data.get('splits', [])], data.get('edits', {}), data.get('kept', {}), data.get('opinion', {}))


def site_value(defect, name):
    """網站上的現值；D、E、R、U 欄位空白時取第一級的 DERU。"""
    value = str(defect.fields.get(name, '') or '').strip()
    if not value and name in ('d', 'e', 'r', 'u'):
        value = (defect.deru + '____')['deru'.index(name)]
    return value


def review_path(folder, bridge, current_date):
    safe = ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in f'{current_date}-{bridge}')
    return Path(folder) / f'{safe}-審查.json'


def saved_sessions(scan_dir, review_dir):
    """之前掃描過的橋（新到舊）：本期掃描檔、審查進度（待審查、已修改筆數）、最後更新時間。"""
    found = []
    for path in Path(scan_dir).glob('*-檢測.json') if Path(scan_dir).is_dir() else []:
        try:
            header = json.loads(path.read_text(encoding='utf-8')).get('header', {})
        except (OSError, ValueError):
            continue
        bridge, date = header.get('橋梁名稱', ''), header.get('檢測日期', '')
        review_file = review_path(review_dir, bridge, date)
        pending = edits = None
        updated = path.stat().st_mtime
        if review_file.exists():
            try:
                review = Review.load(review_file)
                pending = sum(1 for d in review.decisions.values() if d.status == PENDING)
                edits = len(review.edits)
                updated = max(updated, review_file.stat().st_mtime)
            except (OSError, ValueError, TypeError):
                pass
        filled = unsure = None
        fill_file = review_file.with_name(review_file.name.replace('-審查.json', '-填寫.json'))
        if fill_file.exists():
            try:
                entries = json.loads(fill_file.read_text(encoding='utf-8')).get('entries', {})
                filled = sum(1 for e in entries.values() if e.get('status') == '已完成')
                unsure = sum(1 for e in entries.values() if e.get('status') == '待確認')
                updated = max(updated, fill_file.stat().st_mtime)
            except (OSError, ValueError):
                pass
        found.append({'bridge': bridge, 'date': date, 'path': path, 'pending': pending, 'edits': edits,
                      'filled': filled, 'unsure': unsure, 'updated': datetime.fromtimestamp(updated)})
    return sorted(found, key=lambda item: item['updated'], reverse=True)


def session_files(item, scan_dir, review_dir, sessions=()):
    """一筆歷史紀錄的檔案：本期掃描、基本資料、審查、填寫進度；前期掃描只在同橋沒有其他紀錄時一併刪除。"""
    bridge, date = item['bridge'], item['date']
    safe = lambda text: ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in text)  # noqa: E731
    scan_dir, review = Path(scan_dir), review_path(review_dir, bridge, date)
    files = [Path(item['path']), scan_dir / safe(f'{date}-{bridge}-基本資料.json'), review,
             review.with_name(review.name.replace('-審查.json', '-填寫.json'))]
    others = [s for s in sessions if s['bridge'] == bridge and Path(s['path']) != Path(item['path'])]
    if not others:
        files += sorted(scan_dir.glob(f'*-{safe(bridge)}-前期.json'))
    return [f for f in dict.fromkeys(files) if f.exists()]


def clear_sessions(items, scan_dir, review_dir):
    """刪除選取的歷史紀錄；回傳刪除的檔案。"""
    sessions = saved_sessions(scan_dir, review_dir)
    chosen = {Path(i['path']) for i in items}
    remaining = [s for s in sessions if Path(s['path']) not in chosen]
    removed = []
    for item in items:
        for path in session_files(item, scan_dir, review_dir, remaining):
            path.unlink(missing_ok=True)
            removed.append(path)
    return removed
