"""縣市橋梁清單同步（2026-11-23 使用者需求）：把一座橋的檢測整理寫進使用者指定路徑的清單 Excel。

欄位（依表頭名稱找欄，不依位置；年度與換行不影響）：
- 核對（系統資料）：二代系統ID、橋梁編號、橋梁名稱、道路等級、路線、跨數、橋長、最大淨寬、跨越物體、欄杆高度——
  清單空白直接填入；已有值但與系統不同：標橘底並加備註寫系統值，不覆寫。
- 程式整理（每次同步更新）：
  前期觀測劣化 U=1／U=2／U≧3：前期有、本期仍存在的缺失（本期尚未建立時：前期全部缺失），以檢測員意見的寫法
  （同樣態合併、位置縮寫）去掉「建議…」；以前期為主：敘述（套 Excel 規則修正）、U、年份都用前期；
  本期新增劣化 U=1／U=2／U≧3：本期才有的缺失；
  過去維修紀錄：網站歷年維修紀錄（依年度）＋本期標「已維修」的前期缺失；
  已維修構件狀況：前期以前的維修，原構件在之後的檢測又出現劣化（例：110年維修 113年原構件出現劣化）；
  前期淨高（清單空白才填）、本期淨高（本期意見有變化才填，橘底提醒）；
  應維修構件、應維修處理建議（U≧3）、河工設施損壞、護欄高度不足、橋下堆積物、疑似鹽害橋梁、六張基本照片拍攝；
  系統填寫完成（日期；橋名綠底）。
- 不動：序號、下橋路徑、輔助設備、基本資料變動、橋面板厚、大梁高度、橋名牌、限重牌、其他檢測特殊情形與公式欄。
寫入前備份；檔案開著（被 Excel 鎖住）時回報錯誤，不寫入。本模組不依賴 Qt。
"""
import math
import re
import shutil
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .opinion import Entry, compress, display_name, locations, number, pattern_of

DONE_HEAD = '系統填寫完成'
CHECK_FILL = 'FFF8CBAD'          # 與系統不同、本期淨高有變化：橘底
DONE_FILL = 'FFB7E1CD'           # 系統填寫完成：橋名綠底
RIVER_PARTS = ('河道', '護床工', '橋墩或橋基保護設施', '橋墩保護設施', '橋基保護設施')
RIVER_TEXT = ('護岸', '護床', '固床', '河道', '護坦', '保護工')


def key(text):
    """表頭比對：去空白、全形、年度、括號說明（前期梁下/橋下淨高(m)(113年) → 前期梁下/橋下淨高(m)）。"""
    text = unicodedata.normalize('NFKC', str(text or ''))
    text = re.sub(r'\s+', '', text)
    text = re.sub(r'\(?\d{2,3}年\)?', '', text)
    text = text.replace('須修正', '').replace('檢測新增', '新增').replace('≦', '<=').replace('≧', '>=')
    return re.sub(r'\(列出過去維修紀錄\)|\(橋長\*淨寬\)|\(V\)', '', text)


COLUMNS = {   # 程式使用的名稱 → 清單表頭（key 比對）
    'id': '二代系統ID', 'code': '橋梁編號', 'name': '橋梁名稱', 'grade': '道路等級', 'route': '路線', 'spans': '跨數',
    'length': '橋長', 'width': '最大淨寬', 'crossover': '跨越物體', 'railing': '欄杆高度(m)',
    'prior1': '前期觀測劣化(U=1)', 'prior2': '前期觀測劣化(U=2)', 'prior3': '前期觀測劣化(U≧3)',
    'new1': '新增劣化(U=1)', 'new2': '新增劣化(U=2)', 'new3': '新增劣化(U≧3)',
    'repairs': '過去維修紀錄', 'repaired': '已維修構件狀況',
    'height_prior': '前期梁下/橋下淨高(m)', 'height_now': '本期梁下/橋下淨高(m)',
    'members': '應維修構件(ch5.2)', 'advice': '應維修處理建議', 'river': '河工設施損壞(ch5.4)',
    'guard': '護欄高度不足(ch5.5)', 'waste': '橋下堆積物(ch5.6)', 'salt': '疑似鹽害橋梁',
    'photos': '六張基本照片拍攝', 'done': DONE_HEAD,
}
TICKS = ('guard', 'waste', 'salt', 'photos')     # 是／否欄位（2026-11-24 使用者：直接填「是」「否」，不用 V）


def yes_no(value, known):
    """是否判讀：known 為 False（判斷不出來）回傳空白，否則「是」或「否」。"""
    return ('是' if value else '否') if known else ''
CHECKED = ('id', 'code', 'name', 'grade', 'route', 'spans', 'length', 'width', 'crossover', 'railing')


def roc_year(date):
    found = re.match(r'(\d{4})', str(date or ''))
    return int(found[1]) - 1911 if found else None


# ---------------------------------------------------------------- 劣化整理

def entry_of(defect, key_text=''):
    """matching.Defect → opinion.Entry（U 用欄位值，沒有用 DERU 第 4 碼）。"""
    fields = dict(defect.fields)
    deru = str(defect.deru or '')
    u = number(fields.get('u') or (deru[3:4] if len(deru) >= 4 else 0))
    return Entry(key_text or f'{defect.period}:{defect.method_id or id(defect)}', defect.component, fields,
                 number(fields.get('d') or deru[:1]), number(fields.get('e') or deru[1:2]), u, {}, dict(defect.context))


def described(entry, prior=None):
    """備註沒有寫樣態（本期只填類型、備註空白）時改用前期的位置與敘述，U、DERU 仍用本期。"""
    from .editing import split_remark
    if prior is None or split_remark(entry.remark)[1].strip(' 。.'):
        return entry
    old = dict(prior.fields)
    if not split_remark(old.get('remark', ''))[1].strip(' 。.'):
        return entry
    fields = dict(entry.fields, remark=old.get('remark', ''), d_location=old.get('d_location', '') or entry.fields.get('d_location', ''))
    return Entry(entry.key, entry.component, fields, entry.d, entry.e, entry.u, old, entry.context)


def member_lines(entries):
    """缺失 → 「A1河道混凝土裂縫」（位置＋構件＋樣態，同檢測員意見寫法；不標年份）。"""
    lines = []
    for entry in entries:
        pattern, codes = pattern_of(entry), locations(entry)
        lines.append(compress(codes) + display_name(entry.component, pattern, codes, entry.context) + pattern)
    return lines


def u_bucket(u):
    return 1 if u <= 1 else 2 if u == 2 else 3


def defect_lines(items):
    """[(Entry, 年度)] → {1: 'U=1 文字', 2: ..., 3: ...}；檢測員意見的寫法去掉建議：同 U 分欄、同樣態合併、
    位置縮寫（S1~S3），每條標年份（合併的取最早年份）。沒有的欄位為「無」。"""
    groups = {1: {}, 2: {}, 3: {}}
    for entry, year in items:
        pattern = pattern_of(entry)
        codes = locations(entry)
        group = groups[u_bucket(entry.u)].setdefault(pattern, {'parts': {}, 'years': set()})
        name = display_name(entry.component, pattern, codes, entry.context)
        part = group['parts'].setdefault(name, [])
        part += [c for c in codes if c not in part]
        if year:
            group['years'].add(year)
    result = {}
    for bucket, found in groups.items():
        lines = []
        for pattern, group in found.items():
            subject = '、'.join(compress(codes) + name for name, codes in group['parts'].items() if codes or name)
            year = f"({min(group['years'])}年)" if group['years'] else ''
            lines.append(f'{subject}{pattern}{year}')
        result[bucket] = numbered(lines)
    return result


def numbered(lines, empty='無'):
    lines = [line for line in dict.fromkeys(lines) if line]
    return '\n'.join(f'{i}.{line}' for i, line in enumerate(lines, 1)) if lines else empty


# ---------------------------------------------------------------- 維修紀錄

def position_code(record):
    """維修紀錄的構件位置（a002 → A2、p003 → P3、s001 → S1）；讀不出來回傳空白。"""
    for value in (record.position, record.number):
        found = re.fullmatch(r'([aApPsS])0*(\d+)', str(value or '').strip())
        if found:
            return f'{found[1].upper()}{int(found[2])}'
    return ''


def same_member(component, record):
    """維修紀錄的構件名稱（翼牆／擋土牆）是否就是這個構件（橋台或翼牆或擋土牆）。"""
    name, target = re.sub(r'\s', '', record.component), re.sub(r'\s', '', component or '')
    return bool(name) and (name == target or any(part and part in target for part in re.split(r'[／/、]', name)))


def repair_text(records, repaired_now=()):
    """過去維修紀錄：依年度（民國）列出：「110 換行 1.翼牆／擋土牆 A2 混凝土修復」；
    repaired_now：[(年度, 文字)] 本期標「已維修」的前期缺失。沒有回傳「無」。"""
    years = {}
    for record in records:
        if record.year is None:
            continue
        text = ' '.join(x for x in (record.component, position_code(record), record.method) if x)
        years.setdefault(record.year, []).append(text)
    for year, text in repaired_now:
        years.setdefault(year, []).append(text)
    if not years:
        return '無'
    blocks = []
    for year in sorted(years):
        lines = list(dict.fromkeys(years[year]))
        blocks.append(f'{year}\n' + '\n'.join(f'{i}.{line}' for i, line in enumerate(lines, 1)))
    return '\n'.join(blocks)


def reappeared(records, later, before_year, triggers=None):
    """前期（before_year）以前的維修，原構件在之後的檢測又出現劣化。
    later：[(年度, Defect)] 之後各期的缺失；triggers：{檢測日期: [Defect]} 維修前那次檢測的缺失（找原本的位置）。
    回傳 (文字, 提示清單)。沒有維修紀錄回傳「無」，有維修但沒有再出現回傳「良好」。"""
    lines, notes = [], []
    old = [r for r in records if r.year is not None and before_year and r.year < before_year]
    for record in old:
        code = position_code(record)
        original = [d for d in (triggers or {}).get(record.inspect_date, []) if same_member(d.component, record)
                    and (not code or code in short_codes(d))]
        codes = {code} if code else {c for d in original for c in short_codes(d)}
        notes.append(f"{record.year}年維修 {record.component} {code or '（位置不明）'}：{record.method}"
                     + (f"（維修前 {record.inspect_date} 檢測：{'、'.join(d.fields.get('remark', '') for d in original)}）"
                        if original else ''))
        for year, defect in later:
            if same_member(defect.component, record) and (not codes or codes & set(short_codes(defect))) \
                    and same_kind(record.deterioration, defect.deterioration):
                lines.append(f"{record.year}年維修 {year}年原構件出現劣化（{record.component}"
                             f"{' ' + '、'.join(sorted(codes)) if codes else ''}：{pattern_of(entry_of(defect))}）")
    if lines:
        return numbered(lines), notes
    return ('良好' if old else '無'), notes


def same_kind(a, b):
    """維修的劣化類型與之後缺失的類型相同（一方包含另一方，例：混凝土剝落 ⊂ 混凝土剝落、破碎、鋼筋外露、銹蝕）；
    任一方沒有類型時視為相同。"""
    a, b = (re.sub(r'\s', '', str(x or '')) for x in (a, b))
    return not a or not b or a in b or b in a


def short_codes(defect):
    """缺失的主構件位置代碼（A1、P2、S1…）。"""
    return locations(entry_of(defect))


# ---------------------------------------------------------------- 其他欄位

PIER_HEIGHT = re.compile(r'(?i)\bP\s*0*(\d+)\s*[=:]\s*(\d+(?:\.\d+)?)')
SPAN_HEIGHT = re.compile(r'(?i)\bH\s*[=:]\s*(\d+(?:\.\d+)?)')
HEIGHT_FORMAT = '0.00'


def format_height(value):
    """淨高統一寫法（2026-11-24 使用者）：多跨「P01=3.50 換行 P02=3.40」（P 補兩位數、兩位小數，有 P 就不讀 H）；
    單跨數字取兩位小數（數字儲存格用 0.00 格式顯示）。讀不懂原樣回傳。"""
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    text = unicodedata.normalize('NFKC', str(value or '')).strip()
    piers = PIER_HEIGHT.findall(text)
    if piers:
        return '\n'.join(f'P{int(n):02d}={Decimal(v):.2f}' for n, v in piers)
    span = SPAN_HEIGHT.search(text)
    plain = re.fullmatch(r'(\d+(?:\.\d+)?)\s*m?', text, re.I)
    if span or plain:
        return f'{Decimal((span or plain)[1]):.2f}'
    return value


def heights_value(opinion):
    """意見中的淨高（只讀「淨高」之後的文字，避免讀到「基礎裸露(H=0.45m)」）：多跨只取 P、單跨取 H；
    寫法依 format_height。沒有回傳空白。"""
    text = unicodedata.normalize('NFKC', str(opinion or ''))
    at = text.find('淨高')
    if at < 0:
        return ''
    found = format_height(text[at:])
    return found if isinstance(found, str) and found != text[at:] else ''


def same_value(a, b):
    """清單值與系統值是否相同（數字比大小、文字去空白與全形）。"""
    a, b = (unicodedata.normalize('NFKC', str(x if x is not None else '')).strip() for x in (a, b))
    try:
        return Decimal(a) == Decimal(b)
    except InvalidOperation:
        return re.sub(r'\s', '', a) == re.sub(r'\s', '', b)


def heights_differ(listed, current):
    def parse(text):
        text = unicodedata.normalize('NFKC', str(text or ''))
        pairs = {f'P{int(n)}': v for n, v in PIER_HEIGHT.findall(text)}
        if not pairs:
            span = SPAN_HEIGHT.search(text) or re.fullmatch(r'\s*(\d+(?:\.\d+)?)\s*', text)
            pairs = {'H': span[1]} if span else {}
        return {k: Decimal(v) for k, v in pairs.items()}
    return bool(current) and parse(listed) != parse(current)


def railing_height(scan):
    """欄杆高度：掃描檔各構件列的「護欄高度」欄（cm 換成 m），多個取最小；沒有回傳空白。"""
    found = []
    for row in (scan or {}).get('rows', []):
        for label, value in (row.get('columns') or {}).items():
            if '護欄' not in label or '高度' not in label:
                continue
            try:
                height = Decimal(re.sub(r'[^\d.]', '', str(value)) or 'x')
            except InvalidOperation:
                continue
            if 'cm' in label.lower() or height > 10:
                height = height / 100
            if height > 0:
                found.append(height)
    return f'{min(found).normalize():f}' if found else ''


def repaired_mark(defect):
    return defect.repaired or '已維修' in str(defect.fields.get('remark', ''))


def corrected(defect, rules=()):
    """前期缺失套用 Excel 規則修正敘述（2026-11-24 使用者需求）：樣態同義詞（老化 → AC剝離）、
    依規則改成標準樣態（混凝土剝落，鋼筋外露 → 混凝土剝落鋼筋外露）、位置寫法（S2G5 → 2G5）。
    U、年份仍為前期；規則判斷不出來時只套同義詞。工法、單價不使用。"""
    entry = entry_of(defect)
    if not rules:
        return entry
    from .fill import TEXT_FIELDS, shape_remark, standardize, tidy_locations
    from .policy import automatic_plan
    component, context = defect.component, dict(defect.context or {})
    fields = {k: str(defect.fields.get(k) or '').strip() for k in TEXT_FIELDS}
    try:
        fields = standardize(fields, component, str(entry.u))
        fields = shape_remark(fields, component, context, rules)
        decision = automatic_plan(dict(fields), component, [], rules, context)
        if not decision.review:
            fields.update({k: str(v or '').strip() for k, v in decision.fields.items() if k in TEXT_FIELDS})
        else:
            fields['remark'] = joined_pattern(fields.get('remark', ''), component, rules)
        fields = tidy_locations(fields, component, context)
    except Exception:  # 規則判斷失敗：用原本的敘述
        return entry
    return Entry(entry.key, component, dict(entry.fields, **{k: fields[k] for k in ('remark', 'd_location')}),
                 entry.d, entry.e, entry.u, {}, entry.context)


def joined_pattern(remark, component, rules):
    """規則判斷不出唯一方案時：敘述去掉逗號後是這個構件某條規則樣態的一部分，就去掉逗號
    （河道「混凝土剝落，鋼筋外露」→「混凝土剝落鋼筋外露」，規則有護岸、固床工兩種時）。"""
    from .data import component_repairs
    from .editing import split_remark
    location, description, tail = split_remark(remark)
    joined = re.sub(r'[，,、]', '', description)
    if joined == description or not joined:
        return remark
    patterns = [re.sub(r'\s', '', r.get('pattern', '')) for r in component_repairs(rules, component)]
    if not any(joined.rstrip('。. ') in p for p in patterns):
        return remark
    return ' '.join(x for x in (location, joined + tail) if x)


def prior_items(defects, prior_year, rules=()):
    """前期觀測劣化：前期缺失（已維修除外），套 Excel 規則修正敘述，標前期年份。"""
    return [(corrected(d, rules), prior_year) for d in defects if not repaired_mark(d)]


def defect_id(defect):
    """同一筆缺失在不同次讀取（不同物件）時的識別：構件列＋網站 ID＋備註。"""
    return (tuple(defect.row), str(defect.der_id or ''), str(defect.method_id or ''), str(defect.fields.get('remark', '')))


def railing_standard(date):
    """護欄高度標準（2026-11-24 使用者）：2026 年（民國 115 年）起 110cm，之前 80cm。"""
    found = re.match(r'(\d{4})', str(date or ''))
    return Decimal('1.1') if not found or int(found[1]) >= 2026 else Decimal('0.8')


def railing_low(railing, date):
    try:
        return bool(railing) and Decimal(str(railing)) < railing_standard(date)
    except InvalidOperation:
        return False


@dataclass
class Sources:
    """一座橋同步需要的資料（由主程式收集）。"""
    inventory_id: str
    header: dict                         # 本期（沒有本期時為前期）檢測頁表頭
    basic: dict = field(default_factory=dict)      # 基本資料（route、road_lv、crossover、spans、total_length、width_max）
    prior_items: list = field(default_factory=list)    # 前期觀測劣化 [(Entry, 年度)]
    new_items: list | None = None        # 本期新增 [(Entry, 年度)]；None：本期尚未建立（不更新新增欄）
    current_entries: list = field(default_factory=list)  # 本期全部缺失 Entry（應維修、河工、護欄用）
    repairs: list = field(default_factory=list)      # pages.RepairRecord
    repaired_now: list = field(default_factory=list)  # [(年度, 文字)]
    later_defects: list = field(default_factory=list)  # [(年度, Defect)] 前期與本期的缺失
    before_year: int | None = None       # 前期年度（更早的維修要核對是否再出現）
    triggers: dict = field(default_factory=dict)      # {維修前檢測日期: [Defect]}
    prior_height: str = ''
    current_height: str = ''
    railing: str = ''
    photos_done: bool | None = None      # True：本年度六張都在；False：沒有；None：讀不到（不填）
    done: bool = False
    only_basic: bool = False             # 只核對系統資料（⑥「比對橋梁ID與基本資料」）：不寫劣化與其他欄位


def build(sources):
    """回傳 (系統值 {欄: 值}（核對用）, 程式整理值 {欄: 值}, 提示)。"""
    header, basic = sources.header, sources.basic
    system = {'id': sources.inventory_id, 'code': header.get('橋梁編號', ''), 'name': header.get('橋梁名稱', ''),
              'grade': basic.get('road_lv') or header.get('道路等級', ''), 'route': basic.get('route', ''),
              'spans': basic.get('spans') or header.get('總橋孔數', ''), 'length': basic.get('total_length', ''),
              'width': basic.get('width_max', ''), 'crossover': basic.get('crossover', ''), 'railing': sources.railing}
    if sources.only_basic:
        return system, {}, []
    values = {}
    prior = defect_lines(sources.prior_items)
    for bucket in (1, 2, 3):
        values[f'prior{bucket}'] = prior[bucket]
    if sources.new_items is not None:
        new = defect_lines(sources.new_items)
        for bucket in (1, 2, 3):
            values[f'new{bucket}'] = new[bucket]
    notes = []
    if sources.repairs is not None:                    # None：維修紀錄讀不到，這兩欄不更新
        values['repairs'] = repair_text(sources.repairs, sources.repaired_now)
        values['repaired'], notes = reappeared(sources.repairs, sources.later_defects, sources.before_year,
                                               sources.triggers)
    if sources.prior_height:
        values['height_prior'] = sources.prior_height
    if sources.current_height:
        values['height_now'] = sources.current_height
    if sources.new_items is not None:
        urgent = [e for e in sources.current_entries if e.u >= 3]
        names = [display_name(e.component, '', locations(e), e.context) or e.component for e in urgent]
        values['members'] = '、'.join(dict.fromkeys(n for n in names if n)) or ''
        values['advice'] = '、'.join(dict.fromkeys(e.fields.get('method', '') for e in urgent if e.fields.get('method')))
        river = [e for e in sources.current_entries if e.component in RIVER_PARTS
                 or any(w in e.remark for w in RIVER_TEXT)]
        values['river'] = numbered(member_lines(river), '')
        written = any('護欄' in e.component and '高度不足' in e.remark for e in sources.current_entries)
        values['guard'] = yes_no(railing_low(sources.railing, header.get('檢測日期')) or written,
                                 bool(sources.railing) or written)
    elif sources.railing:                              # 尚無本期：依前期欄杆高度與當年標準先判斷
        values['guard'] = yes_no(railing_low(sources.railing, header.get('檢測日期')), True)
    from .opinion import WASTE_WORDS
    said = str(header.get('橋下有無堆放雜物', '')).strip()
    waste = said in ('是', '有') or any(w in e.remark for e in sources.current_entries for w in WASTE_WORDS)
    values['waste'] = yes_no(waste, waste or said in ('否', '無'))       # 缺失寫到雜物、垃圾也算「是」
    salt = str(header.get('是否疑似鹽害橋梁', '')).strip()
    values['salt'] = yes_no(salt == '是', salt in ('是', '否'))
    if sources.photos_done is not None:                # 讀不到照片簿（None）時不填
        values['photos'] = yes_no(sources.photos_done, True)
    if sources.done:
        values['done'] = datetime.now().strftime('%Y-%m-%d')
    return system, values, notes


# ---------------------------------------------------------------- 欄寬與列高（文字完整顯示）

LINE_HEIGHT = 15.0       # 標楷體 10pt 一行（含行距）
DATA_FONT, DATA_SIZE = '標楷體', 10   # 欄位文字字型與大小（2026-11-24 使用者）
MAX_WIDTH = 42           # 劣化等長文字欄的最大欄寬（超過就換行）
MIN_WIDTH = 6


def units(text):
    """顯示寬度：中文（全形）2、英數 1（Excel 欄寬單位約等於 1 個英數字）。"""
    return sum(2 if unicodedata.east_asian_width(ch) in 'WFA' else 1 for ch in str(text))


def cell_lines(value, width):
    """儲存格在欄寬 width 下自動換行後的行數（公式算 1 行）。"""
    if value in (None, '') or str(value).startswith('='):
        return 1
    room = max(width - 1.5, 2)                       # 扣掉儲存格左右邊距
    return sum(max(1, math.ceil(units(line.strip()) / room)) for line in str(value).split('\n'))


def fit_columns(sheet, last_row=None):
    """欄寬依文字調整：每欄取最長一行的寬度（表頭依換行後各行），介於 MIN_WIDTH～MAX_WIDTH，超過的換行顯示。"""
    from openpyxl.utils import get_column_letter
    last_row = last_row or sheet.max_row
    for col in range(1, sheet.max_column + 1):
        longest = 0
        for row in range(1, last_row + 1):
            value = sheet.cell(row, col).value
            if value in (None, '') or str(value).startswith('='):
                continue
            longest = max([longest] + [units(line.strip()) for line in str(value).split('\n')])
        if longest:
            sheet.column_dimensions[get_column_letter(col)].width = min(max(longest + 2, MIN_WIDTH), MAX_WIDTH)


FREEZE = 'E2'          # 凍結：表頭＋序號～橋梁名稱
FREEZE_LIMIT = 80      # 凍結的欄寬合計超過這個寬度（約一個螢幕）就無法左右捲動


def fix_freeze(sheet):
    """凍結窗格的欄太多（寬度超過一個螢幕，左右捲動失效）時改成只凍結表頭與序號～橋梁名稱。回傳是否修改。"""
    from openpyxl.utils import get_column_letter
    if not sheet.freeze_panes:
        return False
    column, row = sheet[sheet.freeze_panes].column, sheet[sheet.freeze_panes].row
    width = sum((sheet.column_dimensions[get_column_letter(i)].width or 8.43) for i in range(1, column))
    if width <= FREEZE_LIMIT and row <= 2:             # 凍結太多列（例：凍結到第 44 列）也無法上下捲動
        return False
    sheet.freeze_panes = FREEZE
    return True


def widen(sheet, row):
    """這一列各儲存格的最長一行比欄寬長時，把欄寬放大到可完整顯示（最多 MAX_WIDTH，超過的換行）。"""
    from openpyxl.utils import get_column_letter
    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row, col).value
        if value in (None, '') or str(value).startswith('='):
            continue
        need = min(max(units(line.strip()) for line in str(value).split('\n')) + 2, MAX_WIDTH)
        dimension = sheet.column_dimensions[get_column_letter(col)]
        if (dimension.width or 8.43) < need:
            dimension.width = need


def centered_column(head):
    """淨高、欄杆至河床高度：多跨也維持置中（不改靠左）。"""
    return any(word in str(head or '') for word in ('淨高', '欄杆至河床'))


def fit_rows(sheet, rows=None):
    """列高依自動換行後的行數；同時把文字儲存格設為自動換行。rows 省略時整張表（含表頭）。"""
    from copy import copy

    from openpyxl.utils import get_column_letter
    rows = rows or range(1, sheet.max_row + 1)
    widths = {c: (sheet.column_dimensions[get_column_letter(c)].width or 8.43) for c in range(1, sheet.max_column + 1)}
    for row in rows:
        lines = 1
        for col, width in widths.items():
            cell = sheet.cell(row, col)
            if cell.value in (None, ''):
                continue
            lines = max(lines, cell_lines(cell.value, width))
            if row == 1:                               # 表頭維持原樣（只確保自動換行）
                if cell_lines(cell.value, width) > 1 and not cell.alignment.wrap_text:
                    alignment = copy(cell.alignment)
                    alignment.wrap_text = True
                    cell.alignment = alignment
                continue
            several = '\n' in str(cell.value) and not str(cell.value).startswith('=') \
                and not centered_column(sheet.cell(1, col).value)
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            # 多筆資料（換行列出）靠左；其他（是／否、數字、日期、單一文字）置中；垂直一律置中
            alignment.horizontal, alignment.vertical = ('left' if several else 'center'), 'center'
            cell.alignment = alignment
            if cell.font.name != DATA_FONT or cell.font.sz != DATA_SIZE:   # 欄位文字：標楷體 10
                font = copy(cell.font)
                font.name, font.sz = DATA_FONT, DATA_SIZE
                cell.font = font
        sheet.row_dimensions[row].height = lines * LINE_HEIGHT + 4


# ---------------------------------------------------------------- 寫入 Excel

class ListError(RuntimeError):
    pass


def find_row(workbook, inventory_id, code='', name='', only=None):
    """依二代系統ID（其次橋梁編號、橋名）找清單中的列：回傳 (工作表, 列號, 表頭 {key: 欄號})。
    only：只找這個工作表（批次整理時）。"""
    sheets = [workbook[only]] if only else workbook.worksheets
    for wanted, column_name in ((str(inventory_id), 'id'), (code, 'code'), (name, 'name')):
        if not wanted:
            continue
        for sheet in sheets:
            heads = {key(c.value): c.column for c in sheet[1] if c.value}
            col = heads.get(key(COLUMNS[column_name]))
            if not col:
                continue
            for row in range(2, sheet.max_row + 1):
                if same_value(sheet.cell(row, col).value, wanted) and str(sheet.cell(row, col).value or '').strip():
                    return sheet, row, heads
    return None, None, None


GENERATED = '填報助手產生'


def stamp(value):
    """寫入內容的識別碼（內容改過就不同）。"""
    import hashlib
    text = unicodedata.normalize('NFKC', str(value if value is not None else '')).strip()
    return hashlib.sha1(text.encode('utf-8')).hexdigest()[:8]


def generated_state(cell):
    """儲存格是不是程式產生的：'same'（有程式備註、內容沒改）、'edited'（有程式備註、已人工修正）、None（不是程式產生）。"""
    text = cell.comment.text if cell.comment else ''
    found = re.search(rf'{GENERATED}.*?識別碼 ([0-9a-f]{{8}})', text, re.S)
    if not found:
        return None
    return 'same' if stamp(cell.value) == found[1] else 'edited'


def mark_generated(cell, extra=''):
    """程式寫入的儲存格加專屬備註（內容識別碼）；之後人工修改過的，同步時不覆寫。"""
    note_cell(cell, (extra + '\n' if extra else '') +
              f'{GENERATED} {datetime.now():%Y-%m-%d}｜識別碼 {stamp(cell.value)}（修改這格後，程式同步時不會覆寫）')


def note_cell(cell, text):
    from openpyxl.comments import Comment
    comment = Comment(text, '填報助手')
    comment.width, comment.height = 300, 80
    cell.comment = comment


def paint(cell, color):
    from openpyxl.styles import PatternFill
    cell.fill = PatternFill('solid', fgColor=color, bgColor=color)


def ensure_done_column(sheet, heads):
    """「系統填寫完成」欄（日期）：沒有就加在最後一欄，並加條件式格式：有日期 → 橋名綠底（優先於下橋路徑的黃底）。"""
    from copy import copy

    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
    col = heads.get(key(DONE_HEAD))
    if col:
        return col
    col = sheet.max_column + 1
    head = sheet.cell(1, col, DONE_HEAD)
    head._style = copy(sheet.cell(1, col - 1)._style)
    sheet.column_dimensions[get_column_letter(col)].width = 12
    heads[key(DONE_HEAD)] = col
    name_col = heads.get(key('橋梁名稱'))
    if name_col:
        letter, done = get_column_letter(name_col), get_column_letter(col)
        rule = FormulaRule(formula=[f'${done}2<>""'], fill=PatternFill('solid', fgColor=DONE_FILL, bgColor=DONE_FILL),
                           stopIfTrue=True)
        rule.priority = 1
        rules = sheet.conditional_formatting
        existing = [(str(r.sqref), items) for r, items in rules._cf_rules.items()]
        from openpyxl.formatting.formatting import ConditionalFormattingList
        sheet.conditional_formatting = ConditionalFormattingList()
        sheet.conditional_formatting.add(f'{letter}2:{letter}1000', rule)        # 綠底優先
        for ranges, items in existing:
            for item in items:
                item.priority = (item.priority or 1) + 1
                sheet.conditional_formatting.add(ranges, item)
    return col


SERIAL_DONE, SERIAL_TODO = 'FF92D050', 'FFFFFF00'   # 序號：綠＝已整理、黃＝未完成（批次中斷續掃用）


def open_book(path):
    from openpyxl import load_workbook
    path = Path(path)
    if not path.is_file():
        raise ListError(f'找不到總表檔：{path}')
    try:
        return load_workbook(path)
    except Exception as error:     # 損毀或不是 xlsx
        raise ListError(f'無法開啟總表檔：{error}') from error


def save_book(workbook, path, backup_dir=None):
    from .list_template import GUIDE_SHEET, add_guide
    if GUIDE_SHEET not in workbook.sheetnames:         # 舊總表沒有說明頁：補在最前面
        add_guide(workbook)
    path = Path(path)
    if backup_dir:
        Path(backup_dir).mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, Path(backup_dir) / f'{path.stem}-{datetime.now():%Y%m%d-%H%M%S}{path.suffix}')
    try:
        workbook.save(path)
    except PermissionError as error:
        raise ListError('總表檔正在使用中（Excel 開著？），請關閉後再同步') from error


def mark_serial(sheet, row, heads, done):
    """序號欄標色：已整理綠色、未完成黃色（批次中斷後續掃跳過綠色）。"""
    col = heads.get(key('序號'))
    if col:
        paint(sheet.cell(row, col), SERIAL_DONE if done else SERIAL_TODO)


def serial_done(sheet, row, heads):
    col = heads.get(key('序號'))
    fill = sheet.cell(row, col).fill if col else None
    return bool(fill and fill.fill_type == 'solid' and str(fill.fgColor.rgb).upper().endswith(SERIAL_DONE[2:]))


def write_row(workbook, sources, only=None):
    """一座橋寫入已開啟的清單（不存檔）。回傳說明；找不到這座橋丟出 ListError。"""
    from copy import copy
    header = sources.header
    sheet, row, heads = find_row(workbook, sources.inventory_id, header.get('橋梁編號', ''), header.get('橋梁名稱', ''),
                                 only)
    if sheet is None:
        raise ListError(f"總表中找不到這座橋（二代系統ID {sources.inventory_id}、{header.get('橋梁名稱', '')}），"
                        '請先在總表加一列（至少填二代系統ID 或橋名）')
    system, values, notes = build(sources)
    report = [f"{sheet.title} 第 {row} 列：{header.get('橋梁名稱', '')}"]
    if 'done' in values:
        ensure_done_column(sheet, heads)
    for name, value in system.items():                 # 核對：空白填入；不同標橘底加備註，不覆寫
        col = heads.get(key(COLUMNS[name]))
        if not col or value in (None, ''):
            continue
        cell = sheet.cell(row, col)
        if cell.value in (None, ''):
            cell.value = number_or_text(value)
            report.append(f'{COLUMNS[name]}：填入系統值 {value}')
        elif not same_value(cell.value, value):
            paint(cell, CHECK_FILL)
            note_cell(cell, f'系統值：{value}（{datetime.now():%Y-%m-%d} 填報助手核對）')
            report.append(f'{COLUMNS[name]}：清單 {cell.value} 與系統 {value} 不同（已標橘底，未修改）')
    for name, value in values.items():
        col = heads.get(key(COLUMNS[name]))
        if not col:
            continue
        cell = sheet.cell(row, col)
        state = generated_state(cell)
        if name in TICKS and not value and state is None and cell.value not in (None, ''):
            continue                                       # 判斷不出來時不清掉人工打的 V
        if state == 'edited':                               # 程式產生後人工修正過：保留修正，不覆寫
            report.append(f'{COLUMNS[name]}：已人工修正（與程式產生的內容不同），保留不覆寫')
            continue
        if name == 'height_prior' and cell.value not in (None, '') and state is None:
            continue                                       # 前期淨高：清單已有（非程式產生）就不改
        extra = ''
        if name == 'height_now':
            before = sheet.cell(row, heads[key(COLUMNS['height_prior'])]).value if key(COLUMNS['height_prior']) in heads else ''
            if not heights_differ(before, value):
                continue
            paint(cell, CHECK_FILL)
            extra = f'本期意見淨高與前期不同（前期：{before}）'
        if str(cell.value or '') != str(value):
            cell.value = number_or_text(value)
            if cell.has_style is False and col > 1:
                cell._style = copy(sheet.cell(row, col - 1)._style)
        if name in ('height_prior', 'height_now'):
            cell.number_format = HEIGHT_FORMAT
        mark_generated(cell, extra)
    widen(sheet, row)                                  # 寫入的文字較長：欄寬放大（最多 MAX_WIDTH）、列高跟著調整
    fit_rows(sheet, [row])
    if fix_freeze(sheet):
        report.append(f'{sheet.title}：凍結窗格太大（無法捲動），已改為只凍結表頭與序號～橋梁名稱')
    if not sources.only_basic:                         # 序號綠色＝前期劣化已建檔
        mark_serial(sheet, row, heads, True)
    return report + notes


def set_inventory_id(workbook, name, inventory_id, only=None):
    """手動配對：總表中這座橋（依橋名找）的二代系統ID 欄空白時填入。回傳是否找到。"""
    sheet, row, heads = find_row(workbook, '', '', name, only)
    if sheet is None:
        return False
    col = heads.get(key(COLUMNS['id']))
    if col and sheet.cell(row, col).value in (None, ''):
        sheet.cell(row, col).value = number_or_text(inventory_id)
    return True


def mark_failed(workbook, inventory_id, name, only=None):
    """批次中這座橋沒有完成：序號標黃色（找不到列就略過）。"""
    sheet, row, heads = find_row(workbook, inventory_id, '', name, only)
    if sheet is not None:
        mark_serial(sheet, row, heads, False)


def write(path, sources, backup_dir=None):
    """同步一座橋到清單檔（開檔、寫入、備份、存檔）。回傳說明；找不到這座橋或檔案無法寫入時丟出 ListError。"""
    workbook = open_book(path)
    report = write_row(workbook, sources)
    save_book(workbook, path, backup_dir)
    return report


def number_or_text(value):
    text = str(value)
    if re.fullmatch(r'-?\d+(?:\.\d+)?', text):
        return int(text) if re.fullmatch(r'-?\d+', text) else float(text)
    return text


# ---------------------------------------------------------------- 判讀結果的比對資料（本期，也可以是其他人輸出的）

@dataclass
class ReportData:
    path: Path
    bridge: str
    prior_date: str
    current_date: str
    prior_still: list            # 前期有、本期仍存在的前期缺失（Defect）
    new: list                    # 本期新增（Defect，已套用修改內容）
    current: list                # 本期全部缺失（Defect）
    repaired: list               # 本期標已維修的前期缺失（Defect）
    height: str = ''             # 本期意見淨高


def report_files(folder, bridge):
    """判讀結果資料夾中這座橋的比對 Excel（檔名含「-橋名-前後期比對」，可有前綴），本期日期新到舊。"""
    found = []
    for path in Path(folder).glob('*前後期比對.xlsx') if Path(folder).is_dir() else []:
        hit = re.search(r'(\d{4}-\d{2}-\d{2})-(.+)-前後期比對\.xlsx$', path.name)
        if hit and hit[2] == bridge:
            found.append((hit[1], path))
    return [path for _, path in sorted(found, reverse=True)]


def read_report(path):
    """⑤ 輸出的前後期比對 Excel → ReportData。本期值＝本期欄＋「修改內容」（人工修改）。讀不懂回傳 None。"""
    from openpyxl import load_workbook

    from .editing import FIELD_NAMES
    from .matching import Defect
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        sheet = book['前後期比對'] if '前後期比對' in book.sheetnames else book.worksheets[0]
        rows = list(sheet.iter_rows(values_only=True))
        title = str(rows[0][0] or '') if rows else ''
        heads = [str(h or '').strip() for h in (rows[2] if len(rows) > 2 else ())]
        height = ''
        if '淨高' in book.sheetnames:
            for line in book['淨高'].iter_rows(min_row=2, values_only=True):
                if line and line[1] not in (None, '', '—') and '意見' in str(line[0] or ''):
                    height = heights_value(line[1])
                    break
    finally:
        book.close()
    if '前期備註' not in heads or '本期備註' not in heads:
        return None
    col = {h: i for i, h in enumerate(heads)}
    names = {v: k for k, v in FIELD_NAMES.items()}
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', title)
    bridge = title.split('　')[0].split(' ')[0].strip()
    result = ReportData(Path(path), bridge, dates[0] if dates else '', dates[1] if len(dates) > 1 else '', [], [], [], [],
                        height)

    def value(row, name):
        return str(row[col[name]] if name in col and col[name] < len(row) and row[col[name]] is not None else '').strip()

    def defect(row, period, label, component, context):
        deru = value(row, f'{label} DERU').replace(' ', '')
        remark = value(row, f'{label}備註')
        kind = value(row, f'{label}劣化類型')
        if not remark and not kind:
            return None
        fields = {'remark': remark, 'deterioration': kind, 'amount': value(row, f'{label}數量'),
                  'unit': value(row, f'{label}單位'), 'd': deru[:1], 'e': deru[1:2], 'r': deru[2:3], 'u': deru[3:4]}
        if period == '本期':                                   # 人工修改：「備註：xxx；U：3」
            for part in value(row, '修改內容').split('；'):
                name, _, text = part.partition('：')
                if names.get(name.strip()):
                    fields[names[name.strip()]] = text.strip()
            deru = ''.join(fields.get(k, '') or '_' for k in 'deru')
        return Defect(period, (component, value(row, '位置／編號')), component, value(row, '位置／編號'), deru, kind,
                      fields, context=dict(context))

    for row in rows[3:]:
        if not row or not value(row, '構件'):
            continue
        shown = value(row, '構件')
        facility = re.search(r'（(.+)）$', shown)
        component = re.sub(r'（.+）$', '', shown)
        context = {'設施類別': facility[1]} if facility else {}
        verdict = value(row, '最終判定') or value(row, '程式判定')
        prior = defect(row, '前期', '前期', component, context)
        current = defect(row, '本期', '本期', component, context)
        if current is not None:
            result.current.append(current)
        if prior is not None and current is not None:
            result.prior_still.append(prior)
        elif current is not None:
            result.new.append(current)
        elif prior is not None and '已維修' in verdict:
            result.repaired.append(prior)
    return result
