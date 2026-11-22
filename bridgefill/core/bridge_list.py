"""縣市橋梁清單同步（2026-11-23 使用者需求）：把一座橋的檢測整理寫進使用者指定路徑的清單 Excel。

欄位（依表頭名稱找欄，不依位置；年度與換行不影響）：
- 核對（系統資料）：二代系統ID、橋梁編號、橋梁名稱、道路等級、路線、跨數、橋長、最大淨寬、跨越物體、欄杆高度——
  清單空白直接填入；已有值但與系統不同：標橘底並加備註寫系統值，不覆寫。
- 程式整理（每次同步更新）：
  前期觀測劣化 U=1／U=2／U≧3：前期有、本期仍存在的缺失（本期尚未建立時：前期全部缺失），以檢測員意見的寫法
  （同樣態合併、位置縮寫）去掉「建議…」，每條標年份（前前期也有的標前前期年份）；
  本期新增劣化 U=1／U=2／U≧3：本期才有的缺失；
  過去維修紀錄：網站歷年維修紀錄（依年度）＋本期標「已維修」的前期缺失；
  已維修構件狀況：前期以前的維修，原構件在之後的檢測又出現劣化（例：110年維修 113年原構件出現劣化）；
  前期淨高（清單空白才填）、本期淨高（本期意見有變化才填，橘底提醒）；
  應維修構件、應維修處理建議（U≧3）、河工設施損壞、護欄高度不足、橋下堆積物、疑似鹽害橋梁、六張基本照片拍攝；
  系統填寫完成（日期；橋名綠底）。
- 不動：序號、下橋路徑、輔助設備、基本資料變動、橋面板厚、大梁高度、橋名牌、限重牌、其他檢測特殊情形與公式欄。
寫入前備份；檔案開著（被 Excel 鎖住）時回報錯誤，不寫入。本模組不依賴 Qt。
"""
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
    return re.sub(r'\(列出過去維修紀錄\)|\(橋長\*淨寬\)', '', text)


COLUMNS = {   # 程式使用的名稱 → 清單表頭（key 比對）
    'id': '二代系統ID', 'code': '橋梁編號', 'name': '橋梁名稱', 'grade': '道路等級', 'route': '路線', 'spans': '跨數',
    'length': '橋長', 'width': '最大淨寬', 'crossover': '跨越物體', 'railing': '欄杆高度(m)',
    'prior1': '前期觀測劣化(U=1)', 'prior2': '前期觀測劣化(U=2)', 'prior3': '前期觀測劣化(U≧3)',
    'new1': '新增劣化(U=1)', 'new2': '新增劣化(U=2)', 'new3': '新增劣化(U≧3)',
    'repairs': '過去維修紀錄', 'repaired': '已維修構件狀況',
    'height_prior': '前期梁下/橋下淨高(m)', 'height_now': '本期梁下/橋下淨高(m)',
    'members': '應維修構件(ch5.2)', 'advice': '應維修處理建議', 'river': '河工設施損壞(ch5.4)',
    'guard': '護欄高度不足(V)(ch5.5)', 'waste': '橋下堆積物(V)(ch5.6)', 'salt': '疑似鹽害橋梁(V)',
    'photos': '六張基本照片拍攝', 'done': DONE_HEAD,
}
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

def heights_value(opinion):
    """意見中的淨高：單跨 H=2.2 → 2.2；多跨 → 「P1=3.5 換行 P2=3.4」。沒有回傳空白。"""
    found = re.findall(r'(?i)\b(P\d+|H)\s*[=＝:：]\s*(\d+(?:\.\d+)?)', str(opinion or ''))
    if not found:
        return ''
    if len(found) == 1 and found[0][0].upper() == 'H':
        return found[0][1]
    return '\n'.join(f'{name.upper()}={value}' for name, value in found)


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
        pairs = dict(re.findall(r'(?i)\b(P\d+|H)\s*[=:]\s*(\d+(?:\.\d+)?)', text))
        if not pairs and re.fullmatch(r'\s*\d+(?:\.\d+)?\s*', text):
            pairs = {'H': text.strip()}
        return {k.upper(): Decimal(v) for k, v in pairs.items()}
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


def prior_base(prior_defects, prior_year, older_pairs=(), older_year=None):
    """前期觀測劣化的基礎：前期全部缺失（已維修除外），前前期也有的標前前期年份。
    older_pairs：前前期（prior）與前期（current）配對結果。回傳 {缺失識別: (Defect, 年度)}。"""
    seen = {defect_id(p.current) for p in older_pairs if p.prior is not None and p.current is not None}
    base = {}
    for defect in prior_defects:
        if defect.repaired or '已維修' in str(defect.fields.get('remark', '')):
            continue
        base[defect_id(defect)] = (defect, older_year if older_year and defect_id(defect) in seen else prior_year)
    return base


def defect_id(defect):
    """同一筆缺失在不同次讀取（不同物件）時的識別：構件列＋網站 ID＋備註。"""
    return (tuple(defect.row), str(defect.der_id or ''), str(defect.method_id or ''), str(defect.fields.get('remark', '')))


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
    photos_done: bool = False
    done: bool = False


def build(sources):
    """回傳 (系統值 {欄: 值}（核對用）, 程式整理值 {欄: 值}, 提示)。"""
    header, basic = sources.header, sources.basic
    system = {'id': sources.inventory_id, 'code': header.get('橋梁編號', ''), 'name': header.get('橋梁名稱', ''),
              'grade': basic.get('road_lv') or header.get('道路等級', ''), 'route': basic.get('route', ''),
              'spans': basic.get('spans') or header.get('總橋孔數', ''), 'length': basic.get('total_length', ''),
              'width': basic.get('width_max', ''), 'crossover': basic.get('crossover', ''), 'railing': sources.railing}
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
        low = sources.railing and Decimal(sources.railing) < Decimal('1.1')
        guard = low or any('護欄' in e.component and '高度不足' in e.remark for e in sources.current_entries)
        values['guard'] = 'V' if guard else ''
    values['waste'] = 'V' if str(header.get('橋下有無堆放雜物', '')).strip() in ('是', '有') else ''
    values['salt'] = 'V' if str(header.get('是否疑似鹽害橋梁', '')).strip() == '是' else ''
    if sources.photos_done:
        values['photos'] = 'V'
    if sources.done:
        values['done'] = datetime.now().strftime('%Y-%m-%d')
    return system, values, notes


# ---------------------------------------------------------------- 寫入 Excel

class ListError(RuntimeError):
    pass


def find_row(workbook, inventory_id, code='', name=''):
    """依二代系統ID（其次橋梁編號、橋名）找清單中的列：回傳 (工作表, 列號, 表頭 {key: 欄號})。"""
    for wanted, column_name in ((str(inventory_id), 'id'), (code, 'code'), (name, 'name')):
        if not wanted:
            continue
        for sheet in workbook.worksheets:
            heads = {key(c.value): c.column for c in sheet[1] if c.value}
            col = heads.get(key(COLUMNS[column_name]))
            if not col:
                continue
            for row in range(2, sheet.max_row + 1):
                if same_value(sheet.cell(row, col).value, wanted) and str(sheet.cell(row, col).value or '').strip():
                    return sheet, row, heads
    return None, None, None


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


def write(path, sources, backup_dir=None):
    """同步一座橋到清單檔。回傳說明文字清單；找不到這座橋或檔案無法寫入時丟出 ListError。"""
    from copy import copy

    from openpyxl import load_workbook
    path = Path(path)
    if not path.is_file():
        raise ListError(f'找不到清單檔：{path}')
    try:
        workbook = load_workbook(path)
    except Exception as error:     # 損毀或不是 xlsx
        raise ListError(f'無法開啟清單檔：{error}') from error
    header = sources.header
    sheet, row, heads = find_row(workbook, sources.inventory_id, header.get('橋梁編號', ''), header.get('橋梁名稱', ''))
    if sheet is None:
        raise ListError(f"清單中找不到這座橋（二代系統ID {sources.inventory_id}、{header.get('橋梁名稱', '')}），"
                        '請先在清單加一列（至少填二代系統ID 或橋名）')
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
        if name == 'height_prior' and cell.value not in (None, ''):
            continue                                       # 前期淨高：清單已有就不改
        if name == 'height_now':
            before = sheet.cell(row, heads[key(COLUMNS['height_prior'])]).value if key(COLUMNS['height_prior']) in heads else ''
            if not heights_differ(before, value):
                continue
            paint(cell, CHECK_FILL)
            note_cell(cell, f'本期意見淨高與前期不同（前期：{before}）')
        if str(cell.value or '') != str(value):
            cell.value = number_or_text(value)
            if cell.has_style is False and col > 1:
                cell._style = copy(sheet.cell(row, col - 1)._style)
    if backup_dir:
        Path(backup_dir).mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, Path(backup_dir) / f'{path.stem}-{datetime.now():%Y%m%d-%H%M%S}{path.suffix}')
    try:
        workbook.save(path)
    except PermissionError as error:
        raise ListError('清單檔正在使用中（Excel 開著？），請關閉後再同步') from error
    return report + notes


def number_or_text(value):
    text = str(value)
    if re.fullmatch(r'-?\d+(?:\.\d+)?', text):
        return int(text) if re.fullmatch(r'-?\d+', text) else float(text)
    return text
