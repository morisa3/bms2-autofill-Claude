"""縣市橋梁清單（Google 試算表下載的 xlsx）：各工作表依範本工作表（預設「西螺」）改成相同欄位與格式，並加入公式。

使用者需求（2026-11-22、11-23）：以「西螺」工作表為範本——
- 欄位依範本表頭；其他工作表的資料依「同義表頭」搬到對應欄位（例：歷年持續觀測劣化 → 前期觀測劣化、
  歷年維修 → 過去維修紀錄、梁下淨高 → 前期梁下／橋下淨高；最早的格式 U≦2 → U=2），範本新增的欄位留白；
- 格式依範本：欄寬、表頭樣式、各欄對齊、自動換行、框線；列高依文字行數調整；
- 公式：橋孔面積 ＝ 橋長 × 最大淨寬；欄杆至河床高度 ＝ 梁下淨高（本期有填用本期，否則用前期）＋ 欄杆高度 ＋ 橋面板厚 ＋
  大梁高度（多跨「P1=3.5 換行 P2=3.4」逐墩計算）；
- 表頭備註：下橋路徑（A～D 級判定標準）、輔助設備（依下橋路徑 A／B 級）。

命令列：python tools/bridge_list.py 來源.xlsx 輸出.xlsx [範本工作表名稱，預設 西螺]；
程式中「⑥ 同步到橋梁總表」的建立空白範本、導入前期總表也使用本模組。
"""
import math
import re
import sys
import unicodedata
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter

from .bridge_list import (FREEZE, GENERATED, HEIGHT_FORMAT, MIN_WIDTH, STANDARD_LAYOUT, fix_freeze, fit_columns, fit_rows,
                          format_height, standard_width)

MAX_SPANS = 8              # 公式最多處理幾跨（P1～P8）；實際資料最多 7 跨

PATH_NOTE = ('下橋路徑判定標準：\n'
             '(1) A級（間接檢測）：須配合無人飛行器（UAV）、水下無人機（ROV）等設備進行檢測。\n'
             '(2) B級（直接檢測）：須搭載輔助載具始能辦理檢測，如橋梁檢測車、高空工作車或工作船舶等。\n'
             '(3) C級（直接檢測）：須搭設移動梯（O截梯）等設備方可到達橋下進行檢測。\n'
             '(4) D級（直接檢測）：檢測人員可徒步到達橋下進行檢測。')
TOOL_NOTE = ('輔助設備依下橋路徑填寫：\n'
             'A：下橋路徑 A級，使用無人飛行器（UAV）、水下無人機（ROV）等設備。\n'
             'B：下橋路徑 B級，使用橋梁檢測車、高空工作車或工作船舶等輔助載具。\n'
             '下橋路徑 C、D 級免填。')


def key(text):
    """表頭比對用：去掉空白、換行、全形、年度（115年／(113年)）、「須修正」與括號內說明。"""
    text = unicodedata.normalize('NFKC', str(text or ''))
    text = re.sub(r'\s+', '', text)
    text = re.sub(r'\(?\d{2,3}年\)?', '', text)
    text = text.replace('須修正', '').replace('檢測新增', '新增').replace('≦', '<=').replace('≧', '>=')
    text = re.sub(r'\(V\)|\(是[/／、]?否\)|\(ch\d+(?:\.\d+)?\)', '', text, flags=re.I)
    text = re.sub(r'[/、]', '', text)        # 西螺／北雲林版型表頭
    return text.replace('梁下橋下淨高', '梁下淨高')


# 範本表頭 → 舊表頭（同義）
ALIASES = {
    '前期觀測劣化(U=1)': ('歷年持續觀測劣化(U=1)',),
    '前期觀測劣化(U=2)': ('歷年持續觀測劣化(U=2)', '歷年持續觀測劣化(U≦2)'),
    '前期觀測劣化(U≧3)': ('歷年持續觀測劣化(U≧3)',),
    '新增劣化(U=2)': ('新增劣化(U≦2)',),
    '過去維修紀錄(列出過去維修紀錄)': ('歷年維修',),
    '前期梁下/橋下淨高(m)': ('梁下淨高(m)',),
}
FORMULA_HEADS = ('橋孔面積(橋長*淨寬)', '欄杆至河床高度(m)')


def column(heads, name):
    """表頭中名稱（比對 key）的欄號（1 起算）；找不到回傳 None。"""
    wanted = key(name)
    return next((i for i, h in enumerate(heads, 1) if key(h) == wanted), None)


def source_index(old_heads, template_head):
    """範本欄位在舊工作表的欄索引（0 起算）；沒有回傳 None。"""
    names = (template_head,) + next((v for k, v in ALIASES.items() if key(k) == key(template_head)), ())
    for name in names:
        found = column(old_heads, name)
        if found:
            return found - 1
    return None


def railing_formula(row, columns):
    """欄杆至河床高度公式（Excel 與 Google 試算表都能用的一般函數，不用 LET／REGEX）。
    梁下淨高：本期欄有填用本期，否則用前期。數字直接相加；「P1=3.5 換行 P2=3.4」（或 P1:3.5m）逐行加總，
    結果寫成「P1=4.8 換行 P2=4.7」。欄杆高度、橋面板厚、大梁高度任一未填 → 空白。
    columns：{'淨高': 本期欄（或唯一的淨高欄）, '前期淨高': 前期欄（可省略）, '欄杆高度'、'橋面板厚'、'大梁高度': 欄名}"""
    now = f"{columns['淨高']}{row}"
    before = f"{columns['前期淨高']}{row}" if columns.get('前期淨高') else ''
    s = f'IF({now}="",{before},{now})' if before else now
    u, v, w = (f'{columns[name]}{row}' for name in ('欄杆高度', '橋面板厚', '大梁高度'))
    extra = f'({u}+{v}+{w})'
    text = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({s},"：","="),":","="),"m","")'
    pieces = []
    for k in range(1, MAX_SPANS + 1):
        line = f'TRIM(MID(SUBSTITUTE({text},CHAR(10),REPT(" ",99)),{(k - 1) * 99 + 1},99))'
        lead = '' if k == 1 else 'CHAR(10)&'
        value = f'TEXT(ROUND(VALUE(MID({line},FIND("=",{line})+1,20))+{extra},2),"0.00")'
        pieces.append(f'IF({line}="","",{lead}IFERROR(LEFT({line},FIND("=",{line}))&{value},{line}))')
    spans = '&'.join(pieces)
    return f'=IF(OR({s}="",COUNT({u},{v},{w})<3),"",IFERROR(ROUND({s}+{extra},2),{spans}))'


def area_formula(row, length, width):
    return f'=IF(OR({length}{row}="",{width}{row}=""),"",ROUND({length}{row}*{width}{row},2))'


def tidy_text(sheet):
    """換行後開頭的空白去掉（「1.xxx 換行 空白2.xxx」→ 對齊），公式不動。"""
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and '\n' in cell.value and not cell.value.startswith('='):
                cell.value = '\n'.join(line.strip() for line in cell.value.split('\n')).strip()


def tidy_heights(sheet):
    """淨高欄：P 補兩位數、兩位小數（P01=3.50），多跨只留 P；單跨數字與欄杆至河床高度用 0.00 格式。"""
    for col, cell in enumerate(sheet[1], 1):
        head = str(cell.value or '')
        if '淨高' not in head and '欄杆至河床' not in head:
            continue
        for row in range(2, last_data_row(sheet) + 1):
            target = sheet.cell(row, col)
            if '淨高' in head and target.value not in (None, ''):
                target.value = format_height(target.value)
                if isinstance(target.value, str) and re.fullmatch(r'\d+\.\d{2}', target.value):
                    target.value = float(target.value)
            target.number_format = HEIGHT_FORMAT


TICK_HEADS = {key(k): v for k, v in (('護欄高度不足(ch5.5)', '護欄高度不足\n(ch5.5)'), ('橋下堆積物(ch5.6)', '橋下堆積物\n(ch5.6)'),
                                      ('疑似鹽害橋梁', '疑似鹽害橋梁'), ('六張基本照片拍攝', '六張基本照片拍攝'))}


def tidy_ticks(sheet):
    """是／否欄位：表頭去掉「(V)」，舊的 V／v 改成「是」、X／x 改成「否」（2026-11-24 使用者）。"""
    for col, cell in enumerate(sheet[1], 1):
        standard = TICK_HEADS.get(key(cell.value))
        if not standard:
            continue
        if '(V)' in unicodedata.normalize('NFKC', str(cell.value)).upper():
            cell.value = standard                          # 只改有 (V) 的舊表頭；使用者的寫法（例：護欄\n高度不足）保留
        for row in range(2, last_data_row(sheet) + 1):
            target = sheet.cell(row, col)
            text = str(target.value or '').strip().upper()
            if text in ('V', '✓'):
                target.value = '是'
            elif text == 'X':
                target.value = '否'


def fit(sheet):
    """欄寬依文字、列高依換行後行數（與程式同步時用同一套計算，文字完整顯示）；凍結表頭＋序號～橋梁名稱。"""
    sheet.freeze_panes = FREEZE
    tidy_ticks(sheet)
    tidy_text(sheet)
    tidy_heights(sheet)
    fit_columns(sheet, last_data_row(sheet))
    standard_columns(sheet)                            # 標準欄位用西螺版型欄寬（不隱藏、不縮小）
    fit_rows(sheet, range(1, last_data_row(sheet) + 1))


def standard_columns(sheet):
    """標準欄位設為西螺版型欄寬並取消隱藏；其他欄至少 MIN_WIDTH。"""
    for col in range(1, sheet.max_column + 1):
        head = sheet.cell(1, col).value
        if head in (None, ''):
            continue
        dimension = sheet.column_dimensions[get_column_letter(col)]
        dimension.hidden = False
        dimension.width = standard_width(head) or max(dimension.width or 0, MIN_WIDTH)


def last_data_row(sheet, name_col=4):
    rows = [r for r in range(2, sheet.max_row + 1) if sheet.cell(r, name_col).value not in (None, '')]
    return max(rows) if rows else 1


def same_layout(sheet, template):
    return [key(c.value) for c in sheet[1]] == [key(c.value) for c in template[1]]


def convert_sheet(sheet, template):
    """工作表改成範本的欄位與格式（就地改寫）。回傳 (搬移列數, 沒有對應欄位而捨棄的舊表頭)。"""
    heads = [c.value for c in sheet[1]]
    template_heads = [c.value for c in template[1]]
    last = last_data_row(sheet)
    rows = [[sheet.cell(r, c).value for c in range(1, len(heads) + 1)] for r in range(2, last + 1)]
    sources = [source_index(heads, h) for h in template_heads]
    dropped = [h for i, h in enumerate(heads) if i not in sources and h and key(h) not in map(key, FORMULA_HEADS)]
    name_rules = [(str(r.sqref), items) for r, items in sheet.conditional_formatting._cf_rules.items()
                  if str(r.sqref).startswith('D')]
    path_col = column(template_heads, '下橋路徑(ABCD)')
    sheet.conditional_formatting = ConditionalFormattingList()
    sheet.delete_rows(1, sheet.max_row)
    sheet.delete_cols(1, sheet.max_column + 5)
    for col, head in enumerate(template_heads, 1):
        sheet.cell(1, col, head)._style = copy(template.cell(1, col)._style)
    for r, values in enumerate(rows, 2):
        for col, source in enumerate(sources, 1):
            value = values[source] if source is not None and source < len(values) else None
            sheet.cell(r, col, value)._style = copy(template.cell(2, col)._style)
    for col in range(1, len(template_heads) + 1):
        letter = get_column_letter(col)
        sheet.column_dimensions[letter].width = template.column_dimensions[letter].width or 8.43
    sheet.freeze_panes = FREEZE                        # 一律凍結表頭＋序號～橋梁名稱（原本凍結太多欄會無法左右捲動）
    for ranges, items in name_rules:                   # 「下橋路徑有填 → 橋名變黃」：改指新的下橋路徑欄
        for rule in items:
            rule.formula = [f'${get_column_letter(path_col)}2<>""']
            sheet.conditional_formatting.add(ranges, rule)
    return len(rows), dropped


def add_formulas(sheet):
    """橋孔面積（有此欄時）、欄杆至河床高度公式（超過 8 墩如 25 墩大跨橋改直接計算寫入，避免公式截斷）；回傳處理列數。"""
    from .bridge_list import PIER_HEIGHT, compute_railing_to_bed
    heads = [c.value for c in sheet[1]]
    letter = lambda name: get_column_letter(column(heads, name)) if column(heads, name) else None   # noqa: E731
    area, length, width = letter('橋孔面積(橋長*淨寬)'), letter('橋長'), letter('最大淨寬')
    railing = letter('欄杆至河床高度(m)')
    now = letter('本期梁下淨高(m)') or letter('梁下淨高(m)')
    prior = letter('前期梁下淨高(m)')
    columns = {'淨高': now, '前期淨高': prior,
               '欄杆高度': letter('欄杆高度(m)'), '橋面板厚': letter('橋面板厚(m)'), '大梁高度': letter('大梁高度(m)')}
    count = 0
    for row in range(2, last_data_row(sheet) + 1):
        if area and length and width:
            sheet[f'{area}{row}'] = area_formula(row, length, width)
        if railing and now and all(columns[k] for k in ('欄杆高度', '橋面板厚', '大梁高度')):
            h_now = sheet[f'{now}{row}'].value
            h_prior = sheet[f'{prior}{row}'].value if prior else None
            active_h = h_now if h_now not in (None, '') else h_prior
            piers = PIER_HEIGHT.findall(unicodedata.normalize('NFKC', str(active_h or '')))
            if len(piers) > MAX_SPANS:
                computed = compute_railing_to_bed(active_h, sheet[f"{columns['欄杆高度']}{row}"].value,
                                                  sheet[f"{columns['橋面板厚']}{row}"].value,
                                                  sheet[f"{columns['大梁高度']}{row}"].value)
                sheet[f'{railing}{row}'] = computed or railing_formula(row, columns)
            else:
                sheet[f'{railing}{row}'] = railing_formula(row, columns)
        count += 1
    return count


def heights_to_prior(sheet):
    """前期淨高欄全空、本期淨高欄有值：插入「前期」欄時舊值留在本期欄（西螺），搬到前期欄。回傳搬移列數。"""
    heads = [c.value for c in sheet[1]]
    before, now = column(heads, '前期梁下/橋下淨高(m)'), column(heads, '本期梁下/橋下淨高(m)')
    last = last_data_row(sheet)
    if not before or not now or any(sheet.cell(r, before).value not in (None, '') for r in range(2, last + 1)):
        return 0
    moved = 0
    for row in range(2, last + 1):
        value = sheet.cell(row, now).value
        if value not in (None, ''):
            sheet.cell(row, before).value, sheet.cell(row, now).value = value, None
            moved += 1
    return moved


def add_notes(sheet):
    heads = [c.value for c in sheet[1]]
    for name, note in (('下橋路徑(ABCD)', PATH_NOTE), ('輔助設備(A/B)', TOOL_NOTE)):
        col = column(heads, name)
        if col:
            comment = Comment(note, '填報助手')
            comment.width, comment.height = 420, 160
            sheet.cell(1, col).comment = comment


def convert(source, output, template_name='西螺'):
    workbook = load_workbook(source)
    template = workbook[template_name]
    report = []
    for sheet in [s for s in workbook.worksheets if not is_guide(s)]:
        if sheet is not template and not same_layout(sheet, template):
            count, dropped = convert_sheet(sheet, template)
            report.append(f'{sheet.title}：轉換 {count} 列' + (f"（沒有對應欄位：{'、'.join(map(str, dropped))}）"
                                                            if dropped else ''))
        moved = heights_to_prior(sheet)
        if moved:
            report.append(f'{sheet.title}：淨高 {moved} 列從本期欄移到前期欄（舊值為前期量測值）')
        formulas = add_formulas(sheet)
        add_notes(sheet)
        fit(sheet)                                     # 全部工作表（含範本）：欄寬與列高讓文字完整顯示
        report.append(f'{sheet.title}：公式 {formulas} 列')
    add_guide(workbook)
    workbook.save(output)
    return report



# ---------------------------------------------------------------- 建立空白總表、導入前期總表（⑥ 步驟 1）

def template_heads(year):
    """標準總表表頭（西螺版型 38 欄＋系統填寫完成）；year：本期民國年（例 115）。"""
    return [head.replace('{year}', str(year)) for head, _ in STANDARD_LAYOUT]


def style_template(sheet, rows=2):
    """標準格式：表頭橘底粗體置中，資料標楷體 10、置中、自動換行、細框線（列 2～rows）。"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, sheet.max_column + 1):
        head = sheet.cell(1, col)
        head.font = Font(name='標楷體', size=10, bold=True)
        head.fill = PatternFill('solid', fgColor='FFF9CB9C')
        head.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        head.border = border
        for row in range(2, rows + 1):
            cell = sheet.cell(row, col)
            cell.font = Font(name='標楷體', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
    standard_columns(sheet)
    sheet.row_dimensions[1].height = 34
    sheet.freeze_panes = 'E2'


def new_template(path, year, sheet_names=('橋梁總表',)):
    """建立空白標準總表：只要填序號、橋梁名稱（二代系統ID 知道就填，不知道留空由程式比對）。"""
    from openpyxl import Workbook
    book = Workbook()
    for index, name in enumerate(sheet_names):
        sheet = book.active if index == 0 else book.create_sheet()
        sheet.title = name
        sheet.append(template_heads(year))
        style_template(sheet, 30)
        add_notes(sheet)
    add_guide(book)
    book.save(path)
    return Path(path)


# 導入時清空的本期欄（換年度：上一期的「本期」資料由程式重新整理）；本期淨高有值時先移到前期淨高
CURRENT_HEADS = ('新增劣化(U=1)', '新增劣化(U=2)', '新增劣化(U≧3)', '本期梁下/橋下淨高(m)', '系統填寫完成',
                 '應維修構件(ch5.2)', '應維修處理建議', '河工設施損壞(ch5.4)', '六張基本照片拍攝')
NEW_MARK, REMOVED_MARK = '新增', '解除管養'


def bridge_rows(sheet):
    """工作表中的橋：[(列號, 二代系統ID, 橋名)]。"""
    heads = [c.value for c in sheet[1]]
    id_col, name_col = column(heads, '二代系統ID'), column(heads, '橋梁名稱')
    rows = []
    for row in range(2, sheet.max_row + 1):
        name = str(sheet.cell(row, name_col).value or '').strip() if name_col else ''
        value = sheet.cell(row, id_col).value if id_col else None
        ident = str(int(value)) if isinstance(value, (int, float)) else str(value or '').strip()
        if name or ident:
            rows.append((row, ident, name))
    return rows


def import_prior(prior_path, output, year, current_path=None):
    """導入前期總表：各工作表轉成標準格式（同義表頭搬移），清空本期欄（本期淨高移到前期淨高）、序號顏色。
    current_path（本期橋梁清單，可省略）：清單有、前期沒有的橋加在對應工作表最後，序號註記「新增」；
    前期有、清單沒有的橋序號註記「解除管養」。回傳說明。"""
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    book = load_workbook(prior_path)
    sheets = list(book.worksheets)
    template = book.create_sheet('__範本__')            # 範本放在同一個活頁簿（樣式才能複製），用完刪除
    template.append(template_heads(year))
    template.append([None] * len(template_heads(year)))
    style_template(template, 2)
    report = []
    for sheet in sheets:
        if is_guide(sheet):
            continue
        if not any(column([c.value for c in sheet[1]], h) for h in ('橋梁名稱', '二代系統ID')):
            report.append(f'{sheet.title}：沒有「橋梁名稱」欄，略過')
            continue
        count, dropped = convert_sheet(sheet, template)
        heads = [c.value for c in sheet[1]]
        before, now = column(heads, '前期梁下/橋下淨高(m)'), column(heads, '本期梁下/橋下淨高(m)')
        for row in range(2, sheet.max_row + 1):
            if now and before and sheet.cell(row, now).value not in (None, ''):
                sheet.cell(row, before).value = sheet.cell(row, now).value
            for name in CURRENT_HEADS:
                col = column(heads, name)
                if col:
                    sheet.cell(row, col).value = None
            serial = sheet.cell(row, 1)
            serial.fill = PatternFill(fill_type=None)
            for cell in sheet[row]:                     # 換年度：去掉上一期的程式產生備註，本期重新產生
                if cell.comment and GENERATED in cell.comment.text:
                    cell.comment = None
        report.append(f'{sheet.title}：導入 {count} 座' + (f"（沒有對應欄位：{'、'.join(map(str, dropped))}）" if dropped else ''))
    book.remove(template)
    if current_path:
        report += merge_current(book, load_workbook(current_path, data_only=True))
    for sheet in book.worksheets:
        if is_guide(sheet):
            continue
        add_formulas(sheet)
        add_notes(sheet)
        fit(sheet)
    add_guide(book)
    book.save(output)
    return report


def merge_current(book, current):
    """本期橋梁清單對照前期：新增的橋加在同名工作表（沒有就第一個）最後、序號「n（新增）」；前期有、清單沒有的「（解除管養）」。"""
    from copy import copy
    known = {}
    for sheet in book.worksheets:
        if is_guide(sheet):
            continue
        for row, ident, name in bridge_rows(sheet):
            known[(ident or name)] = (sheet, row)
            if ident:
                known[name] = (sheet, row)
    listed, report = set(), []
    data_sheets = [s for s in book.worksheets if not is_guide(s)]
    for source in current.worksheets:
        if is_guide(source):
            continue
        target = book[source.title] if source.title in book.sheetnames else data_sheets[0]
        heads = [c.value for c in target[1]]
        for _, ident, name in bridge_rows(source):
            hit = known.get(ident) if ident else None
            hit = hit or known.get(name)
            if hit:
                listed.add((hit[0].title, hit[1]))
                continue
            row = target.max_row + 1
            serial = len(bridge_rows(target)) + 1
            target.cell(row, 1, f'{serial}（{NEW_MARK}）')
            if ident and column(heads, '二代系統ID'):
                target.cell(row, column(heads, '二代系統ID'), int(ident) if ident.isdigit() else ident)
            target.cell(row, column(heads, '橋梁名稱'), name)
            for col in range(1, len(heads) + 1):
                target.cell(row, col)._style = copy(target.cell(2, col)._style)
            listed.add((target.title, row))
            report.append(f'{target.title}：新增 {name or ident}（第 {row} 列）')
    for sheet in data_sheets:
        for row, ident, name in bridge_rows(sheet):
            if (sheet.title, row) not in listed:
                serial = sheet.cell(row, 1).value
                sheet.cell(row, 1).value = f'{serial}（{REMOVED_MARK}）'
                report.append(f'{sheet.title}：{name or ident} 本期清單沒有，標「解除管養」')
    return report


# ---------------------------------------------------------------- 說明頁（2026-11-24 使用者需求）

GUIDE_SHEET = '說明'
GUIDE = [
    ('顏色與標記', None),
    ('序號 綠底', '前期劣化資料已由程式建檔完成（⑥ 步驟 2「建立前期劣化資料」）。'),
    ('序號 黃底', '尚未完成：建檔中斷、讀取失敗或缺二代系統ID（補 ID 後按「建立前期劣化資料」選「接續」）。'),
    ('序號「（新增）」', '導入前期總表時，本期清單有、前期沒有的橋（加在工作表最後）。'),
    ('序號「（解除管養）」', '導入前期總表時，前期有、本期清單沒有的橋。'),
    ('橋梁名稱 黃底', '「下橋路徑」已填寫（現場檢測完成）。'),
    ('橋梁名稱 綠底', '程式核對、送出、檢測員意見都已完成，「系統填寫完成」欄有日期（優先於黃底）。'),
    ('儲存格 橘底', '與系統資料不同（備註寫有系統值，程式不覆寫，請確認），或本期淨高與前期不同。'),
    ('備註「填報助手產生」', '程式寫入的儲存格。人工修正後，程式同步時不會覆寫；要讓程式重新產生，刪掉內容或備註即可。'),
    ('下橋路徑（ABCD）判定標準', None),
    ('A級（間接檢測）', '須配合無人飛行器（UAV）、水下無人機（ROV）等設備進行檢測。'),
    ('B級（直接檢測）', '須搭載輔助載具始能辦理檢測，如橋梁檢測車、高空工作車或工作船舶等。'),
    ('C級（直接檢測）', '須搭設移動梯（O截梯）等設備方可到達橋下進行檢測。'),
    ('D級（直接檢測）', '檢測人員可徒步到達橋下進行檢測。'),
    ('輔助設備（A/B）', None),
    ('A', '下橋路徑 A級：使用無人飛行器（UAV）、水下無人機（ROV）等設備。'),
    ('B', '下橋路徑 B級：使用橋梁檢測車、高空工作車或工作船舶等輔助載具。'),
    ('C、D 級', '免填。'),
    ('寫法（填報助手 ⑤ 組合）', '第一行寫等級；C、D 級寫「下橋位置＋方式＋下橋」（例：A1下游側架設二截梯下橋），需要時加「水深需要長筒雨鞋／青蛙裝」；'
                       'A、B 級寫「需用○○檢測」，輔助設備寫「等級 換行 設備」（例：B 換行 橋檢車）；最後可加備註。'),
    ('欄位由誰填寫', None),
    ('程式核對（系統資料）', '二代系統ID、橋梁編號、橋梁名稱、道路等級、路線、跨數、橋長、最大淨寬、跨越物體、欄杆高度、橋面板厚（構件→橋孔資料表→橋面板明細，各孔不同取最多孔的值）：空白時填入系統值，'
                       '不同時標橘底並加備註，不覆寫。'),
    ('程式整理', '前期觀測劣化、本期新增劣化（依 U=1、U=2、U≧3 分欄並標年份）、過去維修紀錄、已維修構件狀況、前期／本期淨高、'
             '應維修構件、應維修處理建議、河工設施損壞、護欄高度不足、橋下堆積物、疑似鹽害橋梁、六張基本照片拍攝、系統填寫完成。'),
    ('公式', '橋孔面積＝橋長×最大淨寬；欄杆至河床高度＝梁下淨高（本期有填用本期，否則前期）＋欄杆高度＋橋面板厚＋大梁高度，'
           '多跨逐墩計算。'),
    ('現場選擇', '下橋路徑、輔助設備：在填報助手 ⑤ 比對資料輸出選擇（或手打），隨比對資料輸出、同步時寫入；也可直接在總表填寫，'
             '填報助手開啟該橋時會讀回。'),
    ('人工填寫', '序號、基本資料變動、大梁高度、橋名牌損壞／遺失／內容錯誤、是否有限重牌、其他檢測特殊情形。'),
    ('寫法', None),
    ('淨高', '多跨寫 P01=3.50（換行 P02=3.40…，P 補兩位數、兩位小數）；單跨寫 2.20。'),
    ('護欄高度不足', '2026 年起低於 110cm、之前低於 80cm 填「是」，否則填「否」。'),
    ('是／否欄位', '護欄高度不足、橋下堆積物、疑似鹽害橋梁、六張基本照片拍攝：直接填「是」或「否」（程式判斷不出來時留空）。'),
    ('劣化', '每條「位置＋構件＋樣態(年份)」，同樣態合併；敘述依劣化說明與維修單價 Excel 規則修正。'),
]


def add_guide(book):
    """活頁簿最前面加（或更新）「說明」工作表：顏色意義、下橋路徑與輔助設備判定、欄位由誰填寫。"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    opened = book.active.title if book.worksheets else ''
    if GUIDE_SHEET in book.sheetnames:
        book.remove(book[GUIDE_SHEET])
    sheet = book.create_sheet(GUIDE_SHEET, 0)
    thin = Side(style='thin', color='FF999999')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    sheet['A1'] = '橋梁總表說明（由公路橋梁填報助手產生）'
    sheet['A1'].font = Font(name='標楷體', size=14, bold=True)
    samples = {'序號 綠底': 'FF92D050', '序號 黃底': 'FFFFFF00', '橋梁名稱 黃底': 'FFFFFF00',
               '橋梁名稱 綠底': 'FFB7E1CD', '儲存格 橘底': 'FFF8CBAD'}
    row = 3
    for name, text in GUIDE:
        if text is None:                                 # 段落標題
            row += 1 if row > 3 else 0
            cell = sheet.cell(row, 1, name)
            cell.font = Font(name='標楷體', size=12, bold=True)
            cell.fill = PatternFill('solid', fgColor='FFF9CB9C')
            sheet.cell(row, 2).fill = PatternFill('solid', fgColor='FFF9CB9C')
            row += 1
            continue
        left, right = sheet.cell(row, 1, name), sheet.cell(row, 2, text)
        for cell in (left, right):
            cell.font = Font(name='標楷體', size=11)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = border
        if name in samples:
            left.fill = PatternFill('solid', fgColor=samples[name])
        sheet.row_dimensions[row].height = max(18, 16 * (len(text) // 45 + 1))
        row += 1
    sheet.column_dimensions['A'].width = 26
    sheet.column_dimensions['B'].width = 95
    sheet.page_setup.orientation = 'landscape'           # 列印：一頁寬
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
    data = [s.title for s in book.worksheets if s.title != GUIDE_SHEET]  # 說明頁放最前面，開啟時仍停在原本的工作表
    target = opened if opened in data else (data[0] if data else GUIDE_SHEET)
    book.active = book.sheetnames.index(target)
    return sheet


def is_guide(sheet):
    return sheet.title == GUIDE_SHEET
