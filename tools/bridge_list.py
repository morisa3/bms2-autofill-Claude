"""縣市橋梁清單（Google 試算表下載的 xlsx）：各工作表依範本工作表（預設「西螺」）改成相同欄位與格式，並加入公式。

使用者需求（2026-11-22、11-23）：以「西螺」工作表為範本——
- 欄位依範本表頭；其他工作表的資料依「同義表頭」搬到對應欄位（例：歷年持續觀測劣化 → 前期觀測劣化、
  歷年維修 → 過去維修紀錄、梁下淨高 → 前期梁下／橋下淨高；最早的格式 U≦2 → U=2），範本新增的欄位留白；
- 格式依範本：欄寬、表頭樣式、各欄對齊、自動換行、框線；列高依文字行數調整；
- 公式：橋孔面積 ＝ 橋長 × 最大淨寬；欄杆至河床高度 ＝ 梁下淨高（本期有填用本期，否則用前期）＋ 欄杆高度 ＋ 橋面板厚 ＋
  大梁高度（多跨「P1=3.5 換行 P2=3.4」逐墩計算）；
- 表頭備註：下橋路徑（A～D 級判定標準）、輔助設備（依下橋路徑 A／B 級）。

用法：python tools/bridge_list.py 來源.xlsx 輸出.xlsx [範本工作表名稱，預設 西螺]
"""
import math
import re
import sys
import unicodedata
from copy import copy

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[1]))
from bridgefill.core.bridge_list import HEIGHT_FORMAT, fit_columns, fit_rows, format_height  # noqa: E402

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
    return text


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


def fit(sheet):
    """欄寬依文字、列高依換行後行數（與程式同步時用同一套計算，文字完整顯示）。"""
    tidy_text(sheet)
    tidy_heights(sheet)
    fit_columns(sheet, last_data_row(sheet))
    fit_rows(sheet, range(1, last_data_row(sheet) + 1))


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
    freeze = sheet.freeze_panes
    frozen = sheet[freeze].column - 1 if freeze else None          # 凍結窗格的舊欄索引
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
    if frozen is not None:                             # 凍結窗格跟著原本那一欄移動
        new_col = next((i for i, source in enumerate(sources, 1) if source == frozen), None)
        sheet.freeze_panes = f'{get_column_letter(new_col)}2' if new_col else freeze
    for ranges, items in name_rules:                   # 「下橋路徑有填 → 橋名變黃」：改指新的下橋路徑欄
        for rule in items:
            rule.formula = [f'${get_column_letter(path_col)}2<>""']
            sheet.conditional_formatting.add(ranges, rule)
    return len(rows), dropped


def add_formulas(sheet):
    """橋孔面積、欄杆至河床高度公式；回傳加入的列數。"""
    heads = [c.value for c in sheet[1]]
    letter = lambda name: get_column_letter(column(heads, name)) if column(heads, name) else None   # noqa: E731
    area, length, width = letter('橋孔面積(橋長*淨寬)'), letter('橋長'), letter('最大淨寬')
    railing = letter('欄杆至河床高度(m)')
    now = letter('本期梁下/橋下淨高(m)') or letter('梁下淨高(m)')
    columns = {'淨高': now, '前期淨高': letter('前期梁下/橋下淨高(m)'),
               '欄杆高度': letter('欄杆高度(m)'), '橋面板厚': letter('橋面板厚(m)'), '大梁高度': letter('大梁高度(m)')}
    count = 0
    for row in range(2, last_data_row(sheet) + 1):
        if area and length and width:
            sheet[f'{area}{row}'] = area_formula(row, length, width)
        if railing and now and all(columns[k] for k in ('欄杆高度', '橋面板厚', '大梁高度')):
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
    for sheet in workbook.worksheets:
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
    workbook.save(output)
    return report


if __name__ == '__main__':
    for line in convert(*sys.argv[1:4]):
        print(line)
