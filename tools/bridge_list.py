"""縣市橋梁清單（Google 試算表下載的 xlsx）：舊格式工作表改成新格式，並加入「欄杆至河床高度」公式。

2026-11-22 使用者需求：以「西螺」工作表為範本——
- 歷年持續觀測劣化、本年度檢測新增劣化各拆成 U=1、U=2、U≧3 三欄；刪除「漏電檢測照片拍攝」；
- 舊資料搬移：歷年(U≦2) → 歷年 U=2、歷年(U≧3) → 歷年 U≧3、新增(U≦2) → 新增 U=2、新增(U≧3) → 新增 U≧3
  （舊表沒有分 U=1、U=2，先放 U=2，請人工確認）；
- 格式依西螺：欄寬、表頭樣式、各欄對齊、自動換行、框線；列高依文字行數調整；
- 欄杆至河床高度 ＝ 梁下淨高 ＋ 欄杆高度 ＋ 橋面板厚 ＋ 大梁高度（公式；多跨「P1=3.5 換行 P2=3.4」逐墩計算）。

用法：python tools/bridge_list.py 來源.xlsx 輸出.xlsx [範本工作表名稱，預設 西螺]
"""
import math
import re
import sys
import unicodedata
from copy import copy

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter

NEW_HEADS = ('序號', '二代系統ID', '橋梁編號', '橋梁名稱', '道路等級', '路線', '橋長', '跨越物體', '下橋路徑', '輔助設備',
             '歷年持續觀測劣化(U=1)', '歷年持續觀測劣化(U=2)', '歷年持續觀測劣化(U≧3)',
             '檢測新增劣化(U=1)', '檢測新增劣化(U=2)', '檢測新增劣化(U≧3)', '歷年維修', '已維修構件狀況',
             '梁下淨高(m)', '欄杆至河床高度(m)', '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)')
MAX_SPANS = 8              # 公式最多處理幾跨（P1～P8）；實際資料最多 7 跨
LINE_HEIGHT = 14.25        # 標楷體 10pt 每行高度（西螺工作表的列高為 14.25 的倍數）


def key(text):
    """表頭比對用：去掉空白與換行、全形轉半形、年度數字（115年檢測新增劣化 → 檢測新增劣化）。"""
    text = unicodedata.normalize('NFKC', str(text or ''))
    text = re.sub(r'\s+', '', text)
    text = re.sub(r'^\d+年', '', text)
    return text.replace('≦', '<=').replace('≧', '>=').replace('須修正', '')


def is_new_format(sheet):
    heads = [key(c.value) for c in sheet[1]]
    return key('歷年持續觀測劣化(U=1)') in heads


def old_to_new(heads):
    """舊表頭（33 欄）→ 新欄位位置的對應 {新欄索引: 舊欄索引}；None 表示新欄位沒有舊資料。"""
    old = {key(h): i for i, h in enumerate(heads)}
    moved = {key('歷年持續觀測劣化(U=2)'): key('歷年持續觀測劣化(U≦2)'),
             key('歷年持續觀測劣化(U≧3)'): key('歷年持續觀測劣化(U≧3)'),
             key('檢測新增劣化(U=2)'): key('檢測新增劣化(U≦2)'),
             key('檢測新增劣化(U≧3)'): key('檢測新增劣化(U≧3)')}
    return old, moved


def railing_formula(row, columns):
    """欄杆至河床高度公式（Excel 與 Google 試算表都能用的一般函數，不用 LET／REGEX）。
    梁下淨高為數字：直接相加；為「P1=3.5 換行 P2=3.4」（或 P1:3.5m）：逐行加總，結果寫成「P1=4.8 換行 P2=4.7」。
    欄杆高度、橋面板厚、大梁高度任一未填 → 空白。"""
    s, u, v, w = (f'{columns[name]}{row}' for name in ('梁下淨高(m)', '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)'))
    extra = f'({u}+{v}+{w})'
    text = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({s},"：","="),":","="),"m","")'
    pieces = []
    for k in range(1, MAX_SPANS + 1):
        line = f'TRIM(MID(SUBSTITUTE({text},CHAR(10),REPT(" ",99)),{(k - 1) * 99 + 1},99))'
        lead = '' if k == 1 else 'CHAR(10)&'
        value = f'ROUND(VALUE(MID({line},FIND("=",{line})+1,20))+{extra},2)'
        pieces.append(f'IF({line}="","",{lead}IFERROR(LEFT({line},FIND("=",{line}))&{value},{line}))')
    spans = '&'.join(pieces)
    return f'=IF(OR({s}="",COUNT({u}:{w})<3),"",IFERROR(ROUND({s}+{extra},2),{spans}))'


def text_lines(value, width):
    """儲存格在欄寬 width 下大約幾行（中文字算 2、英數 1；欄寬單位約等於 1 個英數字）。"""
    if value in (None, '') or str(value).startswith('='):      # 公式（欄杆至河床高度）行數同梁下淨高，不另計
        return 1
    per_line = max(int((width or 8.43) * 1.1), 2)
    total = 0
    for line in str(value).split('\n'):
        size = sum(2 if unicodedata.east_asian_width(ch) in 'WF' else 1 for ch in line.strip())
        total += max(1, math.ceil(size / per_line))
    return total


def fit_rows(sheet, last_row):
    """列高依文字行數（自動換行後）調整；表頭兩行。"""
    widths = {i: sheet.column_dimensions[get_column_letter(i)].width for i in range(1, sheet.max_column + 1)}
    for row in range(2, last_row + 1):
        lines = max(text_lines(sheet.cell(row, col).value, widths[col]) for col in widths)
        sheet.row_dimensions[row].height = max(lines, 1) * LINE_HEIGHT + (2.25 if lines == 1 else 6)   # 多行留一點餘裕
    sheet.row_dimensions[1].height = 2 * LINE_HEIGHT


def last_data_row(sheet, name_col=4):
    rows = [r for r in range(2, sheet.max_row + 1) if sheet.cell(r, name_col).value not in (None, '')]
    return max(rows) if rows else 1


def convert_sheet(sheet, template):
    """舊格式工作表 → 新格式（就地改寫）。回傳搬移的列數。"""
    heads = [c.value for c in sheet[1]]
    old, moved = old_to_new(heads)
    last = last_data_row(sheet)
    rows = [[sheet.cell(r, c).value for c in range(1, len(heads) + 1)] for r in range(2, last + 1)]
    template_heads = [c.value for c in template[1]]
    freeze = sheet.freeze_panes
    frozen_head = heads[sheet[freeze].column - 1] if freeze else None
    rules = list(sheet.conditional_formatting._cf_rules.items())
    sheet.conditional_formatting = ConditionalFormattingList()
    sheet.delete_rows(1, sheet.max_row)
    sheet.delete_cols(1, sheet.max_column + 5)
    for col, head in enumerate(template_heads, 1):
        cell = sheet.cell(1, col, head)
        cell._style = copy(template.cell(1, col)._style)
    for r, values in enumerate(rows, 2):
        for col, head in enumerate(template_heads, 1):
            name = key(head)
            source = old.get(moved.get(name, name))
            value = values[source] if source is not None and source < len(values) else None
            if value is None and name == key('檢測新增劣化(U=1)') and any(v not in (None, '') for v in values[10:16]):
                value = '無'                              # 同西螺：新增 U=1 先填「無」（歷年 U=1 留白）
            cell = sheet.cell(r, col, value)
            cell._style = copy(template.cell(2, col)._style)
    for col in range(1, len(template_heads) + 1):
        letter = get_column_letter(col)
        source = template.column_dimensions[letter]
        sheet.column_dimensions[letter].width = source.width or 8.43
    if frozen_head:
        new_col = next((i for i, h in enumerate(template_heads, 1) if key(h) == key(frozen_head)), None)
        sheet.freeze_panes = f'{get_column_letter(new_col)}2' if new_col else freeze
    for ranges, items in rules:                       # 條件式格式（例：I 欄有填時橋名變黃）：只保留橋名欄的規則
        if str(ranges.sqref).startswith('D'):
            for rule in items:
                sheet.conditional_formatting.add(str(ranges.sqref), rule)
    return len(rows)


def add_formulas(sheet):
    heads = [c.value for c in sheet[1]]
    columns = {}
    for name in ('梁下淨高(m)', '欄杆至河床高度(m)', '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)'):
        index = next(i for i, h in enumerate(heads, 1) if key(h) == key(name))
        columns[name] = get_column_letter(index)
    target = heads.index(next(h for h in heads if key(h) == key('欄杆至河床高度(m)'))) + 1
    count = 0
    for row in range(2, last_data_row(sheet) + 1):
        sheet.cell(row, target).value = railing_formula(row, columns)
        count += 1
    return count


def convert(source, output, template_name='西螺'):
    workbook = load_workbook(source)
    template = workbook[template_name]
    report = []
    for sheet in workbook.worksheets:
        if sheet is not template and not is_new_format(sheet):
            report.append(f'{sheet.title}：轉換 {convert_sheet(sheet, template)} 列')
        formulas = add_formulas(sheet)
        if sheet is not template:
            fit_rows(sheet, last_data_row(sheet))
        report.append(f'{sheet.title}：欄杆至河床高度公式 {formulas} 列')
    workbook.save(output)
    return report


if __name__ == '__main__':
    for line in convert(*sys.argv[1:4]):
        print(line)
