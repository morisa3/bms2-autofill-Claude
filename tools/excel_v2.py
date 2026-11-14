"""把 v1 劣化維修規則 Excel 轉為 v2 建議格式。

原本 7 欄（劣化類型、劣化樣態、維修工法、維修單位、單價、備註、劣化ID）一字不改，
只在右側新增結構化欄位，並在檔尾新增對照工作表。v1.00 依表頭名稱讀取且略過
缺少必要表頭的工作表，所以轉換後的檔案仍可給 v1.00 使用。

用法：
    python tools/excel_v2.py 來源.xlsx 輸出.xlsx
"""
import re
import sys
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = '標楷體'
ORIGINAL_HEADERS = ('劣化類型', '劣化樣態', '維修工法', '維修單位', '單價', '備註', '劣化ID')
NEW_COLUMNS = (
    ('處理方式', 12, '自動：程式可直接套用。\n手動：一律人工選擇。\n手動或條件：符合條件欄位時自動，否則人工。'),
    ('任一關鍵字', 26, '備註含其中任一個詞即符合；多個詞以「、」分隔。'),
    ('必要關鍵字', 20, '備註必須全部包含；多個詞以「、」分隔。與「任一關鍵字」同時有值時，兩者都要成立。'),
    ('排除關鍵字', 18, '備註含其中任一個詞就不套用此規則。原本寫死在程式中，之後由此欄管理。'),
    ('前置判讀', 18, '先確認構件材質或型式（例：齒型、角鋼、人造橡膠支承墊），再判讀關鍵字。'),
    ('條件欄位', 12, '需要數值判斷時填寫，例：裂縫寬度、護欄高度。'),
    ('最小值', 9, '含等於（≧）。'),
    ('最大值', 9, '不含等於（＜）。'),
    ('條件單位', 9, 'mm、m 等。'),
    ('單獨關鍵字', 14, '備註只寫這個詞、沒有更明確的寫法時使用（例：只寫「植生」）。'),
    ('單獨關鍵字U值', 10, '配合「單獨關鍵字」：U 值符合時才優先使用。空白表示不看 U 值。'),
    ('前期關鍵字', 14, '前期備註含此詞時優先使用本規則（例：前期寫 AC老化）。'),
    ('優先序', 8, '同一劣化樣態有多筆候選時，數字小的優先；空白為一般。'),
    ('修正劣化類型為', 22, '符合本規則但劣化類型填錯時，自動改成此類型。'),
    ('備註範本', 30, '自動產生備註時使用。{位置}、{樣態} 會替換成實際文字；尺寸由檢測員補上數值。'),
    ('轉換狀態', 12, '自動轉換：已由原「備註」完整轉成結構化欄位。\n需確認：有無法完全轉換的條件，請看「轉換說明」。'),
    ('轉換說明', 40, '需確認的原因，或轉換時的補充說明。'),
)
COUNT_UNITS = ('處', '組', '座', '孔', '個', '塊')

HEADER_FILL = PatternFill('solid', fgColor='FFC000')     # 沿用原檔表頭色
NEW_HEADER_FILL = PatternFill('solid', fgColor='9DC3E6')  # 新增欄位表頭
CHECK_FILL = PatternFill('solid', fgColor='FFFF00')       # 需確認／需填寫
EXAMPLE_FILL = PatternFill('solid', fgColor='E2EFDA')     # 範例或程式擷取的草稿
THIN = Side(style='thin', color='808080')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')


def words(text):
    """取出 [..] 中的詞；[A][B][C] 連寫視為並列。"""
    return [w.strip() for w in re.findall(r'\[([^\]]*)\]', text) if w.strip()]


def joined(values):
    return '、'.join(dict.fromkeys(values))


def parse_note(note, unit, pattern):
    """把原「備註」轉成結構化欄位。回傳 (欄位 dict, 是否需確認, 說明)。"""
    note = (note or '').strip()
    out = {name: '' for name, _, _ in NEW_COLUMNS}
    out['處理方式'] = '自動'
    review, remarks = False, []
    rest = note

    def take(regex):
        nonlocal rest
        found = re.search(regex, rest)
        if found:
            rest = (rest[:found.start()] + rest[found.end():]).strip('，, ')
        return found

    if not note:
        return out, True, '原備註空白，請補上適用條件'
    if note == '手動處理' or not pattern:
        out['處理方式'] = '手動'
        return out, False, '原備註：手動處理'

    if take(r'^(優先使用|優先選擇)[，,]?'):
        out['優先序'] = 1

    guard = take(r'(?:(?:手動選擇|優先選擇)[，,]?)?或?讀取到護欄高度'
                 r'(?:(大於等於|大於)([\d.]+)m?)?(?:小於([\d.]+)m?)?時使用')
    if guard and (guard.group(2) or guard.group(3)):
        out.update({'處理方式': '手動或條件', '條件欄位': '護欄高度', '條件單位': 'm'})
        if guard.group(2):
            out['最小值'] = float(guard.group(2))
        if guard.group(3):
            out['最大值'] = float(guard.group(3))
        if guard.group(1) == '大於':
            review = True
            remarks.append(f'原文「大於{guard.group(2)}小於{guard.group(3)}」兩端都不含等於；'
                           f'新格式最小值含等於，請確認 {guard.group(2)}m 剛好時歸哪一筆')
        else:
            remarks.append('讀不到護欄高度時轉人工選擇')

    crack = take(r'裂縫寬度(大於|小於)\[([\d.]+)\s*(mm|cm)\]使用')
    if crack:
        value, unit_text = float(crack.group(2)), crack.group(3)
        out.update({'條件欄位': '裂縫寬度', '條件單位': unit_text})
        if crack.group(1) == '大於':
            out['最小值'] = value
            remarks.append(f'依 v0.9.7.6 修正：剛好 {value}{unit_text} 歸入「大於」這一筆')
        else:
            out['最大值'] = value

    pre = take(r'先判讀是否為\[([^\]]+)\][，,]?[在再]判讀')
    if pre:
        out['前置判讀'] = pre.group(1)

    alone = take(r'或僅有關鍵字\[([^\]]+)\]時[，,]?(?:則|若)U值為(\d)優先使用')
    if alone:
        out['單獨關鍵字'], out['單獨關鍵字U值'] = alone.group(1), int(alone.group(2))

    only = take(r'若僅寫\[([^\]]+)\]則優先使用此項')
    if only:
        out['單獨關鍵字'] = only.group(1)

    history = take(r'若過去資料為\[([^\]]+)\][，,]?則優先(?:使用|填寫\[[^\]]+\])')
    if history:
        out['前期關鍵字'] = history.group(1)
        out['優先序'] = 1

    fix = take(r'但劣化類型填成\[([^\]]+)\][，,]?則(?:直接)?修正成\[([^\]]+)\](?:後使用)?')
    if fix:
        out['修正劣化類型為'] = fix.group(2)
        remarks.append(f'原劣化類型「{fix.group(1)}」')

    by_unit = take(r'依照\[單位\]套用工法與單價')
    if by_unit:
        remarks.append('依維修單位選擇本列或同樣態的其他單位')

    damaged = take(r'(?:若傷及混凝土層時使用[，,]?|\(傷及RC層\))')
    if damaged:
        review = True
        remarks.append('適用條件「傷及混凝土層（RC層）」無法從備註文字判斷，需人工確認或補排除關鍵字')

    owner = take(r'根據構件位置在基礎前補上\[橋台\]或\[橋墩\][，,]?構件英文A為橋台[，,]?英文P為橋墩')
    if owner:
        out['備註範本'] = '{位置} {橋台或橋墩}基礎掏空。(m*m*m)*1'
        remarks.append('{橋台或橋墩}：位置 A 開頭寫橋台、P 開頭寫橋墩')

    keys = take(r'^(?:若判讀到)?(?:有)?關鍵字(.+?)即可使用')
    if keys:
        expr = keys.group(1)
        if '與' in expr:
            left, right = expr.rsplit('與', 1)
            out['必要關鍵字'] = joined(words(right))
            if '或' in left:
                out['任一關鍵字'] = joined(words(left))
            else:
                out['必要關鍵字'] = joined(words(left) + words(right))
        else:
            out['任一關鍵字'] = joined(words(expr))
    elif out['前置判讀'] and rest:
        # 「先判讀是否為[齒型]，在判讀關鍵字[變形]或[開裂]」
        after = re.sub(r'^關鍵字', '', rest)
        out['任一關鍵字'] = joined(words(after))
        rest = ''

    rest = rest.strip('，, 。')
    if rest:
        review = True
        remarks.append(f'未轉換的原文：{rest}')

    if not out['備註範本']:
        out['備註範本'] = template(unit, pattern)
    return out, review, '；'.join(remarks)


def template(unit, pattern):
    if not pattern or not unit:
        return ''
    if unit == '公尺':
        size = '(mm*m)*1' if '裂縫' in pattern else '(m)*1'
    elif unit == '平方公尺':
        size = '(m*m)*1'
    elif unit == '立方公尺':
        size = '(m*m*m)*1'
    elif unit in COUNT_UNITS:
        size = f'(1{unit})'
    else:
        return ''  # 「式」的尺寸寫法依實況，不給預設
    return f'{{位置}} {{樣態}}。{size}'


def style_header(cell, fill, size=14):
    cell.font = Font(name=FONT, size=size, bold=True)
    cell.fill = fill
    cell.border = BORDER
    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')


def style_body(cell, fill=None):
    cell.font = Font(name=FONT, size=11)
    cell.border = BORDER
    cell.alignment = WRAP
    if fill:
        cell.fill = fill


def convert_rule_sheet(ws, stats):
    headers = [c.value for c in ws[1]]
    if tuple(headers[:7]) != ORIGINAL_HEADERS:
        raise ValueError(f'{ws.title}：表頭與 v1 不同，停止轉換')
    first_new = len(ORIGINAL_HEADERS) + 1
    status_col = first_new + [n for n, _, _ in NEW_COLUMNS].index('轉換狀態')
    for offset, (name, width, tip) in enumerate(NEW_COLUMNS):
        cell = ws.cell(1, first_new + offset, name)
        style_header(cell, NEW_HEADER_FILL, size=12)
        cell.comment = Comment(tip, '轉換程式')
        ws.column_dimensions[get_column_letter(first_new + offset)].width = width
    for c in ws[1][:7]:
        c.comment = None
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:{get_column_letter(first_new + len(NEW_COLUMNS) - 1)}{ws.max_row}'

    choices = DataValidation(type='list', formula1='"自動,手動,手動或條件"', allow_blank=True)
    states = DataValidation(type='list', formula1='"自動轉換,需確認,已確認"', allow_blank=True)
    ws.add_data_validation(choices)
    ws.add_data_validation(states)

    for row in range(2, ws.max_row + 1):
        values = [ws.cell(row, c).value for c in range(1, 8)]
        if not any(values):
            continue
        fields, review, remark = parse_note(values[5], values[3], values[1])
        fields['轉換狀態'] = '需確認' if review else '自動轉換'
        fields['轉換說明'] = remark
        for offset, (name, _, _) in enumerate(NEW_COLUMNS):
            cell = ws.cell(row, first_new + offset, fields[name] if fields[name] != '' else None)
            style_body(cell, CHECK_FILL if review and name in ('轉換狀態', '轉換說明') else None)
        choices.add(ws.cell(row, first_new))
        states.add(ws.cell(row, status_col))
        stats['rows'] += 1
        stats['review'] += review
    return status_col


def table_sheet(wb, title, columns, rows, note, fill_rows=None):
    """建立新的對照工作表：第 1 列說明、第 2 列表頭、第 3 列起資料。"""
    ws = wb.create_sheet(title)
    ws.cell(1, 1, note).font = Font(name=FONT, size=11, italic=True, color='595959')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    ws.cell(1, 1).alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[1].height = 48
    for i, (name, width) in enumerate(columns, 1):
        style_header(ws.cell(2, i, name), NEW_HEADER_FILL, size=12)
        ws.column_dimensions[get_column_letter(i)].width = width
    for r, values in enumerate(rows, 3):
        fill = (fill_rows or {}).get(r - 3)
        for c, value in enumerate(values, 1):
            style_body(ws.cell(r, c, value), fill)
    ws.freeze_panes = 'A3'
    return ws


def add_reference_sheets(wb):
    table_sheet(wb, '樣態同義詞', [('標準樣態', 22), ('同義寫法', 36), ('來源', 40)], [
        ('混凝土裂縫', '明顯裂縫、混凝土結構裂縫', '需求說明本期情況 1 的範例'),
        ('混凝土剝落鋼筋外露', '鋼筋外露鏽蝕', '需求說明本期情況 2 的範例'),
        ('路面混凝土層破洞', '路面坑洞', '2026-11-14 修訂：1-05、14-36 改名'),
        ('生銹', '生鏽、生锈', 'v0.9.7.6 修正：三種寫法視為相同'),
        ('防衝鋼板表面生銹', '防撞鋼板生銹、防衝鋼板生鏽', '需求說明 2.1 範例；橋墩／橋基保護設施只寫「生鏽」的情況由 Excel 11-04 處理'),
    ], '備註寫法不同但意思相同時，程式先轉成「標準樣態」再比對。左側寫一個標準寫法，右側列出所有同義寫法，以「、」分隔。'
       '綠底為依需求說明與舊版修正紀錄整理的草稿，請確認或增修。',
        fill_rows={i: EXAMPLE_FILL for i in range(5)})

    table_sheet(wb, '互斥樣態', [('樣態A', 14), ('樣態B', 14), ('說明', 50)], [
        ('木本', '草本', '植生不可互換'),
        ('齒型', '角鋼', '伸縮縫型式不同'),
        ('AC', '混凝土', '鋪面材料不同'),
        ('瀝青', '混凝土', '鋪面材料不同'),
        ('鋼管', '鋼筋', '構材不同'),
        ('鋼板', '鋼管', '構材不同'),
    ], '前期與本期若分別出現 A 與 B，一律不配對。以下 6 組原本寫死在 v1.00 程式（alignment.py）中，搬到這裡之後可自行增減。',
        fill_rows={i: EXAMPLE_FILL for i in range(6)})

    table_sheet(wb, '劣化類型等價', [('群組', 8), ('劣化類型', 34), ('適用構件', 24), ('說明', 40)], [
        (1, '保護設施損壞、移動或遺失', '橋墩／橋基保護設施', '需求說明 2.1 範例；Excel 11-04'),
        (1, '防撞鋼板損傷', '橋墩／橋基保護設施', ''),
        (1, '其他損傷', '橋墩／橋基保護設施', 'Excel 11-08'),
        (2, '混凝土結構裂縫', '（全部）', 'v0.9.7.6 修正：混凝土裂縫／混凝土結構裂縫類型對應'),
        (2, '混凝土裂縫', '（全部）', ''),
        (None, None, None, '← 請由此往下補充'),
    ], '檢測員對劣化類型判斷不同、但實際是同一筆缺失時（需求 2.1），同一「群組」編號的劣化類型可以互相配對。'
       '「適用構件」空白或寫（全部）代表所有構件。綠底為草稿，黃底請補充。',
        fill_rows={**{i: EXAMPLE_FILL for i in range(5)}, 5: CHECK_FILL})

    table_sheet(wb, '位置代碼對照', [('寫法', 16), ('標準代碼', 12), ('說明', 50)], [
        ('橋頭', 'A1', '構件身分比對時使用；不改動備註中的實測位置文字'),
        ('橋尾', 'A2', '同上'),
        ('P010', 'P10', '數字前導零一律去除（A001→A1、S001→S1 同理）'),
        ('A1-L', 'A1-L', '左右側標記 -L／-R 一律保留，不可刪除'),
    ], '位置寫法不同時，程式先轉成標準代碼再比對。以下為 v1.00 已有的規則，列出供確認。',
        fill_rows={i: EXAMPLE_FILL for i in range(4)})

    # 依《橋梁檢測手冊》第四章 4.1、4.2 分類；「手冊」為手冊明訂，「草稿」待使用者確認
    components = (
        ('主梁', '上部結構', '是', '否', '手冊', '4.1.1 主構件'),
        ('橫隔梁', '上部結構', '否', '否', '手冊', '4.1.1 次要構件'),
        ('支承／支承墊／阻尼裝置', '上部結構', '否', '否', '草稿', '4.1.1'),
        ('防落設施', '上部結構', '否', '否', '草稿', '4.1.1'),
        ('鋼纜錨碇裝置(橋塔端)', '上部結構', '是', '否', '草稿', ''),
        ('鋼纜錨碇裝置(主梁端)', '上部結構', '是', '否', '草稿', ''),
        ('鋼纜(吊索)-斜張鋼纜', '上部結構', '是', '否', '草稿', ''),
        ('鋼纜保護套管-斜張鋼纜', '上部結構', '否', '否', '草稿', ''),
        ('墩柱與帽梁', '下部結構', '是', '否', '草稿', '4.1.2 橋墩、帽梁'),
        ('橋台', '下部結構', '是', '否', '草稿', '4.1.2'),
        ('橋台基礎', '下部結構', '是', '否', '草稿', '4.1.2'),
        ('橋墩基礎', '下部結構', '是', '否', '草稿', '4.1.2'),
        ('翼牆／擋土牆', '下部結構', '否', '否', '草稿', '4.1.2'),
        ('橋面板', '橋面系統及附屬設施', '是', '是', '草稿', '4.1.3'),
        ('伸縮縫', '橋面系統及附屬設施', '否', '是', '草稿', '4.1.3'),
        ('橋護欄', '橋面系統及附屬設施', '否', '是', '草稿', '4.1.3'),
        ('橋梁排水設施', '橋面系統及附屬設施', '否', '否', '草稿', '4.1.3'),
        ('河道', '河道及相關設施', '否', '否', '草稿', '4.2'),
        ('引道路堤', '河道及相關設施', '否', '是', '草稿', '4.2'),
        ('引道路堤護欄', '河道及相關設施', '否', '是', '草稿', '4.2'),
        ('引道道路堤保護措施', '河道及相關設施', '否', '否', '草稿', '4.2'),
        ('橋墩／橋基保護設施', '河道及相關設施', '否', '否', '草稿', '4.2 基礎保護措施'),
        ('護床工', '河道及相關設施', '否', '否', '草稿', '4.2 基礎保護措施'),
        ('其他', '其他', '否', '是', '草稿', '多為 AC 舖面；手冊：舖面不列入定期檢測，另於巡查辦理'),
    )
    table_sheet(wb, '構件分類', [('構件', 26), ('系統分類', 20), ('主構件', 8), ('可能影響行車安全', 12),
                                ('依據', 8), ('手冊章節／說明', 44)], components,
        '檢測員意見 U＝2 判斷用：橋面缺失看是否影響行車安全；下部結構先看是否為主構件，再看劣化面積（E 值）。'
        '系統分類依《橋梁檢測手冊》4.1、4.2。依據為「手冊」者為手冊明訂；「草稿」（綠底）請確認後改為「已確認」。')
    ws = wb['構件分類']
    answers = DataValidation(type='list', formula1='"是,否"', allow_blank=True)
    basis = DataValidation(type='list', formula1='"手冊,草稿,已確認"', allow_blank=True)
    ws.add_data_validation(answers)
    ws.add_data_validation(basis)
    for row in range(3, ws.max_row + 1):
        draft = ws.cell(row, 5).value == '草稿'
        for col in (3, 4):
            answers.add(ws.cell(row, col))
            if draft:
                ws.cell(row, col).fill = EXAMPLE_FILL
        basis.add(ws.cell(row, 5))

    table_sheet(wb, 'DERU評等', [('評等', 10), ('0', 16), ('1', 12), ('2', 14), ('3', 12), ('4', 12)], [
        ('程度(D)', '無此項目', '良好', '尚可', '差', '嚴重損壞'),
        ('範圍(E)', '無法檢測', '10%以下', '10%~30%', '30%~60%', '60%以上'),
        ('重要性(R)', '無法判定重要性', '微', '小', '中', '大'),
        ('急迫性(U)', '無法判定急迫性', '例行維護', '3年內維護或持續追蹤', '1年內維護', '緊急處置'),
        ('E下限(%)', None, 0, 10, 30, 60),
        ('E上限(%)', None, 10, 30, 60, 100),
    ], '《橋梁檢測手冊》表 4.1。E＝0 為無法直接目視評等（如地面下之橋台基礎、橋墩基礎）。'
       '最後兩列供程式以「劣化面積 ÷ 構件推估面積」核對 E 值。')

    table_sheet(wb, '意見U值', [('U值', 8), ('建議文字', 30), ('說明', 60)], [
        (2, '建議持續追蹤。', '不影響橋梁或行車安全、劣化範圍小（E≦2，單位為「處」多屬此類）'),
        (2, '建議三年內派工修復。', '橋面缺失影響行車安全，或劣化範圍大（E≧3）'),
        (3, '建議一年內派工修復。', ''),
        (4, '需緊急處置。', ''),
    ], '檢測員意見依 U 值產生的結尾建議（對應手冊表 4.1：U＝2 為 3 年內維護或持續追蹤）。U＝2 依「構件分類」與 E 值判斷。')

    table_sheet(wb, '意見排除清單', [('構件', 22), ('樣態關鍵字', 20), ('適用U值', 10), ('說明', 40)], [
        ('伸縮縫', '漏水', 2, '竹子坑橋最終版刪除'),
        ('橋墩基礎', '混凝土剝落', 2, '同上'),
        ('（全部）', '蜂窩', 2, '同上'),
        ('（全部）', 'AC剝離', 2, '同上'),
        ('橋墩基礎', '基礎裸露', None, '第 1 條已敘明基礎裸露時刪除重複條目（竹子坑橋第 12 條、龍門大橋第 16 條）'),
    ], '列在這裡的缺失不寫入檢測員意見。「適用U值」空白表示任何 U 值都排除。依竹子坑橋程式原稿與最終版比對整理，請確認或增修。',
        fill_rows={i: EXAMPLE_FILL for i in range(5)})

    table_sheet(wb, '意見構件名稱', [('構件', 26), ('意見中的寫法', 26), ('說明', 46)], [
        ('翼牆／擋土牆', '依備註寫翼牆或擋土牆', '竹子坑橋：翼牆／擋土牆 → 擋土牆'),
        ('支承／支承墊／阻尼裝置', '依備註寫支承、支承墊或阻尼裝置', '竹子坑橋：→ 支承墊'),
        ('橋墩／橋基保護設施', '（省略，樣態已含名稱）', '竹子坑橋：橋墩／橋基保護設施防衝鋼板 → 防衝鋼板'),
        ('其他', '（省略）', '竹子坑橋：其他AC剝離 → AC剝離'),
        ('河道', '（位置「整體-」省略）', '竹子坑橋：整體-下游護床工 → 下游護床工'),
    ], '檢測員意見中，斜線構件改寫成實際構件，樣態已含構件名稱時不重複。依竹子坑橋最終版整理。',
        fill_rows={i: EXAMPLE_FILL for i in range(5)})

    table_sheet(wb, '參數', [('參數', 26), ('最小', 8), ('最大', 8), ('單位', 8), ('說明', 56)], [
        ('大面積判定 E 值', 3, None, '', 'E≧3（劣化範圍 30% 以上）視為大面積；U＝2 時寫三年內派工修復'),
        ('位置含「-All」時的 E 值', 4, 4, '', '位置寫 -All 而 E 不是 4 時提示'),
        ('引道路堤預設長度', 5, 5, 'm', '引道路堤缺尺寸時：最小淨寬 × 此長度'),
        ('主梁寬度', 0.4, 0.7, 'm', '經驗值；推估橋面板分格與主梁面積'),
        ('懸臂版寬度（單側）', 0.7, 1.3, 'm', '經驗值；梁與梁之間的橋面板依剩餘寬度平均分配'),
        ('墩柱面積係數', 0.7, 0.8, '', '墩柱推估面積上限 ＝ 全寬 × 淨高 × 係數（高估版本）'),
        ('裂縫寬度轉 cm 門檻', 10, None, 'mm', '大於此值時備註改寫為 cm'),
        ('基礎裸露量測誤差', 0.1, None, 'm', 'v1.00 已採用'),
    ], '程式使用的數值參數。經驗數據由使用者提供，用於檢查備註尺寸是否與橋梁基本資料相符。',
        fill_rows={i: EXAMPLE_FILL for i in range(3, 6)})


def add_guide(wb, rule_sheets, status_col, stats):
    ws = wb.create_sheet('說明', 0)
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 60
    lines = [
        ('劣化說明與維修單價 v2 建議版', None),
        ('這份檔案由 1111v1 自動轉換。原本 7 欄內容完全沒有改動，舊版助手 v1.00 仍可讀取。', None),
        ('', None),
        ('一、改了什麼', 'h'),
        ('1. 各構件工作表右側新增 17 欄（藍色表頭），把原本寫在「備註」裡的條件拆成固定欄位，程式不必再猜文意。', None),
        ('2. 檔尾新增 10 張對照工作表：樣態同義詞、互斥樣態、劣化類型等價、位置代碼對照、構件分類、DERU評等、意見U值、意見排除清單、意見構件名稱、參數；另有「修訂紀錄」列出對原規則的修改。', None),
        ('3. 原本寫死在程式中的規則（例：木本／草本不可互配）搬到 Excel，之後自行增修即可，不用改程式。', None),
        ('', None),
        ('二、顏色', 'h'),
        ('橘色表頭：原有欄位，請勿修改欄名。', None),
        ('藍色表頭：新增欄位。', None),
        ('黃底：需要你確認或填寫。', CHECK_FILL),
        ('綠底：依需求說明、舊版修正紀錄或竹子坑橋資料整理的草稿，請確認。', EXAMPLE_FILL),
        ('', None),
        ('三、建議的確認順序', 'h'),
        ('1. 各構件工作表：用「轉換狀態」篩選「需確認」，看「轉換說明」處理，完成後改成「已確認」。', None),
        ('2. 構件分類：確認綠底草稿（主構件、可能影響行車安全），確認後依據改為「已確認」。', None),
        ('3. 參數：確認經驗數據（主梁寬度、懸臂版寬度、墩柱面積係數）。', None),
        ('4. 劣化類型等價：補上實務上視為同一筆的劣化類型。', None),
        ('5. 其他綠底工作表：確認草稿內容。', None),
        ('', None),
        ('四、寫法規則', 'h'),
        ('多個詞以全形頓號「、」分隔。', None),
        ('數值條件：最小值含等於（≧），最大值不含等於（＜）。例：裂縫寬度最小值 0.3 表示 ≧0.3mm。', None),
        ('備註範本：{位置}、{樣態} 會換成實際文字，尺寸數值由檢測員補上。', None),
        ('', None),
        ('五、轉換統計', 'h'),
    ]
    for r, (text, kind) in enumerate(lines, 1):
        cell = ws.cell(r, 1, text)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        size, bold = (16, True) if r == 1 else (13, True) if kind == 'h' else (11, False)
        cell.font = Font(name=FONT, size=size, bold=bold)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        if isinstance(kind, PatternFill):
            cell.fill = kind
        # 合併後約 52 個中文字寬；長句依行數加高，避免換行後被截斷
        wraps = -(-len(text) // 52) if text else 1
        ws.row_dimensions[r].height = max(20 if size > 11 else 16, 16 * wraps + 2)

    start = len(lines) + 1
    for c, name in enumerate(('工作表', '規則筆數', '需確認', '說明'), 1):
        style_header(ws.cell(start, c, name), NEW_HEADER_FILL, size=12)
    letter = get_column_letter(status_col)
    for i, title in enumerate(rule_sheets, start + 1):
        ref = f"'{title}'!{letter}:{letter}"
        style_body(ws.cell(i, 1, title))
        style_body(ws.cell(i, 2, f'=COUNTA({ref})-1'))
        style_body(ws.cell(i, 3, f'=COUNTIF({ref},"需確認")'))
        style_body(ws.cell(i, 4, ''))
    total = start + len(rule_sheets) + 1
    style_body(ws.cell(total, 1, '合計'))
    style_body(ws.cell(total, 2, f'=SUM(B{start + 1}:B{total - 1})'))
    style_body(ws.cell(total, 3, f'=SUM(C{start + 1}:C{total - 1})'))
    style_body(ws.cell(total, 4, '「需確認」會隨你把轉換狀態改成「已確認」而減少'))
    for c in range(1, 5):
        ws.cell(total, c).font = Font(name=FONT, size=11, bold=True)


# 使用者確認後的規則修訂，在轉換前套用到原 7 欄。
# 備註改寫成 v1.00 也能解析的寫法（數值帶單位、用「大於等於」），新舊版共用同一份檔案。
EDITS = (
    ('2026-11-14', '1-05', {'劣化樣態': '路面混凝土層破洞', '備註': '有關鍵字[破洞]或[坑洞]即可使用'},
     '原「路面坑洞（傷及RC層）」：改名，關鍵字為破洞或坑洞；舊寫法「路面坑洞」自動改寫成新樣態'),
    ('2026-11-14', '14-36', {'劣化樣態': '路面混凝土層破洞', '備註': '有關鍵字[破洞]或[坑洞]即可使用'},
     '原「路面坑洞（傷及混凝土層）」：改名，關鍵字為破洞或坑洞；舊寫法「路面坑洞」自動改寫成新樣態'),
    ('2026-11-14', '2-40', {'備註': '優先選擇，或讀取到護欄高度大於等於0.85m小於1.10m時使用'},
     '單管加高：≧0.85m 且 ＜1.10m；讀不到高度時預設此筆'),
    ('2026-11-14', '2-41', {'備註': '讀取到護欄高度大於等於0.60m小於0.85m時使用'},
     '雙管加高：≧0.60m 且 ＜0.85m'),
    ('2026-11-14', '2-42', {'備註': '讀取到護欄高度大於等於0.30m小於0.60m時使用'},
     '三管加高：≧0.30m 且 ＜0.60m'),
)
# 新增規則：(日期, 工作表, 原 7 欄, 修訂說明, 需確認原因)
ADDITIONS = (
    ('2026-11-14', '引道路堤護欄或橋護欄',
     ('其他損傷', '無設置護欄', '紐澤西護欄施作＋單管護欄加高', '公尺', 9900,
      '有關鍵字[無設置護欄]或[未設置護欄]或[無護欄]即可使用', '2-54'),
     '新增：無設置護欄。單價＝紐澤西護欄重作 4900（2-26）＋單管加高 5000（2-40）',
     None),
    ('2026-11-14', '引道路堤護欄或橋護欄',
     ('其他損傷', '護欄高度不足', '紐澤西護欄施作＋單管護欄加高', '公尺', 9900,
      '讀取到護欄高度小於0.30m時使用', '2-55'),
     '新增：護欄高度 ＜0.30m。工法與單價同 2-54',
     None),
)


def apply_edits(wb):
    """套用 EDITS 與 ADDITIONS，回傳 (修訂紀錄, 需確認的劣化ID→原因)。"""
    log, pending = [], {}
    by_id = {}
    for ws in wb.worksheets:
        for row in range(2, ws.max_row + 1):
            rule_id = ws.cell(row, 7).value
            if rule_id:
                by_id[str(rule_id).strip()] = (ws, row)
    for date, rule_id, changes, reason in EDITS:
        if rule_id not in by_id:
            raise ValueError(f'找不到劣化ID {rule_id}，停止轉換')
        ws, row = by_id[rule_id]
        for header, value in changes.items():
            col = ORIGINAL_HEADERS.index(header) + 1
            old = ws.cell(row, col).value
            ws.cell(row, col).value = value
            log.append((date, ws.title, rule_id, header, old, value, reason))
    for date, title, values, reason, review in ADDITIONS:
        if values[6] in by_id:
            raise ValueError(f'劣化ID {values[6]} 已存在，停止轉換')
        ws = wb[title]
        row = ws.max_row + 1
        template_row = row - 1
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            source = ws.cell(template_row, col)
            cell.font, cell.border, cell.alignment = copy(source.font), copy(source.border), copy(source.alignment)
        if review:
            ws.cell(row, 5).fill = CHECK_FILL
        log.append((date, title, values[6], '（新增整列）', None, '、'.join(str(v) for v in values if v), reason))
        if review:
            pending[values[6]] = review
    return log, pending


def add_edit_log(wb, log):
    rows = [(date, sheet, rule_id, field, '' if old is None else old, new, reason)
            for date, sheet, rule_id, field, old, new, reason in log]
    table_sheet(wb, '修訂紀錄',
                [('日期', 12), ('工作表', 22), ('劣化ID', 9), ('欄位', 12), ('修改前', 40), ('修改後', 44), ('說明', 40)],
                rows, '使用者確認後對原規則的修改。轉換程式 tools/excel_v2.py 的 EDITS、ADDITIONS 記錄同一份內容，可重複產生。')


def convert(source, target):
    wb = load_workbook(source)
    rule_sheets = [ws.title for ws in wb.worksheets]
    log, pending = apply_edits(wb)
    stats = {'rows': 0, 'review': 0}
    status_col = None
    for ws in wb.worksheets:
        status_col = convert_rule_sheet(ws, stats)
    for ws in wb.worksheets:
        for row in range(2, ws.max_row + 1):
            rule_id = str(ws.cell(row, 7).value or '').strip()
            if rule_id in pending:
                state, note = ws.cell(row, status_col), ws.cell(row, status_col + 1)
                if state.value != '需確認':
                    stats['review'] += 1
                state.value = '需確認'
                note.value = '；'.join(filter(None, (pending[rule_id], note.value)))
                state.fill = note.fill = CHECK_FILL
    add_reference_sheets(wb)
    add_edit_log(wb, log)
    add_guide(wb, rule_sheets, status_col, stats)
    for ws in wb.worksheets:
        if ws.title in rule_sheets:
            continue  # 規則表欄位多、以螢幕編輯為主，維持原檔列印設定
        # 說明與對照表列印時縮放成一頁寬，避免欄位被切到下一頁
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 0
    wb.save(target)
    return stats


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    result = convert(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"轉換 {result['rows']} 筆規則，需確認 {result['review']} 筆")
