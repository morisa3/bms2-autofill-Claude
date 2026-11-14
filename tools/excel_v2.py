"""把 v1 劣化維修規則 Excel 轉為 v2 建議格式。

原本 7 欄（劣化類型、劣化樣態、維修工法、維修單位、單價、備註、劣化ID）一字不改，
只在右側新增結構化欄位，並在檔尾新增對照工作表。v1.00 依表頭名稱讀取且略過
缺少必要表頭的工作表，所以轉換後的檔案仍可給 v1.00 使用。

用法：
    python tools/excel_v2.py 來源.xlsx 輸出.xlsx
"""
import re
import sys
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

    guard = take(r'(手動選擇|優先選擇)?[，,]?或?讀取到護欄高度大於([\d.]+)小於([\d.]+)時使用')
    if guard:
        out.update({'處理方式': '手動或條件', '條件欄位': '護欄高度', '最小值': float(guard.group(2)),
                    '最大值': float(guard.group(3)), '條件單位': 'm'})
        review = True
        remarks.append(f'原文「大於{guard.group(2)}小於{guard.group(3)}」兩端都不含等於；'
                       f'新格式最小值含等於，請確認 {guard.group(2)}m 剛好時歸哪一筆')

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
        ('生銹', '生鏽、生锈', 'v0.9.7.6 修正：三種寫法視為相同'),
        ('防衝鋼板表面生銹', '防撞鋼板生銹、防衝鋼板生鏽', '需求說明 2.1 範例；橋墩／橋基保護設施只寫「生鏽」的情況由 Excel 11-04 處理'),
    ], '備註寫法不同但意思相同時，程式先轉成「標準樣態」再比對。左側寫一個標準寫法，右側列出所有同義寫法，以「、」分隔。'
       '綠底為依需求說明與舊版修正紀錄整理的草稿，請確認或增修。',
        fill_rows={i: EXAMPLE_FILL for i in range(4)})

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

    table_sheet(wb, '構件分類', [('構件', 26), ('位置', 10), ('主構件', 8), ('可能影響行車安全', 12), ('說明', 34)], [
        (name, where, None, None, '') for name, where in (
            ('引道路堤', '橋面'), ('引道路堤護欄', '橋面'), ('引道道路堤保護措施', '下部'), ('河道', '下部'),
            ('橋台基礎', '下部'), ('橋台', '下部'), ('翼牆／擋土牆', '下部'), ('護床工', '下部'),
            ('伸縮縫', '橋面'), ('橋梁排水設施', '橋面'), ('橋護欄', '橋面'), ('支承／支承墊／阻尼裝置', '上部'),
            ('防落設施', '上部'), ('橋墩／橋基保護設施', '下部'), ('橋墩基礎', '下部'), ('墩柱與帽梁', '下部'),
            ('主梁', '上部'), ('橫隔梁', '上部'), ('橋面板', '上部'), ('其他', '橋面'),
            ('鋼纜錨碇裝置(橋塔端)', '上部'), ('鋼纜錨碇裝置(主梁端)', '上部'),
            ('鋼纜(吊索)-斜張鋼纜', '上部'), ('鋼纜保護套管-斜張鋼纜', '上部'))
    ], '檢測員意見 U＝2 判斷用：橋面缺失看是否影響行車安全；下部結構先看是否為主構件，再看劣化面積。'
       '「位置」為草稿（綠底），「主構件」「可能影響行車安全」請填「是」或「否」（黃底）。')
    ws = wb['構件分類']
    answers = DataValidation(type='list', formula1='"是,否"', allow_blank=True)
    ws.add_data_validation(answers)
    for row in range(3, ws.max_row + 1):
        ws.cell(row, 2).fill = EXAMPLE_FILL
        for col in (3, 4):
            ws.cell(row, col).fill = CHECK_FILL
            answers.add(ws.cell(row, col))

    table_sheet(wb, '意見U值', [('U值', 8), ('建議文字', 30), ('說明', 60)], [
        (2, '建議持續追蹤。', '小面積、單位為「處」、不影響安全時使用'),
        (2, '建議三年內派工修復。', '橋面缺失影響行車安全，或主構件／大面積劣化'),
        (3, '建議一年內派工修復。', ''),
        (4, '需緊急處置。', ''),
    ], '檢測員意見依 U 值產生的結尾建議。U＝2 有兩種寫法，依「構件分類」與「參數」的大面積門檻判斷。')

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

    table_sheet(wb, '參數', [('參數', 26), ('數值', 10), ('單位', 8), ('說明', 50)], [
        ('大面積門檻', None, '平方公尺', '劣化面積達此值以上算大面積（U＝2 寫三年內）'),
        ('引道路堤預設長度', 5, 'm', '引道路堤缺尺寸時：最小淨寬 × 此長度'),
        ('裂縫寬度轉 cm 門檻', 10, 'mm', '大於此值時備註改寫為 cm'),
        ('基礎裸露量測誤差', 0.1, 'm', 'v1.00 已採用'),
    ], '程式使用的數值參數。黃底為尚未決定、請填寫的值。', fill_rows={0: CHECK_FILL})


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
        ('2. 檔尾新增 9 張對照工作表：樣態同義詞、互斥樣態、劣化類型等價、位置代碼對照、構件分類、意見U值、意見排除清單、意見構件名稱、參數。', None),
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
        ('2. 構件分類：填「主構件」「可能影響行車安全」。', None),
        ('3. 參數：填「大面積門檻」。', None),
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


def convert(source, target):
    wb = load_workbook(source)
    rule_sheets = [ws.title for ws in wb.worksheets]
    stats = {'rows': 0, 'review': 0}
    status_col = None
    for ws in wb.worksheets:
        status_col = convert_rule_sheet(ws, stats)
    add_reference_sheets(wb)
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
