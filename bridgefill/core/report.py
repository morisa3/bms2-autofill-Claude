"""比對資料輸出（第 5 步③）：前後期配對與審查結果 → Excel。

每列一組配對：前期與本期的 DERU、劣化類型、備註、數量、單位，程式判定與信心，
以及審查狀態、最終判定、審查說明。只寫本機檔案，不修改網站。
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .editing import describe
from .matching import KIND_MISSING, KIND_NEW
from .review import CHANGED, KIND_CONFIRMED_MISSING, PENDING, Review

FONT = '標楷體'
# 2026-11-26 使用者範本：本期為修改完成的資料（Excel 規則＋人工修改後將送出的值），編號同核對頁簡化（A1、P1、S3…）
SIDE = ('DERU', '劣化類型', '備註', '數量', '單位')
HEADERS = (('構件', 11.4), ('編號', 13.6), ('前期 DERU', 9), ('前期劣化類型', 13.8), ('前期備註', 23.3), ('前期數量', 10.5),
           ('前期單位', 9), ('前期有無\n維修紀錄', 11.6), ('本期 DERU', 9), ('本期劣化類型', 13.1), ('本期備註', 18.6),
           ('本期數量', 10.9), ('本期單位', 11.6), ('判定', 12), ('變化', 12))
# 程式讀取用（總表同步、他人比對資料）：照樣輸出，欄寬縮到看不到
HIDDEN = (('本期 ID', 0.5), ('前期 ID', 0.5), ('程式判定', 0.5), ('最終判定', 0.5), ('審查', 0.5), ('審查說明', 0.5), ('信心', 0.5))
COLORS = {'本期新增': 'FFFFFF00', '漏建劣化': 'FFFFC000', 'U值上升': 'FFFF5050', 'U值下降': 'FFBDD7EE',
          '已維修': 'FF92D050', '本期無發現該劣化': 'FF92D050'}
FILLS = {PENDING: PatternFill('solid', fgColor='FFF2CC'), CHANGED: PatternFill('solid', fgColor='DDEBF7')}
THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CELL = Alignment(vertical='center', horizontal='left', wrap_text=True)
MIDDLE = Alignment(vertical='center', horizontal='center', wrap_text=True)
CENTERED = ('編號', '前期有無\n維修紀錄', '本期數量', '變化')


def side(defect):
    """前期（網站原始值）：(DERU, 劣化類型, 備註, 數量, 單位)。"""
    if defect is None:
        return ('', '', '', '', '')
    f = defect.fields
    return (' '.join(defect.deru), defect.deterioration, str(f.get('remark', '') or '').strip(),
            f.get('amount', ''), f.get('unit', ''))


def final_values(pair, review, planned=None):
    """本期修改完成的值（將送出／已送出）：planned(pair) 有值用它（填寫清單預定值），否則網站值＋人工修改。"""
    if pair.current is None:
        return {}
    values = (planned(pair) if planned else None) or {}
    keys = ('d', 'e', 'r', 'u', 'deterioration', 'remark', 'amount', 'unit', 'method', 'price', 'd_location')
    return {k: str(values.get(k) if values.get(k) not in (None, '') else review.value(pair, k) or '').strip() for k in keys}


def current_side(pair, review, planned=None):
    if pair.current is None:
        return ('', '', '', '', '')
    v = final_values(pair, review, planned)
    return (' '.join(v[k] or '_' for k in 'deru'), v['deterioration'] or pair.current.deterioration, v['remark'],
            v['amount'], v['unit'])


def size_numbers(text):
    import re
    body = re.findall(r'[\(（]([^()（）]*)[\)）](?:\s*[*＊×]\s*(\d+(?:\.\d+)?))?', str(text or ''))
    numbers = sorted(n for inner, times in body for n in re.findall(r'\d+(?:\.\d+)?', inner) + ([times] if times else []))
    return numbers


def number(value):
    try:
        return float(str(value).strip())
    except ValueError:
        return None


def verdict(pair, review, planned=None):
    """(判定, 變化)：依範本——既有劣化沿用／U值上升／U值下降／本期新增／漏建劣化；
    變化：不變、數量變更、改類型（數量修正）、已維修、本期無發現該劣化。"""
    from .matching import KIND_REPAIRED
    from .review import KIND_GONE
    kind = review.decision(pair).kind
    if pair.prior is None:
        return '本期新增', ''
    if pair.current is None:
        if kind == KIND_GONE:                     # 審查判定優先於程式判定
            return '', '本期無發現該劣化'
        if kind == KIND_REPAIRED:
            return '', '已維修'
        return '漏建劣化', ''
    now = final_values(pair, review, planned)
    before, after = u_value(pair.prior.deru), u_value(''.join(now[k] or '_' for k in 'deru'))
    judged = '既有劣化沿用'
    if before is not None and after is not None and after != before:
        judged = 'U值上升' if after > before else 'U值下降'
    prior = pair.prior.fields
    quantity = (number(prior.get('amount')) != number(now['amount']) or str(prior.get('unit') or '').strip() != now['unit']
                or size_numbers(prior.get('remark')) != size_numbers(now['remark']))
    retyped = bool(now['deterioration']) and now['deterioration'] != pair.prior.deterioration
    if retyped:
        return judged, '改類型' + ('\n數量修正' if quantity else '')
    return judged, '數量變更' if quantity else '不變'


def report_name(bridge, current_date):
    safe = ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in f'{current_date}-{bridge}')
    return f'{safe}-前後期比對.xlsx'


def header_cell(cell, name):
    cell.value = name
    cell.font = Font(name=FONT, size=11, bold=True)
    cell.fill = PatternFill('solid', fgColor='9DC3E6')
    cell.border = BORDER
    cell.alignment = MIDDLE


def comparison_row(pair, review, planned=None):
    """一列前後期比對（範本 15 欄＋隱藏欄）。"""
    any_side = pair.current or pair.prior
    judged, change = verdict(pair, review, planned)
    decision = review.decision(pair)
    repaired = ('有' if getattr(pair.prior, 'repaired', False) else '無') if pair.prior else ''
    visible = (pair.display, pair.code or any_side.label, *side(pair.prior), repaired, *current_side(pair, review, planned),
               judged, change)
    hidden = ((pair.current.method_id if pair.current else ''), (getattr(pair.prior, 'der_id', '') if pair.prior else ''),
              pair.kind, decision.kind, decision.status, decision.note, pair.confidence)
    return visible + hidden, COLORS.get(judged) or COLORS.get(change)


def write_report(pairs, target, review=None, title='', sheets=(), heights=None, access=None, planned=None):
    """pairs：已排序的配對；review：審查紀錄；planned(pair)：本期修改完成的值（填寫清單預定值），沒有時用網站值＋人工修改。
    sheets：另外輸出的工作表（U值增減、新增劣化、未建（漏建）劣化）。heights：淨高；access：下橋路徑。回傳寫出的路徑。"""
    review = review or Review()
    wb = Workbook()
    ws = wb.active
    ws.title = '前後期比對'
    counts = review.counts(pairs)
    ws['A1'] = (title or f'{review.bridge}　前期 {review.prior_date} → 本期 {review.current_date}') + \
        f'　｜　共 {len(pairs)} 組：' + '、'.join(f'{k} {v}' for k, v in counts.items())
    ws['A1'].font = Font(name=FONT, size=12, bold=True)
    columns = HEADERS + HIDDEN
    for col, (name, width) in enumerate(columns, 1):
        header_cell(ws.cell(2, col), name)
        ws.column_dimensions[ws.cell(2, col).column_letter].width = width
    ws.row_dimensions[2].height = 31.5
    for r, pair in enumerate(pairs, 3):
        values, color = comparison_row(pair, review, planned)
        fill = PatternFill('solid', fgColor=color) if color else None
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = MIDDLE if columns[col - 1][0] in CENTERED else CELL
            if fill and col <= len(HEADERS):
                cell.fill = fill
    ws.freeze_panes = 'C3'
    ws.auto_filter.ref = f'A2:{ws.cell(2, len(HEADERS)).column_letter}{max(ws.max_row, 2)}'
    ws.page_setup.orientation = 'landscape'
    if heights:
        summary = next((row for row in heights if row[1] or row[2]), None)
        ws.cell(1, 8).value = f'淨高（{summary[0]}）：本期 {summary[1] or "—"}｜前期 {summary[2] or "—"}' if summary else ''
    for name in SHEETS:
        if name in sheets:
            SHEETS[name](wb.create_sheet(name), pairs, review, planned)
    if heights:
        height_sheet(wb.create_sheet('淨高'), heights)
    if access and any(access):
        access_sheet(wb.create_sheet('下橋路徑'), access)
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)
    return target


# ---------------------------------------------------------------- 另外輸出的工作表

def u_value(deru):
    value = (deru or '')[3:4]
    return int(value) if value.isdigit() else None


def simple_sheet(ws, columns, rows, note):
    ws['A1'] = note
    ws['A1'].font = Font(name=FONT, size=12, bold=True)
    for col, (name, width) in enumerate(columns, 1):
        cell = ws.cell(3, col, name)
        cell.font = Font(name=FONT, size=11, bold=True)
        cell.fill = PatternFill('solid', fgColor='9DC3E6')
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = width
    for r, values in enumerate(rows, 4):
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = CELL
    ws.freeze_panes = 'A4'
    if rows:
        ws.auto_filter.ref = f'A3:{ws.cell(3, len(columns)).column_letter}{ws.max_row}'


def u_sheet(ws, pairs, review, planned=None):
    rows = []
    for pair in pairs:
        if not (pair.prior and pair.current):
            continue
        now = current_side(pair, review, planned)
        before, after = u_value(pair.prior.deru), u_value(now[0].replace(' ', ''))
        if before is None or after is None or before == after:
            continue
        rows.append((pair.display, pair.code, now[1], ' '.join(pair.prior.deru), now[0], f'{before}->{after}',
                     '增' if after > before else '減', pair.prior.remark.strip(), now[2]))
    simple_sheet(ws, (('構件', 12.3), ('編號', 12.5), ('劣化類型', 19.1), ('前期 DERU', 7.9), ('本期 DERU', 9), ('U 值', 8.1),
                      ('增減', 8.3), ('前期備註', 26.3), ('本期備註', 30)),
                 rows, f'U 值增減：{len(rows)} 筆（本期以修改後的 DERU 為準）')


def new_sheet(ws, pairs, review, planned=None):
    rows = []
    for pair in pairs:
        if pair.prior is not None or pair.current is None:
            continue
        now = current_side(pair, review, planned)
        rows.append((pair.display, pair.code, now[0].replace(' ', ''), *now[1:]))
    simple_sheet(ws, (('構件', 12.3), ('編號', 9.4), ('DERU', 7), ('劣化類型', 14), ('備註', 31.5), ('數量', 7.4), ('單位', 12.3)),
                 rows, f'新增劣化（本期有、前期沒有）：{len(rows)} 筆')


def missing_sheet(ws, pairs, review, planned=None):
    rows = []
    for pair in pairs:
        if pair.current is None and pair.prior is not None and verdict(pair, review, planned)[0] == '漏建劣化':
            rows.append(comparison_row(pair, review, planned)[0][:13])
    simple_sheet(ws, tuple((name, width) for name, width in HEADERS[:13]),
                 rows, f'未建（漏建）劣化（前期有、本期未建立）：{len(rows)} 筆')


SHEETS = {'U值增減': u_sheet, '新增劣化': new_sheet, '未建（漏建）劣化': missing_sheet}


def access_sheet(ws, access):
    """下橋路徑與輔助設備（⑤ 設定，同步到橋梁總表的兩格）：access＝(下橋路徑文字, 輔助設備文字)。"""
    for col, (name, width) in enumerate((('項目', 18), ('內容', 60)), 1):
        cell = ws.cell(1, col, name)
        cell.font = Font(name=FONT, size=11, bold=True)
        cell.fill = PatternFill('solid', fgColor='9DC3E6')
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = width
    for r, (name, value) in enumerate((('下橋路徑(ABCD)', access[0]), ('輔助設備(A/B)', access[1])), 2):
        for col, text in enumerate((name, value or ''), 1):
            cell = ws.cell(r, col, text)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = CELL
        ws.row_dimensions[r].height = 16 * max(1, str(value or '').count('\n') + 1)


def height_sheet(ws, rows):
    """淨高：檢測員意見（本期、前期）、程式產生的意見草稿、基本資料。"""
    for col, (name, width) in enumerate((('項目', 26), ('本期', 46), ('前期', 46), ('說明', 46)), 1):
        cell = ws.cell(1, col, name)
        cell.font = Font(name=FONT, size=11, bold=True)
        cell.fill = PatternFill('solid', fgColor='9DC3E6')
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[cell.column_letter].width = width
    for r, row in enumerate(rows, 2):
        for col, value in enumerate(row, 1):
            cell = ws.cell(r, col, value or '—')
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = CELL
    ws.freeze_panes = 'A2'
