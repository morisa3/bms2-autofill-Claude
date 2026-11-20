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
HEADERS = (('構件', 14), ('位置／編號', 16), ('程式判定', 10), ('信心', 6), ('分數', 6),
           ('前期 DERU', 8), ('前期劣化類型', 16), ('前期備註', 30), ('前期數量', 7), ('前期單位', 7),
           ('本期 DERU', 8), ('本期劣化類型', 16), ('本期備註', 30), ('本期數量', 7), ('本期單位', 7), ('本期 ID', 9),
           ('配對方式', 8), ('理由', 32), ('審查', 8), ('最終判定', 10), ('審查說明', 20), ('修改內容', 30), ('人工判斷', 9))
FILLS = {PENDING: PatternFill('solid', fgColor='FFF2CC'), CHANGED: PatternFill('solid', fgColor='DDEBF7')}
THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CELL = Alignment(vertical='center', horizontal='left', wrap_text=True)


def side(defect):
    if defect is None:
        return ('', '', '', '', '')
    f = defect.fields
    return (' '.join(defect.deru), defect.deterioration, str(f.get('remark', '') or '').strip(),
            f.get('amount', ''), f.get('unit', ''))


def report_name(bridge, current_date):
    safe = ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in f'{current_date}-{bridge}')
    return f'{safe}-前後期比對.xlsx'


def write_report(pairs, target, review=None, title='', sheets=(), heights=None):
    """pairs：已排序的配對；review：審查紀錄（None 時全部視為程式判定）。
    sheets：另外輸出的工作表（U值增減、新增劣化、未建（漏建）劣化）。
    heights：淨高資料 [(項目, 本期, 前期, 說明)]，輸出「淨高」工作表並在第一張標題下摘要。回傳寫出的路徑。"""
    review = review or Review()
    wb = Workbook()
    ws = wb.active
    ws.title = '前後期比對'
    counts = review.counts(pairs)
    ws['A1'] = (title or f'{review.bridge}　前期 {review.prior_date} → 本期 {review.current_date}') + \
        f'　｜　共 {len(pairs)} 組：' + '、'.join(f'{k} {v}' for k, v in counts.items())
    ws['A1'].font = Font(name=FONT, size=13, bold=True)
    if heights:
        summary = next((row for row in heights if row[1] or row[2]), None)
        ws['A2'] = f'淨高（{summary[0]}）：本期 {summary[1] or "—"}｜前期 {summary[2] or "—"}（詳見「淨高」工作表）' \
            if summary else '淨高：本期、前期意見都沒有淨高（詳見「淨高」工作表）'
        ws['A2'].font = Font(name=FONT, size=11, bold=True, color='1F4E79')
    for col, (name, width) in enumerate(HEADERS, 1):
        cell = ws.cell(3, col, name)
        cell.font = Font(name=FONT, size=11, bold=True)
        cell.fill = PatternFill('solid', fgColor='9DC3E6')
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = width
    verdict = DataValidation(type='list', formula1='"正確,錯誤"', allow_blank=True)
    ws.add_data_validation(verdict)
    for r, pair in enumerate(pairs, 4):
        any_side = pair.current or pair.prior
        component = pair.display
        decision = review.decision(pair)
        values = ((component, any_side.label, pair.kind, pair.confidence, pair.score or '') +
                  side(pair.prior) + side(pair.current) +
                  ((pair.current.method_id if pair.current else ''),
                   ('人工' if pair.manual else '程式') + ('（位置變更）' if pair.moved else ''), '；'.join(pair.reasons),
                   decision.status, decision.kind, decision.note, describe(review.edit_of(pair)), ''))
        fill = FILLS.get(decision.status)
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = CELL
            if fill and col in (1, 2, 19, 20):
                cell.fill = fill
        verdict.add(ws.cell(r, len(HEADERS)))
    ws.freeze_panes = 'C4'
    ws.auto_filter.ref = f'A3:{ws.cell(3, len(HEADERS)).column_letter}{max(ws.max_row, 3)}'
    ws.page_setup.orientation = 'landscape'
    for name in SHEETS:
        if name in sheets:
            SHEETS[name](wb.create_sheet(name), pairs, review)
    if heights:
        height_sheet(wb.create_sheet('淨高'), heights)
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


def u_sheet(ws, pairs, review):
    rows = []
    for pair in pairs:
        if not (pair.prior and pair.current):
            continue
        current_deru = ''.join(review.value(pair, k) or '_' for k in 'deru')
        before, after = u_value(pair.prior.deru), u_value(current_deru)
        if before is None or after is None or before == after:
            continue
        rows.append((pair.current.display, pair.current.label, pair.prior.deru, current_deru,
                     f'{before} → {after}', '增加' if after > before else '減少',
                     pair.prior.remark.strip(), review.value(pair, 'remark'), review.decision(pair).kind))
    simple_sheet(ws, (('構件', 14), ('位置／編號', 16), ('前期 DERU', 9), ('本期 DERU', 9), ('U 值', 8), ('增減', 6),
                      ('前期備註', 30), ('本期備註', 30), ('最終判定', 10)),
                 rows, f'U 值增減：{len(rows)} 筆（本期以修改後的 DERU 為準）')


def side_sheet(kinds, label):
    def write(ws, pairs, review):
        rows = []
        for pair in pairs:
            decision = review.decision(pair)
            if decision.kind not in kinds:
                continue
            defect = pair.current or pair.prior
            deru = ''.join(review.value(pair, k) or '_' for k in 'deru') if pair.current else defect.deru
            rows.append((defect.display, defect.label, deru, defect.deterioration,
                         review.value(pair, 'remark') if pair.current else defect.remark.strip(),
                         defect.fields.get('amount', ''), defect.fields.get('unit', ''),
                         decision.kind, decision.status, decision.note))
        simple_sheet(ws, (('構件', 14), ('位置／編號', 16), ('DERU', 8), ('劣化類型', 16), ('備註', 36), ('數量', 7),
                          ('單位', 7), ('最終判定', 10), ('審查', 8), ('審查說明', 24)),
                     rows, f'{label}：{len(rows)} 筆')
    return write


SHEETS = {'U值增減': u_sheet,
          '新增劣化': side_sheet({KIND_NEW}, '新增劣化（本期有、前期沒有）'),
          '未建（漏建）劣化': side_sheet({KIND_CONFIRMED_MISSING, KIND_MISSING}, '未建（漏建）劣化（前期有、本期未建立）')}


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
