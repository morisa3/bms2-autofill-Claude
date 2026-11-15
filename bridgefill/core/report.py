"""比對資料輸出（第 5 步③）：前後期配對與審查結果 → Excel。

每列一組配對：前期與本期的 DERU、劣化類型、備註、數量、單位，程式判定與信心，
以及審查狀態、最終判定、審查說明。只寫本機檔案，不修改網站。
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .review import CHANGED, PENDING, Review

FONT = '標楷體'
HEADERS = (('構件', 14), ('位置／編號', 16), ('程式判定', 10), ('信心', 6), ('分數', 6),
           ('前期 DERU', 8), ('前期劣化類型', 16), ('前期備註', 30), ('前期數量', 7), ('前期單位', 7),
           ('本期 DERU', 8), ('本期劣化類型', 16), ('本期備註', 30), ('本期數量', 7), ('本期單位', 7), ('本期 ID', 9),
           ('理由', 32), ('審查', 8), ('最終判定', 10), ('審查說明', 20), ('人工判斷', 9))
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


def write_report(pairs, target, review=None, title=''):
    """pairs：已排序的配對；review：審查紀錄（None 時全部視為程式判定）。回傳寫出的路徑。"""
    review = review or Review()
    wb = Workbook()
    ws = wb.active
    ws.title = '前後期比對'
    counts = review.counts(pairs)
    ws['A1'] = (title or f'{review.bridge}　前期 {review.prior_date} → 本期 {review.current_date}') + \
        f'　｜　共 {len(pairs)} 組：' + '、'.join(f'{k} {v}' for k, v in counts.items())
    ws['A1'].font = Font(name=FONT, size=13, bold=True)
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
        decision = review.decision(pair)
        values = ((any_side.component, any_side.label, pair.kind, pair.confidence, pair.score or '') +
                  side(pair.prior) + side(pair.current) +
                  ((pair.current.method_id if pair.current else ''), '；'.join(pair.reasons),
                   decision.status, decision.kind, decision.note, ''))
        fill = FILLS.get(decision.status)
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = CELL
            if fill and col in (1, 2, 18, 19):
                cell.fill = fill
        verdict.add(ws.cell(r, len(HEADERS)))
    ws.freeze_panes = 'C4'
    ws.auto_filter.ref = f'A3:{ws.cell(3, len(HEADERS)).column_letter}{max(ws.max_row, 3)}'
    ws.page_setup.orientation = 'landscape'
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)
    return target
