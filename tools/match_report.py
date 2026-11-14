"""把前期、本期唯讀掃描結果配對，輸出 Excel 供人工核對。

用法：
    python tools/match_report.py 前期.json 本期.json 輸出.xlsx

每一列是一組判定（沿用、範圍加大、範圍縮小、單位數量變更、新增、已維修、疑似漏建），
最右邊「人工判斷」欄可填「正確」或「錯誤」，回傳後用來校正配對規則。
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bridgefill.core.matching import match_scans  # noqa: E402

FONT = '標楷體'
HEADERS = (('構件', 16), ('位置／編號', 20), ('判定', 12), ('信心', 7), ('分數', 6),
           ('前期 DERU', 8), ('前期劣化類型', 18), ('前期備註', 34), ('前期數量', 8), ('前期單位', 8),
           ('本期 DERU', 8), ('本期劣化類型', 18), ('本期備註', 34), ('本期數量', 8), ('本期單位', 8), ('本期 ID', 10),
           ('理由', 40), ('人工判斷', 10), ('說明', 24))
FILLS = {'需確認': PatternFill('solid', fgColor='FFF2CC'), '新增': PatternFill('solid', fgColor='FCE4D6'),
         '疑似漏建': PatternFill('solid', fgColor='F8CBAD')}
THIN = Side(style='thin', color='999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def side(defect):
    if defect is None:
        return ('', '', '', '', '')
    f = defect.fields
    return (defect.deru, defect.deterioration, f.get('remark', '').strip(), f.get('amount', ''), f.get('unit', ''))


def write_report(prior_scan, current_scan, target):
    result = match_scans(prior_scan, current_scan)
    order = {(r['keys'].get('component_form'), r['keys'].get('component_id')): i
             for i, r in enumerate(current_scan.get('rows', []))}
    pairs = sorted(result.pairs, key=lambda p: (order.get(p.row, 9999), p.prior is None, p.current is None))

    wb = Workbook()
    ws = wb.active
    ws.title = '配對結果'
    head = current_scan.get('header', {})
    ws['A1'] = (f"{head.get('橋梁名稱', '')}　前期 {prior_scan.get('header', {}).get('檢測日期', '')} → "
                f"本期 {head.get('檢測日期', '')}　｜　" + '、'.join(f'{k} {v}' for k, v in result.summary.items()))
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
        values = ((any_side.component, any_side.label, pair.kind, pair.confidence, pair.score or '') +
                  side(pair.prior) + side(pair.current) +
                  ((pair.current.method_id if pair.current else ''), '；'.join(pair.reasons), '', ''))
        fill = FILLS.get(pair.kind) or (FILLS['需確認'] if pair.confidence == '需確認' else None)
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, value)
            cell.font = Font(name=FONT, size=11)
            cell.border = BORDER
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            if fill and col <= 5:
                cell.fill = fill
        verdict.add(ws.cell(r, len(HEADERS) - 1))
    ws.freeze_panes = 'D4'
    ws.auto_filter.ref = f'A3:{ws.cell(3, len(HEADERS)).column_letter}{ws.max_row}'
    ws.page_setup.orientation = 'landscape'
    wb.save(target)
    return result


if __name__ == '__main__':
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    read = lambda path: json.loads(Path(path).read_text(encoding='utf-8'))
    summary = write_report(read(sys.argv[1]), read(sys.argv[2]), sys.argv[3]).summary
    print('、'.join(f'{k} {v}' for k, v in summary.items()))
