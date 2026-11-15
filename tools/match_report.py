"""把前期、本期唯讀掃描結果配對，輸出 Excel 供人工核對。

用法：
    python tools/match_report.py 前期.json 本期.json 輸出.xlsx

每一列是一組判定（沿用、範圍加大、範圍縮小、單位數量變更、新增、已維修、疑似漏建），
最右邊「人工判斷」欄可填「正確」或「錯誤」，回傳後用來校正配對規則。
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bridgefill.core import report  # noqa: E402
from bridgefill.core.matching import match_scans  # noqa: E402


def write_report(prior_scan, current_scan, target):
    result = match_scans(prior_scan, current_scan)
    order = {(r['keys'].get('component_form'), r['keys'].get('component_id')): i
             for i, r in enumerate(current_scan.get('rows', []))}
    pairs = sorted(result.pairs, key=lambda p: (order.get(p.row, 9999), p.prior is None, p.current is None))
    head = current_scan.get('header', {})
    title = (f"{head.get('橋梁名稱', '')}　前期 {prior_scan.get('header', {}).get('檢測日期', '')} → "
             f"本期 {head.get('檢測日期', '')}　｜　" + '、'.join(f'{k} {v}' for k, v in result.summary.items()))
    report.write_report(pairs, target, title=title)
    return result


if __name__ == '__main__':
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    read = lambda path: json.loads(Path(path).read_text(encoding='utf-8'))
    summary = write_report(read(sys.argv[1]), read(sys.argv[2]), sys.argv[3]).summary
    print('、'.join(f'{k} {v}' for k, v in summary.items()))
