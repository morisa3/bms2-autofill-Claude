"""列出實際案例中 Excel 規則表沒有建立的缺失（最終送出的寫法套回 Excel 判斷，判不出規則者）。

用法：python tools/missing_rules.py [輸出.xlsx]
資料：fixtures/private/real-2026-11-18/（修改完成__*-填寫.json、未修改__*-檢測.json、未修改__*-前期.json）
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bridgefill.core.data import component_repairs  # noqa: E402
from bridgefill.core.editing import load_rules, rules_file, split_remark  # noqa: E402
from bridgefill.core.policy import automatic_plan  # noqa: E402
from bridgefill.core.reconcile import deterioration_key  # noqa: E402

DATA = Path(__file__).resolve().parents[1] / 'fixtures' / 'private' / 'real-2026-11-18'
FIELDS = ('deterioration', 'd_location', 'method', 'amount', 'unit', 'price', 'remark')


def missing(folder=DATA):
    rules = load_rules(rules_file())
    found = {}
    for final_path in sorted(folder.glob('修改完成__*-填寫.json')):
        bridge = final_path.name.split('__')[1][len('2026-11-18-'):-len('-填寫.json')]
        current = json.loads(next(folder.glob(f'未修改__*-{bridge}-檢測.json')).read_text(encoding='utf-8'))
        rows = {}
        for row in current['rows']:
            for defect in row.get('defects', []):
                rows[f"本期:{defect.get('method_id')}"] = (row['component'], {**(row.get('columns') or {})})
        for key, entry in json.loads(final_path.read_text(encoding='utf-8'))['entries'].items():
            if key not in rows:
                continue
            component, context = rows[key]
            target = {k: str(entry['target'].get(k, '') or '') for k in FIELDS}
            decision = automatic_plan(dict(target), component, [], rules, dict(context))
            if 'Excel 無對應值' not in decision.reason:
                continue
            _, pattern, _ = split_remark(target['remark'])
            pattern = pattern.rstrip('。. ')
            same_type = [r for r in component_repairs(rules, component)
                         if deterioration_key(r.get('deterioration')) == deterioration_key(target['deterioration'])]
            item = found.setdefault((component, target['deterioration'], pattern), dict(
                component=component, deterioration=target['deterioration'], pattern=pattern, examples=[],
                method=target['method'], unit=target['unit'], price=target['price'],
                existing='、'.join(dict.fromkeys(r['pattern'] for r in same_type)) or '（此類型沒有任何規則）'))
            item['examples'].append(f"{bridge}：{target['remark']}")
    return list(found.values())


def write(items, path):
    book = Workbook()
    sheet = book.active
    sheet.title = '未建立的缺失'
    heads = ('構件', '劣化類型', '劣化樣態（最終寫法）', '維修工法（你送出的）', '單位', '單價', '實例備註',
             '同構件同類型已有的樣態', '確認後的樣態寫法', '確認後的工法', '確認後的單位', '確認後的單價', '說明')
    sheet.append(heads)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='DDEBF7')
    for item in items:
        sheet.append((item['component'], item['deterioration'], item['pattern'], item['method'], item['unit'],
                      item['price'], '\n'.join(item['examples']), item['existing'], '', '', '', '', ''))
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')
        for cell in row[8:12]:
            cell.fill = PatternFill('solid', fgColor='FFF2CC')
    for column, width in zip('ABCDEFGHIJKLM', (14, 18, 18, 30, 8, 8, 34, 34, 18, 26, 10, 10, 20)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = 'A2'
    book.save(path)


if __name__ == '__main__':
    result = missing()
    for item in result:
        print(item['component'], '|', item['deterioration'], '|', item['pattern'], '|', item['method'], item['unit'],
              item['price'], '|', '；'.join(item['examples']))
    if len(sys.argv) > 1:
        write(result, sys.argv[1])
        print('已輸出', sys.argv[1], len(result), '項')
