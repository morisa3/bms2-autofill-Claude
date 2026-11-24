"""v2 規則表（刪除備註）的判讀結果須與 v1.00 備註一致。"""
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'tools'))

from openpyxl import load_workbook  # noqa: E402
from excel_v2 import add_new_sheets, apply_edits, convert  # noqa: E402
from bridgefill.core.conditions import condition, keyword_groups, rule_deterioration  # noqa: E402
from bridgefill.core.data import read_repairs  # noqa: E402
from bridgefill.core.rules_v2 import read_rules, synthesize_note  # noqa: E402

SOURCE = ROOT / 'data' / '劣化說明與維修單價1111v1.xlsx'
# 1111v1 已刪除（規則表改為直接維護 1116）；轉換程式的測試在有 v1 檔時才執行
NEEDS_V1 = unittest.skipUnless(SOURCE.exists(), '1111v1 已刪除，轉換程式僅供紀錄')
CONTEXTS = ([{}] + [{'護欄高度(m)': h} for h in ('0.1', '0.3', '0.6', '0.85', '1.1')] +
            [{'U': u} for u in '1234'] + [{'伸縮縫型式': t} for t in ('齒型', '角鋼')] +
            [{'支承型式': '人造橡膠支承墊'}])
REMARKS = ['', 'A1 植生。(1m)', 'A1 草本植物。(1m)', 'A1 混凝土裂縫。(0.2mm*1m)', 'A1 混凝土裂縫。(0.3mm*1m)',
           'A1 混凝土裂縫。(0.5mm*1m)', 'A1 生鏽']
# 舊備註把說明文字中的 [單位]、[橋台]/[橋墩]、[AC剝離] 也當成關鍵字；v2 只保留真正的判讀關鍵字
CLEANED = {'1-06', '1-07', '1-08', '5-04', '14-39'}


def behaviour(rule):
    return (keyword_groups(rule['note']), [condition(rule, c, r) for c in CONTEXTS for r in REMARKS],
            [rule_deterioration(rule, x) for x in ('保護設施損壞、移動或遺失', '其他損傷', '')])


@NEEDS_V1
class EquivalenceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as folder:
            edited, target = Path(folder) / 'edited.xlsx', Path(folder) / 'v2.xlsx'
            wb = load_workbook(SOURCE)
            apply_edits(wb)
            add_new_sheets(wb)
            wb.save(edited)
            convert(SOURCE, target)
            key = lambda r: (r['source'].split()[0], r['rule_id'])   # noqa: E731
            cls.old = {key(r): r for r in read_repairs(edited)}
            cls.new = {key(r): r for r in read_rules(target)}

    def test_same_rules_and_values(self):
        self.assertEqual(set(self.old) - set(self.new), set())
        for k, rule in self.old.items():
            self.assertEqual([rule[f] for f in ('pattern', 'method', 'unit', 'price', 'deterioration')],
                             [self.new[k][f] for f in ('pattern', 'method', 'unit', 'price', 'deterioration')], k)
            # 「適用構件」只會縮小範圍（主梁與橫隔梁、橋面板依網站選單區隔）
            self.assertTrue(set(self.new[k]['components']) <= set(rule['components']), k)

    def test_same_judgement(self):
        changed = {k[1] for k in self.old if behaviour(self.old[k]) != behaviour(self.new[k])}
        self.assertEqual(changed, CLEANED)


class CurrentRulesTests(unittest.TestCase):
    """使用中的規則表 1116：可讀取、劣化ID 不重複、主梁與橫隔梁／橋面板的劣化類型已區隔。"""

    def test_rules_file(self):
        from collections import Counter
        from bridgefill.core.data import component_repairs
        path = ROOT / 'data' / '劣化說明與維修單價1124.xlsx'
        rules = read_rules(path)
        self.assertGreater(len(rules), 500)
        workbook = load_workbook(path, read_only=True)
        ids = Counter()
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(values_only=True)
            headers = list(next(rows, ()))
            if '劣化ID' in headers:
                ids.update(str(r[headers.index('劣化ID')]) for r in rows if r and r[headers.index('劣化ID')])
        workbook.close()
        self.assertEqual([k for k, v in ids.items() if v > 1], [])
        main = {r['deterioration'] for r in component_repairs(rules, '主梁') if '剝落' in r['deterioration']}
        self.assertNotIn('混凝土剝落、破碎、鋼筋外露、銹蝕', main)
        self.assertTrue(component_repairs(rules, '護床工'))


class SynthesizeTests(unittest.TestCase):
    def test_manual_and_keywords(self):
        self.assertEqual(synthesize_note({'處理方式': '手動', '任一關鍵字': '坑洞'}), '手動處理')
        self.assertEqual(synthesize_note({'任一關鍵字': '碎裂、破碎', '必要關鍵字': '鋼筋外露'}),
                         '有關鍵字[碎裂]或[破碎]與[鋼筋外露]即可使用')

    def test_guardrail_range(self):
        note = synthesize_note({'條件欄位': '護欄高度', '最小值': 0.6, '最大值': 0.85, '條件單位': 'm'})
        self.assertEqual(condition({'note': note}, {'護欄高度(m)': '0.7'}), (True, False))
        self.assertEqual(condition({'note': note}, {'護欄高度(m)': '0.85'})[0], False)


if __name__ == '__main__':
    unittest.main()
