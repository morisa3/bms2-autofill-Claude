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
            self.assertEqual([rule[f] for f in ('pattern', 'method', 'unit', 'price', 'deterioration', 'components')],
                             [self.new[k][f] for f in ('pattern', 'method', 'unit', 'price', 'deterioration',
                                                       'components')], k)

    def test_same_judgement(self):
        changed = {k[1] for k in self.old if behaviour(self.old[k]) != behaviour(self.new[k])}
        self.assertEqual(changed, CLEANED)


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
