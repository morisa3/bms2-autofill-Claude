"""Excel「樣態同義詞」「位置代碼對照」（2026-11-19 使用者指定）。"""
import unittest
from pathlib import Path

from bridgefill.core.fill import standardize, tidy_locations
from bridgefill.core.synonyms import configure, read_tables, standard_location, standard_pattern

RULES = Path(__file__).resolve().parents[1] / 'data' / '劣化說明與維修單價1123.xlsx'


class SynonymTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        configure(RULES)
        cls.table = read_tables(RULES)

    def test_patterns(self):
        cases = [('A1 老化。(5m*5m)', '引道路堤', '', 'A1 AC剝離。(5m*5m)'),
                 ('A1 引道路堤老化。(12m*12m)', '引道路堤', '', 'A1 AC剝離。(12m*12m)'),
                 ('A1 龜裂。(12m*12m)', '引道路堤', '', 'A1 AC龜裂。(12m*12m)'),
                 ('6G3-ALL 火害。(1m*20m)', '主梁', '', '6G3-ALL 表面燻黑。(1m*20m)'),
                 ('S6-L,R 鋼筋未剪。(80處)', '橋面板', '', 'S6-L,R 鋼筋未切。(80處)'),
                 ('A1-L,R 破損鋼筋外露鏽蝕。(10處)', '引道路堤護欄', '', 'A1-L,R 混凝土破損鋼筋外露。(10處)'),
                 ('A1 鋼筋外露鏽蝕。(1處)', '橋台', '', 'A1 混凝土剝落鋼筋外露。(1處)'),
                 ('P1-上游 植物生長。(1處)', '墩柱與帽梁', '1', 'P1-上游 草本植物生長。(1處)'),
                 ('P1-上游 植物生長。(1處)', '墩柱與帽梁', '2', 'P1-上游 木本植物生長。(1處)')]
        for remark, component, u, expected in cases:
            self.assertEqual(standard_pattern(remark, component, u, self.table['patterns']), expected, remark)

    def test_does_not_touch_other_patterns(self):
        for remark, component, u in (('A1 混凝土表面龜裂。(1m)', '橋台', ''), ('A1 草本植物生長。(1處)', '橋台', '2'),
                                     ('A1 龜裂。(1m)', '橋台', '')):        # 龜裂 → AC龜裂 只用於引道路堤、其他
            self.assertEqual(standard_pattern(remark, component, u, self.table['patterns']), remark)

    def test_location_rule(self):
        fields = {'d_location': '整體-ALL', 'remark': '整體-ALL 排水設施阻塞。(56處)'}
        self.assertEqual(standard_location(fields, '橋梁排水設施', self.table['locations'])['remark'], '整體-L,R 排水設施阻塞。(56處)')
        self.assertEqual(tidy_locations(fields, '橋護欄')['d_location'], '整體-L,R')
        self.assertEqual(tidy_locations(fields, '主梁')['d_location'], '整體-ALL')       # 其他構件不改

    def test_standardize_fields(self):
        self.assertEqual(standardize({'remark': 'S1 鋼筋未剪。(80處)'}, '橋面板')['remark'], 'S1 鋼筋未切。(80處)')


if __name__ == '__main__':
    unittest.main()
