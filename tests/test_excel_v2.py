import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'tools'))

from excel_v2 import convert, parse_note, template  # noqa: E402
from bridgefill.core.data import read_repairs  # noqa: E402

SOURCE = ROOT / 'data' / '劣化說明與維修單價1111v1.xlsx'


class ParseNoteTests(unittest.TestCase):
    def parse(self, note, unit='處', pattern='樣態'):
        fields, review, remark = parse_note(note, unit, pattern)
        return {k: v for k, v in fields.items() if v not in ('', None)}, review, remark

    def test_any_and_required_keywords(self):
        fields, review, _ = self.parse('有關鍵字[碎裂]或[破碎]與[鋼筋外露]即可使用')
        self.assertEqual(fields['任一關鍵字'], '碎裂、破碎')
        self.assertEqual(fields['必要關鍵字'], '鋼筋外露')
        self.assertFalse(review)

    def test_all_required_keywords(self):
        fields, _, _ = self.parse('有關鍵字[破損]與[鋼筋外露]即可使用')
        self.assertEqual(fields['必要關鍵字'], '破損、鋼筋外露')
        self.assertNotIn('任一關鍵字', fields)

    def test_crack_width_boundary_follows_v0976(self):
        wide, _, remark = self.parse('裂縫寬度大於[0.3mm]使用', '公尺', '混凝土裂縫')
        narrow, _, _ = self.parse('裂縫寬度小於[0.3mm]使用', '公尺', '混凝土裂縫')
        self.assertEqual((wide['條件欄位'], wide['最小值'], wide['條件單位']), ('裂縫寬度', 0.3, 'mm'))
        self.assertEqual(narrow['最大值'], 0.3)
        self.assertIn('0.3mm 歸入', remark)
        self.assertEqual(wide['備註範本'], '{位置} {樣態}。(mm*m)*1')

    def test_alone_keyword_with_u_value(self):
        fields, _, _ = self.parse('有關鍵字[木本植物]或[木本植生]即可使用，或僅有關鍵字[植生]時，若U值為2優先使用')
        self.assertEqual((fields['單獨關鍵字'], fields['單獨關鍵字U值']), ('植生', 2))

    def test_type_correction_and_precheck(self):
        fields, _, _ = self.parse('若判讀到關鍵字[生銹]或[生鏽]即可使用，但劣化類型填成[其他損傷]，則修正成[防撞鋼板損傷]後使用')
        self.assertEqual(fields['修正劣化類型為'], '防撞鋼板損傷')
        fields, _, _ = self.parse('優先使用，先判讀是否為[人造橡膠支承墊]，再判讀關鍵字[變形]或[開裂]')
        self.assertEqual((fields['前置判讀'], fields['任一關鍵字'], fields['優先序']), ('人造橡膠支承墊', '變形、開裂', 1))

    def test_unparseable_conditions_need_review(self):
        _, review, remark = self.parse('若傷及混凝土層時使用，有關鍵字[坑洞]即可使用')
        self.assertTrue(review)
        self.assertIn('傷及混凝土層', remark)
        _, review, _ = self.parse('手動選擇，或讀取到護欄高度大於0.60小於0.85時使用', '公尺', '護欄高度不足')
        self.assertTrue(review)

    def test_templates_by_unit(self):
        self.assertEqual(template('平方公尺', 'AC剝離'), '{位置} {樣態}。(m*m)*1')
        self.assertEqual(template('立方公尺', '路面坑洞'), '{位置} {樣態}。(m*m*m)*1')
        self.assertEqual(template('公尺', '漏水'), '{位置} {樣態}。(m)*1')
        self.assertEqual(template('式', '基礎裸露'), '')


class ConvertTests(unittest.TestCase):
    def test_v1_reader_reads_identical_rules_from_converted_file(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            stats = convert(SOURCE, target)
            self.assertEqual(stats['rows'], 493)
            self.assertEqual(read_repairs(SOURCE), read_repairs(target))


if __name__ == '__main__':
    unittest.main()
