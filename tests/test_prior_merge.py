"""本期資料不完整時依前期補齊（2026-11-17 使用者指定的規則）。"""
import unittest

from bridgefill.core.matching import Defect, Pair
from bridgefill.core.prior_merge import merge_with_prior, normalize_size, only_size
from bridgefill.core.fill import READY, REVIEW, plan_item
from bridgefill.core.review import Review
from test_fill import RULE, current


def fields(remark, amount, unit, kind='混凝土剝落', location=''):
    return dict(deterioration=kind, remark=remark, amount=amount, unit=unit, d_location=location)


class MergeTests(unittest.TestCase):
    def test_prior_size_equals_current_amount(self):
        merged, why = merge_with_prior(fields('', '20', '平方公尺'), fields('A2 沖蝕 (10m*2m)', '1', '式'))
        self.assertEqual((merged['remark'], merged['amount'], merged['unit'], merged['d_location']),
                         ('A2 沖蝕。(10m*2m)', '20', '平方公尺', 'A2'))
        self.assertIn('與本期數量相同', why)
        self.assertEqual(merge_with_prior(fields('', '21', '平方公尺'), fields('A2 沖蝕 (10m*2m)', '1', '式')), (None, ''))

    def test_current_size_equals_prior_amount(self):
        merged, why = merge_with_prior(fields('(10m*2m)', '', '平方公尺'), fields('A2 沖蝕', '20', '平方公尺'))
        self.assertEqual((merged['remark'], merged['amount']), ('A2 沖蝕。(10m*2m)', '20'))
        self.assertIn('換算與前期數量相同', why)

    def test_same_type_and_unit_new_size_only(self):
        merged, why = merge_with_prior(fields('5*7', '', '平方公尺'), fields('A1 混凝土剝落 (4m*5m)', '20', '平方公尺'))
        self.assertEqual((merged['remark'], merged['amount'], merged['d_location']), ('A1 混凝土剝落。(5m*7m)', '35', 'A1'))
        merged, _ = merge_with_prior(fields('A3 5*7', '', '平方公尺'), fields('A1 混凝土剝落 (4m*5m)', '20', '平方公尺'))
        self.assertEqual(merged['remark'], 'A3 混凝土剝落。(5m*7m)')                 # 本期寫的位置優先
        merged, _ = merge_with_prior(fields('12', '', '公尺'), fields('A1 漏水 (9m)', '9', '公尺'))
        self.assertEqual((merged['remark'], merged['amount']), ('A1 漏水。(12m)', '12'))

    def test_not_merged(self):
        prior = fields('A1 混凝土剝落 (4m*5m)', '20', '平方公尺')
        self.assertIsNone(merge_with_prior(fields('5*7', '', '平方公尺', kind='其他損傷'), prior)[0])   # 類型不同
        self.assertIsNone(merge_with_prior(fields('5*7', '', '處'), prior)[0])                         # 單位不同
        self.assertIsNone(merge_with_prior(fields('A1 混凝土剝落 (5m*7m)', '35', '平方公尺'), prior)[0])  # 本期已完整

    def test_size_helpers(self):
        self.assertEqual(normalize_size('5*7*2', '平方公尺'), '(5m*7m)*2')
        self.assertEqual(normalize_size('5', '處'), '(5處)')
        self.assertEqual(normalize_size('5*7', '公尺'), '(5m)*7')                      # 公尺：5m 共 7 處
        self.assertEqual(normalize_size('0.3mm*2m', '公尺'), '(0.3mm*2m)')
        self.assertEqual(only_size('A2 (10m*2m)'), ('A2', '(10m*2m)'))
        self.assertIsNone(only_size('A2 混凝土剝落 (10m*2m)'))


class PlanTests(unittest.TestCase):
    def item(self, cur_remark, cur_amount, prior_remark, prior_amount, prior_unit):
        pair = current(remark=cur_remark, amount=cur_amount, unit='平方公尺')
        pair.current.fields['d_location'] = ''
        prior = Defect('前期', pair.current.row, '橋台', '', '3322', '混凝土剝落',
                       dict(pair.current.fields, remark=prior_remark, amount=prior_amount, unit=prior_unit),
                       der_id='400')
        return plan_item(Pair('沿用', prior, pair.current, confidence='高'), Review(),
                         [dict(RULE, unit='平方公尺', price='3200')])

    def test_fill_plan_uses_merged_prior(self):
        item = self.item('', '20', 'A1 混凝土剝落 (10m*2m)', '1', '式')
        self.assertEqual((item.status, item.target['amount'], item.target['price']), (READY, '20', '3200'))
        self.assertTrue(item.target['remark'].startswith('A1 混凝土剝落。'))
        self.assertIn('已導入前期位置與敘述', item.reason)
        item = self.item('5*7', '', 'A1 混凝土剝落 (4m*5m)', '20', '平方公尺')
        self.assertEqual((item.status, item.target['amount'], item.target['remark']), (READY, '35', 'A1 混凝土剝落。(5m*7m)'))
        self.assertEqual(self.item('', '21', 'A1 混凝土剝落 (10m*2m)', '1', '式').status, REVIEW)


if __name__ == '__main__':
    unittest.main()


class RealCaseRules(unittest.TestCase):
    """2026-11-18 使用者整理的規則（德安橋、萬年橋實際案例）。"""

    def merge(self, current, prior):
        from bridgefill.core.prior_merge import merge_with_prior
        base = {'deterioration': '其他損傷', 'd_location': '', 'method': '', 'amount': '', 'unit': '', 'price': '', 'remark': ''}
        return merge_with_prior({**base, **current}, {**base, **prior})

    def test_blank_unit_uses_size_conversion(self):
        fields, why = self.merge({'deterioration': '路堤損傷', 'amount': '25'},
                                 {'deterioration': '路堤損傷', 'remark': 'A1 老化 (5m*5m)', 'amount': '25', 'unit': '平方公尺'})
        self.assertEqual((fields['remark'], fields['unit'], fields['d_location']), ('A1 老化。(5m*5m)', '平方公尺', 'A1'))

    def test_everything_blank_imports_prior(self):
        fields, _ = self.merge({}, {'remark': 'P2 積土植生 (5m)', 'amount': '5', 'unit': '公尺'})
        self.assertEqual((fields['remark'], fields['amount'], fields['unit']), ('P2 積土植生。(5m)', '5', '公尺'))

    def test_prior_lump_sum_becomes_quantity(self):
        fields, _ = self.merge({'amount': '10', 'unit': '式'},
                               {'remark': '整體-上游 河道淤積 (5m*2m*1m)', 'amount': '1', 'unit': '式'})
        self.assertEqual((fields['remark'], fields['unit']), ('整體-上游 河道淤積。(5m*2m*1m)', '立方公尺'))

    def test_only_crack_size_keeps_prior_location(self):
        # 使用者範例：只填 0.5mm*2m*2，前期 A1-L,R 裂縫。(2m)*2 → A1-L,R 裂縫。(0.5mm*2m)*2（樣態再由 Excel 規則改成混凝土裂縫）
        fields, _ = self.merge({'deterioration': '混凝土結構裂縫', 'remark': '0.5mm＊2m＊2'},
                               {'deterioration': '混凝土結構裂縫', 'remark': 'A1-L,R 裂縫。(2m)*2', 'amount': '4', 'unit': '公尺'})
        self.assertEqual((fields['remark'], fields['d_location'], fields['unit']), ('A1-L,R 裂縫。(0.5mm*2m)*2', 'A1-L,R', '公尺'))

    def test_description_without_location(self):
        fields, _ = self.merge({'remark': '漏水。', 'amount': '5'}, {'remark': 'P1 積土植生 (5m)', 'amount': '5', 'unit': '公尺'})
        self.assertEqual((fields['remark'], fields['unit']), ('P1 漏水。(5m)', '公尺'))
        fields, _ = self.merge({'remark': '龜裂。(12m＊12m)', 'amount': '144'},
                               {'remark': 'A1引道路堤老化(12m*12m)', 'amount': '144', 'unit': '平方公尺'})
        self.assertTrue(fields['remark'].startswith('A1 龜裂。'))

    def test_location_and_format_tidy(self):
        from bridgefill.core.fill import tidy_locations
        self.assertEqual(tidy_locations({'d_location': '7D-5-1', 'remark': '7D-5-1 破損。(1處)'}, '橫隔梁')['remark'],
                         '7D5-1 破損。(1處)')
        self.assertEqual(tidy_locations({'d_location': 'S2G5-10', 'remark': 'S2G5-10 剝落。(3處)'}, '主梁')['d_location'], '2G5-10')
        self.assertEqual(tidy_locations({'remark': 'A2-3 支撐座 混凝土裂縫(0.1m＊0.2mm)＊1'}, '支承')['remark'],
                         'A2-3 支撐座混凝土裂縫。(0.2mm*0.1m)*1')
        self.assertEqual(tidy_locations({'remark': 'A1-0R.0L 裂縫。(2m)*2'}, '翼牆／擋土牆')['remark'], 'A1-0R,0L 裂縫。(2m)*2')
