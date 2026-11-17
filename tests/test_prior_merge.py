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
