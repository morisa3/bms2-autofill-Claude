"""第 6 步填寫清單：預定值（程式建議＋人工修改）、狀態判斷、進度紀錄。"""
import tempfile
import unittest
from pathlib import Path

from bridgefill.core.fill import (DONE, KEPT, PENDING, READY, REVIEW, SAME, SKIP, FillLog, build_plan, compare,
                                  plan_item)
from bridgefill.core.matching import Defect, Pair
from bridgefill.core.review import Review

RULE = dict(pattern='混凝土剝落', method='混凝土修復', unit='處', price='800', deterioration='混凝土剝落',
            note='有關鍵字[剝落]即可使用', rule_id='6-04', components=['橋台'], component='橋台',
            automatic_ready=True)


def current(remark='A1 混凝土剝落。(1處)', deru='3322', method='', price='', unit='處', amount='1', options=None):
    fields = {'deterioration': '混凝土剝落', 'd_location': 'A1', 'method': method, 'amount': amount, 'unit': unit,
              'price': price, 'remark': remark, 'd': deru[0], 'e': deru[1], 'r': deru[2], 'u': deru[3]}
    d = Defect('本期', ('abutment', '1'), '橋台', '橋台位置：橋頭', deru, '混凝土剝落', fields, method_id='9001',
               der_id='501', inspect_id='9', options=options or {'unit': ['處', '平方公尺'],
                                                                   'deterioration': ['混凝土剝落', '其他損傷']})
    return Pair('新增', current=d, confidence='高')


class PlanTests(unittest.TestCase):
    def test_rule_fills_method_and_price(self):
        item = plan_item(current(), Review(), [RULE])
        self.assertEqual(item.status, READY)
        self.assertEqual(set(item.changes), {'method', 'price'})
        self.assertEqual((item.target['method'], item.target['price']), ('混凝土修復', '800'))
        self.assertEqual(item.describe(), '維修工法：（空） → 混凝土修復；單價：（空） → 800')

    def test_already_matching_is_same(self):
        item = plan_item(current(method='混凝土修復', price='800'), Review(), [RULE])
        self.assertEqual((item.status, item.changes), (SAME, {}))

    def test_user_edit_wins_over_rule(self):
        pair = current()
        review = Review()
        review.set_edit(pair, {'price': '900', 'u': '3'})
        item = plan_item(pair, review, [RULE])
        self.assertEqual((item.status, item.target['price'], item.target['u'], item.target['method']),
                         (READY, '900', '3', '混凝土修復'))

    def test_manual_values_skip_rules(self):
        """人工確認的值送出時不再套用規則（2026-11-28：人工選的工法剛好和網站相同，卻被規則改回）。"""
        pair = current(method='紐澤西護欄＋單管護欄加高', price='5000')          # 網站現值＝人工要的值
        self.assertEqual(plan_item(pair, Review(), [RULE]).target['method'], '混凝土修復')   # 沒人工：套規則
        review = Review()
        values = {k: pair.current.fields[k] for k in ('deterioration', 'd_location', 'method', 'amount', 'unit',
                                                        'price', 'remark', 'd', 'e', 'r', 'u')}
        review.set_edit(pair, values, manual=True)
        item = plan_item(pair, review, [RULE])
        self.assertEqual((item.status, item.target['method'], item.target['price']), (SAME, '紐澤西護欄＋單管護欄加高', '5000'))
        self.assertIn('不套用 Excel 規則', item.reason)
        review.set_edit(pair, dict(values, u='3'), manual=True)                    # 人工另外改了 U：只送人工的值
        item = plan_item(pair, review, [RULE])
        self.assertEqual((item.status, set(item.changes)), (READY, {'u'}))
        self.assertEqual(Review.load(review.save(Path(tempfile.mkdtemp()) / 'r.json')).manual, review.manual)

    def test_unclear_rule_needs_review_unless_edited(self):
        pair = current(remark='A1 看不懂的描述')
        item = plan_item(pair, Review(), [RULE])
        self.assertEqual((item.status, item.changes), (REVIEW, {}))
        review = Review()
        review.set_edit(pair, {'method': '人工指定'})
        item = plan_item(pair, review, [RULE])
        self.assertEqual((item.status, set(item.changes)), (READY, {'method'}))
        self.assertIn('只填寫人工修改', item.reason)

    def test_confirmed_review_item_changes_status_and_reason(self):
        pair = current(remark='A1 看不懂的描述')
        review = Review()
        review.keep(pair)
        item = plan_item(pair, review, [RULE])
        self.assertEqual((item.status, item.changes), (KEPT, {}))
        self.assertIn('已人工確認維持網站現值', item.reason)
        review.set_edit(pair, {'method': '人工指定'})              # 之後修改仍以修改為準
        self.assertEqual(plan_item(pair, review, [RULE]).status, READY)
        review.keep(pair, False)
        review.set_edit(pair, {})
        self.assertEqual(plan_item(pair, review, [RULE]).status, REVIEW)

    def test_good_condition_and_missing_option(self):
        self.assertEqual(plan_item(current(deru='1___'), Review(), [RULE]).status, SKIP)
        pair = current()
        review = Review()
        review.set_edit(pair, {'unit': '公尺'})
        item = plan_item(pair, review, [RULE])
        self.assertEqual(item.status, REVIEW)
        self.assertIn('下拉選單沒有「公尺」', item.reason)

    def test_prior_wrong_ignores_prior_data(self):
        from bridgefill.core.fill import proposal_for
        pair = current()
        prior = Defect('前期', pair.current.row, '橋台', '', '3322', '混凝土剝落',
                       dict(pair.current.fields, remark='A1 混凝土剝落。(2處)', amount='2'), der_id='400')
        pair = Pair('範圍縮小', prior, pair.current, confidence='高')
        review = Review()
        review.decide(pair, '前期選擇或填寫錯誤')
        self.assertEqual(proposal_for(pair, [RULE], True).fields['amount'], '1')
        self.assertEqual(plan_item(pair, review, [RULE]).target['amount'], '1')

    def test_no_rules_only_edits(self):
        self.assertEqual(plan_item(current(), Review(), []).status, REVIEW)

    def test_prior_only_pairs_are_not_filled(self):
        prior = Pair('疑似漏建', prior=current().current)
        self.assertEqual(build_plan([prior, current()], Review(), [RULE])[0].pair.current.method_id, '9001')
        self.assertEqual(len(build_plan([prior, current()], Review(), [RULE])), 1)


class LogTests(unittest.TestCase):
    def test_log_marks_done_and_pending(self):
        pair = current()
        with tempfile.TemporaryDirectory() as folder:
            log = FillLog(Path(folder) / 'f.json')
            item = build_plan([pair], Review(), [RULE], log)[0]
            log.mark(item.key, DONE, item.target)
            again = FillLog(Path(folder) / 'f.json')
            self.assertEqual(build_plan([pair], Review(), [RULE], again)[0].status, DONE)
            again.mark(item.key, PENDING, item.target)
            self.assertEqual(build_plan([pair], Review(), [RULE], FillLog(Path(folder) / 'f.json'))[0].status, PENDING)

    def test_compare(self):
        self.assertEqual(compare({'d': '3', 'remark': ' a '}, {'d': '3', 'remark': 'a', 'u': ''}), [])
        self.assertEqual(compare({'d': '3'}, {'d': '2'}), ['d'])

    def test_compare_ignores_semicolon_rewrite(self):
        """網站儲存時改寫損壞位置、備註中的半形分號：回讀時忽略分隔符號（3.0.001.0，平和橋、雙東大橋）。"""
        from bridgefill.core.fill import differences
        target = {'d_location': '1D1-5,6,8;5-6', 'remark': '1D1-5,6,8;5-6 混凝土剝落鋼筋外露。(4處)', 'price': '800'}
        for site in ('1D1-5,6,8；5-6', '1D1-5,6,85-6', '1D1-5,6,8,5-6', '1D1-5,6,8; 5-6'):
            self.assertEqual(compare({**target, 'd_location': site, 'remark': site + ' 混凝土剝落鋼筋外露。(4處)'}, target), [])
        self.assertEqual(compare({**target, 'd_location': '1D1-5,6,9;5-6'}, target), ['d_location'])
        self.assertEqual(compare({**target, 'price': '8,00'}, target), ['price'])          # 其他欄位仍嚴格比對
        self.assertIn('損壞位置：網站「A1-R」≠ 預定「A1-R;A2-L,R」',
                      differences({'d_location': 'A1-R'}, {'d_location': 'A1-R;A2-L,R'}, ('d_location',)))


if __name__ == '__main__':
    unittest.main()
