"""核對時修改本期欄位：備註樣態替換、Excel 規則排序、修改紀錄。"""
import unittest

from bridgefill.core.editing import best_rule, pattern_tags, replace_pattern, split_remark
from bridgefill.core.review import Review
from test_manual import make, river_case, find

RULES = [dict(pattern='草本植物生長', method='其他(清除植生)', unit='處', price='500', deterioration='植物生長',
              components=['橋台'], component='橋台'),
         dict(pattern='草本植物生長', method='其他(清除植生)', unit='平方公尺', price='100', deterioration='植物生長',
              components=['橋台'], component='橋台'),
         dict(pattern='混凝土剝落', method='混凝土修復', unit='處', price='800', deterioration='混凝土剝落',
              components=['橋台'], component='橋台')]


class RemarkTests(unittest.TestCase):
    def test_split_and_replace_keep_location_and_size(self):
        self.assertEqual(split_remark('A1-上下游 植生 (2處)'), ('A1-上下游', '植生', '(2處)'))
        self.assertEqual(split_remark('P2、P3掛淤'), ('P2、P3', '掛淤', ''))
        self.assertEqual(replace_pattern('A1-上下游 植生 (2處)', '草本植物生長'), 'A1-上下游 草本植物生長。(2處)')
        self.assertEqual(replace_pattern('S1-L,R 護欄高度不足。(45m) * 2', '護欄缺口'), 'S1-L,R 護欄缺口。(45m) * 2')
        self.assertEqual(replace_pattern('整體-上下游 淤積。(200m*30m*2m)', '掛淤'), '整體-上下游 掛淤。(200m*30m*2m)')

    def test_tags_rank_similar_patterns_first_and_rule_prefers_same_unit(self):
        tags = pattern_tags(RULES, '橋台', {'remark': 'A1 植物生長 (2處)', 'd_location': ''})
        self.assertEqual(tags, ['草本植物生長', '混凝土剝落'])
        self.assertEqual(best_rule(RULES, '草本植物生長', '平方公尺')['price'], '100')
        self.assertEqual(best_rule(RULES, '草本植物生長', '公尺')['price'], '500')


class EditTests(unittest.TestCase):
    def test_only_changed_fields_are_kept(self):
        pairs = river_case()
        pair = find(pairs, prior='502', current='9003')
        review = Review()
        kept = review.set_edit(pair, {'remark': 'S1G1 混凝土剝落。(1處)', 'u': '3', 'amount': '2'})
        self.assertEqual(kept, {'u': '3', 'amount': '2'})
        self.assertEqual(review.value(pair, 'u'), '3')
        self.assertEqual(review.value(pair, 'd'), '2')                 # 未改：取網站值（DERU）
        review.set_edit(pair, {'u': '2', 'amount': '1'})
        self.assertNotIn(pair.key, review.edits)
        with self.assertRaises(ValueError):
            review.set_edit(find(pairs, prior='501'), {'u': '1'})


if __name__ == '__main__':
    unittest.main()
