"""人工改配對與跨構件提示。"""
import tempfile
import unittest
from pathlib import Path

from bridgefill.core.manual import LEFT, MANUAL, SPLIT, apply_manual, candidates
from bridgefill.core.matching import (KIND_KEEP, KIND_LARGER, KIND_MISSING, KIND_NEW, Defect, defect_id, match)
from bridgefill.core.review import CHANGED, PENDING, Review


def make(period, component, row, remark, ident, amount='1', unit='處'):
    fields = {'deterioration': '其他損傷', 'remark': remark, 'amount': amount, 'unit': unit, 'd_location': ''}
    d = Defect(period, row, component, '', '2222', '其他損傷', fields)
    if period == '本期':
        d.method_id = ident
    else:
        d.der_id = ident
    return d


def river_case():
    """前期河道「P2、P3掛淤」，本期拆到橋墩基礎 P2、P3 兩筆；另有一組同列正常配對。"""
    p1 = make('前期', '河道', ('river', '1'), 'P2、P3掛淤', '501', '2')
    p2 = make('前期', '主梁', ('girder', '1'), 'S1G1 混凝土剝落 (1處)', '502')
    c1 = make('本期', '橋墩基礎', ('pier', '2'), 'P2 掛淤。(1處)', '9001')
    c2 = make('本期', '橋墩基礎', ('pier', '3'), 'P3 掛淤。(2處)', '9002', '2')
    c3 = make('本期', '主梁', ('girder', '1'), 'S1G1 混凝土剝落。(1處)', '9003')
    c4 = make('本期', '支承／支承墊／阻尼裝置', ('bearing', '1'), 'P5 支承墊生銹。(1處)', '9004')
    return match([p1, p2], [c1, c2, c3, c4]).pairs


def find(pairs, prior=None, current=None):
    return next(p for p in pairs if defect_id(p.prior) == (prior or '') and defect_id(p.current) == (current or ''))


class CrossHintTests(unittest.TestCase):
    def test_moved_silt_is_hinted_not_paired(self):
        pairs = river_case()
        silt = find(pairs, prior='501')
        self.assertEqual(silt.kind, KIND_MISSING)
        self.assertTrue(any('可能對應本期：橋墩基礎 P2 掛淤' in r and '位置變更' in r for r in silt.reasons))
        self.assertTrue(any('可能對應前期：河道' in r for r in find(pairs, current='9001').reasons))
        self.assertFalse(any('可能對應' in r for r in find(pairs, current='9004').reasons))   # 支承不在相關構件

    def test_unrelated_location_codes_are_not_hinted(self):
        p = make('前期', '河道', ('river', '1'), 'P7掛淤', '501')
        c = make('本期', '橋墩基礎', ('pier', '2'), 'P2 掛淤。(1處)', '9001')
        self.assertFalse(any('可能對應' in r for pair in match([p], [c]).pairs for r in pair.reasons))


class CandidateTests(unittest.TestCase):
    def test_order_same_row_then_component_then_other_unmatched_then_taken(self):
        pairs = river_case()
        options = candidates(pairs, find(pairs, prior='501'))
        tiers = [o.tier for o in options]
        self.assertEqual(tiers, sorted(tiers))
        other = [o for o in options if o.tier == 3]
        self.assertEqual([defect_id(o.defect) for o in other[:2]], ['9001', '9002'])   # 樣態相同排最前
        taken = [o for o in options if o.tier == 4]
        self.assertEqual([defect_id(o.defect) for o in taken], ['9003'])
        self.assertIn('主梁', taken[0].partner)

    def test_current_side_lists_priors(self):
        pairs = river_case()
        options = candidates(pairs, find(pairs, current='9001'))
        self.assertTrue(all(o.defect.period == '前期' for o in options))
        self.assertEqual(defect_id(options[0].defect), '501')


class ApplyTests(unittest.TestCase):
    def test_one_to_many_across_components(self):
        pairs = apply_manual(river_case(), {'501': ['9001', '9002']})
        linked = [p for p in pairs if p.manual]
        self.assertEqual(sorted(defect_id(p.current) for p in linked), ['9001', '9002'])
        self.assertTrue(all(p.moved and p.kind == KIND_LARGER for p in linked))     # 前期 2 → 本期合計 3
        self.assertIn(MANUAL, linked[0].reasons)
        self.assertTrue(any('位置變更（河道 → 橋墩基礎）' in r for r in linked[0].reasons))
        self.assertTrue(any('一對多' in r for r in linked[0].reasons))
        self.assertFalse([p for p in pairs if defect_id(p.prior) == '501' and not p.current])
        self.assertEqual(len(pairs), 4)

    def test_stealing_a_paired_defect_leaves_partner_alone(self):
        pairs = apply_manual(river_case(), {'501': ['9003']})
        left = find(pairs, prior='502')
        self.assertEqual((left.kind, left.confidence, left.reasons), (KIND_MISSING, '需確認', [LEFT]))

    def test_split(self):
        pairs = apply_manual(river_case(), splits=[['502', '9003']])
        self.assertEqual(find(pairs, prior='502').reasons, [SPLIT])
        self.assertEqual(find(pairs, current='9003').kind, KIND_NEW)


class ReviewLinkTests(unittest.TestCase):
    def test_link_split_save_and_reload(self):
        auto = river_case()
        review = Review('測試橋', '2026-11-05', '2024-12-05').sync(auto)
        self.assertEqual(review.decision(find(auto, prior='501')).status, PENDING)
        review.link(auto, '501', ['9001', '9002'])
        pairs = apply_manual(auto, review.links, review.splits)
        review.sync(pairs)
        linked = find(pairs, prior='501', current='9001')
        self.assertEqual((review.decision(linked).kind, review.decision(linked).status), (KIND_LARGER, CHANGED))

        review.split(pairs, linked)       # 拆掉其中一筆：另一筆仍是人工配對（變成一對一）
        self.assertEqual(review.links, {'501': ['9002']})
        pairs = apply_manual(auto, review.links, review.splits)
        review.sync(pairs)
        self.assertEqual(find(pairs, current='9001').kind, KIND_NEW)
        self.assertEqual(find(pairs, prior='501', current='9002').kind, KIND_KEEP)

        with tempfile.TemporaryDirectory() as folder:
            path = review.save(Path(folder) / 'r.json')
            again = Review.load(path)
        self.assertEqual((again.links, again.splits), (review.links, review.splits))

    def test_relinking_moves_current_between_priors(self):
        auto = river_case()
        review = Review()
        review.link(auto, '501', ['9001'])
        review.link(auto, '502', ['9001', '9003'])
        self.assertEqual(review.links, {'502': ['9001', '9003']})


if __name__ == '__main__':
    unittest.main()
