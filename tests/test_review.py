import tempfile
import unittest
from pathlib import Path

from bridgefill.core.matching import KIND_KEEP, KIND_MISSING, KIND_NEW, KIND_REPAIRED, KIND_UNIT, match
from bridgefill.core.review import (ACCEPTED, CHANGED, KIND_GONE, PENDING, Review, choices_for, review_path)
from test_matching import current, prior


def sample():
    p1, c1 = prior('A1 剝落 (1處)'), current('A1 剝落。(1處)')
    p2 = prior('P2、P3掛淤')
    c2 = current('S2 木本植物生長。(1處)')
    p1.der_id, p2.der_id, c1.method_id, c2.method_id = '501', '502', '9001', '9002'
    p2.row = c2.row = ('waterway', '7')
    return match([p1, p2], [c1, c2]).pairs


class ReviewTests(unittest.TestCase):
    def test_keys_use_site_ids(self):
        self.assertEqual(sorted(p.key for p in sample()), ['前期:502', '本期:9001', '本期:9002'])

    def test_sync_marks_confident_as_accepted_and_doubtful_as_pending(self):
        pairs = sample()
        review = Review().sync(pairs)
        by_key = {p.key: review.decision(p) for p in pairs}
        self.assertEqual(by_key['本期:9001'].status, ACCEPTED)
        self.assertEqual(by_key['前期:502'].status, PENDING)
        self.assertEqual(review.counts(pairs)[PENDING], 2)

    def test_decide_validates_choices(self):
        pairs = sample()
        review = Review().sync(pairs)
        missing = next(p for p in pairs if p.kind == KIND_MISSING)
        self.assertEqual(review.decide(missing, KIND_GONE, '河水沖刷，已無掛淤').status, CHANGED)
        self.assertEqual(review.decide(missing, KIND_MISSING).status, ACCEPTED)
        with self.assertRaises(ValueError):
            review.decide(missing, KIND_KEEP)
        self.assertIn(KIND_REPAIRED, choices_for(missing))
        new = next(p for p in pairs if p.kind == KIND_NEW)
        self.assertEqual(choices_for(new), (KIND_NEW,))

    def test_save_load_and_resume_after_rescan(self):
        pairs = sample()
        review = Review('產業一號橋', '2026-11-03', '2025-12-04').sync(pairs)
        missing = next(p for p in pairs if p.kind == KIND_MISSING)
        review.decide(missing, KIND_GONE, '河水沖刷')
        keep = next(p for p in pairs if p.kind == KIND_KEEP)
        review.decide(keep, KIND_UNIT)
        with tempfile.TemporaryDirectory() as folder:
            path = review.save(review_path(folder, '產業一號橋', '2026-11-03'))
            self.assertEqual(path.name, '2026-11-03-產業一號橋-審查.json')
            restored = Review.load(path).sync(sample())    # 重新掃描、重新配對後
        again = {p.key: restored.decision(p) for p in sample()}
        self.assertEqual((again['前期:502'].kind, again['前期:502'].note), (KIND_GONE, '河水沖刷'))
        self.assertEqual(again['本期:9001'].kind, KIND_UNIT)
        self.assertEqual(restored.bridge, '產業一號橋')

    def test_file_name_is_safe(self):
        self.assertEqual(review_path('x', '無名橋(鹿/13)', '2026-10-15').name, '2026-10-15-無名橋(鹿_13)-審查.json')


if __name__ == '__main__':
    unittest.main()
