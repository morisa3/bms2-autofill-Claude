import json
import sys
import unittest
from pathlib import Path

from bridgefill.core.matching import (KIND_KEEP, KIND_LARGER, KIND_MISSING, KIND_NEW, KIND_REPAIRED, KIND_SMALLER,
                                      KIND_UNIT, Defect, match, match_scans)

ROOT = Path(__file__).resolve().parent.parent


def defect(period, remark, kind='其他損傷', amount='1', unit='處', row=('girder', '1'), d='2', location='',
           repaired=False):
    fields = {'deterioration': kind, 'remark': remark, 'amount': amount, 'unit': unit, 'd_location': location}
    return Defect(period, row, '主梁', '', f'{d}222', kind, fields, repaired)


def prior(*args, **kwargs):
    return defect('前期', *args, **kwargs)


def current(*args, **kwargs):
    return defect('本期', *args, **kwargs)


def only(result, kind):
    found = [p for p in result.pairs if p.kind == kind]
    assert len(found) == 1, (kind, result.pairs)
    return found[0]


class MatchingTests(unittest.TestCase):
    def test_same_defect_is_kept_with_high_confidence(self):
        pair = only(match([prior('S1G1-3 混凝土剝落 (1處)')], [current('S1G1-3 混凝土剝落。(1處)')]), KIND_KEEP)
        self.assertEqual((pair.confidence, pair.reasons[0]), ('高', '樣態相同'))

    def test_synonyms_and_size_order(self):
        pair = only(match([prior('A1 滲水 (9m)', unit='公尺', amount='9')],
                          [current('A1 漏水。(9m)', unit='公尺', amount='9')]), KIND_KEEP)
        self.assertEqual(pair.confidence, '高')
        only(match([prior('A1 坑洞 (9m*1m)', unit='平方公尺', amount='9')],
                   [current('A1 坑洞。(1m*9m)', unit='平方公尺', amount='9')]), KIND_KEEP)

    def test_larger_smaller_and_unit_change(self):
        only(match([prior('A2 滲水白華 (1m*9m)', unit='平方公尺', amount='9')],
                   [current('A2 滲水白華。(3m*5m)', unit='平方公尺', amount='15')]), KIND_LARGER)
        only(match([prior('A1 木本植物生長 (2處)', amount='2')], [current('A1 木本植物生長。(1處)')]), KIND_SMALLER)
        only(match([prior('A1 明顯裂縫 (0.2mm*3m)*5', unit='公尺', amount='15')],
                   [current('A1 混凝土裂縫。', unit='式', amount='1')]), KIND_UNIT)

    def test_type_correction_from_other_damage_is_not_doubtful(self):
        pair = only(match([prior('A1 坑洞 (9m*1m)', kind='其他損傷', unit='平方公尺', amount='9')],
                          [current('A1 坑洞。(1m*9m)', kind='路堤損傷', unit='平方公尺', amount='9')]), KIND_KEEP)
        self.assertEqual(pair.confidence, '高')
        self.assertIn('劣化類型修正（其他損傷 → 路堤損傷）', pair.reasons)

    def test_exclusive_patterns_are_never_paired(self):
        result = match([prior('S2 木本植物生長 (1處)')], [current('S2 草本植物生長。(1處)')])
        self.assertEqual(sorted(p.kind for p in result.pairs), [KIND_NEW, KIND_MISSING])

    def test_unmatched_prior_is_repaired_or_missing(self):
        result = match([prior('A1 剝落 (1處)', repaired=True), prior('A2 掛淤 (1處)')], [])
        self.assertEqual({p.kind for p in result.pairs}, {KIND_REPAIRED, KIND_MISSING})
        self.assertEqual(only(result, KIND_MISSING).confidence, '需確認')

    def test_rows_are_never_crossed(self):
        result = match([prior('S1 剝落 (1處)', row=('slab', '1'))], [current('S1 剝落。(1處)', row=('slab', '2'))])
        self.assertEqual(sorted(p.kind for p in result.pairs), [KIND_NEW, KIND_MISSING])

    def test_same_location_different_pattern_is_hinted(self):
        result = match([prior('4G1-20R 銹蝕蜂窩 (1處)')], [current('4G1-20R 混凝土剝落鋼筋外露。(1處)')])
        self.assertIn('可能對應前期：4G1-20R 銹蝕蜂窩 (1處)', only(result, KIND_NEW).reasons)
        self.assertIn('可能對應本期：4G1-20R 混凝土剝落鋼筋外露。(1處)', only(result, KIND_MISSING).reasons)

    def test_girder_code_with_or_without_span_prefix(self):
        pair = only(match([prior('S1G1-0,13 混凝土剝落鋼筋外露鏽蝕 (2處)', amount='2')],
                          [current('1G1-0,13 混凝土剝落鋼筋外露。(2處)', amount='2')]), KIND_KEEP)
        self.assertFalse([r for r in pair.reasons if r.startswith('位置')])

    def test_location_field_polluted_with_pattern(self):
        pair = only(match([prior('P1-1橋墩基礎裸露 0.6m', unit='式', location='P1-1橋墩基礎裸露')],
                          [current('P1-1 基礎裸露。(1m)', unit='式', location='P1-1')]), KIND_KEEP)
        self.assertIn('樣態相近', pair.reasons)

    def test_on_site_minimum_type_only(self):
        """現場只填劣化類型與 DERU：以類型配對，但只能標為需確認。"""
        result = match([prior('A1 混凝土剝落 (1處)', kind='混凝土剝落、破碎、鋼筋外露、銹蝕')],
                       [current('', kind='混凝土剝落、破碎、鋼筋外露、銹蝕')])
        pair = only(result, KIND_KEEP)
        self.assertEqual(pair.confidence, '需確認')
        self.assertIn('本期備註未填，以劣化類型判斷', pair.reasons)


PRIVATE = ROOT / 'fixtures' / 'private'


@unittest.skipUnless((PRIVATE / 'v1').is_dir() and (PRIVATE / 'scan').is_dir(), '沒有私人比對資料')
class RealMatchingTests(unittest.TestCase):
    def test_zhuzikeng_accuracy_against_confirmed_pairs(self):
        sys.path.insert(0, str(ROOT / 'tools'))
        from evaluate_matching import evaluate
        right, wrong, _ = evaluate(PRIVATE / 'v1' / '竹子坑橋-比對未修改.json',
                                   PRIVATE / 'v1' / '竹子坑橋-比對修改完成.json', show=False)
        self.assertGreaterEqual(right / (right + wrong), 0.94)

    def test_chanye_scan_pairs(self):
        read = lambda pattern: json.loads(next((PRIVATE / 'scan').glob(pattern)).read_text(encoding='utf-8'))
        result = match_scans(read('*-前期.json'), read('*-檢測.json'))
        summary = result.summary
        self.assertEqual(sum(1 for p in result.pairs if p.current), 47)       # 本期 47 筆全部有判定
        self.assertEqual(sum(1 for p in result.pairs if p.prior), 46)         # 前期 46 筆全部有判定
        self.assertLessEqual(summary['需確認'], 6)
        typo = [p for p in result.pairs if p.prior and p.current and 'S3-L,R 木本' in p.prior.remark][0]
        self.assertEqual(typo.confidence, '需確認')                            # 本期 S3 那列寫成 S4


if __name__ == '__main__':
    unittest.main()
