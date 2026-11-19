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


class MergedTests(unittest.TestCase):
    def test_cover_shortage_merged_into_exposed_rebar(self):
        from bridgefill.core.matching import KIND_MERGED
        from bridgefill.core.review import choices_for
        result = match([prior('S1G1 保護層不足 (1處)'), prior('S1G1 鋼筋外露 (1處)')], [current('S1G1 鋼筋外露。(1處)')])
        merged = only(result, KIND_MERGED)
        self.assertEqual(merged.confidence, '需確認')
        self.assertIn('視為同一缺失', merged.reasons[0])
        self.assertIn(KIND_MERGED, choices_for(merged))
        # 不同位置、或本期沒有同組的缺失：仍是疑似漏建
        self.assertFalse([p for p in match([prior('S1G1 保護層不足 (1處)')], [current('S1G2 混凝土裂縫。(1處)')]).pairs
                          if p.kind == KIND_MERGED])


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

    def test_only_leftovers_at_same_location_are_paired_for_review(self):
        """使用者確認：4G1-20R 銹蝕蜂窩 與 混凝土剝落鋼筋外露 是同一筆。"""
        pair = only(match([prior('4G1-20R 銹蝕蜂窩 (1處)')], [current('4G1-20R 混凝土剝落鋼筋外露。(1處)')]), KIND_KEEP)
        self.assertEqual(pair.confidence, '需確認')
        self.assertIn('樣態或類型不同，但位置與單位相同，判為同一筆', pair.reasons)

    def test_same_location_different_pattern_with_other_leftovers_is_hinted(self):
        result = match([prior('4G1-20R 銹蝕蜂窩 (1處)'), prior('4G1-30R 白華 (1處)')],
                       [current('4G1-20R 混凝土剝落鋼筋外露。(1處)')])
        self.assertIn('可能對應前期：4G1-20R 銹蝕蜂窩 (1處)', only(result, KIND_NEW).reasons)

    def test_remark_location_must_match_its_component(self):
        """使用者確認：S3 那列本期寫成 S4 是寫錯。"""
        p = prior('S3-L,R 木本植物生長 (4處)', amount='4')
        c = current('S4-L,R 木本植物生長。(4處)', amount='4')
        p.label = c.label = '橋孔編號：S003｜橋面板編號：S3'
        pair = only(match([p], [c]), KIND_KEEP)
        self.assertEqual(pair.confidence, '需確認')
        self.assertIn('本期備註位置 S4 與所在構件 S3 不符', pair.reasons)
        ok = current('4G1-20R 剝落。(1處)')
        ok.label = '橋孔位置：S004｜編號：S4G1'
        self.assertFalse([r for r in only(match([prior('4G1-20R 剝落 (1處)')], [ok]), KIND_KEEP).reasons
                          if '不符' in r])

    def test_girder_code_with_or_without_span_prefix(self):
        pair = only(match([prior('S1G1-0,13 混凝土剝落鋼筋外露鏽蝕 (2處)', amount='2')],
                          [current('1G1-0,13 混凝土剝落鋼筋外露。(2處)', amount='2')]), KIND_KEEP)
        self.assertFalse([r for r in pair.reasons if r.startswith('位置')])

    def test_location_field_polluted_with_pattern(self):
        pair = only(match([prior('P1-1橋墩基礎裸露 0.6m', unit='式', location='P1-1橋墩基礎裸露')],
                          [current('P1-1 基礎裸露。(1m)', unit='式', location='P1-1')]), KIND_KEEP)
        self.assertIn('樣態相近', pair.reasons)

    def test_efflorescence_word_order(self):
        only(match([prior('1D-2-1 白華且滲水 (2m*1m)', unit='平方公尺', amount='2')],
                   [current('1D-2-1 滲水白華。(1m*2m)', unit='平方公尺', amount='2')]), KIND_KEEP)

    def test_on_site_minimum_type_only(self):
        """現場只填劣化類型與 DERU：同構件同類型只有一筆時視為同一缺失（2026-11-18 使用者規則：單一缺失留空直接導入前期）。"""
        result = match([prior('A1 混凝土剝落 (1處)', kind='混凝土剝落、破碎、鋼筋外露、銹蝕')],
                       [current('', kind='混凝土剝落、破碎、鋼筋外露、銹蝕')])
        pair = only(result, KIND_KEEP)
        self.assertEqual(pair.confidence, '高')
        self.assertIn('本期備註未填，同構件同劣化類型只有一筆，視為同一缺失', pair.reasons)


PRIVATE = ROOT / 'fixtures' / 'private'


@unittest.skipUnless((PRIVATE / 'v1').is_dir() and (PRIVATE / 'scan').is_dir(), '沒有私人比對資料')
class RealMatchingTests(unittest.TestCase):
    def test_zhuzikeng_accuracy_against_confirmed_pairs(self):
        sys.path.insert(0, str(ROOT / 'tools'))
        from evaluate_matching import evaluate
        right, wrong, _ = evaluate(PRIVATE / 'v1' / '竹子坑橋-比對未修改.json',
                                   PRIVATE / 'v1' / '竹子坑橋-比對修改完成.json', show=False)
        self.assertEqual(wrong, 0)    # 含使用者 2026-11-15 更正的 2 筆答案

    def test_mumaliao_accuracy_against_confirmed_pairs(self):
        sys.path.insert(0, str(ROOT / 'tools'))
        from evaluate_matching import evaluate
        found = list((PRIVATE / 'v1' / '木馬寮橋').glob('compare-*.json'))
        if not found:
            self.skipTest('沒有木馬寮橋比對檔')
        right, wrong, _ = evaluate(found[0], found[0], show=False)    # 鎖定後的比對檔即使用者確認結果
        self.assertEqual((right, wrong), (43, 0))

    def test_chanye_scan_pairs(self):
        read = lambda pattern: json.loads(next((PRIVATE / 'scan').glob(pattern)).read_text(encoding='utf-8'))
        result = match_scans(read('*-前期.json'), read('*-檢測.json'))
        summary = result.summary
        self.assertEqual(sum(1 for p in result.pairs if p.current), 47)       # 本期 47 筆全部有判定
        self.assertEqual(sum(1 for p in result.pairs if p.prior), 46)         # 前期 46 筆全部有判定
        self.assertEqual(summary['需確認'], 3)
        typo = [p for p in result.pairs if p.prior and p.current and 'S3-L,R 木本' in p.prior.remark][0]
        self.assertIn('本期備註位置 S4 與所在構件 S3 不符', typo.reasons)       # 本期 S3 那列寫成 S4

    def test_chanye_pairs_confirmed_by_user_stay_the_same(self):
        """使用者逐筆核對過的結果（fixtures/private/產業一號橋-配對人工判斷.xlsx）當回歸測試。"""
        from openpyxl import load_workbook
        book = PRIVATE / '產業一號橋-配對人工判斷.xlsx'
        if not book.exists():
            self.skipTest('沒有人工判斷檔')
        sheet = load_workbook(book, data_only=True).active
        heads = [c.value for c in sheet[3]]
        confirmed = [dict(zip(heads, r)) for r in sheet.iter_rows(min_row=4, values_only=True)]
        read = lambda pattern: json.loads(next((PRIVATE / 'scan').glob(pattern)).read_text(encoding='utf-8'))
        result = match_scans(read('*-前期.json'), read('*-檢測.json'))
        made = {((p.prior.remark.strip() if p.prior else None), (p.current.remark.strip() if p.current else None),
                 p.kind) for p in result.pairs}
        for row in confirmed:
            if row.get('人工判斷') == '正確':
                self.assertIn((row['前期備註'], row['本期備註'], row['判定']), made, row)
        # 使用者說明為同一筆的 4G1-20R
        self.assertIn(('4G1-20R 銹蝕蜂窩 (1處)', '4G1-20R 混凝土剝落鋼筋外露。(1處)', KIND_KEEP), made)


if __name__ == '__main__':
    unittest.main()


class BlankRemarkTests(unittest.TestCase):
    """2026-11-19 實際案例：本期只填 DERU、劣化類型（數量），備註未填 → 與同構件同類型的前期配對。"""

    def defect(self, period, kind, remark='', amount='', unit='', deru='2211', row=('expansion', '1')):
        from bridgefill.core.matching import Defect
        fields = dict(deterioration=kind, remark=remark, amount=amount, unit=unit, d_location='', method='', price='')
        return Defect(period, row, '伸縮縫', '位置：橋台(A)｜編號：A1', deru, kind, fields)

    def test_single_blank_pairs_with_prior(self):
        from bridgefill.core.matching import match
        priors = [self.defect('前期', '伸縮縫間雜物堆積', 'A1 積土 (9.3m)', '9.3', '公尺', '2411'),
                  self.defect('前期', '漏水', 'A1 滲水 (9.3m)', '9.3', '公尺', '2422')]
        currents = [self.defect('本期', '伸縮縫間雜物堆積', deru='2411'), self.defect('本期', '漏水', deru='2422')]
        pairs = match(priors, currents).pairs
        self.assertEqual(len(pairs), 2)
        for pair in pairs:
            self.assertIsNotNone(pair.prior)
            self.assertEqual(pair.prior.deterioration, pair.current.deterioration)
            self.assertEqual((pair.kind, pair.confidence), (KIND_KEEP, '高'))      # 單位空白不算單位變更

    def test_several_blanks_pair_by_deru(self):
        from bridgefill.core.matching import match
        priors = [self.defect('前期', '漏水', 'A1-L 滲水 (2m)', '2', '公尺', '2311'),
                  self.defect('前期', '漏水', 'A1-R 滲水 (3m)', '3', '公尺', '2422')]
        currents = [self.defect('本期', '漏水', deru='2422'), self.defect('本期', '漏水', deru='2311')]
        pairs = match(priors, currents).pairs
        by_deru = {p.current.deru: p.prior.remark for p in pairs}
        self.assertEqual(by_deru, {'2422': 'A1-R 滲水 (3m)', '2311': 'A1-L 滲水 (2m)'})
        self.assertTrue(all(p.confidence == '需確認' for p in pairs))
