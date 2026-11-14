import tempfile
from pathlib import Path
import unittest
from bridgefill.core.conditions import aliases, condition
from bridgefill.core.data import component_repairs, repair_candidates
from bridgefill.core.policy import automatic_plan
from bridgefill.core.progress import ProgressStore


class ContextRuleTests(unittest.TestCase):
    def test_expansion_type_is_context_not_remark_keyword(self):
        rules = [dict(pattern='齒型伸縮縫高低差', method='齒型修復', unit='公尺', price='40000',
                      note='先判讀是否為[齒型]，在判讀關鍵字[高低差]'),
                 dict(pattern='角鋼伸縮縫高低差', method='角鋼修復', unit='公尺', price='13000',
                      note='先判讀是否為[角鋼]，在判讀關鍵字[高低差]')]
        for kind, expected in [('齒型', '齒型修復'), ('角鋼型', '角鋼修復')]:
            result = automatic_plan({'remark':'A1 高低差。(5m)'}, '伸縮縫', [], rules, {'伸縮縫型式':kind})
            self.assertFalse(result.review)
            self.assertEqual(result.fields['method'], expected)
        self.assertTrue(automatic_plan({'remark':'A1 高低差。(5m)'}, '伸縮縫', [], rules).review)
        self.assertFalse(repair_candidates(rules, {'remark':'A1 齒型。(5m)'}))

    def test_guardrail_height_boundaries_and_manual_options(self):
        rules = [dict(pattern='護欄高度不足', method='單管', unit='公尺', price='5000',
                      note='優先選擇(護欄高度大於0.85m小於1.1m適用)'),
                 dict(pattern='護欄高度不足', method='雙管', unit='公尺', price='7500',
                      note='手動選擇(讀取到既有護欄高度小於0.6m時使用)'),
                 dict(pattern='護欄高度不足', method='三管', unit='公尺', price='11000',
                      note='手動選擇(讀取到既有護欄高度小於0.3m時使用)')]
        for value in ('0.9','90cm','900mm'):
            result = automatic_plan({'remark':'A1 護欄高度不足。(20m)'}, '橋護欄', [], rules, {'橋護欄高度':value})
            self.assertFalse(result.review)
            self.assertEqual(result.fields['method'], '單管')
        for value in ('0.85','1.1','1.15','0.5','0.2','未知'):
            self.assertTrue(automatic_plan({'remark':'A1 護欄高度不足。(20m)'}, '橋護欄', [], rules, {'橋護欄高度':value}).review)
        self.assertEqual(condition(rules[1], {'橋護欄高度':'0.5'}), (True, True))
        self.assertEqual(condition(rules[0], {'橋護欄高度(cm)':'90'}), (True, False))
        self.assertEqual(condition(dict(rules[0], note='寬度小於1m'), {'橋護欄高度':'0.5'}), (None, True))
        self.assertEqual(condition(dict(rules[0], note='護欄高度0.85m<H<1.1m'), {'橋護欄高度':'0.5'}), (None, True))

    def test_alternative_keywords_and_component_lists(self):
        self.assertEqual(aliases('先判讀是否為[齒型]，關鍵字[漏水|滲水]'), ['漏水','滲水'])
        rules = [dict(pattern='漏水', method='止水', unit='公尺', price='700', note='有關鍵字[漏水|滲水]即可使用')]
        self.assertFalse(automatic_plan({'remark':'A1 滲水。(2m)'}, '伸縮縫', [], rules).review)
        scoped = [dict(rules[0], components=['引道路堤']), dict(rules[0], components=['伸縮縫', '橋台'])]
        self.assertEqual(component_repairs(scoped, '伸縮縫'), [scoped[1]])


class ResumeModeTests(unittest.TestCase):
    def test_include_skipped_preserves_done_pending_and_removes_row_cache(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ProgressStore(Path(directory) / 'progress.json')
            store.begin('job','rules')
            for name in ('done','pending','skipped'):
                store.mark(name, name, 'hash')
            store.mark_row('row')
            store.retry_skipped()
            self.assertFalse(store.done_row('row'))
            self.assertFalse(store.record('skipped'))
            self.assertEqual(store.record('done')['status'], 'done')
            self.assertEqual(store.record('pending')['status'], 'pending')

    def test_rule_upgrade_preserves_explicit_skips_until_restart_or_include(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ProgressStore(Path(directory) / 'progress.json')
            store.begin('job','v1')
            store.mark('skipped','skipped','hash')
            store.mark('done','done','hash')
            store.mark('pending','pending','hash')
            store.begin('job','v2')
            self.assertEqual(store.record('skipped')['status'], 'skipped')
            self.assertFalse(store.record('done'))
            self.assertEqual(store.record('pending')['status'], 'pending')
            store.reset_completed()
            self.assertFalse(store.record('skipped'))
            self.assertEqual(store.record('pending')['status'], 'pending')
