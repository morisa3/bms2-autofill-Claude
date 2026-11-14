import tempfile
import unittest
from pathlib import Path
from bridgefill.core.case_identity import case_key,case_job,matching_metadata
from bridgefill.core.progress import ProgressStore

class CaseIdentityTests(unittest.TestCase):
    def test_submission_routes_and_filters_do_not_change_case(self):
        a='https://bms2.iot.gov.tw/bms3/inspect/inspect/edit/inspect_id,149753/inventory_id,14942/sel_type,1/'
        b='https://bms2.iot.gov.tw/bms3/inspect/inspect/der_content/inventory_id,14942/inspect_id,149753/sel_type,/#top'
        self.assertEqual(case_key(a),case_key(b))
        self.assertEqual(case_job('base',a),case_job('base',b))
        self.assertNotEqual(case_key(a),case_key(b.replace('149753','149754')))
        self.assertNotEqual(case_key(a),case_key(b.replace('14942','14943')))
    def test_query_keys_and_duplicate_or_missing_keys(self):
        self.assertIsNotNone(case_key('https://example.test/main?inspect_id=1&inventory_id=2'))
        self.assertIsNone(case_key('https://example.test/main?inspect_id=1'))
        self.assertIsNone(case_key('https://example.test/inspect_id,1/?inspect_id=2&inventory_id=3'))
    def test_legacy_metadata_requires_bridge_code_and_date(self):
        a={'橋梁編號':'13-030','檢測日期':'2026-08-11','橋梁名稱':'龍門橋'}
        self.assertTrue(matching_metadata(a,dict(a,檢測日期='2026/8/11')))
        for b in ({},dict(a,檢測日期='2024-10-12'),dict(a,橋梁編號='13-031'),dict(a,橋梁名稱='其他橋')):
            self.assertFalse(matching_metadata(a,b))
    def test_legacy_checkpoint_key_retains_pending_and_completed_records(self):
        with tempfile.TemporaryDirectory() as d:
            store=ProgressStore(Path(d)/'progress.json');key='1'*64
            store.begin(['old-url'],[],job_key=key)
            store.job['records']['pending']={'status':'pending'};store.job['records']['done']={'status':'done'};store.flush()
            restored=ProgressStore(Path(d)/'progress.json');restored.begin(['changed-url'],[],job_key=key)
            self.assertEqual(restored.job_key,key)
            self.assertEqual(restored.record('pending')['status'],'pending')
            self.assertEqual(restored.record('done')['status'],'done')
