from pathlib import Path
import tempfile
import unittest
from bridgefill.core.progress import ProgressStore


class ManualQueueTests(unittest.TestCase):
    def test_queue_survives_restart_rules_change_and_job_isolation(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'progress.json'
            store = ProgressStore(path)
            store.begin('bridge1','rules1')
            store.mark_row('row')
            store.defer('record','fingerprint','row','policy_review')
            self.assertFalse(store.done_row('row'))
            loaded = ProgressStore(path)
            loaded.begin('bridge1','rules2')
            self.assertEqual(len(loaded.manual_records()), 1)
            loaded.reset_completed()
            self.assertEqual(len(loaded.manual_records()), 1)
            loaded.retry_skipped()
            self.assertEqual(len(loaded.manual_records()), 1)
            loaded.begin('bridge2','rules1')
            self.assertFalse(loaded.manual_records())
            loaded.begin('bridge1','rules2')
            self.assertEqual(loaded.record('record')['reason'], 'policy_review')
            self.assertNotIn('form', path.read_text())

    def test_pending_fingerprint_is_preserved_and_manual_flag_survives_submit(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ProgressStore(Path(directory) / 'progress.json')
            store.begin('bridge','rules')
            store.mark('record','pending','previous-submission')
            store.defer('record','different-current-form','row','pending_mismatch')
            self.assertEqual(store.record('record')['status'], 'pending')
            self.assertEqual(store.record('record')['fingerprint'], 'previous-submission')
            store.mark('record','pending','manual-approved-submission')
            self.assertEqual(len(store.manual_records()), 1)
            self.assertEqual(store.record('record')['fingerprint'], 'manual-approved-submission')
            store.mark('record','done','verified')
            self.assertFalse(store.manual_records())

    def test_completing_record_removes_queue_and_keeps_other_deferred_records(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ProgressStore(Path(directory) / 'progress.json')
            store.begin('bridge','rules')
            store.defer('first','a','row','policy_review')
            store.defer('second','b','row','policy_review')
            store.mark('first','done','verified')
            self.assertEqual(set(store.manual_records()), {'second'})

    def test_auto_pending_is_atomically_queued_until_verification(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'progress.json'
            store = ProgressStore(path)
            store.begin('bridge','rules')
            store.mark('record','pending','submitted', manual_row='row')
            loaded = ProgressStore(path)
            loaded.begin('bridge','rules')
            self.assertEqual(loaded.record('record')['status'], 'pending')
            self.assertEqual(loaded.manual_records()['record']['row_key'], 'row')
            loaded.mark('record','done','verified')
            self.assertFalse(loaded.manual_records())
