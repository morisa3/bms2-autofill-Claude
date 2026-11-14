from pathlib import Path
import tempfile
import unittest
from bridgefill.core.data import History
from bridgefill.core.policy import automatic_plan
from bridgefill.core.progress import ProgressStore


class PolicyTests(unittest.TestCase):
    def setUp(self):
        self.current = {'deterioration': '混凝土剝落', 'd_location': 'A1', 'amount': '2', 'unit': '平方公尺',
                        'method': '修復', 'price': '500', 'remark': 'A1 混凝土剝落。(1m＊2m)'}
        self.rules = [{'pattern': '混凝土剝落', 'method': '修復', 'unit': '平方公尺', 'price': '500'}]

    def test_normalization_and_dimension_quantity(self):
        result = automatic_plan(self.current, '橋護欄', [], self.rules)
        self.assertFalse(result.review)
        self.assertEqual(result.fields['remark'], 'A1 混凝土剝落。(1m*2m)')
        self.assertEqual(result.fields['price'], '500')

    def test_missing_excel_pauses_even_when_nothing_to_change(self):
        self.assertTrue(automatic_plan(self.current, '橋護欄', [], []).review)

    def test_unique_previous_only_fills_blank_and_preserves_der(self):
        old = History('橋護欄', dict(self.current), 3)
        current = dict(self.current, remark='', method='')
        result = automatic_plan(current, '橋護欄', [old], self.rules)
        self.assertFalse(result.review)
        self.assertEqual(result.fields['method'], '修復')
        self.assertEqual(result.fields['remark'], 'A1 混凝土剝落。(1m*2m)')

    def test_ambiguous_previous_is_not_selected(self):
        old = History('橋護欄', dict(self.current), 3)
        self.assertTrue(automatic_plan(dict(self.current, remark=''), '橋護欄', [old, old], self.rules).review)

    def test_ambiguous_excel_needs_manual_choice(self):
        rules = self.rules + [dict(self.rules[0], price='800')]
        self.assertTrue(automatic_plan(dict(self.current, price=''), '橋護欄', [], rules).review)


class ProgressTests(unittest.TestCase):
    def test_restart_rules_change_and_reset_preserve_pending(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'progress.json'
            store = ProgressStore(path)
            store.begin('inspect1', 'rules1')
            store.mark('record1', 'done', 'hash1')
            store.mark('record2', 'pending', 'hash2')
            store.mark_row('row1')
            new = ProgressStore(path)
            new.begin('inspect1', 'rules1')
            self.assertTrue(new.done_row('row1'))
            self.assertEqual(new.record('record1')['status'], 'done')
            new.reset_completed()
            self.assertFalse(new.done_row('row1'))
            self.assertFalse(new.record('record1'))
            self.assertEqual(new.record('record2')['status'], 'pending')
            new.mark('record3', 'done', 'hash3')
            new.begin('inspect1', 'rules2')
            self.assertFalse(new.record('record3'))
            self.assertEqual(new.record('record2')['status'], 'pending')
            new.begin('inspect2', 'rules2')
            self.assertFalse(new.record('record2'))

    def test_corrupt_progress_is_not_silently_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'progress.json'
            path.write_text('broken', encoding='utf-8')
            with self.assertRaises(ValueError):
                ProgressStore(path)
            self.assertEqual(path.read_text(), 'broken')

if __name__ == '__main__':
    unittest.main()
