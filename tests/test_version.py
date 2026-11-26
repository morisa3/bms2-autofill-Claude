"""版本號檔案不可為空，且與版本紀錄、操作說明一致（4.5.001.0 曾因 version.py 被清空而無法啟動）。"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class VersionTests(unittest.TestCase):
    def test_version_matches_docs(self):
        from bridgefill.version import VERSION
        self.assertRegex(VERSION, r'^4\.\d+\.\d{3}\.0$')
        log = (ROOT / 'docs' / '版本紀錄.md').read_text(encoding='utf-8')
        self.assertEqual(re.search(r'^\| (\d+\.\d+\.\d{3}\.0) \|', log, re.M)[1], VERSION)
        guide = (ROOT / 'docs' / '操作說明.md').read_text(encoding='utf-8')
        self.assertIn(f'適用版本：{VERSION}', guide)

    def test_main_window_imports(self):
        from bridgefill.ui import main_window
        self.assertTrue(main_window.APP_RELEASE)
