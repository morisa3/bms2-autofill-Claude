"""主視窗端對端測試：在模擬網站上掃描並配對，審查、存檔、接續。"""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')

try:
    from PySide6.QtCore import QCoreApplication, QEvent, QEventLoop, QTimer
    from PySide6.QtWidgets import QApplication
    import bridgefill.ui.main_window as screen
except ImportError as error:
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

from bridgefill.core.review import ACCEPTED, CHANGED, KIND_GONE, PENDING
from test_scan_window import MockSite, start_mock_site

MAIN = '/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/'


class MainWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

        def no_dialog(_parent, title, text, *args):
            raise AssertionError(f'不應出現對話框：{title}：{text}')
        cls._dialogs = (screen.QMessageBox.information, screen.QMessageBox.warning)
        screen.QMessageBox.information = screen.QMessageBox.warning = staticmethod(no_dialog)
        cls.server, cls.base = start_mock_site()

    @classmethod
    def tearDownClass(cls):
        screen.QMessageBox.information, screen.QMessageBox.warning = cls._dialogs
        cls.server.shutdown()

    def wait(self, ms):
        loop = QEventLoop()
        QTimer.singleShot(ms, loop.quit)
        loop.exec()

    def open_window(self, folder):
        rewrite = lambda url: url.replace('https://bms2.iot.gov.tw', self.base)
        window = screen.MainWindow(self.base + MAIN, Path(folder) / 'profile', Path(folder) / 'scan',
                                   Path(folder) / 'review', delay_ms=0, rewrite=rewrite)
        window.show()
        self.wait(1500)
        return window

    def close(self, window):
        window.close()
        window.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        self.wait(300)

    def test_scan_match_review_save_and_resume(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            pairs = list(window.pairs)
            self.assertTrue(pairs)
            self.assertEqual(window.table.rowCount(), len(pairs))
            self.assertTrue((Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json').exists())
            self.assertTrue((Path(folder) / 'scan' / '2024-12-05-測試橋-前期.json').exists())
            review_file = window.review_file
            self.assertTrue(review_file.exists())
            self.assertIn('共', window.summary.text())

            # 篩選：待審查
            window.filter.setCurrentText(PENDING)
            pending = list(window.visible)
            self.assertTrue(all(window.review.decision(p).status == PENDING for p in pending))

            # 逐筆審查：開第一筆待審查，改判並寫說明
            window.filter.setCurrentText('全部')
            target = next(p for p in window.pairs if p.prior and not p.current)    # 只有前期的一筆
            window.open_card(window.visible.index(target))
            self.assertIs(window.tabs.currentWidget(), window.card)
            self.assertIn(KIND_GONE, [window.card.kind.itemText(i) for i in range(window.card.kind.count())])
            window.card.kind.setCurrentText(KIND_GONE)
            window.card.note.setText('河水沖刷')
            window.card.decided.emit(window.card.kind.currentText(), window.card.note.text())
            saved = json.loads(review_file.read_text(encoding='utf-8'))
            self.assertEqual(saved['decisions'][target.key]['kind'], KIND_GONE)
            self.assertEqual(saved['decisions'][target.key]['status'], CHANGED)
            self.close(window)

            # 重新開啟、讀取上次掃描檔：審查結果接續
            window = self.open_window(folder)
            window.load_saved(Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json')
            again = next(p for p in window.pairs if p.key == target.key)
            self.assertEqual((window.review.decision(again).kind, window.review.decision(again).note),
                             (KIND_GONE, '河水沖刷'))
            window.filter.setCurrentText(CHANGED)
            self.assertEqual([p.key for p in window.visible], [target.key])
            self.close(window)

    def test_site_panel_can_be_hidden(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.toggle_site()
            self.assertFalse(window.view.isVisible())
            self.assertEqual(window.buttons['site'].text(), '顯示網站')
            window.toggle_site()
            self.assertTrue(window.view.isVisible())
            self.close(window)


if __name__ == '__main__':
    unittest.main()
