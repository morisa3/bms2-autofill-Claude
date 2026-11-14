"""內嵌瀏覽器無畫面測試：以自製網頁模擬登入頁與登入後頁面，不連線、不含實際資料。"""
import os
import subprocess
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import unittest
from pathlib import Path

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')

try:
    from PySide6.QtCore import QCoreApplication, QEvent, QEventLoop, QTimer, QUrl
    from PySide6.QtWidgets import QApplication
    from bridgefill.ui.login_test import LoginTestWindow
except ImportError as error:  # 沒有 QtWebEngine 或系統圖形函式庫時略過
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

LOGIN_HTML = """<!doctype html><meta charset="utf-8"><title>登入</title>
<form action="https://bms2.iot.gov.tw/bms3/public/public/login/" method="post">
<input name="account"><input type="password" name="password">
<div class="g-recaptcha"></div><div class="captcha_box"><input id="captcha_code"></div>
</form>"""
HOME_HTML = """<!doctype html><meta charset="utf-8"><title>首頁</title>
<ul id="menu"><li><a href="#">基本資料</a><ul><li><a href="#">資料搜尋</a></li></ul></li>
<li><a href="#">檢測資料</a></li></ul>"""


class LoginWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def run_page(self, html, folder):
        page = Path(folder) / 'page.html'
        page.write_text(html, encoding='utf-8')
        window = LoginTestWindow(QUrl.fromLocalFile(str(page)).toString(), Path(folder) / 'profile', Path(folder))
        loop = QEventLoop()
        timer = QTimer(singleShot=True)
        timer.timeout.connect(loop.quit)
        window.view.loadFinished.connect(lambda _: QTimer.singleShot(800, loop.quit))
        timer.start(20000)
        loop.exec()
        log = window.log_path.read_text(encoding='utf-8')
        state = window.state
        window.close()
        window.deleteLater()
        # 等視窗與瀏覽器設定檔真正釋放，再讓暫存資料夾被刪除
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        wait = QEventLoop()
        QTimer.singleShot(300, wait.quit)
        wait.exec()
        return state, log

    def test_login_page_is_detected_with_both_captchas(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            state, log = self.run_page(LOGIN_HTML, folder)
        self.assertEqual(state, '登入頁')
        self.assertIn('驗證：reCAPTCHA、圖片驗證碼', log)

    def test_logged_in_page_reads_top_level_menu_only(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            state, log = self.run_page(HOME_HTML, folder)
        self.assertEqual(state, '已登入')
        self.assertIn('選單 2 項：基本資料、檢測資料', log)
        self.assertNotIn('資料搜尋', log)

    def test_log_never_contains_query_string(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            _, log = self.run_page(HOME_HTML, folder)
        self.assertNotIn('?', log)


class CookieHandler(BaseHTTPRequestHandler):
    """第一次造訪發一個 30 天的 Cookie（模擬登入）；有 Cookie 時回傳帶選單的首頁。"""

    def do_GET(self):
        logged_in = 'session=ok' in (self.headers.get('Cookie') or '')
        body = (HOME_HTML if logged_in else LOGIN_HTML).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        if not logged_in:
            self.send_header('Set-Cookie', 'session=ok; Max-Age=2592000; Path=/')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass


OPEN_ONCE = """
import sys
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication
from bridgefill.ui.login_test import LoginTestWindow
app = QApplication([])
window = LoginTestWindow(sys.argv[1], sys.argv[2], sys.argv[3])
window.view.loadFinished.connect(lambda _: QTimer.singleShot(800, window.close))
QTimer.singleShot(20000, window.close)
window.show()
app.lastWindowClosed.connect(lambda: QTimer.singleShot(1000, app.quit))
app.exec()
print(window.state)
"""


class CookiePersistenceTests(unittest.TestCase):
    def test_login_cookie_survives_restarting_the_program(self):
        """兩次各自啟動獨立程式：第二次能直接進入登入後頁面，代表 Cookie 已寫入磁碟。"""
        server = ThreadingHTTPServer(('127.0.0.1', 0), CookieHandler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        url = f'http://127.0.0.1:{server.server_port}/'
        root = Path(__file__).resolve().parent.parent
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                states = []
                for _ in range(2):
                    result = subprocess.run([sys.executable, '-c', OPEN_ONCE, url, str(Path(folder) / 'profile'), folder],
                                            cwd=root, capture_output=True, text=True, timeout=60,
                                            env={**os.environ, 'PYTHONPATH': str(root)})
                    states.append(result.stdout.strip().splitlines()[-1] if result.stdout.strip() else result.stderr[-500:])
                logs = ''.join(p.read_text(encoding='utf-8') for p in sorted(Path(folder).glob('login-test-*.log')))
        finally:
            server.shutdown()
        self.assertEqual(states, ['登入頁', '已登入'])
        self.assertIn('沿用上次登入', logs)


if __name__ == '__main__':
    unittest.main()
