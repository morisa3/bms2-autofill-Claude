"""內嵌瀏覽器行為：選單點一下展開、背景頁面的網站提示視窗不會卡住掃描。"""
import os
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')

try:
    from PySide6.QtCore import QCoreApplication, QEvent, QEventLoop, QTimer
    from PySide6.QtWidgets import QApplication
    from bridgefill.ui.browser import SitePage, SiteView, make_profile
    from bridgefill.ui.fetcher import PageFetcher
except ImportError as error:
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

MENU = """<!doctype html><meta charset="utf-8"><style>
#menu li ul { display: none; } #menu li:hover > ul { display: block; }</style>
<ul id="menu">
 <li><a href="javascript:return false;">基本資料</a><ul><li><a href="/search/">資料搜尋</a></li></ul></li>
 <li><a href="javascript:return false;">檢測資料</a><ul>
   <li><a href="javascript:return false;">定期檢測</a><ul><li><a href="/inspect/search/">進階查詢</a></li></ul></li>
 </ul></li>
</ul><p id="outside">內容</p>"""
DIALOGS = """<!doctype html><meta charset="utf-8"><script>
alert('連線逾時');
document.addEventListener('DOMContentLoaded', () => { document.body.dataset.answer = String(confirm('確定刪除？')); });
</script><body>頁面</body>"""

TRICKY = ('<!doctype html><meta charset="utf-8"><body><p>備註：A1 剝落。&lt;/textarea&gt; 😀 &amp; 中</p>'
          '<textarea>原本就有的文字框</textarea></body>')


class Site(BaseHTTPRequestHandler):
    def do_GET(self):
        body = {'/menu': MENU, '/bms3/inspect/inspect/der_list/x,1/': DIALOGS,
                '/bms3/inspect/inspect/der_list/y,1/': TRICKY}.get(self.path, 'x')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(body.encode('utf-8'))

    def log_message(self, *args):
        pass


class ConsolePage(SitePage):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.console = []

    def javaScriptConsoleMessage(self, level, message, line, source):
        self.console.append(message)


class BrowserBehaviorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), Site)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'
        cls.folder = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        cls.profile = make_profile(Path(cls.folder.name) / 'profile')

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()

    def wait(self, ms):
        loop = QEventLoop()
        QTimer.singleShot(ms, loop.quit)
        loop.exec()

    def js(self, page, script):
        box, loop = {}, QEventLoop()
        page.runJavaScript(script, 0, lambda value: (box.setdefault('v', value), loop.quit()))
        QTimer.singleShot(5000, loop.quit)
        loop.exec()
        return box.get('v')

    def test_top_menu_opens_on_click_without_script_error(self):
        view = SiteView(self.profile)
        page = ConsolePage(self.profile, view)
        view.setPage(page)
        view.resize(1000, 600)
        view.show()
        view.open(self.base + '/menu')
        self.wait(1500)
        shown = "getComputedStyle(document.querySelector({!r})).display"
        top, nested = '#menu > li:nth-child(2) > ul', '#menu > li:nth-child(2) > ul > li > ul'
        self.assertEqual(self.js(page, shown.format(top)), 'none')
        self.js(page, "document.querySelector('#menu > li:nth-child(2) > a').click()")
        self.wait(300)
        self.assertEqual(self.js(page, shown.format(top)), 'block')
        self.js(page, "document.querySelector('#menu > li:nth-child(2) > ul > li > a').click()")
        self.wait(300)
        self.assertEqual((self.js(page, shown.format(top)), self.js(page, shown.format(nested))), ('block', 'block'))
        self.js(page, "document.querySelector('#menu > li:nth-child(1) > a').click()")
        self.wait(300)
        self.assertEqual(self.js(page, shown.format(top)), 'none')          # 開另一個選單時收起前一個
        self.js(page, "document.querySelector('#outside').click()")
        self.wait(300)
        self.assertEqual(self.js(page, shown.format('#menu > li:nth-child(1) > ul')), 'none')
        self.assertFalse([m for m in page.console if 'Illegal return' in m], page.console)
        self.assertTrue(self.js(page, "location.pathname") == '/menu')      # 沒有跳頁
        view.release()

    def test_hidden_page_dismisses_alert_and_cancels_confirm(self):
        fetcher = PageFetcher(self.profile, delay_ms=0, timeout_ms=10000,
                              rewrite=lambda url: url.replace('https://bms2.iot.gov.tw', self.base))
        html = fetcher.get('https://bms2.iot.gov.tw/bms3/inspect/inspect/der_list/x,1/')
        self.assertIn('data-answer="false"', html)
        self.assertEqual(fetcher.page.dialogs, [('提示', '連線逾時'), ('確認（已取消）', '確定刪除？')])
        fetcher.release()


    def test_direct_load_keeps_content_with_relative_links(self):
        fetcher = PageFetcher(self.profile, delay_ms=0, timeout_ms=10000,
                              rewrite=lambda url: url.replace('https://bms2.iot.gov.tw', self.base))
        html = fetcher.get('inspect/inspect/der_list/y,1/')        # 實際網頁的寫法：沒有斜線，以 <base> 為基準
        self.assertIn('備註：A1 剝落。&lt;/textarea&gt; 😀 &amp; 中', html)
        self.assertEqual(fetcher.requests, ['https://bms2.iot.gov.tw/bms3/inspect/inspect/der_list/y,1/'])
        fetcher.release()

if __name__ == '__main__':
    unittest.main()
