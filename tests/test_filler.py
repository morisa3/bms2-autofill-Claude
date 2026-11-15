"""第 6 步：在本機模擬網站上實際填寫、送出、回讀核對；確認防錯（ID 鎖定、掃描後變動、結果不明不重送）。"""
import os
import re
import sys
import tempfile
import threading
import unittest
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')

try:
    from PySide6.QtWidgets import QApplication
    from bridgefill.ui.browser import make_profile
    from bridgefill.ui.filler import FillError, SiteFiller
except ImportError as error:
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

from bridgefill.core.fill import DONE, PENDING, FillLog, plan_item
from bridgefill.core.review import Review
from test_fill import RULE, current

START = {'d': '3', 'e': '3', 'r': '2', 'u': '2', 'deterioration': '混凝土剝落', 'd_location': 'A1', 'method': '',
         'amount': '1', 'unit': '處', 'price': '', 'remark': 'A1 混凝土剝落。(1處)'}


def select(name, value, options):
    items = ''.join(f'<option value="{o}"{" selected" if o == value else ""}>{o}</option>' for o in options)
    return f'<select name="{name}" id="{name}">{items}</select>'


class Site(BaseHTTPRequestHandler):
    state, posts, method_id, save = dict(START), [], '9001', True

    def page(self):
        s = Site.state
        return f"""<html><body><form name="form1" id="form1" method="post" enctype="multipart/form-data"
action="/bms3/inspect/damage/update_damage/_popup,1/" onsubmit="return true;">
<input type="hidden" name="id" value="{Site.method_id}"><input type="hidden" name="der_id" value="501">
{''.join(select(k, s[k], ['', '0', '1', '2', '3', '4']) for k in 'deru')}
{select('deterioration', s['deterioration'], ['', '混凝土剝落', '其他損傷'])}
<input name="d_location" id="d_location" value="{s['d_location']}">
<input name="method" id="method" value="{s['method']}"><input name="amount" id="amount" value="{s['amount']}">
{select('unit', s['unit'], ['', '處', '平方公尺'])}
<input name="price" id="price" value="{s['price']}"><textarea name="remark" id="remark">{s['remark']}</textarea>
<div align="center"><input type="submit" name="_submit" value="送出更新"></div></form>
<a href="/bms3/inspect/der/delete/1/">刪除</a></body></html>"""

    def do_GET(self):
        body = self.page() if re.search(r'edit_damage/der_id,501/', self.path) else 'not found'
        self.send_response(200 if 'form1' in body else 404)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(body.encode('utf-8'))

    def do_POST(self):
        data = self.rfile.read(int(self.headers['Content-Length']))
        message = BytesParser(policy=default).parsebytes(
            f'Content-Type: {self.headers["Content-Type"]}\r\n\r\n'.encode() + data)
        fields = {part.get_param('name', header='content-disposition'): part.get_payload(decode=True).decode('utf-8')
                  for part in message.iter_parts()}
        Site.posts.append((self.path, fields))
        if Site.save:
            Site.state.update({k: fields.get(k, '') for k in START})
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write('<script>alert("資料已更新")</script>更新完成'.encode('utf-8'))

    def log_message(self, *args):
        pass


class FillerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), Site)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'
        cls.folder = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        cls.profile = make_profile(Path(cls.folder.name) / 'profile')
        cls.filler = SiteFiller(cls.profile, delay_ms=0, timeout_ms=8000,
                                rewrite=lambda url: url.replace('https://bms2.iot.gov.tw', cls.base))

    @classmethod
    def tearDownClass(cls):
        cls.filler.release()
        cls.server.shutdown()

    def setUp(self):
        Site.state, Site.posts, Site.method_id, Site.save = dict(START), [], '9001', True
        self.log = FillLog(Path(self.folder.name) / f'{self._testMethodName}.json')

    def item(self, **edits):
        pair = current(deru='3322')
        review = Review()
        if edits:
            review.set_edit(pair, edits)
        return plan_item(pair, review, [RULE])

    def test_fill_submit_and_verify(self):
        item = self.item(u='3')
        self.assertEqual(set(item.changes), {'method', 'price', 'u'})
        self.filler.fill(item, self.log)
        self.assertEqual(len(Site.posts), 1)
        self.assertIn('update_damage', Site.posts[0][0])
        self.assertEqual((Site.state['method'], Site.state['price'], Site.state['u']), ('混凝土修復', '800', '3'))
        self.assertEqual(self.log.get(item.key)['status'], DONE)
        self.assertIn('資料已更新', self.log.get(item.key)['message'])
        self.assertFalse(self.filler.page.blocked)
        # 再填一次：網站已是預定值，不送出
        self.filler.fill(self.item(u='3'), self.log)
        self.assertEqual(len(Site.posts), 1)

    def test_site_changed_after_scan_is_not_overwritten(self):
        Site.state['remark'] = 'A1 別人改過。(1處)'
        with self.assertRaisesRegex(FillError, '掃描後有變動'):
            self.filler.fill(self.item(), self.log)
        self.assertEqual(Site.posts, [])

    def test_id_mismatch_stops(self):
        Site.method_id = '9999'
        with self.assertRaisesRegex(FillError, 'ID 不符'):
            self.filler.fill(self.item(), self.log)
        self.assertEqual(Site.posts, [])

    def test_unsaved_result_is_pending_and_never_resubmitted(self):
        Site.save = False
        item = self.item()
        with self.assertRaisesRegex(FillError, '回讀不一致'):
            self.filler.fill(item, self.log)
        self.assertEqual((len(Site.posts), self.log.get(item.key)['status']), (1, PENDING))
        with self.assertRaisesRegex(FillError, '不會自動重送'):
            self.filler.fill(self.item(), self.log)
        self.assertEqual(len(Site.posts), 1)
        # 人工在網站改好後：核對一致即記為完成，仍不送出
        Site.state.update(method='混凝土修復', price='800')
        self.filler.fill(self.item(), self.log)
        self.assertEqual((len(Site.posts), self.log.get(item.key)['status']), (1, DONE))

    def test_delete_links_and_unarmed_submit_are_blocked(self):
        item = self.item()
        self.filler.get(item.url)
        self.filler.page.blocked.clear()
        self.filler.page.runJavaScript("document.querySelector('a').click()")
        self.filler.pause(800)
        self.filler.page.runJavaScript("document.getElementById('form1').submit()")
        self.filler.pause(800)
        self.assertEqual(Site.posts, [])
        self.assertTrue(any('delete' in u for u in self.filler.page.blocked))
        self.assertTrue(any('update_damage' in u for u in self.filler.page.blocked))


if __name__ == '__main__':
    unittest.main()
