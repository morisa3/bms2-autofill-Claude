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
from urllib.parse import parse_qs

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


MAIN_URL = 'https://bms2.iot.gov.tw/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/'


class Site(BaseHTTPRequestHandler):
    state, posts, method_id, save = dict(START), [], '9001', True
    comment = '1. 舊意見\n橋下淨高：H=3m'
    basic = {}

    def main(self):
        b = {'chief': '', 'frontside': '', 'salt_injury': '', 'scour_type': '', 'scour_site': '', 'waste': '', **Site.basic}
        return f"""<html><body><form name="form1" method="post" action="/bms3/inspect/inspect/update/"
onsubmit="return true;"><input name="chief" id="chief" value="{b['chief']}">
{select('frontside', b['frontside'], ['', '里程樁號較小端', '竣工圖上之橋頭端', '其他']).replace(' id="frontside"', '')}
{select('salt_injury', b['salt_injury'], ['', '是', '否', '無法判斷'])}
{select('scour_type', b['scour_type'], ['', '有', '無', '無法判斷'])}
<textarea name="scour_site" id="scour_site">{b['scour_site']}</textarea>
{select('waste', b['waste'], ['', '否', '是'])}
<textarea name="comment" id="comment">{Site.comment}</textarea>
<input type="submit" name="submit" value="送出修改內容"></form>
<form method="post" action="/bms3/inspect/inspect/upload_other_photo/"><input type="submit" value="確定送出"></form>
</body></html>"""

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
        body = (self.page() if re.search(r'edit_damage/der_id,501/', self.path) else
                self.main() if 'der_content/inspect_id,9/' in self.path else 'not found')
        self.send_response(200 if 'form1' in body else 404)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(body.encode('utf-8'))

    def do_POST(self):
        data = self.rfile.read(int(self.headers['Content-Length']))
        if self.path.startswith('/bms3/inspect/inspect/update/'):
            fields = {k: v[0] for k, v in parse_qs(data.decode('utf-8')).items()}
            Site.posts.append((self.path, fields))
            if Site.save:
                Site.comment = fields.get('comment', '')
                Site.basic = {k: fields.get(k, '') for k in ('chief', 'frontside', 'salt_injury', 'scour_type',
                                                             'scour_site', 'waste')}
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write('<script>alert("修改完成")</script>'.encode('utf-8'))
            return
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
        Site.comment = '1. 舊意見\n橋下淨高：H=3m'
        Site.basic = {}
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

    def test_write_opinion(self):
        new = '1. A1橋台混凝土剝落鋼筋外露，建議持續追蹤。\n橋下淨高：H=3.2m'
        self.assertEqual(self.filler.write_opinion(MAIN_URL, Site.comment, new, self.log), '已送出並回讀核對一致')
        self.assertEqual((len(Site.posts), Site.comment.replace('\r\n', '\n')), (1, new))
        self.assertIn('/inspect/inspect/update/', Site.posts[0][0])
        self.assertEqual(self.log.get('檢測員意見')['status'], DONE)
        self.filler.write_opinion(MAIN_URL, '1. 舊意見', new, self.log)   # 已是這段：不送出
        self.assertEqual(len(Site.posts), 1)

    def test_write_opinion_with_basic_info(self):
        basic = {'單位主管': '顏宏宇', '近端定義': '里程樁號較小端', '是否疑似鹽害橋梁': '否', '有無沖刷現象': '有',
                 '沖刷位置/備註': 'P3,P4', '橋下有無堆放雜物': '否'}
        empty = dict.fromkeys(basic, '')
        self.filler.write_opinion(MAIN_URL, Site.comment, '1. 新意見', self.log, basic, empty)
        self.assertEqual(Site.basic, {'chief': '顏宏宇', 'frontside': '里程樁號較小端', 'salt_injury': '否',
                                      'scour_type': '有', 'scour_site': 'P3,P4', 'waste': '否'})
        self.assertEqual(self.log.get('檢測員意見')['status'], DONE)
        # 網站上的主管在掃描後被改過：停止
        Site.basic['chief'] = '別人'
        with self.assertRaisesRegex(FillError, '單位主管在掃描後有變動'):
            self.filler.write_opinion(MAIN_URL, '1. 新意見', '1. 再改', self.log, dict(basic, 單位主管='宋昌國'),
                                      basic)

    def test_opinion_changed_on_site_is_not_overwritten(self):
        with self.assertRaisesRegex(FillError, '掃描後有變動'):
            self.filler.write_opinion(MAIN_URL, '1. 掃描時的意見', '1. 新意見', self.log)
        self.assertEqual(Site.posts, [])

    def test_opinion_unsaved_is_pending_and_not_resubmitted(self):
        Site.save = False
        with self.assertRaisesRegex(FillError, '回讀的檢測員意見不一致'):
            self.filler.write_opinion(MAIN_URL, Site.comment, '1. 新意見', self.log)
        with self.assertRaisesRegex(FillError, '不會自動重送'):
            self.filler.write_opinion(MAIN_URL, Site.comment, '1. 新意見', self.log)
        self.assertEqual(len(Site.posts), 1)


if __name__ == '__main__':
    unittest.main()
