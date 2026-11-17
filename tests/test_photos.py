"""橋梁照片簿：六格名稱、可選日期、拖入與交換、照片簿列表讀取，以及在本機模擬網站上實際上傳與防重傳。"""
import os
import sys
import tempfile
import threading
import unittest
from datetime import date
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from bridgefill.core.bridge_photos import (LAND, WATER, PhotoPlan, add_photo_url, date_choices, exif_date,
                                           parse_photo_list, slot_names)
from bridgefill.site.pages import parse_basic_data

JPEG = b'\xff\xd8\xff\xe1\x00\x20Exif\x00\x00MM2026:11:03 09:12:30\x00\xff\xd9'


class CoreTests(unittest.TestCase):
    def test_slot_names_follow_river_cross(self):
        self.assertEqual(slot_names('是'), WATER)
        self.assertEqual(slot_names('否'), LAND)
        self.assertEqual(slot_names(''), WATER)
        self.assertEqual(LAND, ('橋頭近端', '橋尾遠端', '右側側照', '左側側照', '往外右側', '往外左側'))

    def test_river_cross_from_basic_page(self):
        select = '<select id="river_cross" name="river_cross"><option value=""></option><option selected>否</option><option>是</option></select>'
        self.assertEqual(parse_basic_data(f'<html><body>{select}</body></html>').values['river_cross'], '否')
        hidden = '<input name="river_cross" type="hidden" value="是"> 是'
        self.assertEqual(parse_basic_data(f'<html><body>{hidden}</body></html>').values['river_cross'], '是')

    def test_date_choices(self):
        choices = date_choices('2026-11-04', ['2026:11:03'], today=date(2026, 11, 17), around=1)
        self.assertEqual(choices, ['2026-11-03', '2026-11-04', '2026-11-05', '2026-11-17'])

    def test_exif_date_and_place_swap(self):
        with tempfile.TemporaryDirectory() as folder:
            photo = Path(folder) / 'a.jpg'
            photo.write_bytes(JPEG)
            other = Path(folder) / 'b.png'
            other.write_bytes(b'png')
            self.assertEqual(exif_date(photo), '2026-11-03')
            plan = PhotoPlan.create('否')
            rest = plan.place([photo, other, Path(folder) / 'note.txt'], '2026-11-04')
            self.assertEqual(rest, [])
            self.assertEqual([s.path for s in plan.slots[:2]], [str(photo), str(other)])
            self.assertEqual(plan.slots[0].date, '2026-11-03')        # EXIF
            self.assertEqual(plan.slots[1].date, '2026-11-04')        # 沒有 EXIF：檢測日期
            plan.swap(0, 3)
            self.assertEqual((plan.slots[0].path, plan.slots[3].path), ('', str(photo)))
            plan.place([other], '2026-11-04', start=3)                # 單張拖到有照片的格子：取代
            self.assertEqual(plan.slots[3].path, str(other))
            # 切換成跨水橋：依位置沿用照片
            water = PhotoPlan.create('是', plan)
            self.assertEqual([s.title for s in water.slots], list(WATER))
            self.assertEqual(water.slots[3].path, str(other))
            self.assertEqual(PhotoPlan.from_dict(water.to_dict()).slots[3].path, str(other))

    def test_parse_list_and_add_url(self):
        page = Site.list_page([{'date': '2026-11-04', 'title': '往外右側', 'note': ''}])
        self.assertEqual(parse_photo_list(page), [{'date': '2026-11-04', 'title': '往外右側', 'note': ''}])
        self.assertEqual(add_photo_url(page),
                         'https://bms2.iot.gov.tw/bms3/basic/inventory_photo/add/inventory_id,8/_popup,1/_close,2/')

    def test_real_photo_book(self):
        path = Path(__file__).resolve().parents[1] / 'fixtures' / 'private' / 'pages' / '產業一號橋-照片簿.html'
        if not path.exists():
            self.skipTest('沒有實際網頁')
        from bridgefill.core.bridge_photos import current_page, same_period
        page = path.read_text(encoding='utf-8')
        photos = parse_photo_list(page)
        self.assertEqual(len(photos), 20)                        # 第 1 頁 20 張（新到舊，含往年）
        self.assertEqual(photos[0], {'date': '2026-11-03', 'title': '橋頭近端', 'note': ''})
        self.assertEqual(current_page(page), 1)
        self.assertTrue(add_photo_url(page).endswith('/inventory_photo/add/inventory_id,31858/_popup,1/_close,2/'))
        self.assertTrue(same_period(photos, '橋頭近端', '2026-11-20'))    # 本期已上傳過
        self.assertFalse(same_period(photos, '橋頭近端', '2027-11-20'))   # 往年的不算


class Site(BaseHTTPRequestHandler):
    photos, posts, save, page = [], [], True, 1

    @staticmethod
    def list_page(photos):
        rows = ''.join(f'<tr><td><img src="x.jpg"></td><td>{p["date"]}</td><td>{p["title"]}</td><td>{p["note"]}</td>'
                       f'<td><input type="button" value="修改" onclick="location=\'/bms3/basic/inventory_photo/edit/1/\'">'
                       f'<input type="button" value="刪除" onclick="location=\'/bms3/basic/inventory_photo/delete/1/\'"></td></tr>'
                       for p in photos)
        # 比照實際網頁（2026-11-17 產業一號橋照片簿）：新增圖片以 easy_popup 開彈出頁、表頭在 thead 用 td、有分頁
        return f"""<html><body><input type="button" value="返回橋梁基本資料" onclick="location='/bms3/basic/inventory/edit/8/'">
<input type="button" name="button1" value="新增圖片" onclick="easy_popup('/bms3/basic/inventory_photo/add/inventory_id,8/_popup,1/_close,2/','1030,600');">
<table><tr><td><table id="main"><thead><tr><td>照片</td><td>拍照日期</td><td>照片標題</td><td>照片說明</td><td>&nbsp;</td></tr></thead>
<tbody>{rows}</tbody></table></td></tr></table>
<ul class="pagination"><li class="page-item active"><span class="current page-link">{Site.page}</span></li>
<li><a onclick="return setPage('2');">2</a></li></ul></body></html>"""

    def do_GET(self):
        if '/inventory_photo/list/' in self.path:
            body = self.list_page(Site.photos)
        elif '/inventory_photo/add/' in self.path:
            body = """<html><body><form name="form1" id="form1" method="post" enctype="multipart/form-data"
action="/bms3/basic/inventory_photo/insert/inventory_id,8/" onsubmit="return confirm('確定新增？');">
<table><tr><th>*照片</th><td><input type="file" name="photo"></td></tr>
<tr><th>拍照日期</th><td><input type="text" name="photo_date" readonly></td></tr>
<tr><th>照片標題</th><td><input type="text" name="subject"></td></tr>
<tr><th>照片說明</th><td><textarea name="memo"></textarea></td></tr></table>
<input type="button" value="回列表"><div><input type="submit" name="submit" value="確定送出"></div></form></body></html>"""
        else:
            body = 'not found'
        self.send_response(200 if body != 'not found' else 404)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(body.encode('utf-8'))

    def do_POST(self):
        data = self.rfile.read(int(self.headers['Content-Length']))
        message = BytesParser(policy=default).parsebytes(
            f'Content-Type: {self.headers["Content-Type"]}\r\n\r\n'.encode() + data)
        fields = {}
        for part in message.iter_parts():
            name = part.get_param('name', header='content-disposition')
            payload = part.get_payload(decode=True) or b''
            fields[name] = (part.get_filename(), payload) if part.get_filename() is not None else payload.decode('utf-8')
        Site.posts.append((self.path, fields))
        if Site.save:
            Site.photos.append({'date': fields.get('photo_date', ''), 'title': fields.get('subject', ''),
                                'note': fields.get('memo', '')})
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write('<script>alert("新增完成")</script>'.encode('utf-8'))

    def log_message(self, *args):
        pass


LIST_URL = 'https://bms2.iot.gov.tw/bms3/basic/inventory_photo/list/inventory_id,8/'


class UploadTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
        if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
            os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')
        try:
            from PySide6.QtWidgets import QApplication
            from bridgefill.ui.browser import make_profile
            from bridgefill.ui.filler import SiteFiller
        except ImportError as error:
            raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')
        cls.app = QApplication.instance() or QApplication([])
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), Site)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'
        cls.folder = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        cls.profile = make_profile(Path(cls.folder.name) / 'profile')
        cls.filler = SiteFiller(cls.profile, delay_ms=0, timeout_ms=8000,
                                rewrite=lambda url: url.replace('https://bms2.iot.gov.tw', cls.base))
        cls.photo = Path(cls.folder.name) / '橋頭.jpg'
        cls.photo.write_bytes(JPEG)

    @classmethod
    def tearDownClass(cls):
        cls.filler.release()
        cls.server.shutdown()

    def setUp(self):
        from bridgefill.core.fill import FillLog
        Site.photos, Site.posts, Site.save, Site.page = [], [], True, 1
        self.log = FillLog(Path(self.folder.name) / f'{self._testMethodName}.json')
        self.plan = PhotoPlan.create('否')
        self.plan.place([self.photo], '2026-11-04')
        self.plan.slots[0].note = '說明文字'

    def test_upload_and_verify(self):
        from bridgefill.core.fill import DONE
        slot = self.plan.slots[0]
        self.assertEqual(self.filler.upload_photo(LIST_URL, slot, self.log), '已上傳')
        self.assertEqual(len(Site.posts), 1)
        path, fields = Site.posts[0]
        self.assertIn('/inventory_photo/insert/', path)
        self.assertEqual((fields['subject'], fields['photo_date'], fields['memo']), ('橋頭近端', '2026-11-03', '說明文字'))
        self.assertEqual(fields['photo'], ('橋頭.jpg', JPEG))
        self.assertEqual(self.log.get('照片:橋頭近端')['status'], DONE)
        self.assertFalse(any('delete' in u or 'edit' in u for u in self.filler.page.submitted))
        # 再上傳一次：照片簿已有，不重傳
        self.assertEqual(self.filler.upload_photo(LIST_URL, slot, self.log), '網站已有')
        self.assertEqual(len(Site.posts), 1)

    def test_unclear_result_is_never_reuploaded(self):
        from bridgefill.core.fill import PENDING
        from bridgefill.ui.filler import FillError
        Site.save = False
        slot = self.plan.slots[0]
        with self.assertRaisesRegex(FillError, '沒有出現'):
            self.filler.upload_photo(LIST_URL, slot, self.log)
        self.assertEqual(self.log.get('照片:橋頭近端')['status'], PENDING)
        with self.assertRaisesRegex(FillError, '不會自動重傳'):
            self.filler.upload_photo(LIST_URL, slot, self.log)
        self.assertEqual(len(Site.posts), 1)

    def test_same_title_other_date(self):
        from bridgefill.ui.filler import FillError
        # 往年同標題的照片：照常上傳
        Site.photos = [{'date': '2025-12-04', 'title': '橋頭近端', 'note': ''}]
        self.assertEqual(self.filler.upload_photo(LIST_URL, self.plan.slots[0], self.log), '已上傳')
        # 本期附近已有同標題（日期不同）：停止，不重複上傳也不刪除
        Site.photos, Site.posts = [{'date': '2026-11-04', 'title': '橋尾遠端', 'note': ''}], []
        slot = self.plan.slots[1]
        slot.path, slot.date = str(self.photo), '2026-11-05'
        with self.assertRaisesRegex(FillError, '本期已上傳過'):
            self.filler.upload_photo(LIST_URL, slot, self.log)
        self.assertEqual(Site.posts, [])

    def test_not_first_page_stops(self):
        from bridgefill.ui.filler import FillError
        Site.page = 2
        with self.assertRaisesRegex(FillError, '第 1 頁'):
            self.filler.upload_photo(LIST_URL, self.plan.slots[0], self.log)
        self.assertEqual(Site.posts, [])

if __name__ == '__main__':
    unittest.main()
