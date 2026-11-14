"""唯讀掃描視窗的端對端測試：本機模擬網站（自製簡化網頁），確認讀取內容、筆數核對與不會開啟禁止的網址。"""
import json
import os
import re
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
    from bridgefill.ui.scan_test import ScanTestWindow
except ImportError as error:
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

from bridgefill.site.scan import is_read_only_url
from test_pages import VIEW_PAGE, WITH_STATUS, WITHOUT_STATUS, defect_list, history_page, main_page
from test_scan import form

MEMBER_PAGE = """<html><body><div id="main">各構件</div><script>
function xajax_set_action_ajax(action) {
  fetch('/ajax/' + action).then(r => r.text()).then(t => {
    setTimeout(() => { document.getElementById('main').innerHTML = t; }, 150);
  });
}</script></body></html>"""
SPANS = """<table><tr><th>細部內容</th><th>橋孔編號</th><th>結構型式</th><th>車道數</th><th>跨距(M)</th><th>全寬(M)</th></tr>
<tr><td><a href="#">橋面板(1)</a><a href="javascript:xajax_set_action_ajax('basic/girder/list/span_id,71/extended_id,5/member_action,/')">主梁(1)</a></td>
<td>S001</td><td>梁式橋</td><td>2</td><td>25</td><td>9.1</td></tr>
<tr><td><a href="#">橋面板(1)</a><a href="javascript:xajax_set_action_ajax('basic/girder/list/span_id,72/extended_id,5/member_action,/')">主梁(1)</a></td>
<td>S002</td><td>梁式橋</td><td>2</td><td>30</td><td>9.1</td></tr></table>"""


def girders(*codes):
    rows = ''.join(f'<tr><td>{c}</td><td>I型梁</td></tr>' for c in codes)
    return f'<table id="main"><tr><th>主梁編號</th><th>主梁型式</th></tr>{rows}</table>'


class MockSite(BaseHTTPRequestHandler):
    """模擬網站。第一級清單只在帶有來源頁（Referer）時回傳內容，模擬網站以 iframe 載入彈出視窗的情況。"""
    pages = {}
    requests = []
    referers = {}

    def do_GET(self):
        MockSite.requests.append(self.path)
        referer = self.headers.get('Referer', '')
        MockSite.referers[self.path] = referer
        body = next((page for key, page in MockSite.pages.items() if re.search(key, self.path)), None)
        if '/der_list/' in self.path and not referer:
            body = '<html><body>請由系統操作</body></html>'
        self.send_response(200 if body is not None else 404)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write((body or 'not found').encode('utf-8'))

    def log_message(self, *args):
        pass


class ScanWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        # 對話框在無畫面環境會一直等人按；出現就代表流程有問題，直接判定失敗
        import bridgefill.ui.scan_test as screen

        def no_dialog(_parent, title, text, *args):
            raise AssertionError(f'不應出現對話框：{title}：{text}')
        cls._information = screen.QMessageBox.information
        screen.QMessageBox.information = staticmethod(no_dialog)
        main = main_page(WITH_STATUS + WITHOUT_STATUS)
        MockSite.pages = {
            r'^/bms3/inspect/inspect/der_content/inspect_id,9/': main,
            r'component_id,11/': defect_list(('3', '3', '混凝土剝落', '1001', '501', True),
                                             ('2', '3', '基礎沖刷', '1002', '502', False)),
            r'component_id,21/': defect_list(('2', '2', '路堤損傷', '2001', '601', False)),
            r'edit_damage/der_id,501/': form('1001', 'A1 混凝土剝落。(1m*1m)'),
            r'edit_damage/der_id,502/': form('1002', 'A1 基礎裸露。(1處)'),
            r'edit_damage/der_id,601/': form('2001', 'A1 AC龜裂。(8m*5m)'),
            r'^/bms3/basic/inventory/edit/8/$': '<input id="width_min" value="8"><input id="width_max" value="8">'
                                                '<input id="distribution" value="25+30"><input id="height_below" value="7.5">',
            r'^/bms3/basic/extended/list/': "<input onclick=\"location='/bms3/basic/member/list/extended_id,5/'\">",
            r'^/bms3/basic/member/list/extended_id,5/$': MEMBER_PAGE,
            r'^/ajax/baisc/span/list/': SPANS,
            r'^/ajax/basic/girder/list/span_id,71/': girders('S1G1', 'S1G2', 'S1G3'),
            r'^/ajax/basic/girder/list/span_id,72/': girders('S2G1', 'S2G2'),
            r'^/bms3/inspect/inspect/result/inventory_id,8/': history_page(('9', '2026-11-05'), ('7', '2024-12-05')),
            r'der_content/inspect_id,7/': main_page(WITH_STATUS).replace('2026-11-05', '2024-12-05'),
            r'view_damage/der_id,50[12]/': VIEW_PAGE,
        }
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), MockSite)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'

    @classmethod
    def tearDownClass(cls):
        import bridgefill.ui.scan_test as screen
        screen.QMessageBox.information = cls._information
        cls.server.shutdown()

    def wait(self, ms):
        loop = QEventLoop()
        QTimer.singleShot(ms, loop.quit)
        loop.exec()

    def test_scan_and_basic_scan_read_only(self):
        MockSite.requests.clear()
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            rewrite = lambda url: url.replace('https://bms2.iot.gov.tw', self.base)
            window = ScanTestWindow(f'{self.base}/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/', Path(folder) / 'profile',
                                    Path(folder) / 'scan', delay_ms=0, rewrite=rewrite)
            window.show()
            self.wait(1500)
            self.assertEqual(window.view.url().toString(), f'{self.base}/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/')  # 開在指定頁
            window.scan(True)
            table = [[window.table.item(r, c).text() for c in range(window.table.columnCount())]
                     for r in range(window.table.rowCount())]
            window.scan_basic()
            window.scan(True, prior=True)
            log = window.output.toPlainText()
            results = {p.name: json.loads(p.read_text(encoding='utf-8')) for p in (Path(folder) / 'scan').glob('*.json')}
            fetched = list(window.fetcher.requests)
            window.close()
            window.deleteLater()
            QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
            self.wait(300)

        self.assertEqual(len(table), 4)
        by_place = {row[1]: row for row in table}
        self.assertEqual(by_place['橋台位置：橋頭｜編號：A001'][4:], ['2', '2', '✔'])
        self.assertEqual(by_place['橋台位置：橋尾｜編號：A002'][5:], ['—', '良好，未開啟'])
        inspection = results['2026-11-05-測試橋-檢測.json']
        self.assertEqual(inspection['summary'], {'構件列數': 4, '缺失筆數': 3, '核對異常列數': 0, '網頁問題': 0})
        remarks = [d['fields']['remark'] for r in inspection['rows'] for d in r['defects']]
        self.assertEqual(remarks, ['A1 混凝土剝落。(1m*1m)', 'A1 基礎裸露。(1處)', 'A1 AC龜裂。(8m*5m)'])
        basic = results['2026-11-05-測試橋-基本資料.json']
        self.assertEqual(basic['values']['height_below'], '7.5')
        self.assertEqual([s['girders'] for s in basic['structures'][0]['spans']],
                         [['S1G1', 'S1G2', 'S1G3'], ['S2G1', 'S2G2']])
        self.assertIn('主梁根數 [3, 2]', log)
        prior = results['2024-12-05-測試橋-前期.json']
        self.assertEqual(prior['summary']['核對異常列數'], 0)
        self.assertIn('前期：2024-12-05（檢測系統 ID 7', log)
        # 第一級清單都以缺失主頁面為來源頁（本期掃描用本期主頁面，前期掃描用前期主頁面），且沒有讀取異常需要存檔
        der_lists = [p for p in MockSite.referers if '/der_list/' in p]
        self.assertTrue(der_lists)
        self.assertTrue(all(re.search(r'/inspect/inspect/der_content/inspect_id,[79]/', MockSite.referers[p])
                            for p in der_lists), MockSite.referers)
        self.assertTrue(re.search(r'der_content/inspect_id,7/', MockSite.referers[
            next(p for p in der_lists if 'component_id,11/' in p)]))   # 最後一次（前期）來自前期主頁面
        # 模擬資料：主頁面總橋孔數 3，但跨距分配與橋孔資料表都只有 2 孔，應提示並只存基本資料的診斷檔
        self.assertIn('孔數不一致：跨距分配 2、橋孔資料表 2、總橋孔數 3', log)
        dumps = [line for line in log.splitlines() if '讀取異常的頁面' in line]
        self.assertEqual(len(dumps), 1)
        self.assertIn('-基本資料-', dumps[0])
        # 只開啟唯讀網址：所有隱藏頁面的讀取都在白名單，網站伺服器沒收到任何刪除／更新請求
        self.assertTrue(all(is_read_only_url(u.split('#')[0]) for u in fetched))
        self.assertFalse(any(re.search(r'delete|update|add', p) for p in MockSite.requests))


    def test_unreadable_rows_are_saved_for_diagnosis(self):
        saved = dict(MockSite.pages)
        MockSite.pages[r'component_id,21/'] = '<html><body>沒有資料</body></html>'
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                rewrite = lambda url: url.replace('https://bms2.iot.gov.tw', self.base)
                window = ScanTestWindow(f'{self.base}/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/', Path(folder) / 'profile',
                                        Path(folder) / 'scan', delay_ms=0, rewrite=rewrite)
                window.show()
                self.wait(1500)
                window.scan(False)
                indexes = list((Path(folder) / 'scan' / 'debug').glob('*/index.txt'))
                index = indexes[0].read_text(encoding='utf-8') if indexes else ''
                dumped = sorted(p.name for p in indexes[0].parent.glob('*.html')) if indexes else []
                window.close()
                window.deleteLater()
                QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
                self.wait(300)
        finally:
            MockSite.pages = saved
        self.assertEqual(len(indexes), 1)
        self.assertIn('component_id,21/', index)
        self.assertNotIn('component_id,11/', index)    # 正常的列不存
        self.assertEqual(dumped, ['01.html'])


if __name__ == '__main__':
    unittest.main()
