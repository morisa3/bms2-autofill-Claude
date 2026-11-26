"""Google 試算表同步（core.google_sync）：Apps Script 呼叫、302 轉導、防覆寫與無損匯出 38 欄 Excel。"""
import http.server
import json
import socketserver
import tempfile
import threading
import unittest
from pathlib import Path

from openpyxl import load_workbook

from bridgefill.core import bridge_list as bl
from bridgefill.core.google_sync import GoogleSheetClient, GoogleSyncError, build_sync_payload, build_xlsx_from_cloud
from bridgefill.core.opinion import Entry

HEADERS_38 = [h.replace('{year}', '115') for h, _ in bl.STANDARD_LAYOUT]


class MockAppsScript(http.server.BaseHTTPRequestHandler):
    token = 'secret-115'
    received = []
    export_data = None

    def log_message(self, *args):
        pass

    def do_POST(self):
        length = int(self.headers.get('Content-Length') or 0)
        body = json.loads(self.rfile.read(length).decode('utf-8'))
        self.__class__.received.append(body)
        if body.get('token') != self.__class__.token:
            out = {'ok': False, 'error': '金鑰驗證失敗'}
        elif body.get('action') == 'ping':
            out = {'ok': True, 'title': '115年北雲林自檢橋梁總表', 'sheets': ['西螺', '莿桐', '二崙']}
        elif body.get('action') == 'list_bridges':
            out = {
                'ok': True, 'sheets': ['西螺'],
                'bridges': [
                    {'sheet': '西螺', 'row': 2, 'serial': '1', 'id': '37996', 'code': '14-004CS',
                     'name': '行度橋', 'done': '', 'access_path': 'D\nA1下游徒步下橋', 'access_aux': ''},
                    {'sheet': '西螺', 'row': 3, 'serial': '2', 'id': '9535', 'code': '14-005CS',
                     'name': '新社二號橋', 'done': '2026-11-25', 'access_path': '', 'access_aux': ''},
                ],
            }
        elif body.get('action') == 'sync_bridge':
            out = {
                'ok': True, 'sheet': '西螺', 'row': 2,
                'report': [f"西螺 第 2 列：{body.get('name', '')}", '前期觀測劣化(U=2)：已有人工內容，保留不覆寫'],
            }
        elif body.get('action') == 'export_all':
            out = self.__class__.export_data or {'ok': True, 'title': '總表', 'sheets': []}
        else:
            out = {'ok': False, 'error': 'unknown'}
        raw = json.dumps(out, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)


class GoogleSyncTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.httpd = socketserver.TCPServer(('127.0.0.1', 0), MockAppsScript)
        cls.url = f'http://127.0.0.1:{cls.httpd.server_address[1]}/exec'
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()

    def setUp(self):
        MockAppsScript.received.clear()

    def test_ping_auth_and_list_bridges(self):
        client = GoogleSheetClient(self.url, 'secret-115')
        pong = client.ping()
        self.assertEqual(pong['title'], '115年北雲林自檢橋梁總表')
        self.assertEqual(pong['sheets'], ['西螺', '莿桐', '二崙'])
        listing = client.list_bridges()
        self.assertEqual(len(listing['bridges']), 2)
        self.assertEqual(listing['bridges'][0]['id'], '37996')

        bad = GoogleSheetClient(self.url, 'wrong')
        with self.assertRaises(GoogleSyncError):
            bad.ping()

    def test_sync_bridge_payload_and_response(self):
        sources = bl.Sources(
            inventory_id='37996',
            header={'橋梁編號': '14-004CS', '橋梁名稱': '行度橋', '檢測日期': '2026-11-25', '橋下有無堆放雜物': '否'},
            basic={'route': '農雲螺003', 'road_lv': '農路', 'spans': '1', 'total_length': '6', 'width_max': '26'},
            prior_items=[(Entry('1', '橋台', {'remark': 'A1 混凝土蜂窩。(1處)', 'u': '2'}, 2, 2, 2), 113)],
            new_items=[],
            railing='0.8',
            photos_done=True,
            done=True,
        )
        payload = build_sync_payload(sources)
        self.assertEqual(payload['inventory_id'], '37996')
        self.assertEqual(payload['values']['guard'], '是')
        self.assertEqual(payload['values']['photos'], '是')
        self.assertIn('height_prior', payload['column_keys'])
        self.assertEqual(payload['column_keys']['height_prior'], '前期梁下淨高(m)')

        client = GoogleSheetClient(self.url, 'secret-115')
        report = client.sync_bridge(sources)
        self.assertTrue(any('行度橋' in line for line in report))
        self.assertTrue(any('保留不覆寫' in line for line in report))

    def test_export_to_xlsx_preserves_38_columns_and_25_piers(self):
        piers_25 = '\n'.join(f'P{i:02d}={3 + i * 0.1:.2f}' for i in range(1, 26))
        row1 = [''] * 38
        row1[0], row1[1], row1[3], row1[6] = '1', '31316', '媽祖大橋', '26'
        row1[20], row1[23], row1[24], row1[25] = piers_25, '1.45', '0.25', '1.30'
        row1[30], row1[33] = '否', '是'
        fills = ['#ffffff'] * 38
        fills[3] = '#b7e1cd'
        notes = [''] * 38
        notes[33] = '填報助手產生 2026-11-25｜識別碼 12345678'
        MockAppsScript.export_data = {
            'ok': True, 'title': '115年北雲林自檢橋梁總表',
            'sheets': [{'name': '虎尾溪', 'headers': HEADERS_38, 'rows': [row1], 'fills': [fills], 'notes': [notes]}],
        }
        client = GoogleSheetClient(self.url, 'secret-115')
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder) / 'export.xlsx'
            summary = client.export_to_xlsx(out, year=115)
            self.assertEqual(summary, ['虎尾溪：匯出 1 座'])
            book = load_workbook(out)
            self.assertEqual(book.sheetnames, ['說明', '虎尾溪'])
            ws = book['虎尾溪']
            self.assertEqual(ws.max_column, 38)
            self.assertEqual(ws.freeze_panes, 'E2')
            self.assertEqual(ws.cell(2, 4).value, '媽祖大橋')
            self.assertEqual(ws.cell(2, 4).fill.fgColor.rgb, 'FFB7E1CD')
            self.assertEqual(ws.cell(2, 4).font.name, '標楷體')
            self.assertIn('填報助手產生', ws.cell(2, 34).comment.text)
            rtb_lines = str(ws.cell(2, 23).value).splitlines()
            self.assertEqual(len(rtb_lines), 25)
            self.assertEqual(rtb_lines[0], 'P01=6.10')
            self.assertEqual(rtb_lines[-1], 'P25=8.50')


if __name__ == '__main__':
    unittest.main()
