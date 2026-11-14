import unittest
from pathlib import Path

from bridgefill.site.scan import (Stopped, extended_ids, is_read_only_ajax, is_read_only_url, needs_detail,
                                  scan_basic, scan_inspection, scan_prior)
from bridgefill.site.pages import parse_inspection
from test_pages import LINK, VIEW_PAGE, WITH_STATUS, WITHOUT_STATUS, defect_list, history_page, main_page

BASE = 'https://bms2.iot.gov.tw/bms3'


def form(method_id, remark):
    return f"""<form id="form1"><input type="hidden" name="id" value="{method_id}">
<select id="d"><option selected>3</option></select><input id="d_location" value="A1">
<textarea id="remark">{remark}</textarea></form>"""


class FakeSite:
    """以網址片段對應網頁；第二級網址同時含 der_id 與 component_id，所以先比對 der_id。"""

    def __init__(self, pages):
        self.pages = sorted(pages.items(), key=lambda kv: not kv[0].startswith('der_id'))
        self.requests = []

    def __call__(self, url):
        self.requests.append(url)
        for key, page in self.pages:
            if key in url:
                return page
        return '<html></html>'


class ReadOnlyTests(unittest.TestCase):
    def test_only_listed_read_pages_are_allowed(self):
        allowed = [f'{BASE}/inspect/inspect/der_list/component_form,abutment/component_id,1/inspect_id,2/',
                   f'{BASE}/inspect/damage/edit_damage/der_id,1/inspect_id,2/_popup,1/',
                   f'{BASE}/inspect/damage/view_damage/der_id,1/inspect_id,2/_popup,1/',
                   f'{BASE}/basic/inventory/edit/15952/',
                   f'{BASE}/basic/inventory/edit/31858/pre_action,/from,/',
                   f'{BASE}/basic/extended/list/inventory_id,15952/',
                   f'{BASE}/inspect/inspect/der_content/inspect_id,130942/inventory_id,15153/from,result/',
                   f'{BASE}/inspect/inspect/result/inventory_id,15153/from,inventory/']
        refused = [f'{BASE}/inspect/der/delete/6386975/del_from,der_list/',
                   f'{BASE}/inspect/damage/update_damage/_popup,1/',
                   f'{BASE}/basic/inventory/add/',
                   f'{BASE}/basic/inventory/edit/15952/save/',
                   'https://example.com/other', '']
        self.assertTrue(all(is_read_only_url(u) for u in allowed))
        self.assertFalse(any(is_read_only_url(u) for u in refused))
        self.assertTrue(is_read_only_ajax('baisc/span/list/extended_id,16644/member_action,/'))
        self.assertTrue(is_read_only_ajax('basic/girder/list/span_id,69841/extended_id,16644/member_action,/'))
        self.assertFalse(is_read_only_ajax('basic/girder/delete/span_id,1/'))
        self.assertFalse(is_read_only_ajax("baisc/span/list/extended_id,1/');alert(1);//"))

    def test_extended_ids(self):
        page = ("location='/bms3/basic/member/list/extended_id,16644/' "
                "location='/bms3/basic/member/list/extended_id,16700/' location='/bms3/basic/member/list/extended_id,16644/'")
        self.assertEqual(extended_ids(page), ['16644', '16700'])


class ScanTests(unittest.TestCase):
    def site(self):
        return FakeSite({
            'component_id,11/': defect_list(('3', '3', '混凝土剝落', '1001', '501', True),
                                             ('2', '3', '基礎沖刷', '1002', '502', False)),
            'component_id,21/': defect_list(('2', '2', '路堤損傷', '2001', '601', False)),
            'der_id,501/': form('1001', 'A1 混凝土剝落。(1m*1m)'),
            'der_id,502/': form('1002', 'A1 基礎裸露。(1處)'),
            'der_id,601/': form('2001', 'A1 AC龜裂。(8m*5m)'),
        })

    def test_scan_reads_lists_and_forms_and_skips_good_rows(self):
        site = self.site()
        result = scan_inspection(main_page(WITH_STATUS + WITHOUT_STATUS), site)
        rows = {r.keys['component_id']: r for r in result.rows}
        self.assertEqual([d['fields']['remark'] for d in rows['11'].defects],
                         ['A1 混凝土剝落。(1m*1m)', 'A1 基礎裸露。(1處)'])
        self.assertEqual([d['repaired'] for d in rows['11'].defects], [True, False])
        self.assertTrue(rows['11'].ok and rows['21'].ok)
        self.assertEqual(rows['12'].check, '良好，未開啟')
        self.assertEqual(rows['22'].check, '無法檢測，未開啟')
        self.assertFalse(any('component_id,12/' in u or 'component_id,22/' in u for u in site.requests))
        self.assertTrue(all(is_read_only_url(u) for u in site.requests))
        self.assertEqual(result.summary, {'構件列數': 4, '缺失筆數': 3, '核對異常列數': 0, '網頁問題': 0})

    def test_without_forms_reads_first_level_only(self):
        site = self.site()
        scan_inspection(main_page(WITH_STATUS + WITHOUT_STATUS), site, read_forms=False)
        self.assertFalse(any('edit_damage' in u for u in site.requests))

    def test_mismatched_form_id_is_reported(self):
        site = self.site()
        site.pages.insert(0, ('der_id,601/', form('9999', 'x')))
        rows = {r.keys['component_id']: r for r in scan_inspection(main_page(WITHOUT_STATUS), site).rows}
        self.assertFalse(rows['21'].ok)
        self.assertIn('9999', rows['21'].check)

    def test_unreadable_form_is_reported(self):
        site = self.site()
        site.pages.insert(0, ('der_id,601/', '<html>逾時</html>'))
        rows = {r.keys['component_id']: r for r in scan_inspection(main_page(WITHOUT_STATUS), site).rows}
        self.assertFalse(rows['21'].ok)
        self.assertIn('第二級表單讀取失敗', rows['21'].check)
        self.assertEqual(rows['21'].defects[0]['fields'], {})

    def test_one_failed_page_does_not_stop_the_scan(self):
        site = self.site()

        def flaky(url):
            if 'component_id,11/' in url:
                raise TimeoutError('開啟逾時')
            return site(url)
        rows = {r.keys['component_id']: r for r in scan_inspection(main_page(WITH_STATUS + WITHOUT_STATUS), flaky).rows}
        self.assertFalse(rows['11'].ok)
        self.assertIn('第一級讀取失敗：開啟逾時', rows['11'].check)
        self.assertTrue(rows['21'].ok)
        self.assertEqual(rows['21'].defects[0]['fields']['remark'], 'A1 AC龜裂。(8m*5m)')

    def test_prior_scan_finds_previous_inspection_and_reads_view_pages(self):
        current = main_page(WITHOUT_STATUS)  # 連結帶 inspect_id,9／inventory_id,8
        prior_main = main_page(WITH_STATUS).replace('2026-11-05', '2024-12-05')
        site = FakeSite({'/inspect/inspect/result/inventory_id,8/': history_page(('9', '2026-11-05'), ('7', '2024-12-05')),
                         'der_content/inspect_id,7/': prior_main,
                         'component_id,11/': defect_list(('3', '3', '混凝土剝落', '1001', '501', True),
                                                          ('2', '3', '基礎沖刷', '1002', '502', False)),
                         'view_damage/der_id,501/': VIEW_PAGE,
                         'view_damage/der_id,502/': VIEW_PAGE})
        prior, result = scan_prior(current, site)
        self.assertEqual((prior.inspect_id, result.header['檢測日期']), ('7', '2024-12-05'))
        row = [r for r in result.rows if r.keys['component_id'] == '11'][0]
        self.assertTrue(row.ok, row.check)
        self.assertEqual(row.defects[0]['fields']['method'], '持續追蹤')
        self.assertFalse(any('edit_damage' in u for u in site.requests))   # 前期只讀檢視頁
        self.assertTrue(all(is_read_only_url(u) for u in site.requests))

    def test_prior_scan_without_older_inspection(self):
        site = FakeSite({'/inspect/inspect/result/': history_page(('9', '2026-11-05'))})
        with self.assertRaises(ValueError):
            scan_prior(main_page(WITHOUT_STATUS), site)

    def test_good_row_with_several_records_is_still_opened(self):
        page = main_page('<span>橋台</span><table id="abutment"><tr><th>編號</th><th>DERU</th><th>動作</th></tr>'
                         '<tr><td>A1</td><td>1 _ _ _</td><td>' + LINK.format(form='abutment', cid=11, n=2) +
                         '</td></tr></table>')
        row = parse_inspection(page).rows[0]
        self.assertTrue(needs_detail(row))

    def test_component_filter_and_stop(self):
        site = self.site()
        result = scan_inspection(main_page(WITH_STATUS + WITHOUT_STATUS), site, components={'引道路堤'})
        self.assertEqual({r.component for r in result.rows}, {'引道路堤'})
        with self.assertRaises(Stopped):
            scan_inspection(main_page(WITH_STATUS), site, stopped=lambda: True)


class BasicScanTests(unittest.TestCase):
    def test_units_mismatch_and_failures_are_reported_without_stopping(self):
        pages = {'/basic/inventory/edit/8/': '<input id="width_min" value="11.6"><input id="distribution" '
                                             'value="2040M+1@35M+2@45M+1@20"><input id="height_below" value="3.35">',
                 '/basic/extended/list/': "location='/bms3/basic/member/list/extended_id,5/'"}
        fetch = lambda url: next(v for k, v in pages.items() if k in url)
        spans = ('<table><tr><th>橋孔編號</th><th>跨距(M)</th><th>全寬(M)</th></tr>' +
                 ''.join(f'<tr><td>S00{i}</td><td>30</td><td>12</td></tr>' for i in range(1, 7)) + '</table>')

        def call(page, action, ready):
            if 'span/list' in action:
                return spans
            raise TimeoutError('逾時')
        data = scan_basic('8', fetch, call, expected_spans=6)
        self.assertEqual(data['values']['width_min'], '11.6')
        self.assertEqual(data['distribution'], ['2040', '35', '45', '45', '20'])
        self.assertEqual(len(data['structures'][0]['spans']), 6)
        self.assertIn('孔數不一致：跨距分配 5、橋孔資料表 6、總橋孔數 6；跨距以橋孔資料表為準', data['problems'])

    def test_basic_page_failure_does_not_stop_span_reading(self):
        def fetch(url):
            if 'inventory/edit' in url:
                raise TimeoutError('逾時')
            return "location='/bms3/basic/member/list/extended_id,5/'"
        data = scan_basic('8', fetch, lambda *a: '<table></table>')
        self.assertIn('基本資料頁讀取失敗：逾時', data['problems'])
        self.assertEqual(data['structures'], [{'extended_id': '5', 'spans': []}])


PRIVATE = Path(__file__).resolve().parent.parent / 'fixtures' / 'private' / 'pages'


@unittest.skipUnless(PRIVATE.is_dir(), '沒有私人網頁檔')
class RealScanTests(unittest.TestCase):
    def test_longmen_protect_row_reads_three_defects_and_form(self):
        read = lambda name: (PRIVATE / name).read_text(encoding='utf-8')
        site = FakeSite({'component_id,16261/': read('第一級缺失頁面-多筆缺失.html'),
                         'der_id,6386654/': read('第二級缺失頁面-填寫頁.html')})
        result = scan_inspection(read('龍門大橋-缺失主頁面.html'), site, components={'引道道路堤保護措施'})
        row = [r for r in result.rows if r.keys['component_id'] == '16261'][0]
        self.assertEqual([d['method_id'] for d in row.defects], ['2297410', '2297412', '2297627'])
        self.assertEqual(row.defects[0]['fields']['remark'], 'A1-R 混凝土剝落。(2m*6m)')
        # 只存了第一筆的第二級網頁，另外兩筆讀不到，必須標為異常
        self.assertFalse(row.ok)
        self.assertIn('2297412', row.check)
        self.assertIn('2297627', row.check)


if __name__ == '__main__':
    unittest.main()
