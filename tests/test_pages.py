"""網頁解析測試。

SyntheticPageTests 使用自製的簡化網頁，涵蓋各種表格特殊情況，可公開。
RealPageTests 使用使用者存下的實際網頁（fixtures/private，不上傳），沒有檔案時略過。
"""
import unittest
from decimal import Decimal
from pathlib import Path

from bridgefill.site import pages as P

LINK = ('<a href="https://bms2.iot.gov.tw/bms3/inspect/inspect/der_list/component_form,{form}/component_id,{cid}/'
        'inspect_id,9/inventory_id,8/from,result/view,/" class="popup">詳細資料({n})</a>')


def main_page(tables, comment='', extra=''):
    return f"""<html><body>
<table><tr><th>橋梁名稱</th><td>測試橋</td><th>檢測日期</th><td><input id="date" value="2026-11-05"></td></tr>
<tr><th>道路等級</th><td>鄉道</td><th>總橋孔數</th><td>3</td></tr></table>
<input id="chief" value="主管"><select id="salt_injury"><option>是</option><option selected>否</option></select>
<select name="frontside"><option selected>里程樁號較小端</option></select>
<textarea id="comment">{comment}</textarea>{extra}{tables}</body></html>"""


WITH_STATUS = """<span>橋台</span><table id="abutment"><tbody>
<tr><th>狀態</th><th>橋台位置</th><th>編號</th><th>DERU</th><th>動作</th><th>維修狀態</th></tr>
<tr><td><img src="yes.jpg"></td><td>橋頭</td><td>A001</td><td>3 3 2 2</td><td>""" + LINK.format(form='abutment', cid=11, n=2) + """</td><td>部分未維修</td></tr>
<tr><td><img src="yes.jpg"></td><td>橋尾</td><td>A002</td><td>1 _ _ _</td><td>""" + LINK.format(form='abutment', cid=12, n=1) + """</td><td></td></tr>
</tbody></table>"""

WITHOUT_STATUS = """<span>引道路堤</span><table id="approach_bank"><tbody>
<tr><th>引道路堤位置</th><th>回填料質</th><th>DERU</th><th>動作</th></tr>
<tr><td>橋頭</td><td>土方</td><td>2 2 1 1</td><td>""" + LINK.format(form='approach_bank', cid=21, n=1) + """</td></tr>
<tr><td>橋尾</td><td>土方</td><td>_ 0 _ _</td><td>""" + LINK.format(form='approach_bank', cid=22, n=1) + """</td></tr>
</tbody></table>"""

ROWSPAN_TYPO = """<span>伸縮縫</span><table id="expansion"><tbody>
<tr><th>位置</th><th>編號</th><th>DREU</th><th>動作</th></tr>
<tr><td rowspan="2">橋台(A)</td><td>001</td><td>2 4 1 1</td><td>""" + LINK.format(form='expansion', cid=31, n=1) + """</td></tr>
<tr><td>002</td><td>2 4 2 2</td><td>""" + LINK.format(form='expansion', cid=32, n=1) + """</td></tr>
</tbody></table>"""

CABLES = """<span id="x_stay_cable">鋼纜(吊索)-斜張鋼纜</span><table id="stay_cable"><tbody>
<tr><th>編號</th><th>DERU</th><th>動作</th></tr>
<tr><td>C01</td><td>2 2 2 2</td><td>""" + LINK.format(form='stay_cable', cid=41, n=1) + """</td></tr></tbody></table>
<span id="x_stay_cable_cover">鋼纜保護套管-斜張鋼纜</span><table id="stay_cable"><tbody>
<tr><th>編號</th><th>DERU</th><th>動作</th></tr>
<tr><td>C01</td><td>1 _ _ _</td><td>""" + LINK.format(form='stay_cable', cid=42, n=1) + """</td></tr></tbody></table>"""

DEFECT_ROW = """<tr><td><div><span class="der_h">D:</span><span class="der_v">{d}</span><span class="der_h">E:</span><span class="der_v">{e}</span>
<span class="der_h">R:</span><span class="der_v">2</span><span class="der_h">U:</span><span class="der_v">2</span>
<span class="der_h">劣化類型：</span><span class="der_i">{kind}</span></div>
<div><span class="der_h">建議維修工法ID：</span><span class="der_v">{mid}</span></div></td>
<td><a href="https://x/inspect/damage/view_damage/der_id,{der}/inspect_id,9/_popup,1/">檢視</a>
<a href="https://x/inspect/damage/edit_damage/der_id,{der}/inspect_id,9/_popup,1/">編輯</a>
<a href="https://x/inspect/der/delete/{der}/">刪除</a>{repaired}</td></tr>"""


def defect_list(*rows):
    return '<div id="main"><table><tbody>' + ''.join(
        DEFECT_ROW.format(d=d, e=e, kind=k, mid=m, der=der, repaired='<a href="https://x/repair/">已維修</a>' if r else '')
        for d, e, k, m, der, r in rows) + '</tbody></table></div>'


class SyntheticPageTests(unittest.TestCase):
    def test_header_and_columns_follow_header_text_with_or_without_status_column(self):
        page = P.parse_inspection(main_page(WITH_STATUS + WITHOUT_STATUS, comment='梁下淨高\nP1=7.5m\nP2=m'))
        self.assertEqual(page.header['橋梁名稱'], '測試橋')
        self.assertEqual(page.header['檢測日期'], '2026-11-05')
        self.assertEqual(page.header['是否疑似鹽害橋梁'], '否')
        self.assertEqual(page.header['近端定義'], '里程樁號較小端')
        rows = {r.keys['component_id']: r for r in page.rows}
        self.assertEqual(rows['11'].columns, {'橋台位置': '橋頭', '編號': 'A001'})
        self.assertEqual((rows['11'].deru, rows['11'].status, rows['11'].count, rows['11'].repair),
                         ('3 3 2 2', P.DEFECT, 2, '部分未維修'))
        self.assertEqual(rows['12'].status, P.GOOD)
        self.assertEqual(rows['21'].columns, {'引道路堤位置': '橋頭', '回填料質': '土方'})
        self.assertEqual(rows['22'].status, P.UNABLE)
        self.assertEqual(rows['11'].identity, ('abutment', '11'))

    def test_rowspan_and_dreu_typo_keep_columns_aligned(self):
        page = P.parse_inspection(main_page(ROWSPAN_TYPO))
        second = [r for r in page.rows if r.keys['component_id'] == '32'][0]
        self.assertEqual(second.columns, {'位置': '橋台(A)', '編號': '002'})
        self.assertEqual(second.deru, '2 4 2 2')

    def test_shared_cable_table_uses_section_heading(self):
        page = P.parse_inspection(main_page(CABLES))
        self.assertEqual([r.component for r in page.rows], ['鋼纜(吊索)-斜張鋼纜', '鋼纜保護套管-斜張鋼纜'])

    def test_defect_list_reads_each_defect_and_repair_mark(self):
        defects = P.parse_defect_list(defect_list(('3', '3', '混凝土剝落', '1001', '501', True),
                                                  ('2', '3', '基礎沖刷', '1002', '502', False)))
        self.assertEqual([(d.deru, d.deterioration, d.method_id, d.der_id, d.repaired) for d in defects],
                         [('3322', '混凝土剝落', '1001', '501', True), ('2322', '基礎沖刷', '1002', '502', False)])

    def test_count_check_cases(self):
        page = P.parse_inspection(main_page(WITH_STATUS))
        defect_row, good_row = page.rows
        two = P.parse_defect_list(defect_list(('3', '3', 'a', '1', '1', False), ('2', '2', 'b', '2', '2', False)))
        self.assertTrue(P.check_count(defect_row, two).ok)
        missing = P.check_count(defect_row, two[:1])
        self.assertFalse(missing.ok)
        self.assertIn('只讀到 1 筆', missing.message)
        dup = P.check_count(defect_row, [two[0], two[0]])
        self.assertIn('重複 ID：1', dup.message)
        self.assertTrue(P.check_count(good_row, []).ok)  # D=1 的 (1) 是照片紀錄

    def test_defect_form_reads_fields_hidden_ids_and_options(self):
        form = P.parse_defect_form("""<form id="form1"><input type="hidden" name="id" value="2297410">
<input type="hidden" name="der_id" value="6386654"><input type="hidden" name="component_form" value="abutment">
<select id="d"><option>1</option><option selected>2</option></select>
<select id="deterioration"><option value="">請選擇</option><option selected>其他損傷</option><option>混凝土結構裂縫</option></select>
<input id="d_location" value="A1-R"><input id="amount" value="12">
<select id="unit"><option selected>平方公尺</option><option>公尺</option></select>
<textarea id="remark">A1-R 混凝土剝落。(2m*6m)</textarea></form>""")
        self.assertEqual((form.method_id, form.der_id, form.keys['component_form']), ('2297410', '6386654', 'abutment'))
        self.assertEqual(form.fields['d'], '2')
        self.assertEqual(form.fields['deterioration'], '其他損傷')
        self.assertEqual(form.fields['remark'], 'A1-R 混凝土剝落。(2m*6m)')
        self.assertEqual(form.options['deterioration'], ['請選擇', '其他損傷', '混凝土結構裂縫'])

    def test_distribution_and_clearance(self):
        self.assertEqual(P.parse_distribution('25+1@42.5+2@40+30'),
                         [Decimal('25'), Decimal('42.5'), Decimal('40'), Decimal('40'), Decimal('30')])
        with self.assertRaises(ValueError):
            P.parse_distribution('25+abc')
        self.assertEqual(P.parse_clearance('單跨H=7.2m'), {'H': Decimal('7.2')})
        self.assertEqual(P.parse_clearance('P1=7.5m, P2=m'), {'P1': Decimal('7.5'), 'P2': None})
        self.assertEqual(P.parse_clearance('梁下淨高P1=7.5m，P2=7.1m'), {'P1': Decimal('7.5'), 'P2': Decimal('7.1')})
        self.assertEqual(P.parse_clearance('S1GH=3m'), {})


PRIVATE = Path(__file__).resolve().parent.parent / 'fixtures' / 'private' / 'pages'


@unittest.skipUnless(PRIVATE.is_dir(), '沒有私人網頁檔')
class RealPageTests(unittest.TestCase):
    def read(self, name):
        return (PRIVATE / name).read_text(encoding='utf-8')

    def test_longmen_main_page(self):
        page = P.parse_inspection(self.read('龍門大橋-缺失主頁面.html'))
        self.assertEqual((page.header['橋梁名稱'], page.header['檢測日期']), ('龍門大橋', '2026-11-05'))
        self.assertEqual(len(page.rows), 121)
        self.assertEqual(page.problems, [])
        guard = [r for r in page.rows if r.component == '橋護欄'][0]
        self.assertEqual(guard.columns['橋護欄高度'], '1.2')
        clearance = P.parse_clearance(page.header['檢測員意見'])
        self.assertEqual((clearance['P6'], clearance['P1']), (Decimal('10.25'), None))

    def test_shenmu_prior_page_has_three_repair_states(self):
        page = P.parse_inspection(self.read('神木橋-缺失主頁面(有維修、部分未維修、與已維修紀錄).html'))
        states = {r.repair for r in page.rows}
        self.assertTrue({'已維修', '部分未維修', '未維修'} <= states)

    def test_first_level_matches_main_page_count(self):
        longmen = P.parse_inspection(self.read('龍門大橋-缺失主頁面.html'))
        row = [r for r in longmen.rows if r.identity == ('approach_protect', '16261')][0]
        defects = P.parse_defect_list(self.read('第一級缺失頁面-多筆缺失.html'))
        self.assertTrue(P.check_count(row, defects).ok)
        shenmu = P.parse_inspection(self.read('神木橋-缺失主頁面(有維修、部分未維修、與已維修紀錄).html'))
        row = [r for r in shenmu.rows if r.identity == ('abutment_fund', '31890')][0]
        defects = P.parse_defect_list(self.read('第一級缺失頁面-多筆缺失(部分未維修紀錄).html'))
        self.assertEqual(row.repair, '部分未維修')
        self.assertEqual([d.repaired for d in defects], [True, False])
        self.assertTrue(P.check_count(row, defects).ok)

    def test_form_basic_spans_and_girders(self):
        form = P.parse_defect_form(self.read('第二級缺失頁面-填寫頁.html'))
        self.assertEqual((form.method_id, form.fields['remark']), ('2297410', 'A1-R 混凝土剝落。(2m*6m)'))
        basic = P.parse_basic_data(self.read('編輯橋梁基本資料頁.html'))
        self.assertEqual((basic.values['width_min'], basic.values['height_below'], len(basic.spans)), ('8', '8.5', 13))
        spans = P.parse_spans(self.read('點到橋孔資料表.html'))
        self.assertEqual([(s.code, s.length, s.full_width) for s in spans[:2]],
                         [('S001', Decimal('25'), Decimal('9.1')), ('S002', Decimal('42.5'), Decimal('9.1'))])
        self.assertEqual(len(spans), len(basic.spans))
        self.assertEqual(P.parse_girders(self.read('神木橋-主梁網頁.html')), ['S1G1', 'S1G2', 'S1G3'])


if __name__ == '__main__':
    unittest.main()
