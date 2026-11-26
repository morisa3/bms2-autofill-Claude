"""橋梁清單同步（core.bridge_list，2026-11-23）：劣化依 U 分欄並標年份、維修紀錄、原構件再出現、核對與寫入 Excel。"""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from bridgefill.core import bridge_list as bl
from bridgefill.core.matching import Defect
from bridgefill.core.opinion import Entry
from bridgefill.site.pages import RepairRecord

HEADS = ['序號', '二代系統ID', '橋梁編號\n 須修正', '橋梁名稱', '道路等級', '路線', '跨數', '橋長', '最大淨寬',
         '橋孔面積\n(橋長*淨寬)', '跨越物體', '下橋路徑(ABCD)', '輔助設備(A/B)', '前期觀測劣化\n(U=1)', '前期觀測劣化\n(U=2)',
         '前期觀測劣化\n (U≧3)', '115年新增劣化\n (U=1)', '115年新增劣化\n (U=2)', '115年新增劣化\n (U≧3)',
         '過去維修紀錄\n(列出過去維修紀錄)', '已維修構件狀況', '前期梁下/橋下\n淨高(m)(113年)', '本期梁下/橋下\n淨高(m)(115年)',
         '欄杆至河床高度(m)', '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)', '基本資料變動\n (ch5.1)', '應維修構件\n (ch5.2)',
         '應維修處理建議', '河工設施損壞\n (ch5.4)', '護欄高度不足(V)\n (ch5.5)', '橋下堆積物(V)\n (ch5.6)', '疑似鹽害橋梁(V)',
         '六張基本照片拍攝', '橋名牌損壞、遺失、內容錯誤', '是否有限重牌', '其他檢測特殊情形\n ex:水深無法檢測..']


def entry(component, remark, u, location=''):
    return Entry(remark, component, {'remark': remark, 'd_location': location, 'u': str(u), 'method': '混凝土修復'},
                 2, 2, u)


def defect(component, remark, period='前期', der='1', kind=''):
    return Defect(period, ('x', der), component, '', '2222', kind, {'remark': remark, 'u': '2'}, der_id=der)


def sheet_file(folder, values=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '莿桐'
    sheet.append(HEADS)
    row = [1, 22870, '13-005CS', '莿桐N29', '村里道路', '五華村', None, 13.4, None, None, '鹿場課圳幹線', 'D', None]
    sheet.append(row + [None] * (len(HEADS) - len(row)))
    for name, value in (values or {}).items():
        sheet.cell(2, HEADS.index(name) + 1, value)
    path = Path(folder) / '清單.xlsx'
    workbook.save(path)
    return path


class LinesTests(unittest.TestCase):
    def test_grouped_by_u_with_years(self):
        lines = bl.defect_lines([(entry('橋台', 'A1 混凝土剝落。(1處)', 2), 113), (entry('橋台', 'A2 混凝土剝落。(1處)', 2), 111),
                                 (entry('橋護欄', 'S1-L 混凝土裂縫。(1m)', 1), 113),
                                 (entry('主梁', 'S1G2 混凝土剝落鋼筋外露。(1處)', 3), 113)])
        self.assertEqual(lines[2], '1.A1、A2橋台混凝土剝落(111年)')            # 同樣態合併、取最早年份
        self.assertEqual(lines[1], '1.S1護欄混凝土裂縫(113年)')
        self.assertEqual(lines[3], '1.S1主梁混凝土剝落鋼筋外露(113年)')
        self.assertEqual(bl.defect_lines([])[1], '無')

    def test_blank_remark_uses_prior_description(self):
        now = Entry('k', '橋台', {'remark': '', 'u': '3'}, 2, 2, 3)
        old = defect('橋台', 'A1 混凝土剝落。(1處)')
        self.assertEqual(bl.described(now, old).remark, 'A1 混凝土剝落。(1處)')
        self.assertEqual(bl.described(now, old).u, 3)                          # U 用本期

    def test_prior_items_use_prior_and_excel_rules(self):
        """前期觀測劣化以前期為主（前期敘述、U、年份），敘述套 Excel 規則修正；已維修的不列（2026-11-24）。"""
        from bridgefill.core.editing import load_rules
        rules = load_rules(Path(__file__).resolve().parents[1] / 'data' / '劣化說明與維修單價v3.xlsx')
        rules = rules[0] if isinstance(rules, tuple) else rules
        kind = '混凝土剝落、破碎、鋼筋外露、銹蝕'
        exposed = Defect('前期', ('x', '1'), '橋台', '', '2223', kind,
                         {'remark': 'A1-下游背牆 混凝土剝落，鋼筋外露 (1處)', 'deterioration': kind, 'unit': '處', 'u': '3'},
                         der_id='1')
        aged = Defect('前期', ('y', '2'), '引道路堤', '', '2222', '路堤損傷',
                      {'remark': 'A1 老化。(5m*7m)', 'deterioration': '路堤損傷', 'unit': '平方公尺', 'u': '2'}, der_id='2')
        repaired = Defect('前期', ('z', '3'), '橋台', '', '2222', kind, {'remark': 'A2 蜂窩已維修。(1處)', 'u': '2'}, der_id='3')
        items = bl.prior_items([exposed, aged, repaired], 113, rules)
        self.assertEqual(len(items), 2)
        lines = bl.defect_lines(items)
        self.assertEqual(lines[3], '1.A1橋台混凝土剝落鋼筋外露(113年)')         # 規則修正、U 用前期
        self.assertEqual(lines[2], '1.A1引道路堤AC剝離(113年)')                 # 樣態同義詞：老化 → AC剝離
        self.assertEqual(bl.prior_items([exposed], 113)[0][0].remark, exposed.fields['remark'])   # 沒有規則：原文


class RepairTests(unittest.TestCase):
    records = [RepairRecord('2020-07-26', '3', '翼牆／擋土牆', 'a002', '', '混凝土剝落', '混凝土修復', '', '', '2021-11-26',
                            '2021-12-02'),
               RepairRecord('2017-09-25', '3', '橋台', 'a002', '002', '其他損傷', '植物清理', '', '', '2017-12-08', '2017-12-31')]

    def test_repair_text_by_year(self):
        text = bl.repair_text(self.records, [(115, '橋台 A1 蜂窩（115年檢測標示已維修）')])
        self.assertEqual(text, '106\n1.橋台 A2 植物清理\n110\n1.翼牆／擋土牆 A2 混凝土修復\n115\n1.橋台 A1 蜂窩（115年檢測標示已維修）')
        self.assertEqual(bl.repair_text([]), '無')

    def test_reappeared_after_repair(self):
        later = [(113, defect('橋台或翼牆或擋土牆', 'A2-R 混凝土剝落。(1處)', kind='混凝土剝落、破碎、鋼筋外露、銹蝕')),
                 (113, defect('橋台', 'A1 植生。(1處)', kind='植物生長'))]
        triggers = {'2020-07-26': [defect('橋台或翼牆或擋土牆', 'A2 翼牆混凝土剝落。(1處)')]}
        text, notes = bl.reappeared(self.records, later, 113, triggers)
        self.assertIn('110年維修 113年原構件出現劣化（翼牆／擋土牆 A2：混凝土剝落）', text)
        self.assertNotIn('106年維修', text)                                   # 106 年維修的 A2 橋台沒有再出現
        self.assertTrue(any('維修前 2020-07-26 檢測' in n for n in notes))
        self.assertEqual(bl.reappeared(self.records, [], 113)[0], '良好')
        self.assertEqual(bl.reappeared([], later, 113)[0], '無')


class WriteTests(unittest.TestCase):
    def sources(self, **extra):
        values = dict(inventory_id='22870', header={'橋梁編號': '13-005CS', '橋梁名稱': '莿桐N29', '道路等級': '村里道路',
                                                    '橋下有無堆放雜物': '是', '是否疑似鹽害橋梁': '否'},
                      basic={'route': '五華村', 'spans': '3', 'total_length': '13.5', 'width_max': '5', 'crossover': '鹿場課圳幹線'},
                      prior_items=[(entry('橋台', 'A1 混凝土剝落。(1處)', 2), 113)],
                      new_items=[(entry('河道', 'A1 護岸掏空。(1處)', 3), 115)],
                      current_entries=[entry('河道', 'A1 護岸掏空。(1處)', 3)], repairs=[],
                      prior_height='P1=3.3\nP2=3.3', current_height='P1=3.5\nP2=3.3', railing='0.8')
        values.update(extra)
        return bl.Sources(**values)

    def test_write_row(self):
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder, {'前期梁下/橋下\n淨高(m)(113年)': 'P1=3.3\nP2=3.3'})
            report = bl.write(path, self.sources(done=True), Path(folder) / 'backup')
            sheet = load_workbook(path)['莿桐']
            cell = lambda name: sheet.cell(2, [c.value for c in sheet[1]].index(name) + 1)  # noqa: E731
            self.assertEqual(cell('跨數').value, 3)                                       # 空白：填入系統值
            self.assertEqual(cell('橋長').value, 13.4)                                    # 不同：不覆寫
            self.assertIn('系統值：13.5', cell('橋長').comment.text)
            self.assertEqual(cell('橋長').fill.fgColor.rgb, bl.CHECK_FILL)
            self.assertEqual(cell('前期觀測劣化\n(U=2)').value, '1.A1橋台混凝土剝落(113年)')
            self.assertEqual(cell('115年新增劣化\n (U≧3)').value, '1.A1河道護岸掏空(115年)')
            self.assertEqual(cell('115年新增劣化\n (U=1)').value, '無')
            self.assertEqual(cell('本期梁下/橋下\n淨高(m)(115年)').value, 'P1=3.5\nP2=3.3')     # 有變化才填
            self.assertEqual(cell('本期梁下/橋下\n淨高(m)(115年)').alignment.horizontal, 'center')  # 淨高置中、不靠左
            self.assertEqual(cell('應維修構件\n (ch5.2)').value, '河道')
            self.assertEqual(cell('河工設施損壞\n (ch5.4)').value, '1.A1河道護岸掏空')
            self.assertEqual(cell('護欄高度不足(V)\n (ch5.5)').value, '是')                 # 欄杆 0.8 < 1.1
            self.assertEqual(cell('橋下堆積物(V)\n (ch5.6)').value, '是')
            self.assertEqual(cell('疑似鹽害橋梁(V)').value, '否')                             # 是／否直接填
            self.assertEqual(cell('下橋路徑(ABCD)').value, 'D')                            # 手動欄不動
            self.assertRegex(cell('系統填寫完成').value, r'^\d{4}-\d{2}-\d{2}$')
            rules = [(str(r.sqref), [x.formula for x in items]) for r, items in sheet.conditional_formatting._cf_rules.items()]
            self.assertIn(('D2:D1000', [['$AM2<>""']]), rules)                            # 綠底
            self.assertTrue(any('莿桐 第 2 列' in line for line in report))
            self.assertEqual(len(list((Path(folder) / 'backup').iterdir())), 1)

    def test_same_height_not_written_and_missing_bridge(self):
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder, {'前期梁下/橋下\n淨高(m)(113年)': 'P1=3.3\nP2=3.3'})
            bl.write(path, self.sources(current_height='P1=3.3\nP2=3.3', new_items=None))
            sheet = load_workbook(path)['莿桐']
            row = dict(zip([c.value for c in sheet[1]], (c.value for c in sheet[2])))
            self.assertEqual(row['本期梁下/橋下\n淨高(m)(115年)'], 'P1=3.3\nP2=3.3')       # 2026-11-29：本期淨高一律寫入
            cell = sheet.cell(2, [c.value for c in sheet[1]].index('本期梁下/橋下\n淨高(m)(115年)') + 1)
            self.assertIsNone(cell.comment)                                               # 不加備註、不標橘底（不鎖定）
            self.assertNotEqual(str(cell.fill.fgColor.rgb).upper()[-6:], bl.CHECK_FILL[2:])
            self.assertIsNone(row['115年新增劣化\n (U=1)'])                              # 尚無本期：新增欄不動
            with self.assertRaises(bl.ListError):
                bl.write(path, self.sources(inventory_id='999', header={'橋梁名稱': '不存在的橋'}))

    def test_helpers(self):
        self.assertEqual(bl.heights_value('梁下淨高：P1=3.35m、P2=5.15m'), 'P01=3.35\nP02=5.15')
        self.assertEqual(bl.heights_value('梁下淨高：H=2.2m'), '2.20')
        opinion = '4.P1基礎裸露(H=0.45m),P2基礎裸露(H=1.1m)\n橋下淨高:\nP1=7.2m\nP2=7.7m\nP10=10.6m'
        self.assertEqual(bl.heights_value(opinion), 'P01=7.20\nP02=7.70\nP10=10.60')     # 只讀淨高之後、有 P 不讀 H
        self.assertEqual(bl.heights_value('P1基礎裸露(H=0.45m)'), '')                      # 沒有「淨高」不讀
        self.assertEqual(bl.format_height('P1:5.8m\n P2:6.8m'), 'P01=5.80\nP02=6.80')
        self.assertFalse(bl.heights_differ('P01=3.30', 'P1=3.3'))
        self.assertFalse(bl.heights_differ('P1=3.3\n P2=3.3', 'P1=3.30\nP2=3.3'))
        self.assertTrue(bl.heights_differ(2.2, '2.5'))
        self.assertEqual(bl.railing_height({'rows': [{'columns': {'護欄高度(cm)': '80'}}, {'columns': {'護欄高度(m)': '1.1'}}]}),
                         '0.8')
        self.assertEqual(bl.position_code(RepairTests.records[0]), 'A2')
        self.assertEqual(bl.roc_year('2024-04-18'), 113)


if __name__ == '__main__':
    unittest.main()


class PageTests(unittest.TestCase):
    def test_parse_repairs_and_basic_fields(self):
        from bridgefill.site.pages import parse_basic_data, parse_repairs
        from bridgefill.site.scan import is_read_only_url, repair_url
        html = ('<table><tr><th>項次</th><th>橋梁名稱</th><th>檢測日期</th><th>U</th><th>構件名稱</th><th>構件位置</th>'
                '<th>構件編號</th><th>劣化類型</th><th>維修工法</th><th>維修數量</th><th>單位</th><th>開工日期</th><th>完工日期</th>'
                '<th>動作</th></tr><tr><td>1</td><td>藍厝橋</td><td>2020-07-26</td><td>3</td><td>翼牆／擋土牆</td><td>a002</td>'
                '<td></td><td>混凝土剝落</td><td>混凝土修復</td><td></td><td></td><td>2021-11-26</td><td>2021-12-02</td>'
                '<td>檢視 修改 刪除</td></tr></table>')
        records = parse_repairs(html)
        self.assertEqual((records[0].component, records[0].method, records[0].year), ('翼牆／擋土牆', '混凝土修復', 110))
        self.assertTrue(is_read_only_url(repair_url(22792)))
        basic = parse_basic_data('<input name="route_name_type" type="radio" checked><select name="route"><option>－</option>'
                                 '</select><input name="route_text" value="投100線"><input name="road_lv" type="hidden" value="鄉道">'
                                 '<input name="crossover" id="crossover" value="南清水溝溪"><input name="spans" type="hidden" value="6">'
                                 '<input id="width_max" value="11.6">')
        self.assertEqual({k: basic.values.get(k) for k in ('route', 'road_lv', 'crossover', 'spans', 'width_max')},
                         {'route': '投100線', 'road_lv': '鄉道', 'crossover': '南清水溝溪', 'spans': '6', 'width_max': '11.6'})


class LayoutTests(unittest.TestCase):
    def test_write_widens_column_and_row(self):
        """同步寫入較長的文字：欄寬放大、列高依換行後行數調整，文字不被切掉（2026-11-24）。"""
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            book = load_workbook(path)
            book.active.column_dimensions['N'].width = 8                         # 前期觀測劣化(U=1) 原本很窄
            book.save(path)
            items = [(entry('橋台', f'A{i} 混凝土剝落鋼筋外露且有白華與滲水第{i}處。(1處)', 1), 113) for i in range(1, 4)] + \
                    [(entry('主梁', 'S1G1 混凝土剝落鋼筋外露且有白華與滲水。(1處)', 1), 113)]
            bl.write(path, WriteTests().sources(prior_items=items))
            sheet = load_workbook(path).active
            text = sheet['N2'].value
            self.assertGreater(sheet.column_dimensions['N'].width, 30)
            self.assertLessEqual(sheet.column_dimensions['N'].width, bl.MAX_WIDTH)
            lines = bl.cell_lines(text, sheet.column_dimensions['N'].width)
            self.assertGreaterEqual(sheet.row_dimensions[2].height, lines * bl.LINE_HEIGHT)
            self.assertTrue(sheet['N2'].alignment.wrap_text)
            self.assertEqual(sheet['N2'].alignment.horizontal, 'left')                  # 多筆資料靠左
            self.assertNotEqual(sheet['A2'].alignment.horizontal, 'left')               # 單一值不變


class GuardTests(unittest.TestCase):
    def test_railing_standard_by_year(self):
        """護欄高度不足：2026 年起 110cm，之前 80cm；尚無本期時前期欄杆已低於當年標準就先打勾（2026-11-24）。"""
        self.assertTrue(bl.railing_low('0.7', '2024-04-18'))
        self.assertFalse(bl.railing_low('0.9', '2024-04-18'))
        self.assertTrue(bl.railing_low('0.9', '2026-11-19'))
        self.assertFalse(bl.railing_low('1.1', '2026-11-19'))
        base = dict(inventory_id='1', header={'檢測日期': '2024-04-18'}, railing='0.7')
        self.assertEqual(bl.build(bl.Sources(**base))[1].get('guard'), '是')
        self.assertEqual(bl.build(bl.Sources(**dict(base, railing='0.9')))[1].get('guard'), '否')
        self.assertNotIn('guard', bl.build(bl.Sources(**dict(base, railing='')))[1])     # 判斷不出來不填


class GeneratedNoteTests(unittest.TestCase):
    def test_manual_correction_is_kept(self):
        """程式產生的儲存格有專屬備註（識別碼）；人工修正後再同步不覆寫（2026-11-24）。"""
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            bl.write(path, WriteTests().sources())
            book = load_workbook(path)
            cell = book.active['O2']                                           # 前期觀測劣化(U=2)
            self.assertIn(bl.GENERATED, cell.comment.text)
            self.assertEqual(bl.generated_state(cell), 'same')
            cell.value = '1.A1橋台混凝土剝落（人工修正）(113年)'
            book.save(path)
            report = bl.write(path, WriteTests().sources())
            sheet = load_workbook(path).active
            self.assertEqual(sheet['O2'].value, '1.A1橋台混凝土剝落（人工修正）(113年)')   # 保留人工修正
            self.assertTrue(any('已人工修正' in line for line in report))
            self.assertIsNone(sheet['R2'].comment)                                        # 本期資料不鎖定、不加備註（2026-11-29）
            self.assertTrue(sheet['R2'].value)                                             # 照常更新


class FreezeTests(unittest.TestCase):
    def test_too_wide_freeze_is_fixed(self):
        """凍結窗格的欄太多（例：莿桐凍結到第 T 欄），左右捲動失效：同步時改成只凍結表頭與序號～橋梁名稱（2026-11-24）。"""
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            book = load_workbook(path)
            for letter in 'ABCDEFGHIJKLMNOPQRS':
                book.active.column_dimensions[letter].width = 20
            book.active.freeze_panes = 'T2'
            book.save(path)
            report = bl.write(path, WriteTests().sources())
            self.assertEqual(load_workbook(path).active.freeze_panes, 'E2')
            self.assertTrue(any('凍結窗格' in line for line in report))
            book = load_workbook(path)
            book.active.freeze_panes = 'C2'                                     # 合理的凍結不動
            book.save(path)
            bl.write(path, WriteTests().sources())
            self.assertEqual(load_workbook(path).active.freeze_panes, 'C2')


class TickTests(unittest.TestCase):
    def test_ticks(self):
        """ch5.5～六張照片（2026-11-24）：雜物寫在缺失也打 V；判斷不出來時不清掉人工打的 V。"""
        base = dict(inventory_id='22870', header={'橋梁名稱': '莿桐N29', '檢測日期': '2026-11-19', '是否疑似鹽害橋梁': '是'},
                    new_items=[], current_entries=[entry('河道', 'P1 橋下雜物堆積。(1式)', 2)], railing='1.0', photos_done=True)
        values = bl.build(bl.Sources(**base))[1]
        self.assertEqual((values['guard'], values['waste'], values['salt'], values['photos']), ('是', '是', '是', '是'))
        unknown = bl.build(bl.Sources(inventory_id='1', header={}, new_items=[]))[1]
        self.assertEqual((unknown['waste'], unknown['salt'], unknown.get('photos')), ('', '', None))   # 讀不到不填
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder, {'橋下堆積物(V)\n (ch5.6)': 'v'})                  # 人工打的
            bl.write(path, bl.Sources(**dict(base, current_entries=[], railing='1.2', photos_done=False,
                                             header={'橋梁名稱': '莿桐N29', '檢測日期': '2026-11-19'})))
            sheet = load_workbook(path).active
            row = dict(zip([c.value for c in sheet[1]], (c.value for c in sheet[2])))
            self.assertEqual(row['橋下堆積物(V)\n (ch5.6)'], 'v')                          # 不清掉
            self.assertEqual(row['護欄高度不足(V)\n (ch5.5)'], '否')


class PhotoSixTests(unittest.TestCase):
    def test_any_version_counts(self):
        """六張基本照片：任一版本的六個標題本年度都在就算「是」（雲林縣版本不會被當成跨水橋而判「否」）。"""
        from bridgefill.core.bridge_photos import VERSIONS, six_uploaded
        yunlin = [{'title': t, 'date': '2026-03-02'} for t in VERSIONS['雲林縣']]
        self.assertTrue(six_uploaded(yunlin, 2026))
        self.assertTrue(six_uploaded([dict(p, title=p['title'] + ' ') for p in yunlin], 2026))   # 標題多空白也算
        self.assertFalse(six_uploaded(yunlin[:5], 2026))                                       # 少一張
        self.assertFalse(six_uploaded(yunlin, 2025))                                           # 不是本年度


class AlignFontTests(unittest.TestCase):
    def test_center_and_font(self):
        """寫入的欄位：單一值置中、多筆資料靠左，字型標楷體 10（2026-11-24）。"""
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            items = [(entry('橋台', 'A1 混凝土剝落。(1處)', 2), 113), (entry('河道', 'A1 掏空。(1處)', 2), 113)]
            bl.write(path, WriteTests().sources(prior_items=items))
            sheet = load_workbook(path).active
            heads = [c.value for c in sheet[1]]
            cell = lambda name: sheet.cell(2, heads.index(name) + 1)  # noqa: E731
            several, single = cell('前期觀測劣化\n(U=2)'), cell('橋下堆積物(V)\n (ch5.6)')
            self.assertIn('\n', several.value)
            self.assertEqual((several.alignment.horizontal, several.alignment.vertical), ('left', 'center'))
            self.assertEqual((single.value, single.alignment.horizontal, single.alignment.vertical), ('是', 'center', 'center'))
            for target in (several, single, cell('橋梁名稱')):
                self.assertEqual((target.font.name, target.font.sz), ('標楷體', 10))


class Format38AndManySpansTests(unittest.TestCase):
    """v4.0：38 欄北雲林定案格式表頭比對＋超過 8 墩（如媽祖大橋 25 墩）欄杆至河床高度精準計算。"""

    def test_38_col_headers_match_all_columns(self):
        new_heads = [
            '序號', '二代系統ID', '橋梁編號', '橋梁名稱', '道路等級', '路線', '跨數', '橋長', '最大淨寬',
            '跨越物體', '下橋路徑(ABCD)', '輔助設備(A/B)',
            '前期觀測劣化\n (U=1)', '前期觀測劣化\n (U=2)', '前期觀測劣化\n (U≧3)',
            '115年新增劣化\n  (U=1)', '115年新增劣化\n  (U=2)', '115年新增劣化\n  (U≧3)',
            '過去維修紀錄', '已維修構件狀況', '前期梁下淨高(m)', '本期梁下淨高(m)', '欄杆至河床高度(m)',
            '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)', '基本資料變動', '應維修構件', '應維修\n處理建議',
            '河工設施損壞', '護欄高度不足\n(是/否)', '橋下堆積物\n(是/否)', '疑似鹽害橋梁\n(是/否)',
            '六張基本照片\n拍攝(是/否)', '橋名牌損壞、遺失\n或內容錯誤', '是否有限重牌',
            '其他檢測特殊情形\n ex:水深無法檢測..', '系統填寫完成',
        ]
        self.assertEqual(len(new_heads), 38)
        mapped = {bl.key(h): i for i, h in enumerate(new_heads, 1)}
        for name, label in bl.COLUMNS.items():
            self.assertIn(bl.key(label), mapped, f'欄位 {name}（{label}）未對應到 38 欄表頭')

        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            ws = wb.active
            ws.title = '西螺'
            ws.append(new_heads)
            row = [None] * len(new_heads)
            row[0], row[1], row[3] = 1, 31316, '媽祖大橋'
            row[20] = '\n'.join(f'P{i:02d}={3 + i * 0.1:.2f}' for i in range(1, 26))   # 25 墩
            row[23], row[24], row[25] = 1.45, 0.25, 1.30                                # +3.00m
            ws.append(row)
            path = Path(folder) / '38col.xlsx'
            wb.save(path)
            bl.write(path, bl.Sources(inventory_id='31316', header={'橋梁名稱': '媽祖大橋', '檢測日期': '2026-11-25',
                                                                    '橋下有無堆放雜物': '否', '是否疑似鹽害橋梁': '否'},
                                      new_items=[], railing='1.45', photos_done=True))
            out = load_workbook(path)['西螺']
            self.assertEqual(out.cell(2, 31).value, '否')        # 護欄高度不足(是/否)
            self.assertEqual(out.cell(2, 34).value, '是')        # 六張基本照片拍攝(是/否)
            rtb = out.cell(2, 23).value                          # 欄杆至河床高度(m)：25 墩完整計算
            lines = str(rtb).splitlines()
            self.assertEqual(len(lines), 25)
            self.assertEqual(lines[0], 'P01=6.10')
            self.assertEqual(lines[-1], 'P25=8.50')
