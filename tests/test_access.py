"""下橋路徑與輔助設備（core.access，3.0）：組合文字、讀回、存檔、比對資料輸出與同步總表。"""
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from bridgefill.core import access as ac
from bridgefill.core import bridge_list as bl


class AccessTextTests(unittest.TestCase):
    def test_user_examples(self):
        c = ac.Access('C', ['A1下游側'], '架設二截梯', ['長筒雨鞋'])
        self.assertEqual(ac.path_text(c), 'C\nA1下游側架設二截梯下橋\n水深需要長筒雨鞋')
        self.assertEqual(ac.aux_text(c), '')
        b = ac.Access('B', equipment=['橋檢車'])
        self.assertEqual(ac.path_text(b), 'B\n需用橋檢車檢測')
        self.assertEqual(ac.aux_text(b), 'B\n橋檢車')

    def test_multiple_positions_and_note(self):
        d = ac.Access('D', ['A2上游側', 'A1上游側'], '既有檢修梯', note='注意蜂窩')
        self.assertEqual(ac.path_text(d), 'D\nA1上游側、A2上游側既有檢修梯下橋\n注意蜂窩')
        self.assertEqual(ac.path_text(ac.Access('D', ['A1上游側'])), 'D\nA1上游側徒步下橋')
        self.assertEqual(ac.path_text(ac.Access()), '')

    def test_parse_round_trip(self):
        for value in (ac.Access('C', ['A1下游側'], '架設二截梯', ['長筒雨鞋']),
                      ac.Access('B', equipment=['橋檢車', '工作船']),
                      ac.Access('A', equipment=['空拍機'], note='河道深'),
                      ac.Access('D', ['A1上游側', 'A2下游側'], '徒步', ['青蛙裝'], '注意')):
            back = ac.parse(ac.path_text(value), ac.aux_text(value))
            self.assertEqual(ac.path_text(back), ac.path_text(value))
            self.assertEqual(ac.aux_text(back), ac.aux_text(value))

    def test_parse_free_text(self):
        self.assertEqual(ac.parse('D\n上下游又爬梯'), ac.Access('D', note='上下游又爬梯'))
        self.assertEqual(ac.parse('', 'B\n橋檢車').level, 'B')
        self.assertEqual(ac.parse('D 徒步').level, 'D')
        self.assertTrue(ac.parse('', '').empty)

    def test_save_load_with_manual_text(self):
        with tempfile.TemporaryDirectory() as folder:
            value = ac.Access('C', ['A1下游側'], '架設一截梯')
            self.assertIsNone(ac.load(folder, '測試橋', '2026-11-05'))
            ac.save(folder, '測試橋', '2026-11-05', value, 'C\n手改的文字', '')
            self.assertEqual(ac.load(folder, '測試橋', '2026-11-05'), value)
            self.assertEqual(ac.saved_texts(folder, '測試橋', '2026-11-05'), ('C\n手改的文字', ''))
            self.assertEqual(ac.saved_texts(folder, '別的橋', '2026-11-05'), ('', ''))


class AccessListTests(unittest.TestCase):
    def test_written_to_list_cells(self):
        from tests.test_bridge_list_sync import HEADS, sheet_file
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            sources = bl.Sources(22870, {'橋梁名稱': '莿桐N29', '橋梁編號': '13-005CS'}, {}, [], [], [], [], [], [], 113,
                                 [], '', '', None, True, True)
            sources.access_path, sources.access_aux = 'B\n需用橋檢車檢測', 'B\n橋檢車'
            bl.write(path, sources)
            row = [c.value for c in load_workbook(path).active[2]]
            self.assertEqual(row[HEADS.index('下橋路徑(ABCD)')], 'B\n需用橋檢車檢測')
            self.assertEqual(row[HEADS.index('輔助設備(A/B)')], 'B\n橋檢車')
            sources.access_path = sources.access_aux = ''                   # 沒設定：不清掉總表原有內容
            bl.write(path, sources)
            self.assertEqual(load_workbook(path).active[2][HEADS.index('下橋路徑(ABCD)')].value, 'B\n需用橋檢車檢測')


if __name__ == '__main__':
    unittest.main()


class NewBridgeTests(unittest.TestCase):
    def test_new_bridge_prior_text(self):
        """沒有前期紀錄（新建橋梁）：前期觀測劣化三欄寫「本橋為新建橋梁，未有前期資料」（2026-11-25）。"""
        _, values, _ = bl.build(bl.Sources('1', {'橋梁名稱': '新橋'}, new_bridge=True))
        self.assertEqual([values[f'prior{b}'] for b in (1, 2, 3)], [bl.NEW_BRIDGE_TEXT] * 3)
        _, values, _ = bl.build(bl.Sources('1', {'橋梁名稱': '舊橋'}))
        self.assertEqual(values['prior1'], '無')

    def test_user_headers_match(self):
        """使用者改過的表頭（西螺版型）都對得到程式欄位。"""
        from bridgefill.core.list_template import template_heads
        heads = {bl.key(h) for h in template_heads(115)}
        self.assertEqual([v for v in bl.COLUMNS.values() if bl.key(v) not in heads], [])
        for user, ours in (('下橋路徑\n(A/B/C/D)', '下橋路徑(ABCD)'), ('應維修\n構件', '應維修構件(ch5.2)'),
                           ('護欄\n高度不足', '護欄高度不足(V)\n (ch5.5)'), ('橋名牌損壞、遺失\n內容錯誤', '橋名牌損壞、遺失、內容錯誤')):
            self.assertEqual(bl.key(user), bl.key(ours))
        self.assertEqual(bl.standard_width('115年新增劣化\n (U=1)'), 15)


class VolumeRemarkTests(unittest.TestCase):
    """備註沒寫單位的體積、只寫本期深度（橋墩基礎）：不當成新樣態，依前期補成 (長m*寬m*深m)（2026-11-25）。"""
    PRIOR = {'deterioration': '基礎沖刷、裸露、掏空', 'd_location': 'P3', 'method': '持續追蹤', 'amount': '1', 'unit': '式',
             'price': '2000', 'remark': 'P3 基礎裸露。(4m*3m*1.2m)'}

    def test_pattern_strips_measure(self):
        from bridgefill.core.matching import Defect, clean_pattern
        from bridgefill.core.rule_add import defect_pattern
        for remark in ('P3 基礎裸露 4*3*1.2', 'P3基礎裸露4*3*1.5m', 'P3 基礎裸露H=1.2m', 'P1-1 基礎裸露，深1.5m', 'P3 基礎裸露4X3X1.2'):
            defect = Defect('本期', ('x', '1'), '橋墩基礎', '', '2222', '基礎沖刷、裸露、掏空',
                            {'remark': remark, 'd_location': 'P3'}, der_id='1')
            self.assertEqual(defect_pattern(defect), '基礎裸露', remark)
            self.assertEqual(clean_pattern(defect.fields), '基礎裸露', remark)

    def test_merge_volume_with_prior(self):
        from bridgefill.core.prior_merge import merge_with_prior
        for remark in ('P3 基礎裸露 4*3*1.5', 'P3 基礎裸露。4*3*1.5', 'P3 基礎裸露 1.5', 'P3 基礎裸露H=1.5m', 'P3 1.5'):
            fields, why = merge_with_prior(dict(self.PRIOR, remark=remark), self.PRIOR)
            self.assertEqual((fields['remark'], fields['unit'], fields['amount']), ('P3 基礎裸露。(4m*3m*1.5m)', '式', '1'),
                             remark)
        cubic = dict(self.PRIOR, unit='立方公尺', amount='14.4')
        fields, _ = merge_with_prior(dict(cubic, remark='P3 基礎裸露 1.5'), cubic)
        self.assertEqual((fields['remark'], fields['amount']), ('P3 基礎裸露。(4m*3m*1.5m)', '18'))
        self.assertEqual(merge_with_prior(dict(self.PRIOR, remark='P3 基礎裸露加劇'), self.PRIOR), (None, ''))


class DemolishedTests(unittest.TestCase):
    def test_changed_rows(self):
        """已拆除檢核：使用者修正後只重新掃描 ID 或橋名改過的列（2026-11-25）。"""
        from openpyxl import Workbook
        from bridgefill.ui.main_window import changed_rows
        book = Workbook()
        sheet = book.active
        sheet.title = '西螺'
        sheet.append(['序號', '二代系統ID', '橋梁名稱'])
        sheet.append([1, 111, '甲橋'])            # 未修改
        sheet.append([2, 999, '乙橋'])            # 改了 ID
        sheet.append(['3（解除管養）', 333, '丙橋'])
        sheet.append([4, 444, '丁橋(新)'])        # 改了橋名
        items = [{'sheet': '西螺', 'row': r, 'id': i, 'name': n, 'site': n + '(已拆除)'}
                 for r, i, n in ((2, '111', '甲橋'), (3, '222', '乙橋'), (4, '333', '丙橋'), (5, '444', '丁橋'))]
        jobs, notes = changed_rows(book, items)
        self.assertEqual(jobs, [('西螺', '999', '乙橋'), ('西螺', '444', '丁橋(新)')])
        self.assertEqual(len(notes), 2)


class RebuildTests(unittest.TestCase):
    def test_clear_row_keeps_manual_columns(self):
        """重新建立前期劣化資料：清掉程式整理的欄位，保留 ID、橋名、下橋路徑與人工欄位，序號改黃（2026-11-25）。"""
        from tests.test_bridge_list_sync import HEADS, sheet_file
        with tempfile.TemporaryDirectory() as folder:
            values = {'前期觀測劣化\n(U=2)': '1.A1橋台混凝土剝落(113年)', '過去維修紀錄\n(列出過去維修紀錄)': '110年度',
                      '下橋路徑(ABCD)': 'D', '其他檢測特殊情形\n ex:水深無法檢測..': '水深'}
            path = sheet_file(folder, values)
            book = load_workbook(path)
            sheet, row, name = bl.clear_row(book, '22870')
            self.assertEqual((sheet.title, row, name), ('莿桐', 2, '莿桐N29'))
            cells = {h: sheet.cell(2, i + 1).value for i, h in enumerate(HEADS)}
            self.assertIsNone(cells['前期觀測劣化\n(U=2)'])
            self.assertIsNone(cells['過去維修紀錄\n(列出過去維修紀錄)'])
            self.assertIsNone(cells['道路等級'])                       # 系統核對欄位重新讀取
            self.assertEqual((cells['二代系統ID'], cells['橋梁名稱'], cells['下橋路徑(ABCD)']), (22870, '莿桐N29', 'D'))
            self.assertEqual(cells['其他檢測特殊情形\n ex:水深無法檢測..'], '水深')
            self.assertFalse(bl.serial_done(sheet, 2, {bl.key(h): i + 1 for i, h in enumerate(HEADS)}))
            self.assertEqual(bl.clear_row(book, '1'), (None, None, ''))


class StoreFolderTests(unittest.TestCase):
    def test_prior_store_by_government_and_sheet(self):
        """前期掃描存到 scan\\縣市政府\\pre-brigdedata\\工作表，讀取時找得到；舊位置 temp\\scan 也找得到（2026-11-25）。"""
        import json
        from types import SimpleNamespace
        from bridgefill.ui.main_window import MainWindow, PRIOR_FOLDER
        with tempfile.TemporaryDirectory() as folder:
            window = SimpleNamespace(scan_dir=Path(folder) / 'temp', store_dir=Path(folder) / 'scan', login_agency='',
                                     list_sheet='西螺', safe_name=MainWindow.safe_name, log=print)
            for name in ('store_folder', 'stored_files', 'prior_files', 'save_json', 'prebuilt_path', 'prebuilt_prior'):
                setattr(window, name, getattr(MainWindow, name).__get__(window))
            scan = {'header': {'檢測日期': '2024-05-11', '橋梁名稱': '甲橋', '管理機關': '雲林縣莿桐鄉公所'}, 'rows': []}
            saved = window.save_json(scan['header'], '前期', scan)
            self.assertEqual(saved.parent, Path(folder) / 'scan' / '雲林縣政府' / PRIOR_FOLDER / '西螺')
            window.list_sheet = ''
            self.assertIsNotNone(window.prebuilt_prior(SimpleNamespace(date='2024-05-11', agency=''), '甲橋'))
            old = window.scan_dir / '2022-05-01-甲橋-前期.json'
            old.parent.mkdir(parents=True)
            old.write_text(json.dumps(scan), encoding='utf-8')
            self.assertEqual([p.name for p in window.prior_files('甲橋')],
                             ['2022-05-01-甲橋-前期.json', '2024-05-11-甲橋-前期.json'])


class PhotoVersionTests(unittest.TestCase):
    def test_other_version_six_photos_count_as_uploaded(self):
        """雲林縣的橋用跨水橋標題上傳了六張：照片頁六格都算「網站已有」，總表六張基本照片為「是」（2026-11-26）。"""
        from bridgefill.core.bridge_photos import VERSIONS, PhotoPlan, Slot, mark_uploaded, six_uploaded, uploaded_version
        photos = [{'title': t, 'date': '2026-10-01'} for t in VERSIONS['跨水橋']]
        plan = PhotoPlan('', [Slot(t) for t in VERSIONS['雲林縣']], '雲林縣')
        marked = mark_uploaded(plan, photos, 2026)
        self.assertEqual(uploaded_version(photos, 2026), '跨水橋')
        self.assertTrue(six_uploaded(photos, 2026))
        self.assertEqual(len(marked), 6)
        self.assertTrue(all('跨水橋格式六張' in s.status or s.status.startswith('網站已有') for s in plan.slots))
        plan = PhotoPlan('', [Slot(t) for t in VERSIONS['雲林縣']], '雲林縣')
        self.assertEqual(mark_uploaded(plan, photos[:5], 2026), [t for t in VERSIONS['雲林縣'] if t in VERSIONS['跨水橋'][:5]])
        self.assertFalse(six_uploaded(photos[:5], 2026))


class HeightAndRepairTests(unittest.TestCase):
    def test_height_text_multiline(self):
        """「橋下淨高:」換行 P1=3.3m 的寫法：前期淨高不再是空的（2026-11-26）。"""
        from bridgefill.core.opinion import height_text
        text = '1.A1引道路堤AC沉陷，建議持續追縱\n橋下淨高:\nP1=3.3m\nP2=3.3m'
        self.assertEqual(height_text(text), '橋下淨高:\nP1=3.3m\nP2=3.3m')
        self.assertEqual(bl.heights_value(height_text(text)), 'P01=3.30\nP02=3.30')
        self.assertEqual(height_text('梁下淨高：P1=5.90m、P2=6.15m'), '梁下淨高：P1=5.90m、P2=6.15m')

    def test_repaired_wording(self):
        from types import SimpleNamespace
        from bridgefill.core.report import repaired_text
        from bridgefill.ui.main_window import repaired_note
        self.assertEqual(repaired_text(SimpleNamespace(repaired=True)), '已維修（系統有維修紀錄）')
        self.assertIn('其他單位維修', repaired_text(SimpleNamespace(repaired=False)))
        self.assertEqual(repaired_note(SimpleNamespace(repaired=True)), '檢測標示已維修')
        self.assertIn('系統無維修建檔', repaired_note(SimpleNamespace(repaired=False)))


class PhotoTitleTests(unittest.TestCase):
    def tearDown(self):
        from bridgefill.core.bridge_photos import apply_custom
        apply_custom({})

    def test_yunlin_land(self):
        """雲林縣陸橋：往外上游→往外右側、往外下游→往外左側，其餘同雲林縣（2026-11-27）。"""
        from bridgefill.core.bridge_photos import VERSIONS, bridge_kind, six_uploaded
        self.assertEqual(bridge_kind('否', '雲林縣政府'), '雲林縣陸橋')
        self.assertEqual(bridge_kind('是', '雲林縣莿桐鄉公所'), '雲林縣')
        self.assertEqual(VERSIONS['雲林縣陸橋'], ('橋頭近端', '橋尾遠端', '橋左側照', '橋右側照', '往外右側', '往外左側'))
        photos = [{'title': t, 'date': '2026-11-25'} for t in ('橋頭近端', '橋尾遠端', '往外左側', '往外右側', '橋右側照', '橋左側照')]
        self.assertTrue(six_uploaded(photos, 2026))                     # 使用者截圖的雲林縣陸橋照片簿

    def test_custom_titles_per_government(self):
        """各縣市政府自訂標題：依管理機關套用；存檔讀回；不足六個不採用。"""
        from bridgefill.core.bridge_photos import (CUSTOM, PhotoPlan, bridge_kind, load_custom, rename_plan,
                                                   save_custom, six_uploaded)
        titles = ['頭', '尾', '左', '右', '外左', '外右']
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'photo_titles.json'
            save_custom(path, {'南投縣政府': {'陸橋': titles, '跨水橋': ['只有一個']}})
            self.assertEqual(set(CUSTOM['南投縣政府']), {'陸橋'})            # 不足六個不採用
            load_custom(path)
            self.assertEqual(bridge_kind('否', '南投縣竹山鎮公所'), '南投縣政府－陸橋')
            self.assertEqual(bridge_kind('是', '南投縣政府'), '跨水橋')       # 沒自訂跨水橋：用預設
            plan = PhotoPlan.create('否', agency='南投縣政府')
            self.assertEqual([s.title for s in plan.slots], titles)
            self.assertTrue(six_uploaded([{'title': t, 'date': '2026-01-01'} for t in titles], 2026))
            photos = [{'id': str(i), 'edit_url': 'x', 'title': t, 'date': '2026-01-01'} for i, t in enumerate(titles)]
            self.assertEqual(len(rename_plan(photos, 2026, '陸橋')), 6)
            save_custom(path, {})
            self.assertEqual(bridge_kind('否', '南投縣政府'), '陸橋')


class OrganizeStoreTests(unittest.TestCase):
    def test_organize_by_list(self):
        """依總表歸類（2026-11-28）：temp\\scan 與未分縣市的前期、基本資料搬到 scan\\縣市政府\\…\\工作表；同名不同 ID 不搬。"""
        import json
        from bridgefill.core import store_organize as so
        with tempfile.TemporaryDirectory() as folder:
            scan, store = Path(folder) / 'temp' / 'scan', Path(folder) / 'scan'

            def prior(path, ident, agency):
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(json.dumps({'header': {'管理機關': agency}, 'rows': [{'keys': {'inventory_id': ident}}]},
                                           ensure_ascii=False), encoding='utf-8')
            prior(scan / '2024-05-11-甲橋-前期.json', '111', '雲林縣莿桐鄉公所')
            prior(store / '未分縣市' / so.PRIOR_FOLDER / '2022-05-11-甲橋-前期.json', '111', '雲林縣政府')
            prior(scan / '2024-06-01-乙橋-前期.json', '999', '雲林縣政府')          # 同名但 ID 不同
            prior(scan / '0000-00-00-甲橋-前期.json', '', '')                        # 第一次檢測的空白前期：留在 temp
            basic = store / '雲林縣政府' / so.BASIC_FOLDER / '111-甲橋-基本資料.json'
            basic.parent.mkdir(parents=True)
            basic.write_text(json.dumps({'county': '雲林縣'}), encoding='utf-8')
            bridges = [{'sheet': '莿桐', 'name': '甲橋', 'inventory_id': '111'},
                       {'sheet': '西螺', 'name': '乙橋', 'inventory_id': '222'}]
            moves, notes = so.plan(bridges, store, scan)
            self.assertEqual(len(moves), 3)
            self.assertIn('同名的其他橋', notes[0])
            moved, more = so.apply(moves)
            self.assertEqual((moved, more), (3, []))
            target = store / '雲林縣政府' / so.PRIOR_FOLDER / '莿桐'
            self.assertEqual(sorted(p.name for p in target.iterdir()), ['2022-05-11-甲橋-前期.json', '2024-05-11-甲橋-前期.json'])
            self.assertTrue((store / '雲林縣政府' / so.BASIC_FOLDER / '莿桐' / '111-甲橋-基本資料.json').exists())
            self.assertTrue((scan / '0000-00-00-甲橋-前期.json').exists())
            self.assertTrue((scan / '2024-06-01-乙橋-前期.json').exists())
            self.assertEqual(so.plan(bridges, store, scan)[0], [])                   # 再執行一次：不需搬移
