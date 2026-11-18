"""主視窗端對端測試：在模擬網站上掃描並配對，審查、存檔、接續。"""
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
if sys.platform.startswith('linux') and hasattr(os, 'geteuid') and os.geteuid() == 0:
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--no-sandbox')

try:
    from PySide6.QtCore import QCoreApplication, QEvent, QEventLoop, QTimer
    from PySide6.QtWidgets import QApplication
    import bridgefill.ui.main_window as screen
except ImportError as error:
    raise unittest.SkipTest(f'QtWebEngine 無法載入：{error}')

from openpyxl import load_workbook
from bridgefill.core.review import ACCEPTED, CHANGED, KIND_GONE, PENDING
from test_scan_window import MockSite, history_page, start_mock_site

MAIN = '/bms3/inspect/inspect/der_content/inspect_id,9/inventory_id,8/from,result/'
SEARCH_PAGE = """<div id="maincontent"><div><form method="get" action="/bms3/inspect/inspect/result/">
<table><tr><td><input name="inventory_id"></td></tr></table>
<input id="bridge_name" name="bridge_name"><input type="hidden" name="a"><input type="hidden" name="b">
<input type="reset" value="重新輸入"><input type="submit" value="開始查詢"></form></div></div>"""


def search_results(*rows):
    """查詢結果頁，欄位比照歷史定期檢測紀錄；rows：(橋梁系統ID, 檢測系統ID, 日期, 橋名)。"""
    body = ''.join(f"""<tr><td>{i}</td><td>{inv}</td><td>{iid}</td><td>{date}</td><td>某公司</td><td>甲</td>
<td>{name}</td><td>南投縣政府</td><td>南投縣</td><td>鹿谷</td>
<td><a href="inspect/inspect/view/inspect_id,{iid}/inventory_id,{inv}/from,inventory/view,1/">檢視</a>
<a href="inspect/inspect/der_content/inspect_id,{iid}/inventory_id,{inv}/from,result/">修改</a></td></tr>"""
                   for i, (inv, iid, date, name) in enumerate(rows, 1))
    return ('<base href="https://bms2.iot.gov.tw/bms3/"><table><tr><th>項次</th><th>橋梁系統ID</th><th>檢測系統ID</th>'
            '<th>檢測日期</th><th>檢測單位</th><th>檢測人員</th><th>橋梁名稱</th><th>管理機關</th><th>所在縣市</th>'
            f'<th>所在區鄉</th><th></th></tr>{body}</table>')


class MainWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

        def no_dialog(_parent, title, text, *args):
            raise AssertionError(f'不應出現對話框：{title}：{text}')
        cls._no_dialog = staticmethod(no_dialog)
        cls._dialogs = (screen.QMessageBox.information, screen.QMessageBox.warning)
        screen.QMessageBox.information = screen.QMessageBox.warning = staticmethod(no_dialog)
        cls._ask_resume = screen.MainWindow.ask_resume
        screen.MainWindow.ask_resume = lambda self, item: 'rescan'     # 預設重新掃描；接續另外測
        cls.server, cls.base = start_mock_site()
        both = search_results(('8', '9', '2026-11-05', '測試橋'), ('18', '19', '2026-10-01', '測試橋(2)'),
                              ('8', '7', '2024-12-05', '測試橋'))
        only = search_results(('8', '9', '2026-11-05', '測試橋'), ('8', '7', '2024-12-05', '測試橋'))
        # 查詢「測試橋」只查到一座（直接處理）；查詢「測試」查到兩座（請使用者選擇）
        MockSite.pages = {r'^/bms3/inspect/inspect/search/$': SEARCH_PAGE,
                          r'result/\?inventory_id=&bridge_name=%E6%B8%AC%E8%A9%A6%E6%A9%8B&': only,
                          r'result/\?inventory_id=&bridge_name=%E6%B8%AC%E8%A9%A6&': both,
                          **MockSite.pages}

    @classmethod
    def tearDownClass(cls):
        screen.QMessageBox.information, screen.QMessageBox.warning = cls._dialogs
        screen.MainWindow.ask_resume = cls._ask_resume
        cls.server.shutdown()

    def wait(self, ms):
        loop = QEventLoop()
        QTimer.singleShot(ms, loop.quit)
        loop.exec()

    def open_window(self, folder):
        rewrite = lambda url: url.replace('https://bms2.iot.gov.tw', self.base)
        window = screen.MainWindow(self.base + MAIN, Path(folder) / 'profile', Path(folder) / 'scan',
                                   Path(folder) / 'review', delay_ms=0, rewrite=rewrite,
                                   output_dir=Path(folder) / 'output', rules_path=self.rules_copy(folder),
                                   supervisor_path=Path(folder) / 'supervisors.json')
        window.show()
        self.wait(1500)
        return window

    def rules_copy(self, folder):
        """規則表複本（填寫分頁會自動新增規則，不能寫到正式檔）。"""
        target = Path(folder) / 'rules' / '規則v2.xlsx'
        if not target.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(screen.rules_file(), target)
        return target

    def close(self, window):
        window.close()
        window.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        self.wait(300)

    def test_scan_match_review_save_and_resume(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            pairs = list(window.pairs)
            self.assertTrue(pairs)
            self.assertEqual(window.table.rowCount(), len(pairs))
            self.assertTrue((Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json').exists())
            self.assertTrue((Path(folder) / 'scan' / '2024-12-05-測試橋-前期.json').exists())
            review_file = window.review_file
            self.assertTrue(review_file.exists())
            self.assertIn('共', window.summary.text())

            # 篩選：待審查
            window.filter.setCurrentText(PENDING)
            pending = list(window.visible)
            self.assertTrue(all(window.review.decision(p).status == PENDING for p in pending))

            # 逐筆審查：開第一筆待審查，改判並寫說明
            window.filter.setCurrentText('全部')
            target = next(p for p in window.pairs if p.prior and not p.current)    # 只有前期的一筆
            window.open_card(window.visible.index(target))
            self.assertIs(window.tabs.currentWidget(), window.check_page)
            self.assertIs(window.card.pair, target)
            self.assertIn(KIND_GONE, [window.card.kind.itemText(i) for i in range(window.card.kind.count())])
            window.card.kind.setCurrentText(KIND_GONE)
            window.card.note = '河水沖刷'
            window.card.decided.emit(window.card.kind.currentText(), window.card.note)
            saved = json.loads(review_file.read_text(encoding='utf-8'))
            self.assertEqual(saved['decisions'][target.key]['kind'], KIND_GONE)
            self.assertEqual(saved['decisions'][target.key]['status'], CHANGED)

            # ③ 比對資料輸出：Excel 含審查結果
            path = window.export_report()
            self.assertEqual(path, Path(folder) / 'output' / '2026-11-05-測試橋-前後期比對.xlsx')
            sheet = load_workbook(path).active
            headers = [c.value for c in sheet[3]]
            rows = [dict(zip(headers, [c.value for c in r])) for r in sheet.iter_rows(min_row=4)]
            self.assertEqual(len(rows), len(pairs))
            gone = [r for r in rows if r['最終判定'] == KIND_GONE]
            self.assertEqual([(r['審查'], r['審查說明']) for r in gone], [(CHANGED, '河水沖刷')])
            self.close(window)

            # 重新開啟、讀取上次掃描檔：審查結果接續
            window = self.open_window(folder)
            window.load_saved(Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json')
            window.set_overall_filter('全部')
            again = next(p for p in window.pairs if p.key == target.key)
            self.assertEqual((window.review.decision(again).kind, window.review.decision(again).note),
                             (KIND_GONE, '河水沖刷'))
            window.filter.setCurrentText(CHANGED)
            self.assertEqual([p.key for p in window.visible], [target.key])
            self.close(window)

    def test_bridge_name_alone_runs_everything(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            asked = []
            window.choose_bridge = lambda name, bridges: asked.append(name)
            MockSite.requests.clear()
            window.bridge_name.setText('測試橋')
            window.run_by_name()
            self.assertEqual(asked, [])                                  # 只有一座且同名：不必選
            # 先開基本資料、出現確認列，確認前不掃描
            self.assertFalse(window.pairs)
            self.assertFalse(window.confirm_bar.isHidden())
            self.assertIn('8', window.confirm_text.text())
            self.wait(500)
            self.assertIn('/basic/inventory/edit/8/', window.view.url().toString())
            window.confirm_bridge(True)
            self.assertTrue(window.confirm_bar.isHidden())
            self.assertTrue(window.pairs)
            self.assertTrue((Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json').exists())
            self.assertTrue((Path(folder) / 'scan' / '2024-12-05-測試橋-前期.json').exists())
            self.assertTrue(any('/inspect/inspect/search/' in p for p in MockSite.requests))
            self.assertFalse([p for p in MockSite.requests if re.search(r'delete|update|add|save', p)])
            self.wait(500)
            self.assertIn('der_content/inspect_id,9/', window.view.url().toString())   # 左側同步顯示本期
            self.close(window)

    def test_no_current_year_prebuilds_prior(self):
        """網站沒有本年度紀錄：先把最新一期當前期掃描建檔、重新讀歷史紀錄；沒有本期就不掃描。
        之後有本年度紀錄時沿用建檔的前期（不重新掃描），只掃描本期。"""
        key = r'^/bms3/inspect/inspect/result/inventory_id,8/'
        original = MockSite.pages[key]
        shown = []
        try:
            MockSite.pages[key] = history_page(('7', '2024-12-05'))
            screen.QMessageBox.information = staticmethod(lambda _p, title, text, *a: shown.append((title, text)))
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                window = self.open_window(folder)
                window.this_year = lambda: '2026'
                window.choose_bridge = lambda name, bridges: bridges[0]
                window.bridge_name.setText('測試橋')
                window.run_by_name()
                MockSite.requests.clear()
                window.confirm_bridge(True)
                prior = Path(folder) / 'scan' / '2024-12-05-測試橋-前期.json'
                self.assertTrue(prior.exists())
                self.assertTrue(json.loads(prior.read_text(encoding='utf-8'))['prebuilt'])
                self.assertFalse(window.pairs)                                       # 沒有本期：不掃描
                self.assertFalse(list((Path(folder) / 'scan').glob('*-檢測.json')))
                self.assertEqual(sum('result/inventory_id,8/' in p for p in MockSite.requests), 2)   # 建檔後重新讀歷史紀錄
                self.assertEqual(shown[-1][0], '前期已建檔')
                self.assertIn('沿用已建檔的前期', shown[-1][1])
                self.assertIn('還沒有 2026 年', window.summary.text())
                # 網站新增本年度紀錄後再按開始：前期沿用建檔，只掃描本期
                MockSite.pages[key] = original
                window.bridge_name.setText('測試橋')
                window.run_by_name()
                MockSite.requests.clear()
                window.confirm_bridge(True)
                self.assertTrue(window.pairs)
                self.assertFalse([p for p in MockSite.requests if 'der_content/inspect_id,7/' in p])   # 前期不重掃
                self.assertTrue(any('der_content/inspect_id,9/' in p for p in MockSite.requests))
                self.close(window)
        finally:
            MockSite.pages[key] = original
            screen.QMessageBox.information = staticmethod(self._no_dialog)

    def test_multi_bridge_and_schedule(self):
        """多橋同步掃描（不需逐座確認、單座失敗不影響其他）與排程掃描（保持登入、時間到自動掃描）。"""
        shown = []
        screen.QMessageBox.information = staticmethod(lambda _p, title, text, *a: shown.append((title, text)))
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                window = self.open_window(folder)
                window.this_year = lambda: '2026'
                window.run_by_name('查無此橋、測試橋')
                self.assertTrue((Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json').exists())
                self.assertTrue(window.pairs)                                     # 第一座成功的橋已開啟
                self.assertEqual(shown[-1][0], '多橋掃描')
                self.assertIn('查無此橋：失敗', shown[-1][1])
                self.assertIn('測試橋：已掃描', shown[-1][1])
                # 排程：每 30 分鐘讀一次網站保持登入；時間到自動掃描（無人值守，不跳對話框）
                window.set_schedule(['測試橋'], screen.QTime(23, 59))
                self.assertIn('測試橋', window.schedule_label.text())
                MockSite.requests.clear()
                window.last_alive = window.last_alive.addSecs(-31 * 60)
                window.schedule_tick()
                self.assertTrue(any('/inspect/inspect/search/' in p for p in MockSite.requests))
                count = len(shown)
                window.schedule_at = screen.QDateTime.currentDateTime().addSecs(-1)
                window.schedule_tick()
                self.assertFalse(window.schedule_names)                           # 排程已執行
                self.assertEqual(len(shown), count)                               # 排程掃描不跳對話框
                self.assertIn('多橋掃描完成', window.summary.text())
                self.close(window)
        finally:
            screen.QMessageBox.information = staticmethod(self._no_dialog)

    def test_remembered_login_is_filled(self):
        """記住帳號密碼：登入頁開啟時自動填入（不按登入）；測試以記憶體代替 Windows 認證管理員。"""
        from bridgefill.ui import credentials
        login = ('<html><body><form method="post" action="/bms3/public/public/login/"><input id="user" name="user">'
                 '<input id="password" name="password" type="password"><input type="submit" value="確定登入"></form></body></html>')
        MockSite.pages[r'/public/public/login/'] = login
        vault = credentials.MemoryVault()
        credentials.use(vault)
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                window = self.open_window(folder)
                credentials.save('inspector', 'p@ss')
                window.view.open(self.base + '/bms3/public/public/login/')
                self.wait(1500)
                values = []
                window.view.page().runJavaScript(
                    "[document.getElementById('user').value, document.getElementById('password').value].join('|')",
                    0, values.append)
                self.wait(300)
                self.assertEqual(values, ['inspector|p@ss'])
                self.assertIn('/public/public/login/', window.view.url().toString())     # 沒有代按登入
                self.assertNotIn('p@ss', window.output.toPlainText())                      # 紀錄不含密碼
                self.close(window)
        finally:
            credentials.use(None)
            MockSite.pages.pop(r'/public/public/login/', None)

    def test_similar_names_ask_which_bridge(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            seen = {}

            def choose(name, bridges):
                seen['names'] = [b.name for b in bridges]
                return next(b for b in bridges if b.name == '測試橋')
            window.choose_bridge = choose
            window.run_by_name('測試')
            self.assertEqual(seen['names'], ['測試橋', '測試橋(2)'])
            window.confirm_bridge(True)
            self.assertTrue(window.pairs)
            # 基本資料確認「不是」：不掃描
            window.pairs = []
            window.run_by_name('測試')
            window.confirm_bridge(False)
            self.assertEqual(window.pairs, [])
            self.assertIn('不是這座橋', window.output.toPlainText())
            # 取消時不做任何事
            window.pairs = []
            window.choose_bridge = lambda name, bridges: None
            window.run_by_name('測試')
            self.assertEqual(window.pairs, [])
            self.assertIn('已取消', window.output.toPlainText())
            self.close(window)

    def test_history_resumes_without_rescanning(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            pair = next(p for p in window.pairs if p.current)
            window.review.set_edit(pair, {'remark': '改過'})
            window.review.save(window.review_file)
            self.close(window)

            window = self.open_window(folder)
            window.fill_history()
            entries = [a.text() for a in window.history_menu.actions() if '｜' in a.text()]
            self.assertEqual(len(entries), 1)
            self.assertIn('測試橋｜本期 2026-11-05', entries[0])
            self.assertIn('已修改 1', entries[0])
            MockSite.requests.clear()
            window.ask_resume = lambda item: 'continue'
            window.run_by_name('測試橋')                               # 同名且有紀錄：接續，不查詢、不掃描
            self.assertTrue(window.pairs)
            self.assertEqual(window.review.value(next(p for p in window.pairs if p.key == pair.key), 'remark'), '改過')
            self.wait(500)
            self.assertFalse([r for r in MockSite.requests if 'search' in r or 'der_list' in r])
            self.assertIn('der_content/inspect_id,9/', window.view.url().toString())
            self.close(window)

    def test_clear_history(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            window.tabs.setCurrentWidget(window.fill_tab)            # 產生填寫進度檔
            window.fill_log.mark(window.fill_items[0].key, '已完成', window.fill_items[0].target)
            files = sorted(p.name for p in Path(folder).rglob('*.json') if 'layout' not in p.name)
            self.assertEqual(len(files), 5)                           # 本期、前期、基本資料、審查、填寫
            window.fill_history()
            self.assertEqual(window.history_menu.actions()[-1].text(), '清除歷史紀錄…')
            window.choose_history = lambda sessions: []               # 取消：不刪
            window.clear_history()
            self.assertEqual(len([p for p in Path(folder).rglob('*.json') if 'layout' not in p.name]), 5)
            window.choose_history = lambda sessions: sessions
            window.confirm_clear = lambda items: True
            window.clear_history()
            self.assertEqual([p.name for p in Path(folder).rglob('*.json') if 'layout' not in p.name], [])
            self.assertEqual(window.pairs, [])
            self.assertIn('已清除', window.summary.text())
            window.fill_history()
            self.assertEqual(window.history_menu.actions()[0].text(), '（尚無掃描紀錄）')
            self.close(window)

    def test_relink_and_split(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            alone = next(p for p in window.pairs if p.prior and not p.current)
            window.index = window.pairs.index(alone)
            seen = {}

            def choose(pair, options):
                seen['tiers'] = [o.tier for o in options]
                return [next(o.defect for o in options if o.tier in (1, 2, 3))]
            window.choose_link = choose
            window.relink()
            self.assertEqual(seen['tiers'], sorted(seen['tiers']))
            linked = [p for p in window.pairs if p.manual]
            self.assertEqual(len(linked), 1)
            self.assertEqual(window.review.decision(linked[0]).status, CHANGED)
            self.assertIs(window.pairs[window.index], linked[0])
            window.filter.setCurrentText('人工配對')
            self.assertEqual(window.visible, linked)
            saved = json.loads(window.review_file.read_text(encoding='utf-8'))
            self.assertEqual(len(saved['links']), 1)

            window.choose_link = lambda pair, options: 'split'
            window.relink()
            self.assertFalse([p for p in window.pairs if p.manual])
            self.assertEqual(len(json.loads(window.review_file.read_text(encoding='utf-8'))['splits']), 1)
            self.close(window)

    def test_edit_photos_and_finish(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            pair = next(p for p in window.pairs if p.current)
            window.open_card(window.visible.index(pair))
            card = window.card
            card.edit_toggle.setChecked(True)
            self.assertTrue(card.editors['remark'].isVisible())
            card.deru[3].setCurrentText('3')
            card.editors['remark'].setText(card.editors['remark'].text() + '補')
            card.edited.emit(card.values())
            saved = json.loads(window.review_file.read_text(encoding='utf-8'))['edits'][pair.key]
            self.assertEqual(saved['u'], '3')
            self.assertTrue(saved['remark'].endswith('補'))
            self.assertTrue({'u', 'remark'} <= set(saved))                # 只存與網站不同的欄位
            self.assertTrue(all(v != screen.site_value(pair.current, k) for k, v in saved.items()))
            self.assertIn('✎', window.table.item(window.visible.index(pair), 0).text())  # 構件欄標示已修改
            card.restore()                                               # 還原網站值後確認：清除修改、維持現值
            window.index = window.pairs.index(pair)
            card.decided.emit(card.kind.currentText(), '')
            self.assertNotIn(pair.key, window.review.edits)
            self.assertIn(pair.key, window.review.kept)
            window.open_card(window.visible.index(pair)) if pair in window.visible else window.show_card()

            # 照片：本期開第二級編輯頁（可刪照片），網站隱藏時自動顯示
            window.show_site(False)
            window.show_photos('current')
            self.assertTrue(window.view.isVisible())
            self.wait(1500)
            self.assertIn('/inspect/damage/edit_damage/der_id,', window.view.url().toString())
            if pair.prior:                                           # 前後期照片比對：兩個網站並排
                window.show_photos('compare')
                self.assertTrue(window.prior_view.isVisible() and window.view.isVisible())
                self.wait(500)
                self.assertIn('/inspect/damage/view_damage/der_id,', window.prior_view.url().toString())
                self.assertIn('/inspect/damage/view_damage/der_id,', window.view.url().toString())   # 比對時本期也用檢視頁
                sites, work = window.splitter.sizes()
                self.assertLess(sites, (sites + work) * 0.5)          # 只看照片：網站區縮窄、工作區變寬

            window.show_photos('main')                               # 回到缺失主頁面
            self.wait(500)
            self.assertIn('der_content/inspect_id,9/', window.view.url().toString())
            new = next((p for p in window.pairs if p.current and not p.prior), None)
            if new:                                                  # 防呆：新增劣化沒有前期照片
                window.open_card(window.visible.index(new))
                self.assertFalse(card.photo_buttons['prior'].isEnabled())
                self.assertFalse(card.photo_buttons['compare'].isEnabled())

            # 待處理之間移動：到最後提示
            for p in window.pairs:
                window.review.decide(p, window.review.decision(p).kind)
            window.move(1)
            self.assertRegex(window.output.toPlainText(), '待處理全部完成|已到最後一筆')

            # 比對輸出的選項工作表
            window.review.set_edit(pair, {'u': '4' if pair.current.deru[3:4] != '4' else '1'})
            path = window.export_report()
            names = load_workbook(path).sheetnames
            self.assertEqual(names, ['前後期比對', 'U值增減', '新增劣化', '未建（漏建）劣化'])
            window.export_options['新增劣化'].setChecked(False)
            self.assertNotIn('新增劣化', load_workbook(window.export_report()).sheetnames)
            self.close(window)

    def test_fill_tab_lists_items_and_stops_safely(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            window.tabs.setCurrentWidget(window.fill_tab)
            self.assertEqual(len(window.fill_items), sum(1 for p in window.pairs if p.current))
            # ② 送出只列要送出／已送出的項目
            self.assertEqual(window.fill_table.rowCount(), len(window.fill_visible))
            self.assertTrue(all(i.status in ('可填寫', '已完成', '待確認', '失敗') for i in window.fill_visible))
            pair = next(p for p in window.pairs if p.current)
            window.review.set_edit(pair, {'remark': '人工修改。(1處)'})
            window.build_fill_plan()
            item = next(i for i in window.fill_items if i.pair is pair)
            self.assertEqual(item.status, '可填寫')
            self.assertIn('人工修改', item.describe())
            for each in window.pairs:                      # 送出只含已確認的項目：先全部確認
                window.review.decide(each, window.review.decision(each).kind)
            # 多選送出、依構件送出（攔下實際送出，只看選了哪些）
            sent, run_fill = [], window.run_fill
            window.run_fill = lambda items: sent.append(list(items))
            window.refresh_fill_table()
            row = window.fill_visible.index(item)
            window.fill_table.selectAll()
            ready = [i for i in window.fill_visible if i.status == '可填寫' and i.changes]
            self.assertIn(f'（{len(ready)}）', window.buttons['send_one'].text())
            window.fill_selected()
            self.assertEqual(sent[-1], ready)                                   # 已完成等略過
            window.fill_table.clearSelection()
            window.fill_table.selectRow(row)
            window.fill_selected()
            self.assertEqual(sent[-1], [item])
            name = item.pair.current.component
            window.fill_component.setCurrentIndex(window.fill_component.findData(name))
            self.assertTrue(all(i.pair.current.component == name for i in window.fill_visible))
            self.assertTrue(window.buttons['send_component'].isEnabled())
            window.fill_component_items()
            self.assertEqual(sent[-1], [i for i in window.fill_items if i.status == '可填寫' and i.pair.current.component == name])
            window.fill_component.setCurrentIndex(0)
            self.assertFalse(window.buttons['send_component'].isEnabled())    # 全部構件：用整批送出
            window.run_fill = run_fill
            window.confirm_fill = lambda items: True
            MockSite.requests.clear()
            window.run_fill([item])                      # 模擬網站的表單沒有「送出更新」：不送出、記為失敗
            self.assertEqual(item.status, '失敗')
            self.assertRegex(item.message, '未送出|沒有送出')            # 簡化表單：安全停止
            self.assertFalse([r for r in MockSite.requests if 'update_damage' in r])
            self.assertIn('填寫失敗', window.output.toPlainText())
            self.close(window)

    def test_import_prior_type_filter_and_autosave(self):
        from bridgefill.core.matching import Defect, Pair
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            card = window.card
            rules = [dict(pattern='混凝土沖蝕', method='混凝土修復，蜂窩及剝落修補，水刀鑿除', unit='平方公尺', price='3200',
                          deterioration='其他損傷', note='有關鍵字[沖蝕]即可使用', rule_id='6-23', components=['橋台'],
                          component='橋台'),
                     dict(pattern='混凝土剝落鋼筋外露', method='混凝土修復，鋼筋銹蝕處理', unit='處', price='800',
                          deterioration='鋼筋外露、銹蝕', note='有關鍵字[鋼筋外露]即可使用', rule_id='6-11',
                          components=['橋台'], component='橋台')]
            card.set_rules(rules)
            fields = dict(deterioration='其他損傷', d_location='', method='混凝土修復', amount='20', unit='平方公尺',
                          price='2300', remark='', d='2', e='2', r='1', u='1')
            options = {'deterioration': ['其他損傷', '鋼筋外露、銹蝕'], 'unit': ['式', '平方公尺', '處']}
            current = Defect('本期', ('abutment', '2'), '橋台', '橋尾', '2211', '其他損傷', fields, method_id='9009',
                             options=options)
            prior = Defect('前期', ('abutment', '2'), '橋台', '橋尾', '2211', '其他損傷',
                           dict(fields, d_location='A2', method='持續追蹤', amount='1', unit='式', price='2000',
                                remark='A2 沖蝕 (10m*2m)'), der_id='400')
            pair = Pair('單位數量變更', prior, current, confidence='需確認')
            window.pairs, window.index = [pair], 0
            window.review.sync(window.pairs)
            window.review_file = Path(folder) / 'review.json'
            window.show_card()
            card.edit_toggle.setChecked(True)
            # 樣態按鈕只列目前劣化類型（其他損傷）的規則
            tags = [card.tag_layout.itemAt(i).widget().text() for i in range(card.tag_layout.count())]
            self.assertEqual(tags, ['混凝土沖蝕'])
            # 導入前期：依規則修正（樣態、單位、數量、單價），再點樣態只換中間敘述
            card.import_prior()
            self.assertEqual(card.editors['remark'].text(), 'A2 混凝土沖蝕。(2m*10m)')
            card.pick_pattern('混凝土沖蝕')
            values = card.values()
            self.assertEqual((values['unit'], values['amount'], values['d_location'], values['price']),
                             ('平方公尺', '20', 'A2', '3200'))
            self.assertTrue(values['remark'].startswith('A2 混凝土沖蝕。('))
            # 確認並下一筆：本期欄的值就是確認的結果（不再有儲存修改按鈕）
            card.decided.emit(card.kind.currentText(), '')
            self.assertEqual(window.planned_values(pair)['price'], '3200')       # 與程式預定值相同：不另存修改
            self.assertTrue(window.review.decision(pair).at)
            # 只填範圍：點樣態補上前期位置，不洗掉本期範圍；備註與位置、數量雙向同步
            window.show_card()
            card.edit_toggle.setChecked(True)
            card.editors['remark'].setText('3m*4m')
            card.pick_pattern('混凝土沖蝕')
            self.assertEqual(card.editors['remark'].text(), 'A2 混凝土沖蝕。(3m*4m)')
            self.assertEqual(card.editors['amount'].text(), '12')
            card.editors['d_location'].setText('A2-L')
            card.editors['d_location'].textEdited.emit('A2-L')
            self.assertTrue(card.editors['remark'].text().startswith('A2-L 混凝土沖蝕'))
            card.editors['remark'].setText('A1 混凝土沖蝕。(2m*2m)')
            card.editors['remark'].textEdited.emit('A1 混凝土沖蝕。(2m*2m)')
            self.assertEqual((card.editors['d_location'].text(), card.editors['amount'].text()), ('A1', '4'))
            self.close(window)

    def test_one_list_with_overall_status(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            auto = next((p for p in window.pairs if window.review.decision(p).status == '已確認'), None)
            if auto is not None:                          # 高信心自動判定：未確認（可送出或待處理），須確認後才算數
                self.assertFalse(window.review.decision(auto).at)
                self.assertIn(window.overall(auto)[0], ('待處理', '可送出'))
                self.assertRegex(window.overall(auto)[1], '待確認|未確認')
            # 程式高信心判定且可直接填寫：放在「可送出」（未確認）；「確認全部」一次確認，送出只含已確認的
            waiting = window.unconfirmed_sendable()
            for each in waiting:
                self.assertEqual(window.overall(each)[0], '可送出')
                self.assertIn('未確認', window.overall(each)[1])
            if waiting:
                self.assertTrue(window.confirm_all_button.isEnabled())
                self.assertEqual(window.confirm_all(ask=False), len(waiting))
                self.assertTrue(all(window.confirmed(p) for p in waiting))
                self.assertFalse(window.confirm_all_button.isEnabled())
            window.rules = []                             # 沒有規則表：本期每一筆的填寫都需確認
            window.build_fill_plan()
            review = next(i for i in window.fill_items if i.status == '需確認')
            pair = review.pair
            window.review.decide(pair, window.review.decision(pair).kind)    # 審查已確認，但填寫需確認
            window.refresh_table()
            self.assertEqual(window.overall(pair)[0], '待處理')
            self.assertIn('填寫需確認', window.overall(pair)[1])
            # 數字卡片：待處理篩選、顯示筆數
            window.set_overall_filter('待處理')
            self.assertIn(pair, window.visible)
            self.assertTrue(window.status_buttons['待處理'].isChecked())
            self.assertIn(str(len(window.visible)), window.status_buttons['待處理'].text())
            # 點選清單一筆：下方工作區顯示，不換分頁；填寫狀態橫幅
            window.table.setCurrentCell(window.visible.index(pair), 0)
            self.assertIs(window.card.pair, pair)
            self.assertIs(window.tabs.currentWidget(), window.check_page)
            self.assertIn('需確認', window.card.fill_text.text())
            # 還原網站值＋確認並下一筆＝維持網站現值：狀態改為無需修改
            window.card.restore()
            window.card.decided.emit(window.card.kind.currentText(), '')
            self.assertEqual(window.overall(pair)[0], '無需修改')
            self.assertNotIn(pair, window.visible)
            self.assertIn('無需修改', window.summary.text())
            self.close(window)

    def test_codes_columns_and_photo_tab(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            # 清單：沒有狀態欄（顏色表示），位置改為簡短代碼，備註位置與樣態分行，每列兩行高
            headers = [window.table.horizontalHeaderItem(c).text() for c in range(window.table.columnCount())]
            self.assertEqual(headers[:2], ['構件', '編號'])
            self.assertNotIn('狀態', headers)
            self.assertNotIn('將填寫的變更', headers)                 # 看 ② 送出
            self.assertLessEqual(window.table.columnWidth(0), screen.TABLE_WIDTHS[0])   # 構件名稱長時折成兩行
            codes = {window.table.item(r, 1).text() for r in range(window.table.rowCount())}
            self.assertTrue(codes <= {p.code or '—' for p in window.pairs})
            self.assertFalse(any('：' in c for c in codes))
            row = next(r for r in range(window.table.rowCount()) if window.visible[r].current)
            self.assertIn('\n', window.table.item(row, 3).text())
            height = window.table.rowHeight(row)
            self.assertEqual(height, window.table.fontMetrics().lineSpacing() * 2 + 8)
            # 工作區可改放右側（1080p），設定會記住
            window.side_toggle.setChecked(True)
            self.assertEqual(window.check_parts.orientation(), screen.Qt.Orientation.Horizontal)
            window.log_toggle.setChecked(False)
            self.assertFalse(window.output.isVisible())
            saved = json.loads(window.layout_file.read_text(encoding='utf-8'))
            self.assertEqual((saved['work_side'], saved['log_hidden']), (True, True))

            # 橋梁照片：基本資料「否」→ 陸橋六格；拖入照片、交換、存檔後重開仍在
            panel = window.photo_panel
            self.assertEqual(panel.kind.currentText(), '陸橋')
            self.assertEqual([s.title for s in panel.plan.slots][:2], ['橋頭近端', '橋尾遠端'])
            self.assertFalse(panel.upload_button.isEnabled())
            photos = []
            for name in ('a.jpg', 'b.jpg'):
                path = Path(folder) / name
                path.write_bytes(b'\xff\xd8\xff\xd9')
                photos.append(str(path))
            panel.place(photos, 0)
            self.assertEqual([s.date for s in panel.plan.slots[:2]], ['2026-11-05', '2026-11-05'])   # 檢測日期
            self.assertIn('2026-11-05', [panel.slots[0].date.itemText(i) for i in range(panel.slots[0].date.count())])
            panel.swap(0, 1)
            self.assertEqual(panel.plan.slots[0].path, photos[1])
            self.assertTrue(panel.upload_button.isEnabled())
            self.assertTrue(window.photo_plan_path().exists())
            # 多選上傳：新放入的預設勾選；可取消個別、全選、全不選
            self.assertEqual(panel.selected_slots(), panel.plan.slots[:2])
            panel.slots[0].pick.setChecked(False)
            self.assertEqual(panel.selected_slots(), [panel.plan.slots[1]])
            self.assertIn('（1）', panel.upload_button.text())
            self.assertFalse(panel.slots[3].pick.isEnabled())                 # 沒有照片不能勾選
            panel.pick_all(False)
            self.assertFalse(panel.upload_button.isEnabled())
            panel.pick_all(True)
            self.assertEqual(len(panel.selected_slots()), 2)
            panel.plan.slots[1].status = '已上傳'                             # 已上傳的不再選取
            panel.refresh()
            self.assertEqual(panel.selected_slots(), [panel.plan.slots[0]])
            panel.plan.slots[1].status = ''
            self.close(window)
            window = self.open_window(folder)
            window.load_saved(Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json')
            self.assertEqual(window.photo_panel.plan.slots[0].path, photos[1])
            self.close(window)

    def test_keep_and_rescan_after_fill(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            window.tabs.setCurrentWidget(window.fill_tab)
            review = next((i for i in window.fill_items if i.status == '需確認'), None)
            if review is not None:
                window.keep_current(review.pair)
                again = next(i for i in window.fill_items if i.pair.key == review.pair.key)
                self.assertEqual(again.status, '已確認')
                self.assertIn('已人工確認', again.reason)
                saved = json.loads(window.review_file.read_text(encoding='utf-8'))
                self.assertIn(review.pair.key, saved['kept'])
            MockSite.requests.clear()
            self.assertTrue(window.rescan_current())
            self.assertTrue(any('der_content/inspect_id,9/' in r for r in MockSite.requests))
            self.assertFalse([r for r in MockSite.requests if 'search' in r or 'inspect_id,7' in r])   # 前期不重掃
            self.assertTrue(window.fill_items)
            self.assertIs(window.tabs.currentWidget(), window.fill_tab)
            window.fill_log.mark(window.fill_items[0].key, '已完成', window.fill_items[0].target)
            window.fill_history()
            self.assertIn('已填寫 1', window.history_menu.actions()[0].text())
            self.close(window)

    def test_opinion_tab(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            window.tabs.setCurrentWidget(window.opinion_tab)
            standard = [i for i in window.opinion_items if i.kind == 'standard']
            self.assertGreaterEqual(len(standard), 5)                       # 標準文字一律列出
            self.assertEqual(window.opinion_list.count(), len(window.opinion_items))
            self.assertTrue(window.buttons['opinion_write'].isEnabled())
            text = window.opinion_edit.toPlainText()
            self.assertIn('淨高', text)
            # 勾選第 6 個標準文字（特定人員使用）：即時加入
            special = window.opinion_items.index(standard[-1])
            window.opinion_list.item(special).setCheckState(screen.Qt.CheckState.Checked)
            self.assertIn('特定人員使用', window.opinion_edit.toPlainText())
            # 勾選有待選與待填的標準文字（護欄）：下拉選定、填寫後括號消失
            guard = window.opinion_items.index(standard[0])
            window.opinion_list.item(guard).setCheckState(screen.Qt.CheckState.Checked)
            self.assertIn('【有／無】', window.opinion_inputs)
            window.opinion_inputs['【有／無】'].setCurrentText('無')
            window.opinion_inputs['【加高／重作】'].setCurrentText('加高')
            if '【道路等級】' in window.opinion_inputs:
                window.opinion_inputs['【道路等級】'].setText('鄉道')
            line = next(l for l in window.opinion_edit.toPlainText().splitlines() if '護欄高度不足110cm' in l)
            self.assertIn('現況無人行道', line)
            self.assertIn('建議採加高方式改善', line)
            self.assertNotIn('【', line)
            # 直接修改第一條後，取消勾選其他條目：修改保留、條目同步減少
            lines = window.opinion_edit.toPlainText().splitlines()
            lines[0] = '1. 手動修改的第一條'
            window.opinion_edit.setPlainText('\n'.join(lines))
            window.opinion_list.item(special).setCheckState(screen.Qt.CheckState.Unchecked)
            now = window.opinion_edit.toPlainText()
            self.assertTrue(now.startswith('1. 手動修改的第一條'))
            self.assertNotIn('特定人員使用', now)
            # 建議插入：游標所在條目改成「建議一年內派工修復」
            cursor = window.opinion_edit.textCursor()
            cursor.setPosition(0)
            window.opinion_edit.setTextCursor(cursor)
            window.insert_advice('建議一年內派工修復')
            self.assertTrue(window.opinion_edit.toPlainText().startswith('1. 手動修改的第一條，建議一年內派工修復。'))
            window.insert_advice('建議持續追蹤')
            self.assertTrue(window.opinion_edit.toPlainText().startswith('1. 手動修改的第一條，建議持續追蹤。'))
            # 意見框刪掉一條：對應條目取消勾選、自動重新編號（雙向同步）
            lines = window.opinion_edit.toPlainText().splitlines()
            numbered = [line for line in lines if re.match(r'^\d+\.', line)]
            if len(numbered) >= 2:
                before = sum(i.checked for i in window.opinion_items)
                window.opinion_edit.setPlainText('\n'.join(line for line in lines if line != numbered[1]))  # 刪第 2 條
                self.wait(200)
                self.assertEqual(sum(i.checked for i in window.opinion_items), before - 1)
                self.assertTrue(screen.numbering_ok(window.opinion_edit.toPlainText()))
                self.assertTrue(window.opinion_edit.toPlainText().startswith('1. 手動修改的第一條'))
            # 有【…】未選定時不能寫入
            window.opinion_edit.setPlainText(window.opinion_edit.toPlainText() + '\n9. 待填【＿m】')
            asked = []
            window.confirm_opinion = lambda text: asked.append(text) or True
            screen.QMessageBox.information = staticmethod(lambda *args: None)
            try:
                self.assertFalse(window.write_opinion())
            finally:
                screen.QMessageBox.information = staticmethod(self._no_dialog)
            self.assertEqual(asked, [])
            saved_text = window.opinion_edit.toPlainText()
            # 重新開啟：草稿接續
            self.close(window)
            window = self.open_window(folder)
            window.load_saved(Path(folder) / 'scan' / '2026-11-05-測試橋-檢測.json')
            window.set_overall_filter('全部')
            window.tabs.setCurrentWidget(window.opinion_tab)
            reopened = window.opinion_edit.toPlainText()
            self.assertTrue(reopened.startswith('1. 手動修改的第一條，建議持續追蹤。'))
            self.assertIn('. 待填【＿m】', reopened)
            self.assertEqual(len(reopened.splitlines()), len(saved_text.splitlines()))
            self.close(window)

    def test_prior_wrong_choice(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
            window.set_overall_filter('全部')
            pair = next(p for p in window.pairs if p.prior)
            window.open_card(window.visible.index(pair))
            options = [window.card.kind.itemText(i) for i in range(window.card.kind.count())]
            self.assertIn('前期選擇或填寫錯誤', options)
            window.card.kind.setCurrentText('前期選擇或填寫錯誤')
            window.card.decided.emit('前期選擇或填寫錯誤', '前期單價填錯')
            self.assertEqual(window.review.decision(pair).kind, '前期選擇或填寫錯誤')
            window.filter.setCurrentText('前期選擇或填寫錯誤')
            self.assertEqual([p.key for p in window.visible], [pair.key])
            self.close(window)

    def test_basic_info_and_supervisors(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.agency = '南投縣政府'
            window.scan_and_match()
            window.set_overall_filter('全部')
            window.current_header.update({'單位主管': '', '有無沖刷現象': ''})
            window.prior_header.update({'近端定義': '里程樁號較小端', '有無沖刷現象': '無', '沖刷位置/備註': 'P1'})
            window.tabs.setCurrentWidget(window.opinion_tab)
            values = window.basic_values()
            self.assertEqual(values['單位主管'], '顏宏宇')                      # 依管理機關（預設對照）
            self.assertEqual(values['近端定義'], '里程樁號較小端')              # 沿用前期
            self.assertEqual(values['沖刷位置/備註'], '')                       # 無沖刷不填
            self.assertFalse(window.basic_inputs['沖刷位置/備註'].isEnabled())
            self.assertIn('單位主管', window.basic_note.text())
            # 單位主管設定：改成王小明負責南投縣政府
            table = screen.QTableWidget(1, 2)
            table.setItem(0, 0, screen.QTableWidgetItem('王小明'))
            table.setItem(0, 1, screen.QTableWidgetItem('南投縣政府、苗栗縣政府'))
            self.assertEqual(window.save_supervisors(table), {'王小明': ['南投縣政府', '苗栗縣政府']})
            saved = json.loads((Path(folder) / 'supervisors.json').read_text(encoding='utf-8'))
            self.assertEqual(saved, {'王小明': ['南投縣政府', '苗栗縣政府']})
            self.assertEqual(window.basic_values()['單位主管'], '王小明')
            self.assertIn('chiefs', window.menu_actions)
            self.close(window)

    def test_loading_progress_bar(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.show_loading('本期網站', 40)
            self.assertTrue(window.load_bar.isVisible())
            self.assertEqual((window.load_bar.value(), window.load_label.text()), (40, '本期網站載入中'))
            window.show_loading('本期網站', None)
            self.assertFalse(window.load_bar.isVisible())
            self.close(window)

    def test_steps_and_other_functions_menu(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            self.assertEqual([window.tabs.tabText(i) for i in range(5)], list(screen.STEPS))
            self.assertEqual(set(window.menu_actions), {'定期檢測查詢', '基本資料搜尋', 'scan', 'open', 'temp', 'chiefs', 'schedule', 'login'})
            self.assertFalse(window.buttons['export'].isEnabled())      # 尚未配對不能輸出
            self.assertFalse(window.buttons['send_all'].isEnabled())    # 尚未配對不能填寫
            self.close(window)

    def test_site_panel_can_be_hidden(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            self.assertFalse(window.prior_view.isVisible())          # 預設只顯示本期
            self.assertEqual(window.buttons['site_mode'].text(), '網站：本期')
            window.site_actions['both'].trigger()
            self.assertTrue(window.prior_view.isVisible() and window.view.isVisible())
            sites, work = window.splitter.sizes()
            self.assertGreater(sites, work * 2)                      # 兩個網站時網站區約佔 73%
            window.toggle_site()                                     # 隱藏全部
            self.assertFalse(window.prior_view.isVisible() or window.view.isVisible())
            self.assertEqual(window.buttons['site'].text(), '顯示網站')
            self.assertEqual(window.buttons['site_mode'].text(), '網站：前期＋本期')
            window.toggle_site()                                     # 顯示：恢復前期＋本期
            self.assertTrue(window.prior_view.isVisible() and window.view.isVisible())
            window.site_actions['prior'].trigger()
            self.assertFalse(window.view.isVisible())
            self.assertTrue(window.prior_view.isVisible())
            # 預設寬度：一般螢幕依比例；大螢幕每個網站最多到 100% 縮放所需寬度，其餘給工作區
            self.assertEqual(screen.site_width(1450, 'both'), int(1450 * 0.73))
            self.assertEqual(screen.site_width(3840, 'both'), 2 * (860 + 40))
            self.assertEqual(screen.site_width(3840, 'current'), 900)
            # 拖動分隔線後記住比例
            total = sum(window.splitter.sizes())
            window.splitter.setSizes([int(total * 0.4), total - int(total * 0.4)])
            window.remember_layout()
            window.site_actions['both'].trigger()
            window.site_actions['prior'].trigger()
            sites, work = window.splitter.sizes()
            self.assertAlmostEqual(sites / (sites + work), 0.4, delta=0.03)
            window.site_actions['current'].trigger()
            self.wait(300)
            self.assertAlmostEqual(window.view.zoomFactor(), screen.fit_zoom(window.view.width()), places=2)
            self.assertIn(f'本期 {round(window.view.zoomFactor() * 100)}%', window.zoom_label.text())
            self.assertEqual((screen.fit_zoom(400), screen.fit_zoom(731), screen.fit_zoom(2000)), (0.6, 0.85, 1.0))
            self.close(window)


if __name__ == '__main__':
    unittest.main()
