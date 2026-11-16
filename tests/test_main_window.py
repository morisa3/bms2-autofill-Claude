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
from test_scan_window import MockSite, start_mock_site

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
                                   output_dir=Path(folder) / 'output', rules_path=self.rules_copy(folder))
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
            self.assertIs(window.tabs.currentWidget(), window.check_tabs)
            self.assertIs(window.check_tabs.currentWidget(), window.card)
            self.assertIn(KIND_GONE, [window.card.kind.itemText(i) for i in range(window.card.kind.count())])
            window.card.kind.setCurrentText(KIND_GONE)
            window.card.note.setText('河水沖刷')
            window.card.decided.emit(window.card.kind.currentText(), window.card.note.text())
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
            self.assertEqual(set(saved), {'u', 'remark'})                # 只存與網站不同的欄位
            self.assertIn('✎', window.table.item(window.visible.index(pair), 0).text())
            card.restore()
            self.assertNotIn(pair.key, window.review.edits)

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

            window.show_photos('main')                               # 回到缺失主頁面
            self.wait(500)
            self.assertIn('der_content/inspect_id,9/', window.view.url().toString())
            new = next((p for p in window.pairs if p.current and not p.prior), None)
            if new:                                                  # 防呆：新增劣化沒有前期照片
                window.open_card(window.visible.index(new))
                self.assertFalse(card.photo_buttons['prior'].isEnabled())
                self.assertFalse(card.photo_buttons['compare'].isEnabled())

            # 全部審查完回到總表
            for p in window.pairs:
                window.review.decide(p, window.review.decision(p).kind)
            window.check_tabs.setCurrentWidget(card)
            window.move(1)
            self.assertEqual(window.check_tabs.currentIndex(), 0)

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
            window.tabs.setCurrentWidget(window.fill_tab)
            self.assertEqual(len(window.fill_items), sum(1 for p in window.pairs if p.current))
            self.assertEqual(window.fill_table.rowCount(), len(window.fill_items))
            pair = next(p for p in window.pairs if p.current)
            window.review.set_edit(pair, {'remark': '人工修改。(1處)'})
            window.build_fill_plan()
            item = next(i for i in window.fill_items if i.pair is pair)
            self.assertEqual(item.status, '可填寫')
            self.assertIn('人工修改', item.describe())
            window.confirm_fill = lambda items: True
            MockSite.requests.clear()
            window.run_fill([item])                      # 模擬網站的表單沒有「送出更新」：不送出、記為失敗
            self.assertEqual(item.status, '失敗')
            self.assertRegex(item.message, '未送出|沒有送出')            # 簡化表單：安全停止
            self.assertFalse([r for r in MockSite.requests if 'update_damage' in r])
            self.assertIn('填寫失敗', window.output.toPlainText())
            self.close(window)

    def test_keep_and_rescan_after_fill(self):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
            window = self.open_window(folder)
            window.scan_and_match()
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
            window.tabs.setCurrentWidget(window.fill_tab)
            window.fill_tab.setCurrentWidget(window.opinion_tab)
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
            window.tabs.setCurrentWidget(window.fill_tab)
            window.fill_tab.setCurrentWidget(window.opinion_tab)
            reopened = window.opinion_edit.toPlainText()
            self.assertTrue(reopened.startswith('1. 手動修改的第一條，建議持續追蹤。'))
            self.assertIn('. 待填【＿m】', reopened)
            self.assertEqual(len(reopened.splitlines()), len(saved_text.splitlines()))
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
            self.assertEqual([window.tabs.tabText(i) for i in range(3)], list(screen.STEPS))
            self.assertEqual(set(window.menu_actions), {'定期檢測查詢', '基本資料搜尋', 'scan', 'open', 'temp'})
            self.assertFalse(window.buttons['export'].isEnabled())      # 尚未配對不能輸出
            self.assertFalse(window.buttons['fill_all'].isEnabled())    # 尚未配對不能填寫
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
            self.assertEqual((screen.fit_zoom(400), screen.fit_zoom(731), screen.fit_zoom(2000)), (0.6, 0.85, 1.0))
            self.close(window)


if __name__ == '__main__':
    unittest.main()
