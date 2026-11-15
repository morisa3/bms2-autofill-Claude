"""公路橋梁填報助手：主視窗（第 5 步，版型 A 外框＋B 逐筆審查／C 總表）。

使用者只要登入網站、輸入橋名按「開始」，之後分三個步驟：
① 掃描＋核對：查詢橋梁 → 唯讀掃描本期、前期與基本資料 → 前後期配對；「總表」與「逐筆審查」確認或改判，
   審查結果自動存檔、可接續。
② 填寫：單筆修正或整筆自動（第 6 步完成後開放）。
③ 比對資料輸出：把配對與審查結果輸出成 Excel，存到「判讀結果」資料夾。
定期檢測查詢、基本資料搜尋、掃描目前頁面、開啟掃描檔、暫存資料夾等放在「其他功能」選單。
目前全程唯讀，不修改網站。

用法：
    python -m bridgefill.ui.main_window
"""
import json
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QUrl, Qt, Signal
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMenu, QMessageBox, QPlainTextEdit, QPushButton, QToolButton,
                               QSplitter, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from ..core.local_files import APP_DIR
from ..core.manual import apply_manual, candidates
from ..core.matching import defect_id, match_scans
from ..core.report import report_name, write_report
from ..core.review import ACCEPTED, CHANGED, PENDING, Review, choices_for, review_path
from ..site.pages import group_bridges, needs_choice, parse_history, parse_inspection
from ..site.scan import Stopped, history_url, inventory_id, scan_basic, scan_inspection, scan_prior
from ..site.session import LOGIN_URL
from .browser import PROFILE_DIR, SHORTCUTS, SiteView, make_profile
from .fetcher import FetchError, PageFetcher

SCAN_DIR = APP_DIR / 'temp' / 'scan'
REVIEW_DIR = APP_DIR / 'temp' / 'review'
OUTPUT_DIR = APP_DIR / '判讀結果'
STEPS = ('① 掃描＋核對', '② 填寫', '③ 比對資料輸出')
TABLE_COLUMNS = ('審查', '構件', '位置／編號', '判定', '信心', '前期備註', '本期備註', '理由')
FIELDS = (('DERU', None), ('劣化類型', 'deterioration'), ('損壞位置', 'd_location'), ('維修工法', 'method'),
          ('數量', 'amount'), ('單位', 'unit'), ('單價', 'price'), ('備註', 'remark'))
FILTERS = ('全部', '待審查', '需確認', '已改判', '人工配對', '位置變更', '新增', '疑似漏建', '單位數量變更', '範圍加大', '範圍縮小', '沿用', '已維修')
STATUS_COLORS = {PENDING: '#fff2cc', CHANGED: '#ddebf7', ACCEPTED: None}
DIFFERENT = '#fce4d6'


def field_value(defect, key):
    if defect is None:
        return ''
    if key is None:
        return ' '.join(defect.deru)
    return str(defect.fields.get(key, '') or '').strip()


class ReviewCard(QWidget):
    """逐筆審查（版型 B）：前期與本期並排，不同的欄位標色；可確認或改判。"""
    decided = Signal(str, str)      # 判定、說明
    moved = Signal(int)             # -1 上一筆、+1 下一筆
    relink = Signal()               # 改配對

    def __init__(self, parent=None):
        super().__init__(parent)
        self.title = QLabel('尚未選擇')
        self.title.setStyleSheet('font-size:16px; font-weight:bold;')
        self.reasons = QLabel('')
        self.reasons.setWordWrap(True)
        grid = QGridLayout()
        grid.addWidget(QLabel(''), 0, 0)
        for col, text in ((1, '前期'), (2, '本期')):
            head = QLabel(text)
            head.setStyleSheet('font-weight:bold;')
            grid.addWidget(head, 0, col)
        self.cells = {}
        for row, (label, key) in enumerate(FIELDS, 1):
            grid.addWidget(QLabel(label), row, 0, Qt.AlignmentFlag.AlignTop)
            for col in (1, 2):
                cell = QLabel('')
                cell.setWordWrap(True)
                cell.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                cell.setFrameShape(QFrame.Shape.StyledPanel)
                cell.setMinimumHeight(28)
                grid.addWidget(cell, row, col)
                self.cells[(label, col)] = cell
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(2, 1)
        self.kind = QComboBox()
        self.note = QLineEdit()
        self.note.setPlaceholderText('說明（選填，例如：河水沖刷，已無掛淤）')
        self.status = QLabel('')
        accept = QPushButton('確認並下一筆')
        accept.clicked.connect(lambda: self.decided.emit(self.kind.currentText(), self.note.text().strip()))
        relink = QPushButton('改配對…')
        relink.setToolTip('配錯對象、或同一筆缺失前後期記在不同構件（例：河道掛淤改記在橋墩）時，手動指定對應的缺失')
        relink.clicked.connect(self.relink.emit)
        back, forward = QPushButton('上一筆'), QPushButton('下一筆')
        back.clicked.connect(lambda: self.moved.emit(-1))
        forward.clicked.connect(lambda: self.moved.emit(1))
        self.pending_only = QCheckBox('只在待審查之間移動')
        self.pending_only.setChecked(True)
        actions = QHBoxLayout()
        for widget in (QLabel('判定'), self.kind, self.note, relink):
            actions.addWidget(widget)
        actions.setStretch(2, 1)
        moves = QHBoxLayout()
        for widget in (back, forward, self.pending_only):
            moves.addWidget(widget)
        moves.addStretch()
        moves.addWidget(self.status)
        moves.addWidget(accept)
        layout = QVBoxLayout(self)
        layout.addWidget(self.title)
        layout.addWidget(self.reasons)
        layout.addLayout(grid)
        layout.addSpacing(8)
        layout.addLayout(actions)     # 操作列緊接在欄位下方，不必上下來回看
        layout.addLayout(moves)
        layout.addStretch()

    def show_pair(self, pair, decision, position):
        any_side = pair.current or pair.prior
        component = f'{pair.prior.component} → {pair.current.component}' if pair.moved else any_side.component
        self.title.setText(f'{component}｜{any_side.label or "—"}｜程式判定：{pair.kind}（{pair.confidence}）'
                           f'　{position}')
        self.reasons.setText('理由：' + ('；'.join(pair.reasons) if pair.reasons else '—'))
        for label, key in FIELDS:
            before, after = field_value(pair.prior, key), field_value(pair.current, key)
            changed = pair.prior is not None and pair.current is not None and before != after
            for col, text in ((1, before or ('（無）' if pair.prior is None else '')),
                              (2, after or ('（無）' if pair.current is None else ''))):
                cell = self.cells[(label, col)]
                cell.setText(text)
                cell.setStyleSheet(f'background:{DIFFERENT};' if changed else '')
        self.kind.clear()
        self.kind.addItems(choices_for(pair))
        self.kind.setCurrentText(decision.kind)
        self.note.setText(decision.note)
        self.status.setText(f'審查：{decision.status}')


class MainWindow(QMainWindow):
    def __init__(self, url=LOGIN_URL, profile_dir=PROFILE_DIR, scan_dir=SCAN_DIR, review_dir=REVIEW_DIR,
                 delay_ms=500, rewrite=None, output_dir=OUTPUT_DIR):
        super().__init__()
        self.setWindowTitle('公路橋梁填報助手')
        self.resize(1680, 960)
        self.scan_dir, self.review_dir, self.output_dir = Path(scan_dir), Path(review_dir), Path(output_dir)
        self.pairs, self.visible, self.review, self.review_file = [], [], Review(), None
        self.auto_pairs, self.order = [], {}
        self.index = -1
        self.busy = self.stop_requested = False
        self.profile = make_profile(profile_dir, self)
        self.view = SiteView(self.profile, self)
        self.fetcher = PageFetcher(self.profile, self, delay_ms=delay_ms, rewrite=rewrite)

        # 上方：橋名 → 開始；旁邊是隱藏網站、開啟判讀結果資料夾；其餘放「其他功能」
        bar = QHBoxLayout()
        self.bridge_name = QLineEdit()
        self.bridge_name.setPlaceholderText('例如：產業一號橋（按 Enter 或「開始」後自動掃描）')
        self.bridge_name.setMinimumWidth(220)
        self.bridge_name.returnPressed.connect(lambda: self.run_by_name())
        hint = QLabel('輸入後將自動進行掃描')
        hint.setStyleSheet('color:#666;')
        self.start_button = QPushButton('開始')
        self.start_button.setStyleSheet('font-weight:bold;')
        self.start_button.clicked.connect(lambda: self.run_by_name())
        bar.addWidget(QLabel('橋名'))
        bar.addWidget(self.bridge_name)
        bar.addWidget(self.start_button)
        self.buttons = {}
        for key, text, slot in (('site', '隱藏網站', self.toggle_site),
                                ('results', '開啟判讀結果資料夾', self.open_results),
                                ('stop', '停止', self.stop)):
            button = QPushButton(text)
            button.clicked.connect(slot)
            bar.addWidget(button)
            self.buttons[key] = button
        self.buttons['stop'].setEnabled(False)
        bar.addWidget(hint)
        bar.addStretch()
        self.more = QMenu(self)
        self.menu_actions = {}
        for text, target in SHORTCUTS.items():
            self.menu_actions[text] = self.more.addAction(text, lambda u=target: self.view.open(u))
        self.more.addSeparator()
        for key, text, slot in (('scan', '掃描目前頁面（左側開著的本期缺失主頁面）', self.scan_and_match),
                                ('open', '開啟掃描檔…', self.open_saved),
                                ('temp', '開啟暫存資料夾', self.open_folder)):
            self.menu_actions[key] = self.more.addAction(text, slot)
        more = QToolButton()
        more.setText('其他功能')
        more.setMenu(self.more)
        more.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        bar.addWidget(more)
        self.summary = QLabel('登入網站後，輸入橋名按「開始」：程式自動查詢、找出本期與前期、掃描並配對。目前全程唯讀，不修改網站。')
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet('background:#eef4fb; padding:6px; border-radius:6px;')

        # ① 掃描＋核對：總表、逐筆審查
        self.filter = QComboBox()
        self.filter.addItems(FILTERS)
        self.filter.currentTextChanged.connect(self.refresh_table)
        self.table = QTableWidget(0, len(TABLE_COLUMNS))
        self.table.setHorizontalHeaderLabels(TABLE_COLUMNS)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setWordWrap(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for col, width in enumerate((60, 110, 150, 90, 50, 220, 220, 260)):
            self.table.setColumnWidth(col, width)
        header.setStretchLastSection(True)
        self.table.cellDoubleClicked.connect(lambda row, _: self.open_card(row))
        overview = QWidget()
        top = QHBoxLayout()
        top.addWidget(QLabel('篩選'))
        top.addWidget(self.filter)
        top.addWidget(QLabel('（雙擊一列開啟逐筆審查）'))
        top.addStretch()
        box = QVBoxLayout(overview)
        box.addLayout(top)
        box.addWidget(self.table)
        self.card = ReviewCard()
        self.card.decided.connect(self.decide)
        self.card.moved.connect(self.move)
        self.card.relink.connect(self.relink)
        self.check_tabs = QTabWidget()
        self.check_tabs.addTab(overview, '總表')
        self.check_tabs.addTab(self.card, '逐筆審查')

        # ② 填寫（第 6 步完成後開放）
        fill = QWidget()
        fill_box = QVBoxLayout(fill)
        fill_box.addWidget(QLabel('依審查後的判定填寫網站：'))
        for text, tip in (('單筆修正', '在總表選一筆，只填寫該筆'), ('整筆自動', '依審查結果逐筆填寫，填完重新讀取核對')):
            button = QPushButton(text)
            button.setEnabled(False)
            button.setToolTip(tip)
            fill_box.addWidget(button)
            self.buttons[text] = button
        note = QLabel('填寫功能開發中（第 6 步）：會以建議維修工法 ID 鎖定每一筆，填完重新讀取核對，不重複送出。')
        note.setWordWrap(True)
        note.setStyleSheet('color:#666;')
        fill_box.addWidget(note)
        fill_box.addStretch()

        # ③ 比對資料輸出
        export = QWidget()
        export_box = QVBoxLayout(export)
        explain = QLabel('把前後期配對與審查結果輸出成 Excel（含程式判定、最終判定、審查說明），存到「判讀結果」資料夾。')
        explain.setWordWrap(True)
        self.buttons['export'] = QPushButton('輸出比對 Excel')
        self.buttons['export'].clicked.connect(lambda: self.export_report())
        self.buttons['export'].setEnabled(False)
        self.export_status = QLabel('')
        self.export_status.setWordWrap(True)
        self.export_status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        for widget in (explain, self.buttons['export'], self.export_status):
            export_box.addWidget(widget)
        export_box.addStretch()

        self.output = QPlainTextEdit(readOnly=True)
        self.output.setFont(QFont('Consolas', 9))
        self.tabs = QTabWidget()
        for widget, text in zip((self.check_tabs, fill, export), STEPS):
            self.tabs.addTab(widget, text)
        self.tabs.addTab(self.output, '紀錄')

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.addWidget(self.view)
        self.splitter.addWidget(self.tabs)
        self.splitter.setSizes([980, 700])
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addLayout(bar)
        layout.addWidget(self.summary)
        layout.addWidget(self.splitter, 1)
        self.setCentralWidget(root)
        self.view.open(url)

    # ------------------------------------------------------------ 共用
    def log(self, message):
        self.output.appendPlainText(f'{datetime.now():%H:%M:%S} {message}')

    def set_busy(self, busy):
        self.busy = busy
        for key in ('scan', 'open'):
            self.menu_actions[key].setEnabled(not busy)
        self.buttons['export'].setEnabled(not busy and bool(self.pairs))
        self.start_button.setEnabled(not busy)
        self.buttons['stop'].setEnabled(busy)

    def stop(self):
        self.stop_requested = True
        self.log('停止中，完成目前這一頁後停止…')

    def toggle_site(self):
        hidden = self.view.isVisible()
        self.view.setVisible(not hidden)
        self.buttons['site'].setText('顯示網站' if hidden else '隱藏網站')

    def open_folder(self):
        self.scan_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.scan_dir.parent)))

    def open_results(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_dir)))

    def export_report(self):
        """③ 比對資料輸出：配對與審查結果 → 判讀結果資料夾的 Excel。回傳路徑。"""
        if not self.pairs:
            QMessageBox.information(self, '比對資料輸出', '請先輸入橋名按「開始」完成掃描與配對。')
            return None
        target = self.output_dir / report_name(self.review.bridge, self.review.current_date)
        try:
            write_report(self.pairs, target, self.review)
        except PermissionError:
            QMessageBox.warning(self, '比對資料輸出', f'無法寫入 {target.name}，檔案可能正在 Excel 中開啟，請關閉後再試。')
            return None
        self.export_status.setText(f'已輸出：{target}')
        self.log(f'比對資料輸出：{target}')
        return target

    def current_html(self):
        box = {}
        self.view.page().toHtml(lambda html: box.setdefault('html', html))
        if not self.fetcher.wait_until(lambda: 'html' in box, timeout=10000, interval=100):
            raise FetchError('讀取目前頁面逾時')
        return box['html']

    def save_json(self, header, kind, data):
        self.scan_dir.mkdir(parents=True, exist_ok=True)
        name = f"{header.get('檢測日期', '')}-{header.get('橋梁名稱', '')}-{kind}.json"
        path = self.scan_dir / ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in name)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        return path

    # ------------------------------------------------------------ 掃描並配對
    def scan_and_match(self):
        """掃描左側網站目前開著的本期檢測缺失主頁面。"""
        if self.busy:
            return
        try:
            html = self.current_html()
        except FetchError as error:
            self.log(str(error))
            return
        self.run_pipeline(html)

    def run_by_name(self, name=None):
        """輸入橋名即完成：查詢 → 確認橋梁 → 本期／前期 → 掃描 → 配對。"""
        name = (name if name is not None else self.bridge_name.text()).strip()
        if self.busy:
            return
        if not name:
            QMessageBox.information(self, '開始', '請先輸入橋名。')
            return
        self.stop_requested = False
        self.set_busy(True)
        try:
            self.log(f'查詢橋名：{name}')
            bridges = group_bridges(parse_history(self.fetcher.search(name)))
            if not bridges:
                raise ValueError(f'查無「{name}」的定期檢測資料，請確認橋名或改用部分名稱')
            bridge = bridges[0] if not needs_choice(bridges, name) else self.choose_bridge(name, bridges)
            if bridge is None:
                self.log('已取消')
                return
            self.log(f'橋梁：{bridge.name}（橋梁系統ID {bridge.inventory_id}）')
            history = parse_history(self.fetcher.get(history_url(bridge.inventory_id))) or bridge.records
            current = history[0]
            if not current.edit_url:
                raise ValueError(f'最新一期 {current.date} 沒有「修改」連結，無法讀取（可能不是本公司的檢測）')
            self.log(f'本期：{current.date}（檢測系統ID {current.inspect_id}）')
            html = self.fetcher.get(current.edit_url)
            self.view.open(current.edit_url)          # 左側同步顯示本期缺失主頁面
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'中斷：{error}')
            QMessageBox.warning(self, '開始', str(error))
            return
        finally:
            self.set_busy(False)
        self.run_pipeline(html)

    def choose_bridge(self, name, bridges):
        """查詢到多座相似的橋（例：龍門大橋、龍門大橋(2)）時請使用者選擇；取消回傳 None。"""
        dialog = QDialog(self)
        dialog.setWindowTitle('請確認橋梁')
        listing = QListWidget()
        for bridge in bridges:
            latest = bridge.latest.date if bridge.latest else '—'
            listing.addItem(f'{bridge.name}｜橋梁系統ID {bridge.inventory_id}｜{bridge.agency}｜{bridge.place}｜'
                            f'最新檢測 {latest}')
        exact = [i for i, b in enumerate(bridges) if b.name == name]
        listing.setCurrentRow(exact[0] if exact else 0)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        listing.itemDoubleClicked.connect(lambda _: dialog.accept())
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f'「{name}」查到 {len(bridges)} 座橋，請選擇要處理的橋梁：'))
        layout.addWidget(listing)
        layout.addWidget(buttons)
        dialog.resize(720, 320)
        return bridges[listing.currentRow()] if dialog.exec() == QDialog.DialogCode.Accepted else None

    def run_pipeline(self, html):
        """本期主頁面 HTML → 掃描本期、前期、基本資料 → 配對。"""
        page = parse_inspection(html)
        if not page.rows or not page.header.get('檢測日期'):
            QMessageBox.information(self, '掃描並配對', '目前頁面不是「檢測缺失主頁面」。\n'
                                    '請在上方輸入橋名按「開始」，或在左側網站開到本期的檢測缺失主頁面。')
            return
        header = page.header
        self.stop_requested = False
        self.set_busy(True)
        started = datetime.now()
        options = dict(progress=self.progress, stopped=lambda: self.stop_requested)
        try:
            self.log(f"掃描本期：{header['橋梁名稱']} {header['檢測日期']}")
            current = scan_inspection(html, self.fetcher, **options)
            self.save_json(header, '檢測', current.to_dict())
            self.log('由歷史紀錄找出前期並掃描')
            record, prior = scan_prior(html, self.fetcher, **options)
            self.save_json(prior.header, '前期', prior.to_dict())
            self.log(f'前期：{record.date}（{record.inspectors}）')
            inventory = inventory_id(page)
            basic = scan_basic(inventory, self.fetcher, self.fetcher.call,
                               expected_spans=int(header.get('總橋孔數') or 0) or None) if inventory else {}
            if basic:
                self.save_json(header, '基本資料', basic)
                for problem in basic.get('problems', []):
                    self.log(f'基本資料問題：{problem}')
        except Stopped as error:
            self.log(str(error))
            return
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'中斷：{error}')
            QMessageBox.warning(self, '掃描並配對', f'掃描中斷：{error}')
            return
        finally:
            self.set_busy(False)
        failed = [r for r in current.rows + prior.rows if not r.ok]
        for row in failed:
            self.log(f'核對異常：{row.component} {row.label}：{row.check}')
        self.log(f'掃描完成（{(datetime.now() - started).total_seconds():.0f} 秒），核對異常 {len(failed)} 列')
        self.load_pairs(prior.to_dict(), current.to_dict())

    def progress(self, index, total, label):
        self.statusBar().showMessage(f'掃描中 {index}／{total}：{label}')

    def open_saved(self):
        start = str(self.scan_dir if self.scan_dir.exists() else APP_DIR)
        path, _ = QFileDialog.getOpenFileName(self, '選擇本期掃描檔', start, '本期掃描檔 (*-檢測.json)')
        if path:
            self.load_saved(path)

    def load_saved(self, current_path):
        current_path = Path(current_path)
        current = json.loads(current_path.read_text(encoding='utf-8'))
        bridge = current.get('header', {}).get('橋梁名稱', '')
        priors = sorted(current_path.parent.glob(f'*-{bridge}-前期.json'))
        if not priors:
            QMessageBox.information(self, '開啟掃描檔', f'找不到 {bridge} 的前期掃描檔（*-{bridge}-前期.json）。')
            return
        prior = json.loads(priors[-1].read_text(encoding='utf-8'))
        self.log(f'開啟掃描檔：{current_path.name}、{priors[-1].name}')
        self.load_pairs(prior, current)

    # ------------------------------------------------------------ 配對與審查
    def load_pairs(self, prior_scan, current_scan):
        result = match_scans(prior_scan, current_scan)
        self.order = {(r['keys'].get('component_form'), r['keys'].get('component_id')): i
                      for i, r in enumerate(current_scan.get('rows', []))}
        self.auto_pairs = result.pairs
        head = current_scan.get('header', {})
        self.review_file = review_path(self.review_dir, head.get('橋梁名稱', ''), head.get('檢測日期', ''))
        self.review = Review.load(self.review_file) if self.review_file.exists() else Review(
            head.get('橋梁名稱', ''), head.get('檢測日期', ''), prior_scan.get('header', {}).get('檢測日期', ''))
        self.apply_review()
        self.log('配對：' + '、'.join(f'{k} {v}' for k, v in result.summary.items()))
        self.refresh_table()
        self.update_summary()
        self.buttons['export'].setEnabled(True)
        self.tabs.setCurrentWidget(self.check_tabs)
        self.check_tabs.setCurrentIndex(0)

    def apply_review(self):
        """自動配對＋人工指定 → 目前的配對清單；同步審查結果並存檔。"""
        pairs = apply_manual(self.auto_pairs, self.review.links, self.review.splits)
        self.pairs = sorted(pairs, key=lambda p: (self.order.get(p.row, 9999), p.prior is None, p.current is None))
        self.review.sync(self.pairs)
        self.review.save(self.review_file)

    def relink(self):
        """改配對：列出候選（同一構件列 → 同構件其他列 → 其他構件未配對 → 已配給其他缺失），可一對多、可拆開。"""
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        chosen = self.choose_link(pair, candidates(self.pairs, pair))
        if chosen is None:
            return
        if chosen == 'split':
            self.review.split(self.pairs, pair)
            self.log(f'拆開配對：{(pair.prior.remark or "").strip()}／{(pair.current.remark or "").strip()}')
        elif pair.prior:
            self.review.link(self.pairs, defect_id(pair.prior), [defect_id(d) for d in chosen])
            self.log(f'人工配對：前期 {pair.prior.remark.strip()} → 本期 {len(chosen)} 筆')
        else:
            prior = chosen[0]
            existing = [defect_id(p.current) for p in self.pairs if p.prior is prior and p.current]
            self.review.link(self.pairs, defect_id(prior), existing + [defect_id(pair.current)])
            self.log(f'人工配對：前期 {prior.remark.strip()} → 本期 {pair.current.remark.strip()}')
        key = pair.key if pair.current else None
        self.apply_review()
        self.refresh_table()
        self.update_summary()
        # 停在剛處理的那一筆（以本期為準；只有前期時找它新的配對）
        target = next((i for i, p in enumerate(self.pairs) if key and p.key == key), None)
        if target is None and pair.prior:
            target = next((i for i, p in enumerate(self.pairs) if p.prior is pair.prior), None)
        self.index = target if target is not None else min(self.index, len(self.pairs) - 1)
        self.show_card()

    def choose_link(self, pair, options):
        """改配對對話框。回傳：選取的缺失清單、'split'（拆開目前配對）或 None（取消）。"""
        dialog = QDialog(self)
        dialog.setWindowTitle('改配對')
        anchor = pair.prior or pair.current
        many = pair.prior is not None
        listing = QListWidget()
        tier = None
        for option in options:
            if option.tier != tier:
                tier = option.tier
                head = QListWidgetItem(f'── {option.tier_name} ──')
                head.setFlags(Qt.ItemFlag.NoItemFlags)
                listing.addItem(head)
            d = option.defect
            text = f'{d.component}｜{d.label or "—"}｜DERU {"".join(d.deru)}｜{d.remark.strip() or d.deterioration}'
            if option.selected:
                text += '　（目前配對）'
            elif option.partner:
                text += f'　〔目前配給：{option.partner}〕'
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, d)
            if option.partner:
                item.setForeground(QColor('#888888'))
            if many:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if option.selected else Qt.CheckState.Unchecked)
            listing.addItem(item)
        side = '本期' if many else '前期'
        tip = (f'勾選要對應的本期缺失（可勾多筆：一筆前期拆成多筆本期）。' if many else
               '選一筆要對應的前期缺失；若那筆前期已配給其他本期，會變成一對多。')
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f'{anchor.period}：{anchor.component}｜{anchor.remark.strip()}'))
        layout.addWidget(QLabel(f'{tip}\n灰字為已配給其他缺失，選了會把原配對拆開。'))
        layout.addWidget(listing)
        buttons = QDialogButtonBox()
        ok = buttons.addButton('配對', QDialogButtonBox.ButtonRole.AcceptRole)
        split = buttons.addButton('拆開目前配對', QDialogButtonBox.ButtonRole.ActionRole)
        split.setEnabled(bool(pair.prior and pair.current))
        buttons.addButton(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        result = {}
        ok.clicked.connect(dialog.accept)
        split.clicked.connect(lambda: (result.setdefault('split', True), dialog.accept()))
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.resize(860, 480)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        if result.get('split'):
            return 'split'
        items = [listing.item(i) for i in range(listing.count())]
        if many:
            chosen = [i.data(Qt.ItemDataRole.UserRole) for i in items
                      if i.data(Qt.ItemDataRole.UserRole) is not None and i.checkState() == Qt.CheckState.Checked]
        else:
            current = listing.currentItem()
            chosen = [current.data(Qt.ItemDataRole.UserRole)] if current and current.data(Qt.ItemDataRole.UserRole) else []
        if not chosen:
            QMessageBox.information(self, '改配對', f'沒有選取{side}缺失；要取消配對請按「拆開目前配對」。')
            return None
        return chosen

    def update_summary(self):
        if not self.pairs:
            return
        counts = self.review.counts(self.pairs)
        self.summary.setText(f'{self.review.bridge}　本期 {self.review.current_date}／前期 {self.review.prior_date}　'
                             f'共 {len(self.pairs)} 組：待審查 {counts[PENDING]}、已確認 {counts[ACCEPTED]}、'
                             f'已改判 {counts[CHANGED]}。審查結果自動存檔。')

    def matches_filter(self, pair):
        chosen = self.filter.currentText()
        decision = self.review.decision(pair)
        if chosen == '全部':
            return True
        if chosen in (PENDING, CHANGED):
            return decision.status == chosen
        if chosen == '需確認':
            return pair.confidence == '需確認'
        if chosen == '人工配對':
            return pair.manual
        if chosen == '位置變更':
            return pair.moved or any('位置變更' in r for r in pair.reasons)
        return decision.kind == chosen

    def refresh_table(self, *_):
        self.visible = [p for p in self.pairs if self.matches_filter(p)]
        self.table.setRowCount(len(self.visible))
        for row, pair in enumerate(self.visible):
            decision = self.review.decision(pair)
            any_side = pair.current or pair.prior
            kind = decision.kind if decision.kind == pair.kind else f'{decision.kind}（原：{pair.kind}）'
            component = f'{pair.prior.component} → {pair.current.component}' if pair.moved else any_side.component
            values = (decision.status, component, any_side.label, kind, pair.confidence,
                      pair.prior.remark.strip() if pair.prior else '—',
                      pair.current.remark.strip() if pair.current else '—', '；'.join(pair.reasons))
            color = STATUS_COLORS.get(decision.status)
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if color:
                    item.setBackground(QColor(color))
                self.table.setItem(row, col, item)
        self.table.resizeRowsToContents()

    def open_card(self, row):
        if 0 <= row < len(self.visible):
            self.index = self.pairs.index(self.visible[row])
            self.show_card()

    def show_card(self):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        self.card.show_pair(pair, self.review.decision(pair), f'第 {self.index + 1}／{len(self.pairs)} 組')
        self.tabs.setCurrentWidget(self.check_tabs)
        self.check_tabs.setCurrentWidget(self.card)

    def decide(self, kind, note):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        decision = self.review.decide(pair, kind, note)
        self.review.save(self.review_file)
        self.log(f'審查：{(pair.current or pair.prior).component} {kind}（{decision.status}）')
        self.refresh_table()
        self.update_summary()
        self.move(1)

    def move(self, step):
        if not self.pairs:
            return
        index = self.index
        for _ in range(len(self.pairs)):
            index += step
            if not 0 <= index < len(self.pairs):
                return
            if not self.card.pending_only.isChecked() or self.review.decision(self.pairs[index]).status == PENDING:
                self.index = index
                self.show_card()
                return

    def closeEvent(self, event):
        self.stop_requested = True
        self.fetcher.release()
        self.view.release()
        super().closeEvent(event)


def main(argv=None):
    app = QApplication.instance() or QApplication(sys.argv[:1])
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
