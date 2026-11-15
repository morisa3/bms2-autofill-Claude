"""公路橋梁填報助手：主視窗（第 5 步，版型 A 外框＋B 逐筆審查／C 總表）。

使用者只要登入網站、輸入橋名按「開始」，之後分三個步驟：
① 掃描＋核對：查詢橋梁 → 唯讀掃描本期、前期與基本資料 → 前後期配對；「總表」與「逐筆審查」確認或改判，
   審查結果自動存檔、可接續。
② 填寫：依建議維修工法 ID 填寫網站（單筆修正或整筆自動），每筆送出後回讀核對，結果不明不重送。
③ 比對資料輸出：把配對與審查結果輸出成 Excel，存到「判讀結果」資料夾。
定期檢測查詢、基本資料搜尋、掃描目前頁面、開啟掃描檔、暫存資料夾等放在「其他功能」選單。
只有「② 填寫」會修改網站資料，其餘全程唯讀。

用法：
    python -m bridgefill.ui.main_window
"""
import json
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QUrl, Qt, Signal
from PySide6.QtGui import QAction, QActionGroup, QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMenu, QMessageBox, QPlainTextEdit, QPushButton, QToolButton,
                               QSplitter, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from ..core.fill import proposal_for
from ..core.fill import (DONE, FAILED, KEPT, PENDING, READY, REVIEW, SAME, SKIP, FillLog, build_plan,
                         fill_log_path)
from ..core.rule_add import NO_RULE, add_rules, suggest_rule, writable_rules_file
from ..core.editing import rules_file
from ..core.editing import EDIT_FIELDS, best_rule, describe, load_rules, pattern_tags, ranked_rules, replace_pattern
from ..core.local_files import APP_DIR
from ..core.manual import apply_manual, candidates
from ..core.matching import defect_id, match_scans
from ..core.report import report_name, write_report
from ..core.review import (ACCEPTED, CHANGED, PENDING, Review, choices_for, review_path, saved_sessions,
                           site_value)
from ..site.pages import group_bridges, needs_choice, parse_history, parse_inspection
from ..site.scan import Stopped, basic_urls, history_url, inventory_id, scan_basic, scan_inspection, scan_prior
from ..site.session import LOGIN_URL
from .browser import PROFILE_DIR, SHORTCUTS, SiteView, make_profile
from .fetcher import FetchError, PageFetcher
from .filler import FillError, SiteFiller

SCAN_DIR = APP_DIR / 'temp' / 'scan'
REVIEW_DIR = APP_DIR / 'temp' / 'review'
OUTPUT_DIR = APP_DIR / '判讀結果'
SITE_MODES = {'current': '只顯示本期網站', 'prior': '只顯示前期網站', 'both': '前期＋本期網站並排'}
SITE_SHARE = {'both': 0.73, 'current': 0.58, 'prior': 0.58}     # 網站區佔左右寬度的比例（其餘為工作區）
SITE_LABELS = {'current': '網站：本期', 'prior': '網站：前期', 'both': '網站：前期＋本期', 'none': '網站：已隱藏'}
STEPS = ('① 掃描＋核對', '② 填寫', '③ 比對資料輸出')
TABLE_COLUMNS = ('審查', '構件', '位置／編號', '判定', '信心', '前期備註', '本期備註', '理由')
FIELDS = (('DERU', None), ('劣化類型', 'deterioration'), ('損壞位置', 'd_location'), ('維修工法', 'method'),
          ('數量', 'amount'), ('單位', 'unit'), ('單價', 'price'), ('備註', 'remark'))
FILL_COLUMNS = ('狀態', '構件', '位置／編號', '網站 ID', '將填寫的變更', '說明')
FILL_COLORS = {READY: '#e2efda', REVIEW: '#fff2cc', DONE: '#ddebf7', PENDING: '#f8cbad', FAILED: '#f8cbad',
               KEPT: '#e4dfec'}
FILL_BUTTONS = ('fill_one', 'fill_all', 'fill_keep', 'fill_rescan', 'fill_plan')
FILTERS = ('全部', '待審查', '需確認', '已改判', '人工配對', '位置變更', '已修改', '新增', '疑似漏建', '單位數量變更', '範圍加大', '範圍縮小', '沿用', '已維修')
STATUS_COLORS = {PENDING: '#fff2cc', CHANGED: '#ddebf7', ACCEPTED: None}
DIFFERENT = '#fce4d6'
EDITED = '#ddebf7'
SCROLL_TO_PHOTOS = """(() => {
  const all = Array.from(document.querySelectorAll('h1,h2,h3,h4,div,span,td,p'));
  const head = all.find(e => e.childElementCount === 0 && /損壞照片/.test(e.textContent || ''));
  if (head) head.scrollIntoView({block: 'start'});
})()"""


def history_text(item):
    progress = ('未審查' if item['pending'] is None else
                f"待審查 {item['pending']}、已修改 {item['edits']}" if item['pending'] else f"審查完成、已修改 {item['edits']}")
    if item.get('filled') is not None:
        progress += f"｜已填寫 {item['filled']}" + (f"、待確認 {item['unsure']}" if item.get('unsure') else '')
    else:
        progress += '｜未填寫'
    return f"{item['bridge']}｜本期 {item['date']}｜{progress}｜{item['updated']:%m/%d %H:%M}"


def field_value(defect, key):
    if defect is None:
        return ''
    if key is None:
        return ' '.join(defect.deru)
    return str(defect.fields.get(key, '') or '').strip()


class ReviewCard(QWidget):
    """逐筆審查（版型 B）：前期與本期並排，不同的欄位標色；可確認或改判、修改本期欄位、查看兩期照片。"""
    decided = Signal(str, str)      # 判定、說明
    moved = Signal(int)             # -1 上一筆、+1 下一筆
    relink = Signal()               # 改配對
    photos = Signal(str)            # current／prior／compare
    edited = Signal(dict)           # 修改後的本期欄位
    kept = Signal()                 # 確認維持網站現值（填寫清單的需確認）

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rules, self.component_rules, self.pair, self.site_values = [], [], None, {}
        self.title = QLabel('尚未選擇')
        self.title.setStyleSheet('font-size:16px; font-weight:bold;')
        self.title.setWordWrap(True)
        self.reasons = QLabel('')
        self.reasons.setWordWrap(True)

        pictures = QHBoxLayout()
        self.photo_buttons = {}
        for key, text in (('current', '本期照片'), ('prior', '前期照片'), ('compare', '前後期照片比對'),
                          ('main', '回到缺失主頁面')):
            button = QPushButton(text)
            button.clicked.connect(lambda _=False, k=key: self.photos.emit(k))
            pictures.addWidget(button)
            self.photo_buttons[key] = button
        pictures.addStretch()
        self.edit_toggle = QPushButton('修改本期')
        self.edit_toggle.setCheckable(True)
        self.edit_toggle.setToolTip('修改本期的 DERU、劣化類型、樣態、維修工法、數量、單價等（第 6 步填寫時寫入網站）')
        self.edit_toggle.toggled.connect(self.set_editing)
        pictures.addWidget(self.edit_toggle)

        grid = QGridLayout()
        grid.addWidget(QLabel(''), 0, 0)
        self.edit_widgets = []
        for col, text in ((1, '前期'), (2, '本期（網站）'), (3, '修改後')):
            head = QLabel(text)
            head.setStyleSheet('font-weight:bold;')
            grid.addWidget(head, 0, col)
            if col == 3:
                self.edit_widgets.append(head)
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

        # 修改後：DERU 四格、類型與單位下拉、其餘文字框
        self.deru = []
        deru_box = QWidget()
        deru_row = QHBoxLayout(deru_box)
        deru_row.setContentsMargins(0, 0, 0, 0)
        for letter in 'DERU':
            combo = QComboBox()
            combo.addItems(['_', '0', '1', '2', '3', '4'])
            combo.setToolTip(letter)
            deru_row.addWidget(QLabel(letter))
            deru_row.addWidget(combo)
            self.deru.append(combo)
        deru_row.addStretch()
        self.editors = {'deterioration': QComboBox(editable=True), 'd_location': QLineEdit(),
                        'method': QLineEdit(), 'amount': QLineEdit(), 'unit': QComboBox(editable=True),
                        'price': QLineEdit(), 'remark': QLineEdit()}
        for row, (label, key) in enumerate(FIELDS, 1):
            widget = deru_box if key is None else self.editors[key]
            grid.addWidget(widget, row, 3)
            self.edit_widgets.append(widget)
        extra = len(FIELDS) + 1
        self.rule = QComboBox()
        self.rule.activated.connect(self.pick_rule)
        label = QLabel('Excel 規則')
        grid.addWidget(label, extra, 0)
        grid.addWidget(self.rule, extra, 1, 1, 3)
        self.tags = QWidget()
        self.tag_layout = QGridLayout(self.tags)
        self.tag_layout.setContentsMargins(0, 0, 0, 0)
        tag_label = QLabel('樣態')
        grid.addWidget(tag_label, extra + 1, 0)
        grid.addWidget(self.tags, extra + 1, 1, 1, 3)
        save, restore = QPushButton('儲存修改'), QPushButton('還原網站值')
        suggest = QPushButton('帶入程式建議')
        suggest.setToolTip('依 Excel 規則與前期資料計算的建議值（與填寫清單相同），帶入後可再調整')
        suggest.clicked.connect(self.load_proposal)
        save.clicked.connect(lambda: self.edited.emit(self.values()))
        restore.clicked.connect(self.restore)
        self.edit_status = QLabel('')
        self.edit_status.setWordWrap(True)
        buttons = QHBoxLayout()
        keep = QPushButton('確認維持現值')
        keep.setToolTip('看過後不需修改：填寫清單的「需確認」改為「已確認」（之後若修改，仍以修改為準）')
        keep.clicked.connect(self.kept.emit)
        buttons.addWidget(self.edit_status, 1)
        buttons.addWidget(keep)
        buttons.addWidget(suggest)
        buttons.addWidget(restore)
        buttons.addWidget(save)
        button_box = QWidget()
        button_box.setLayout(buttons)
        grid.addWidget(button_box, extra + 2, 1, 1, 3)
        self.edit_widgets += [self.rule, label, self.tags, tag_label, button_box]
        for col in (1, 2, 3):
            grid.setColumnStretch(col, 1)
        for combo in self.deru:
            combo.currentTextChanged.connect(self.mark)
        for editor in self.editors.values():
            (editor.currentTextChanged if isinstance(editor, QComboBox) else editor.textChanged).connect(self.mark)

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
        layout.addLayout(pictures)
        layout.addLayout(grid)
        layout.addSpacing(8)
        layout.addLayout(actions)     # 操作列緊接在欄位下方，不必上下來回看
        layout.addLayout(moves)
        layout.addStretch()
        self.set_editing(False)

    # ------------------------------------------------------------ 顯示
    def set_rules(self, rules):
        self.rules = rules or []

    def set_editing(self, on):
        on = bool(on and self.pair is not None and self.pair.current is not None)
        if self.edit_toggle.isChecked() != on:
            self.edit_toggle.setChecked(on)
        for widget in self.edit_widgets:
            widget.setVisible(on)

    def show_pair(self, pair, decision, position, edits=None):
        self.pair = pair
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
        # 防呆：只有本期（新增）沒有前期照片；只有前期（漏建、已維修）沒有本期照片；比對需兩期都有
        tips = {'current': ('開啟本期第二級編輯頁，捲到損壞照片（可直接刪照片）', '只有前期：本期沒有這筆，無本期照片'),
                'prior': ('開啟前期第二級檢視頁，捲到損壞照片', '本期新增：前期沒有這筆，無前期照片'),
                'compare': ('前期、本期網站並排，各自捲到損壞照片', '需前期、本期都有這筆才能比對')}
        usable = {'current': bool(pair.current and pair.current.damage_url(True)),
                  'prior': bool(pair.prior and pair.prior.damage_url()),
                  'compare': bool(pair.current and pair.prior)}
        for key, ok in usable.items():
            self.photo_buttons[key].setEnabled(ok)
            self.photo_buttons[key].setToolTip(tips[key][0 if ok else 1])
        self.photo_buttons['main'].setToolTip('看完照片後，把顯示中的網站開回本期／前期的檢測缺失主頁面')
        self.edit_toggle.setEnabled(pair.current is not None)
        if pair.current is None:
            self.set_editing(False)
            return
        current = pair.current
        self.site_values = {k: site_value(current, k) for k in EDIT_FIELDS}
        self.component_rules = ranked_rules(self.rules, current.component, current.fields)
        self.rule.clear()
        self.rule.addItem('（選擇 Excel 規則：帶入劣化類型、工法、單位、單價，並替換備註樣態）', None)
        for rule in self.component_rules:
            self.rule.addItem(f"{rule['pattern']}｜{rule['method']}｜{rule['unit']}｜{rule['price']}", rule)
        while self.tag_layout.count():
            self.tag_layout.takeAt(0).widget().deleteLater()
        for i, pattern in enumerate(pattern_tags(self.rules, current.component, current.fields)):
            button = QPushButton(pattern)
            button.setToolTip('把備註的樣態換成這個，並帶入對應的工法、單位、單價')
            button.clicked.connect(lambda _=False, p=pattern: self.pick_pattern(p))
            self.tag_layout.addWidget(button, i // 4, i % 4)
        for key in ('deterioration', 'unit'):
            combo = self.editors[key]
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(current.options.get(key, []))
            combo.blockSignals(False)
        self.load({**self.site_values, **(edits or {})})
        self.set_editing(self.edit_toggle.isChecked())

    # ------------------------------------------------------------ 修改
    def load(self, values):
        for combo, letter in zip(self.deru, 'deru'):
            combo.setCurrentText(values.get(letter) or '_')
        for key, editor in self.editors.items():
            (editor.setCurrentText if isinstance(editor, QComboBox) else editor.setText)(values.get(key, ''))
        self.mark()

    def values(self):
        result = {letter: combo.currentText() for combo, letter in zip(self.deru, 'deru')}
        for key, editor in self.editors.items():
            result[key] = (editor.currentText() if isinstance(editor, QComboBox) else editor.text()).strip()
        return result

    def mark(self, *_):
        """與網站現值不同的欄位標藍。"""
        values, changed = self.values(), 0
        for combo, letter in zip(self.deru, 'deru'):
            diff = values[letter] != self.site_values.get(letter, '_')
            combo.setStyleSheet(f'background:{EDITED};' if diff else '')
            changed += diff
        for key, editor in self.editors.items():
            diff = values[key] != self.site_values.get(key, '')
            editor.setStyleSheet(f'background:{EDITED};' if diff else '')
            changed += diff
        self.edit_status.setText(f'與網站不同 {changed} 欄' if changed else '與網站相同')

    def apply_rule(self, rule):
        if rule['deterioration']:
            self.editors['deterioration'].setCurrentText(rule['deterioration'])
        self.editors['method'].setText(rule['method'])
        self.editors['unit'].setCurrentText(rule['unit'])
        self.editors['price'].setText(str(rule['price']))
        self.editors['remark'].setText(replace_pattern(self.editors['remark'].text(), rule['pattern']))

    def pick_rule(self, index):
        rule = self.rule.itemData(index)
        if rule:
            self.apply_rule(rule)

    def pick_pattern(self, pattern):
        rule = best_rule(self.component_rules, pattern, self.editors['unit'].currentText())
        if rule:
            self.apply_rule(rule)
            self.rule.setCurrentIndex(self.rule.findData(rule))
        else:
            self.editors['remark'].setText(replace_pattern(self.editors['remark'].text(), pattern))

    def load_proposal(self):
        if self.pair is None or self.pair.current is None:
            return
        try:
            decision = proposal_for(self.pair, self.rules) if self.rules else None
        except Exception as error:
            decision, reason = None, f'程式建議失敗：{error}'
        else:
            reason = decision.reason if decision else '沒有讀到 Excel 規則表'
        if decision:
            self.load({**self.values(), **{k: str(v or '').strip() for k, v in decision.fields.items() if k in self.editors}})
        self.edit_status.setText(('程式建議（需確認）：' if decision is None or decision.review else '程式建議：') + reason)

    def restore(self):
        self.load(self.site_values)
        self.edited.emit(self.values())


PAGE_WIDTH = 860            # 網站主頁面（含左側構件列表）約需的寬度（CSS 像素）
ZOOM_RANGE = (0.6, 1.0)


def fit_zoom(width):
    """依網站面板寬度自動縮放，讓第一級、第二級與主頁面不必左右拖動（螢幕解析度不同也適用）。"""
    low, high = ZOOM_RANGE
    return round(min(high, max(low, width / PAGE_WIDTH)), 2)


class FitView(SiteView):
    """面板寬度改變（視窗縮放、拖動分隔線、切換單／雙網站、不同解析度）時自動調整網頁縮放。"""

    def __init__(self, profile, parent=None):
        super().__init__(profile, parent)
        self.loadFinished.connect(lambda _ok: self.fit())

    def fit(self):
        if self.width() > 50:
            self.setZoomFactor(fit_zoom(self.width()))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit()


def open_photos(view, url):
    """開啟第二級缺失頁，載入後捲到下方的「損壞照片」。"""
    def scroll(ok):
        view.loadFinished.disconnect(scroll)
        if ok:
            view.probe(SCROLL_TO_PHOTOS, lambda _: None)
    view.loadFinished.connect(scroll)
    view.open(url)


class MainWindow(QMainWindow):
    def __init__(self, url=LOGIN_URL, profile_dir=PROFILE_DIR, scan_dir=SCAN_DIR, review_dir=REVIEW_DIR,
                 delay_ms=500, rewrite=None, output_dir=OUTPUT_DIR, rules_path=None):
        super().__init__()
        self.setWindowTitle('公路橋梁填報助手')
        self.resize(1680, 960)
        self.scan_dir, self.review_dir, self.output_dir = Path(scan_dir), Path(review_dir), Path(output_dir)
        self.pairs, self.visible, self.review, self.review_file = [], [], Review(), None
        self.auto_pairs, self.order, self.prior_main, self.current_main = [], {}, '', ''
        self.index = -1
        self.busy = self.stop_requested = False
        self.profile = make_profile(profile_dir, self)
        self.view = FitView(self.profile, self)             # 本期網站
        self.prior_view = FitView(self.profile, self)       # 前期網站（照片比對、查看前期資料）
        self.prior_view.hide()
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
        # 歷史紀錄：之前掃描過的橋，選了直接接續審查，不必重新掃描
        self.history_menu = QMenu(self)
        self.history_menu.aboutToShow.connect(self.fill_history)
        history = QToolButton()
        history.setText('歷史紀錄')
        history.setToolTip('之前掃描過的橋：選一筆直接接續審查與修改，不必重新掃描')
        history.setMenu(self.history_menu)
        history.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        bar.addWidget(QLabel('橋名'))
        bar.addWidget(self.bridge_name)
        bar.addWidget(self.start_button)
        bar.addWidget(history)
        self.buttons = {'history': history}
        # 網站顯示：一個選單，按鈕上寫目前狀態
        site = QToolButton()
        site.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.site_menu = QMenu(self)
        self.site_actions = {}
        group = QActionGroup(self)
        for key, text in SITE_MODES.items():
            action = self.site_menu.addAction(text)
            action.setCheckable(True)
            action.setActionGroup(group)
            action.triggered.connect(lambda _=False, k=key: self.set_site_mode(k))
            self.site_actions[key] = action
        site.setMenu(self.site_menu)
        site.setToolTip('選擇要顯示的網站：本期、前期或兩者並排')
        hide = QPushButton('隱藏網站')
        hide.setToolTip('隱藏／顯示網站（顯示時恢復左邊選的本期、前期或兩者並排）')
        hide.clicked.connect(self.toggle_site)
        bar.addWidget(site)
        bar.addWidget(hide)
        self.buttons['site_mode'], self.buttons['site'] = site, hide
        self.chosen_mode = 'current'
        for key, text, slot in (('results', '開啟判讀結果資料夾', self.open_results),
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
        self.card.photos.connect(self.show_photos)
        self.card.edited.connect(self.save_edit)
        self.card.kept.connect(lambda: self.keep_current(self.pairs[self.index] if 0 <= self.index < len(self.pairs) else None))
        self.rules_path = Path(rules_path) if rules_path else rules_file()
        self.rules = load_rules(self.rules_path)
        self.card.set_rules(self.rules)
        self.check_tabs = QTabWidget()
        self.check_tabs.addTab(overview, '總表')
        self.check_tabs.addTab(self.card, '逐筆審查')

        # ② 填寫：依建議維修工法 ID 填寫網站
        fill = QWidget()
        fill_box = QVBoxLayout(fill)
        explain = QLabel('預定值＝程式依 Excel 規則與前期資料的建議，再套用你在逐筆審查的修改（修改優先）。'
                         '只填與網站不同的欄位；每筆送出後重新開啟核對，結果不明不會重送。「需確認」請雙擊到逐筆審查修改後再填。')
        explain.setWordWrap(True)
        self.fill_summary = QLabel('')
        self.fill_table = QTableWidget(0, len(FILL_COLUMNS))
        self.fill_table.setHorizontalHeaderLabels(FILL_COLUMNS)
        self.fill_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.fill_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.fill_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.fill_table.setWordWrap(True)
        for col, width in enumerate((70, 100, 120, 70, 300, 240)):
            self.fill_table.setColumnWidth(col, width)
        self.fill_table.setMinimumWidth(200)
        self.fill_table.horizontalHeader().setStretchLastSection(True)
        self.fill_table.cellDoubleClicked.connect(self.open_fill_item)
        self.fill_filter = QComboBox()
        self.fill_filter.addItems(('全部', READY, REVIEW, KEPT, DONE, PENDING, FAILED, SAME, SKIP))
        self.fill_filter.currentTextChanged.connect(lambda _: self.refresh_fill_table())
        actions, second = QHBoxLayout(), QHBoxLayout()     # 兩列，避免工作區被撐寬
        for key, text, tip, slot in (
                ('fill_one', '單筆修正', '只填寫表格中選取的那一筆（可填寫、待確認或失敗）', lambda: self.fill_selected()),
                ('fill_all', '整筆自動', '填寫全部「可填寫」：逐筆送出並回讀核對；遇到結果不明立即停止', lambda: self.fill_all()),
                ('fill_keep', '確認維持現值', '選取的「需確認」看過後不需修改：改為「已確認」；對「已確認」再按一次則取消',
                 lambda: self.keep_selected()),
                ('fill_rescan', '重新掃描本期', '重新讀取網站本期缺失（前期沿用上次掃描），更新清單與審查', lambda: self.rescan_current()),
                ('fill_plan', '重新整理', '依最新的審查與修改重新計算預定值', lambda: self.build_fill_plan())):
            button = QPushButton(text)
            button.setToolTip(tip)
            button.clicked.connect(slot)
            (actions if key in ('fill_one', 'fill_all', 'fill_keep') else second).addWidget(button)
            self.buttons[key] = button
        actions.addStretch()
        second.addStretch()
        second.addWidget(QLabel('顯示'))
        second.addWidget(self.fill_filter)
        fill_box.addWidget(explain)
        fill_box.addLayout(actions)
        fill_box.addLayout(second)
        fill_box.addWidget(self.fill_summary)
        fill_box.addWidget(self.fill_table)
        self.fill_items, self.fill_visible, self.fill_log, self.filler = [], [], None, None
        self.fill_tab = fill
        for key in FILL_BUTTONS:
            self.buttons[key].setEnabled(False)

        # ③ 比對資料輸出
        export = QWidget()
        export_box = QVBoxLayout(export)
        explain = QLabel('把前後期配對與審查結果輸出成 Excel，存到「判讀結果」資料夾。\n'
                         '「全部配對」一定輸出（含程式判定、最終判定、審查說明、修改內容）；下列項目各輸出成一張工作表：')
        explain.setWordWrap(True)
        self.export_options = {}
        options = QVBoxLayout()
        for key, tip in (('U值增減', '前後期都有、U 值不同（修改後的 U 值為準）'),
                         ('新增劣化', '最終判定為新增'),
                         ('未建（漏建）劣化', '最終判定為確實漏建或疑似漏建')):
            box = QCheckBox(f'{key}：{tip}')
            box.setChecked(True)
            options.addWidget(box)
            self.export_options[key] = box
        self.buttons['export'] = QPushButton('輸出比對 Excel')
        self.buttons['export'].clicked.connect(lambda: self.export_report())
        self.buttons['export'].setEnabled(False)
        self.export_status = QLabel('')
        self.export_status.setWordWrap(True)
        self.export_status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        export_box.addWidget(explain)
        export_box.addLayout(options)
        for widget in (self.buttons['export'], self.export_status):
            export_box.addWidget(widget)
        export_box.addStretch()

        self.output = QPlainTextEdit(readOnly=True)
        self.output.setFont(QFont('Consolas', 9))
        self.tabs = QTabWidget()
        for widget, text in zip((self.check_tabs, fill, export), STEPS):
            self.tabs.addTab(widget, text)
        self.tabs.currentChanged.connect(lambda _: self.tabs.currentWidget() is self.fill_tab and self.build_fill_plan())
        log_box = QWidget()
        log_layout = QVBoxLayout(log_box)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.addWidget(QLabel('紀錄'))
        log_layout.addWidget(self.output)
        self.work = QSplitter(Qt.Orientation.Vertical)     # 右側工作區：上步驟、下紀錄（可拖拉高度）
        self.work.addWidget(self.tabs)
        self.work.addWidget(log_box)
        self.work.setSizes([780, 140])

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.sites = QSplitter(Qt.Orientation.Horizontal)
        self.sites.addWidget(self.prior_view)
        self.sites.addWidget(self.view)
        self.splitter.addWidget(self.sites)
        self.splitter.addWidget(self.work)
        self.splitter.setSizes([980, 700])
        # 查到橋梁後先開基本資料，請使用者確認是這座橋再掃描
        self.pending_bridge = None
        self.confirm_bar = QFrame()
        self.confirm_bar.setStyleSheet('QFrame{background:#fff2cc; border-radius:6px;}')
        confirm = QHBoxLayout(self.confirm_bar)
        self.confirm_text = QLabel('')
        self.confirm_text.setWordWrap(True)
        yes, no = QPushButton('是這座橋，開始掃描'), QPushButton('不是，取消')
        yes.setStyleSheet('font-weight:bold;')
        yes.clicked.connect(lambda: self.confirm_bridge(True))
        no.clicked.connect(lambda: self.confirm_bridge(False))
        confirm.addWidget(self.confirm_text, 1)
        confirm.addWidget(yes)
        confirm.addWidget(no)
        self.confirm_buttons = (yes, no)
        self.confirm_bar.hide()
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addLayout(bar)
        layout.addWidget(self.confirm_bar)
        layout.addWidget(self.summary)
        layout.addWidget(self.splitter, 1)
        self.setCentralWidget(root)
        self.set_sites(current=True)
        self.view.open(url)

    # ------------------------------------------------------------ 共用
    def log(self, message):
        self.output.appendPlainText(f'{datetime.now():%H:%M:%S} {message}')

    def set_busy(self, busy):
        self.busy = busy
        for key in ('scan', 'open'):
            self.menu_actions[key].setEnabled(not busy)
        self.buttons['export'].setEnabled(not busy and bool(self.pairs))
        for key in FILL_BUTTONS:
            self.buttons[key].setEnabled(not busy and bool(self.pairs))
        self.start_button.setEnabled(not busy)
        self.buttons['stop'].setEnabled(busy)

    def stop(self):
        self.stop_requested = True
        self.log('停止中，完成目前這一頁後停止…')

    def site_mode(self):
        prior, current = self.prior_view.isVisibleTo(self), self.view.isVisibleTo(self)
        return 'both' if prior and current else 'prior' if prior else 'current' if current else 'none'

    def set_site_mode(self, mode):
        """選單選擇顯示哪個網站（會記住，「顯示網站」時恢復）。"""
        self.set_sites(prior=mode in ('prior', 'both'), current=mode in ('current', 'both'))

    def toggle_site(self):
        """隱藏全部網站；再按一次依選單目前的選擇顯示。"""
        if self.site_mode() != 'none':
            self.set_sites(prior=False, current=False)
        else:
            self.set_site_mode(self.chosen_mode)

    def show_site(self, visible=True):
        self.set_sites(current=visible)

    def set_sites(self, prior=None, current=None):
        for view, on in ((self.prior_view, prior), (self.view, current)):
            if on is None:
                continue
            view.setVisible(on)
            if on and view is self.prior_view and self.prior_main and self.prior_view.url().isEmpty():
                self.prior_view.open(self.prior_main)       # 第一次顯示前期網站時開前期缺失主頁面
        mode = self.site_mode()
        if mode != 'none':
            self.chosen_mode = mode
        self.site_actions[self.chosen_mode].setChecked(True)
        self.buttons['site_mode'].setText(SITE_LABELS[self.chosen_mode])
        self.buttons['site'].setText('隱藏網站' if mode != 'none' else '顯示網站')
        if mode != 'none':
            total = max(self.splitter.width(), self.width(), 1000)
            share = SITE_SHARE[mode]
            self.splitter.setSizes([int(total * share), int(total * (1 - share))])
            if mode == 'both':
                self.sites.setSizes([1, 1])

    # ------------------------------------------------------------ 照片與修改
    def show_photos(self, which):
        """本期照片（編輯頁，可直接刪照片）、前期照片（檢視頁）、前後期照片比對（另開視窗並排）。"""
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        if which == 'main':
            if self.site_mode() == 'none':
                self.set_site_mode(self.chosen_mode)
            for view, url in ((self.view, self.current_main), (self.prior_view, self.prior_main)):
                if url and view.isVisibleTo(self):
                    view.open(url)
            return
        if (which == 'prior' and not pair.prior) or (which == 'current' and not pair.current) or \
                (which == 'compare' and not (pair.prior and pair.current)):
            QMessageBox.information(self, '照片', '本期新增的缺失沒有前期照片。' if not pair.prior else '這筆本期沒有建立，沒有本期照片。')
            return
        sides = ('prior', 'current') if which == 'compare' else (which,)
        for side in sides:
            defect = pair.current if side == 'current' else pair.prior
            url = defect.damage_url(edit=side == 'current') if defect else ''
            if not url:
                QMessageBox.information(self, '照片', '這筆沒有第二級缺失頁的連結（請重新掃描）。')
                return
            self.set_sites(**{side: True})
            open_photos(self.view if side == 'current' else self.prior_view, url)
            self.log(f"開啟{'本期' if side == 'current' else '前期'}照片：{defect.component} {defect.remark.strip()}")

    def save_edit(self, values):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        kept = self.review.set_edit(pair, values)
        self.review.save(self.review_file)
        self.log(f'修改本期：{pair.current.component}：' + (describe(kept) if kept else '還原為網站值'))
        self.card.mark()
        self.refresh_table()
        self.update_summary()
        if self.fill_items:
            self.build_fill_plan()

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
            write_report(self.pairs, target, self.review,
                         sheets={k for k, box in self.export_options.items() if box.isChecked()})
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
        saved = [item for item in saved_sessions(self.scan_dir, self.review_dir) if item['bridge'] == name]
        if saved:
            answer = self.ask_resume(saved[0])
            if answer is None:
                return
            if answer == 'continue':
                self.resume(saved[0]['path'])
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
            self.log(f'橋梁：{bridge.name}（橋梁系統ID {bridge.inventory_id}），請在左側基本資料確認')
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'中斷：{error}')
            QMessageBox.warning(self, '開始', str(error))
            return
        finally:
            self.set_busy(False)
        self.ask_bridge(bridge)

    def ask_bridge(self, bridge):
        """左側開啟基本資料，上方黃色提示列請使用者確認後才掃描（可捲動網站核對）。"""
        self.pending_bridge = bridge
        self.show_site(True)
        self.view.open(basic_urls(bridge.inventory_id)['basic'])
        latest = bridge.latest.date if bridge.latest else '—'
        self.confirm_text.setText(f'請確認左側基本資料是否為要處理的橋：{bridge.name}｜橋梁系統ID {bridge.inventory_id}｜'
                                  f'{bridge.agency}｜{bridge.place}｜最新檢測 {latest}')
        self.confirm_bar.show()
        self.start_button.setEnabled(False)

    def confirm_bridge(self, accepted):
        bridge, self.pending_bridge = self.pending_bridge, None
        self.confirm_bar.hide()
        self.start_button.setEnabled(True)
        if bridge is None:
            return
        if not accepted:
            self.log('已取消：不是這座橋')
            return
        self.stop_requested = False
        self.set_busy(True)
        try:
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

    def fill_history(self):
        self.history_menu.clear()
        sessions = saved_sessions(self.scan_dir, self.review_dir)
        if not sessions:
            self.history_menu.addAction('（尚無掃描紀錄）').setEnabled(False)
        for item in sessions[:30]:
            self.history_menu.addAction(history_text(item), lambda p=item['path']: self.resume(p))

    def resume(self, path):
        """接續之前的掃描：讀掃描檔與審查檔，不重新掃描；左側開回本期缺失主頁面。"""
        if self.busy:
            return
        self.load_saved(path)
        if self.current_main:
            self.view.open(self.current_main)
            self.show_site(True)

    def ask_resume(self, item):
        """同一座橋已有掃描紀錄：接續（不重新掃描）或重新掃描。回傳 'continue'／'rescan'／None。"""
        box = QMessageBox(self)
        box.setWindowTitle('已有掃描紀錄')
        box.setText(f'「{item["bridge"]}」已在 {item["updated"]:%m/%d %H:%M} 掃描過：\n{history_text(item)}\n\n'
                    '要接續上次的審查與修改，還是重新掃描？（網站資料有變動時才需要重新掃描）')
        keep = box.addButton('接續上次', QMessageBox.ButtonRole.AcceptRole)
        again = box.addButton('重新掃描', QMessageBox.ButtonRole.DestructiveRole)
        box.addButton('取消', QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(keep)
        box.exec()
        return {keep: 'continue', again: 'rescan'}.get(box.clickedButton())

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
        keys = next((r['keys'] for r in prior_scan.get('rows', []) if r['keys'].get('inspect_id')), {})
        current_keys = next((r['keys'] for r in current_scan.get('rows', []) if r['keys'].get('inspect_id')), {})
        self.current_main = (f"https://bms2.iot.gov.tw/bms3/inspect/inspect/der_content/inspect_id,"
                             f"{current_keys['inspect_id']}/inventory_id,{current_keys.get('inventory_id', '')}/from,result/"
                             ) if current_keys else ''
        self.prior_main = (f"https://bms2.iot.gov.tw/bms3/inspect/inspect/der_content/inspect_id,{keys['inspect_id']}/"
                           f"inventory_id,{keys.get('inventory_id', '')}/from,result/") if keys else ''
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

    # ------------------------------------------------------------ ② 填寫
    def build_fill_plan(self):
        if not self.pairs or self.busy:
            return
        self.fill_log = FillLog(fill_log_path(self.review_dir, self.review.bridge, self.review.current_date))
        self.fill_items = build_plan(self.pairs, self.review, self.rules, self.fill_log)
        if self.add_missing_rules():
            self.fill_items = build_plan(self.pairs, self.review, self.rules, self.fill_log)
        self.refresh_fill_table()

    def add_missing_rules(self):
        """Excel 沒有對應規則（新劣化樣態、劣化類型判讀差異）：依規則格式新增到規則表並重新讀取。回傳是否有新增。"""
        path = self.rules_path or rules_file()
        if not path:
            return False
        additions, seen = [], set()
        for item in self.fill_items:
            if not item.reason.startswith(NO_RULE):
                continue
            rule = suggest_rule(item.pair.current, self.rules)
            if not rule:
                continue
            key = (rule['component'], rule['劣化類型'], rule['劣化樣態'])
            if key in seen:
                continue
            seen.add(key)
            rule['origin'] = f'{self.review.bridge} {self.review.current_date} ID {item.method_id}「{item.pair.current.remark.strip()}」；'
            additions.append(rule)
        if not additions:
            return False
        try:
            target = writable_rules_file(path)
            added = add_rules(target, additions)
        except PermissionError:
            self.log(f'無法新增規則：{Path(path).name} 正在 Excel 中開啟，請關閉後按「重新整理」')
            return False
        except Exception as error:   # 規則表寫入失敗不影響其他功能
            self.log(f'無法新增規則：{error}')
            return False
        self.rules_path = target
        self.rules = load_rules(target)
        self.card.set_rules(self.rules)
        for rule_id, sheet, rule in added:
            self.log(f"Excel 新增規則 {rule_id}（{sheet}，{rule['kind']}）：{rule['劣化類型']}／{rule['劣化樣態']}／"
                     f"{rule['維修工法']}／{rule['維修單位']} {rule['單價']}；說明欄黃底待確認")
        self.statusBar().showMessage(f'Excel 已新增 {len(added)} 條規則（{target.name}，黃底待確認）', 10000)
        return bool(added)

    def refresh_fill_table(self):
        chosen = self.fill_filter.currentText()
        self.fill_visible = [i for i in self.fill_items if chosen == '全部' or i.status == chosen]
        self.fill_table.setRowCount(len(self.fill_visible))
        for row, item in enumerate(self.fill_visible):
            current = item.pair.current
            values = (item.status, current.component, current.label, item.method_id, item.describe() or '—',
                      '；'.join(x for x in (item.message, item.reason) if x))
            color = FILL_COLORS.get(item.status)
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if color:
                    cell.setBackground(QColor(color))
                self.fill_table.setItem(row, col, cell)
        self.fill_table.resizeRowsToContents()
        counts = {}
        for item in self.fill_items:
            counts[item.status] = counts.get(item.status, 0) + 1
        self.fill_summary.setText('本期 {} 筆：'.format(len(self.fill_items)) +
                                  '、'.join(f'{k} {v}' for k, v in counts.items()))
        self.buttons['fill_all'].setEnabled(not self.busy and counts.get(READY, 0) > 0)
        self.buttons['fill_one'].setEnabled(not self.busy and bool(self.fill_items))
        for key in ('fill_plan', 'fill_keep', 'fill_rescan'):
            self.buttons[key].setEnabled(not self.busy and bool(self.pairs))

    def open_fill_item(self, row, _col=0):
        if 0 <= row < len(self.fill_visible):
            self.index = self.pairs.index(self.fill_visible[row].pair)
            self.show_card()
            self.card.edit_toggle.setChecked(True)

    def confirm_fill(self, items):
        """送出前確認（會修改網站資料）。"""
        lines = '\n'.join(f'・{i.pair.current.component} {i.pair.current.label}（ID {i.method_id}）：{i.describe()}'
                          for i in items[:8])
        more = f'\n…另 {len(items) - 8} 筆' if len(items) > 8 else ''
        answer = QMessageBox.question(self, '填寫網站', f'將填寫並送出 {len(items)} 筆到網站：\n{lines}{more}\n\n'
                                      '每筆送出後會重新開啟核對。確定開始？')
        return answer == QMessageBox.StandardButton.Yes

    def keep_current(self, pair):
        if pair is None or pair.current is None:
            return
        on = pair.key not in self.review.kept
        self.review.keep(pair, on)
        self.review.save(self.review_file)
        self.log(('確認維持現值：' if on else '取消確認維持現值：') + f'{pair.current.component} {pair.current.remark.strip()}')
        self.card.edit_status.setText('已確認維持網站現值' if on else '已取消確認')
        self.build_fill_plan()

    def keep_selected(self):
        rows = self.fill_table.selectionModel().selectedRows() if self.fill_table.selectionModel() else []
        if not rows:
            QMessageBox.information(self, '確認維持現值', '請先在表格選一筆「需確認」。')
            return
        item = self.fill_visible[rows[0].row()]
        if item.status not in (REVIEW, KEPT):
            QMessageBox.information(self, '確認維持現值', f'這筆是「{item.status}」，不需要確認。')
            return
        self.keep_current(item.pair)

    def rescan_current(self):
        """填寫後重新掃描本期（前期沿用上次的掃描檔），更新配對、審查與填寫清單。"""
        if self.busy or not self.current_main:
            return False
        priors = sorted(self.scan_dir.glob(f'*-{self.review.bridge}-前期.json'))
        if not priors:
            self.log('找不到前期掃描檔，請改用「開始」重新掃描')
            return False
        prior = json.loads(priors[-1].read_text(encoding='utf-8'))
        self.stop_requested = False
        self.set_busy(True)
        try:
            self.log('重新掃描本期（前期沿用上次掃描）')
            html = self.fetcher.get(self.current_main)
            current = scan_inspection(html, self.fetcher, progress=self.progress, stopped=lambda: self.stop_requested)
        except Stopped as error:
            self.log(str(error))
            return False
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'重新掃描中斷：{error}')
            return False
        finally:
            self.set_busy(False)
        self.save_json(current.header, '檢測', current.to_dict())
        failed = [r for r in current.rows if not r.ok]
        for row in failed:
            self.log(f'核對異常：{row.component} {row.label}：{row.check}')
        index_key = self.pairs[self.index].key if 0 <= self.index < len(self.pairs) else None
        self.load_pairs(prior, current.to_dict())
        for pair in self.pairs:                        # 已填進網站的修改不再算「已修改」
            if self.review.edit_of(pair):
                self.review.set_edit(pair, self.review.edit_of(pair))
        self.review.save(self.review_file)
        self.refresh_table()
        self.update_summary()
        self.index = next((i for i, p in enumerate(self.pairs) if p.key == index_key), -1)
        self.tabs.setCurrentWidget(self.fill_tab)
        self.build_fill_plan()
        self.log(f'重新掃描完成：核對異常 {len(failed)} 列')
        return True

    def fill_selected(self):
        rows = self.fill_table.selectionModel().selectedRows() if self.fill_table.selectionModel() else []
        if not rows:
            QMessageBox.information(self, '單筆修正', '請先在表格選一筆。')
            return
        item = self.fill_visible[rows[0].row()]
        if item.status not in (READY, PENDING, FAILED) or not item.changes:
            QMessageBox.information(self, '單筆修正', f'這筆是「{item.status}」，沒有要填寫的變更。'
                                    + ('請雙擊到逐筆審查修改後再填。' if item.status == REVIEW else ''))
            return
        self.run_fill([item])

    def fill_all(self):
        self.run_fill([i for i in self.fill_items if i.status == READY])

    def run_fill(self, items):
        if self.busy or not items or not self.confirm_fill(items):
            return
        if self.filler is None:
            self.filler = SiteFiller(self.profile, self, delay_ms=self.fetcher.delay_ms, rewrite=self.fetcher.rewrite)
        self.stop_requested = False
        self.set_busy(True)
        done = 0
        try:
            for number, item in enumerate(items, 1):
                if self.stop_requested:
                    self.log(f'已停止填寫：完成 {done}／{len(items)} 筆')
                    break
                label = f'{item.pair.current.component} {item.pair.current.label}（ID {item.method_id}）'
                self.statusBar().showMessage(f'填寫中 {number}／{len(items)}：{label}')
                try:
                    message = self.filler.fill(item, self.fill_log)
                    done += 1
                    self.log(f'填寫完成：{label}：{message}')
                except FillError as error:
                    self.log(f'填寫失敗：{label}：{error}')
                    if item.status == PENDING:          # 結果不明：整批停止，等人工確認
                        QMessageBox.warning(self, '填寫網站', f'{label}\n{error}\n\n已停止後續填寫。')
                        break
                self.refresh_fill_table()
        finally:
            self.set_busy(False)
            self.statusBar().showMessage(f'填寫結束：完成 {done}／{len(items)} 筆', 8000)
            self.refresh_fill_table()
        if self.view.isVisibleTo(self) and self.current_main:
            self.view.open(self.current_main)            # 左側重新載入本期主頁面，看得到最新結果
        if done:
            self.rescan_current()                        # 填寫後重新掃描本期，清單與審查依網站最新資料更新

    def update_summary(self):
        if not self.pairs:
            return
        counts = self.review.counts(self.pairs)
        self.summary.setText(f'{self.review.bridge}　本期 {self.review.current_date}／前期 {self.review.prior_date}　'
                             f'共 {len(self.pairs)} 組：待審查 {counts[PENDING]}、已確認 {counts[ACCEPTED]}、'
                             f'已改判 {counts[CHANGED]}；已修改 {sum(1 for p in self.pairs if self.review.edit_of(p))}。'
                             '審查結果自動存檔。')

    def matches_filter(self, pair):
        chosen = self.filter.currentText()
        decision = self.review.decision(pair)
        if chosen == '全部':
            return True
        if chosen in (PENDING, CHANGED):
            return decision.status == chosen
        if chosen == '需確認':
            return pair.confidence == '需確認'
        if chosen == '已修改':
            return bool(self.review.edit_of(pair))
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
            edits = self.review.edit_of(pair)
            remark = self.review.value(pair, 'remark') if pair.current else '—'
            values = (decision.status + ('✎' if edits else ''), component, any_side.label, kind, pair.confidence,
                      pair.prior.remark.strip() if pair.prior else '—',
                      ('✎ ' if 'remark' in edits else '') + remark, '；'.join(pair.reasons))
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
        self.card.show_pair(pair, self.review.decision(pair), f'第 {self.index + 1}／{len(self.pairs)} 組',
                            self.review.edit_of(pair))
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
        """移到上／下一筆。只在待審查之間移動時，到底會從頭找；全部審查完回到總表。"""
        if not self.pairs:
            return
        pending_only = self.card.pending_only.isChecked()
        wanted = lambda i: not pending_only or self.review.decision(self.pairs[i]).status == PENDING  # noqa: E731
        index = self.index + step
        while 0 <= index < len(self.pairs):
            if wanted(index):
                self.index = index
                self.show_card()
                return
            index += step
        if pending_only:
            rest = [i for i in range(len(self.pairs)) if wanted(i)]
            if rest:
                self.index = rest[0] if step > 0 else rest[-1]
                self.show_card()
                return
        if step > 0:
            self.check_tabs.setCurrentIndex(0)
            message = '全部審查完成，已回到總表' if not any(
                self.review.decision(p).status == PENDING for p in self.pairs) else '已到最後一筆，已回到總表'
            self.statusBar().showMessage(message, 8000)
            self.log(message)

    def closeEvent(self, event):
        self.stop_requested = True
        self.fetcher.release()
        if self.filler is not None:
            self.filler.release()
        self.prior_view.release()
        self.view.release()
        super().closeEvent(event)


def main(argv=None):
    app = QApplication.instance() or QApplication(sys.argv[:1])
    window = MainWindow()
    window.showMaximized()          # 依螢幕解析度放到最大，網站與工作區比例、網頁縮放隨之調整
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
