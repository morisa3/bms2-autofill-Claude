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
import re
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QTimer, QUrl, Qt, Signal
from PySide6.QtGui import QAction, QActionGroup, QColor, QTextCursor, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QStackedWidget, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
                               QMenu, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QToolButton,
                               QSplitter, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from ..core.fill import TEXT_FIELDS, adopt_prior_method, plan_item, proposal_for, tidy_locations
from ..core.policy import automatic_plan
from ..core.prior_merge import COUNT_UNITS, merge_with_prior, normalize_size, only_size, quantity, size_part
from ..core.conditions import aliases
from ..core.remarks import parse as parse_remark
from ..core.opinion import (numbering_ok, OpinionItem, absorb, apply_choices, blanks, choices, compose, generate, heights,
                            unresolved)
from ..core.fill import (DONE, FAILED, KEPT, PENDING, READY, REVIEW, SAME, SKIP, FillLog, build_plan,
                         fill_log_path)
from ..core.rule_add import NO_RULE, add_rules, migrate_auto_rules, suggest_rule, writable_rules_file
from ..core.editing import RULES_NAME, rules_file
from ..core.editing import EDIT_FIELDS, best_rule, describe, load_rules, pattern_tags, ranked_rules, replace_pattern, split_remark
from ..core.local_files import APP_DIR
from ..core.basic_info import NAMES as BASIC_NAMES, OPTIONS as BASIC_OPTIONS
from ..core.basic_info import changes as basic_changes, defaults as basic_defaults, options as basic_options
from ..core.inspection import supervisors
from ..core.manual import apply_manual, candidates
from ..core.matching import defect_id, match_scans
from ..core.report import report_name, write_report
from ..core.review import (ACCEPTED, CHANGED, KIND_PRIOR_WRONG, PENDING, Review, choices_for, clear_sessions, review_path,
                           saved_sessions,
                           site_value)
from ..site.pages import group_bridges, needs_choice, parse_history, parse_inspection
from ..core.bridge_photos import PhotoPlan, photo_list_url
from .photo_panel import PhotoPanel
from ..site.scan import find_prior, scan_record, Stopped, basic_urls, history_url, inventory_id, scan_basic, scan_inspection, scan_prior
from ..site.session import LOGIN_URL
from .browser import PROFILE_DIR, SHORTCUTS, SiteView, make_profile, set_lite
from .fetcher import FetchError, PageFetcher
from .filler import FillError, SiteFiller

SCAN_DIR = APP_DIR / 'temp' / 'scan'
REVIEW_DIR = APP_DIR / 'temp' / 'review'
OUTPUT_DIR = APP_DIR / '判讀結果'
SITE_MODES = {'current': '只顯示本期網站', 'prior': '只顯示前期網站', 'both': '前期＋本期網站並排'}
SITE_SHARE = {'both': 0.73, 'current': 0.58, 'prior': 0.58}     # 網站區佔左右寬度的比例（其餘為工作區）
SITE_LABELS = {'current': '網站：本期', 'prior': '網站：前期', 'both': '網站：前期＋本期', 'none': '網站：已隱藏'}
STEPS = ('① 核對與修改', '② 送出資料', '③ 送出檢測員意見', '④ 上傳基本照片', '⑤ 比對資料輸出')
# 合併狀態（審查＋填寫）：清單上方的數字卡片，點了就篩選
TODO, SENDABLE, SENT, SEND_ISSUE, NOTHING = '待處理', '可送出', '已送出', '送出異常', '無需修改'
OVERALL_CARDS = (TODO, SENDABLE, SENT, SEND_ISSUE, NOTHING, '全部')
OVERALL_TIPS = {TODO: '審查待確認，或填寫需確認（理由欄寫明是哪一種）',
                SENDABLE: '已確認，且有要填寫到網站的變更',
                SENT: '已送出並回讀核對一致',
                SEND_ISSUE: '送出結果不明（待確認）或沒有送出（失敗）',
                NOTHING: '網站已符合、不需填寫，或只有前期（漏建、已維修等）',
                '全部': '全部配對'}
OVERALL_COLORS = {TODO: '#fff2cc', SENDABLE: '#e2efda', SENT: '#ddebf7', SEND_ISSUE: '#f8cbad'}
NO_CURRENT_TIP = ('{bridge}：網站上還沒有 {year} 年的定期檢測紀錄。\n\n'
                  '已先把最新一期（{date}）當作前期掃描建檔，本期暫不掃描。\n'
                  '請在網站新增本年度的定期檢測紀錄並填入缺失後，再輸入橋名按「開始」：'
                  '程式會沿用已建檔的前期（不重新掃描），只掃描本期並配對。')
TABLE_COLUMNS = ('構件', '編號', '前期備註', '本期備註', '判定', '信心', '理由')   # 工作區放右側時優先看到構件、編號、前後期備註
FIELDS = (('DERU', None), ('劣化類型', 'deterioration'), ('損壞位置', 'd_location'), ('維修工法', 'method'),
          ('數量', 'amount'), ('單位', 'unit'), ('單價', 'price'), ('備註', 'remark'))
TABLE_WIDTHS = (100, 130, 220, 220, 150, 50, 0)    # 各欄上限（0：不限，最後一欄填滿）；構件較長時折成兩行
FILL_COLUMNS = ('狀態', '構件', '編號', '網站 ID', '將填寫的變更', '說明')
FILL_COLORS = {READY: '#e2efda', REVIEW: '#fff2cc', DONE: '#ddebf7', PENDING: '#f8cbad', FAILED: '#f8cbad',
               KEPT: '#e4dfec'}
ADVICE = ('建議持續追蹤', '建議三年內派工修復', '建議一年內派工修復', '建議通知相關單位立即處置')
FILL_BUTTONS = ('send_one', 'send_component', 'send_all', 'fill_rescan', 'fill_plan', 'opinion_build', 'opinion_write')
FILTERS = ('全部', '待審查', '填寫需確認', '需確認', '已改判', '人工配對', '位置變更', '已修改', '前期選擇或填寫錯誤', '與其他缺失重複合併', '新增', '疑似漏建', '單位數量變更', '範圍加大', '範圍縮小', '沿用', '已維修')
STATUS_COLORS = {PENDING: '#fff2cc', CHANGED: '#ddebf7', ACCEPTED: None}
DIFFERENT = '#fce4d6'
EDITED = '#ddebf7'
SCROLL_TO_PHOTOS = """(() => {
  const all = Array.from(document.querySelectorAll('h1,h2,h3,h4,div,span,td,p'));
  const head = all.find(e => e.childElementCount === 0 && /損壞照片/.test(e.textContent || ''));
  if (head) head.scrollIntoView({block: 'start'});
})()"""


# 前後期照片比對：只留「損壞照片」（隱藏 DER 資料表、標題列），網站面板可以縮窄
PHOTOS_ONLY_JS = """(() => {
  const heads = Array.from(document.querySelectorAll('.panel-heading, h1, h2, h3, h4, span, div'))
    .filter(e => e.childElementCount === 0 && /損壞照片/.test(e.textContent || ''));
  const photos = heads.length ? (heads[0].closest('.panel') || heads[0].parentElement) : null;
  if (!photos) return;
  document.querySelectorAll('.panel').forEach(p => { if (p !== photos && !p.contains(photos)) p.style.display = 'none'; });
  document.querySelectorAll('h1, hr, input[name="back_list"]').forEach(e => { e.style.display = 'none'; });
  document.querySelectorAll('form').forEach(f => { if (!f.contains(photos)) f.style.display = 'none'; });
  photos.scrollIntoView({block: 'start'});
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


def remark_lines(remark):
    """清單用：備註的位置與樣態尺寸分兩行（A1-L,R／土石堆積）。"""
    lines = []
    for line in str(remark or '').strip().splitlines() or ['']:
        location, description, tail = split_remark(line)
        lines.append(f'{location}\n{description}{tail}' if location and (description or tail) else line.strip())
    return '\n'.join(lines) or '—'


def fit_columns(table, widest):
    """欄寬依最寬的內容：每格最多兩行——單行太寬（超過上限）的折成兩行，已分行的取最長那行；欄寬不超過上限。"""
    metrics = table.fontMetrics()
    pad = 14
    for col, limit in enumerate(widest):
        if not limit:
            continue
        header = table.horizontalHeaderItem(col)
        need = metrics.horizontalAdvance(header.text() if header else '') + pad
        for row in range(table.rowCount()):
            item = table.item(row, col)
            lines = (item.text() if item else '').split('\n')
            widths = [metrics.horizontalAdvance(line) for line in lines]
            width = max(widths, default=0)
            if len(lines) == 1 and width + pad > limit:
                width = (width + 1) // 2 + metrics.horizontalAdvance('字')    # 折成兩行（中文可在任一字斷行）
            need = max(need, width + pad)
        table.setColumnWidth(col, min(need, limit))


def component_label(pair):
    """清單用的構件名稱：「其他」不加設施類別（編號欄已寫鋪面等）；位置變更寫「前期 → 本期」。"""
    if pair.moved:
        return f'{pair.prior.component} → {pair.current.component}'
    return (pair.current or pair.prior).component


class ReviewCard(QWidget):
    """逐筆審查（版型 B）：前期與本期並排，不同的欄位標色；可確認或改判、修改本期欄位、查看兩期照片。"""
    decided = Signal(str, str)      # 判定、說明
    moved = Signal(int)             # -1 上一筆、+1 下一筆
    relink = Signal()               # 改配對
    photos = Signal(str)            # current／prior／compare
    edited = Signal(dict)           # 修改後的本期欄位
    kept = Signal()                 # 確認維持網站現值（填寫清單的需確認）

    def __init__(self, parent=None, fill_mode=False):
        super().__init__(parent)
        self.fill_mode = fill_mode          # 填寫分頁用：只修改欄位，不顯示判定與上下筆
        self.rules, self.component_rules, self.pair, self.site_values = [], [], None, {}
        self.title = QLabel('尚未選擇')
        self.title.setStyleSheet('font-size:14px; font-weight:bold;')
        self.title.setWordWrap(True)
        self.reasons = QLabel('')
        self.reasons.setWordWrap(True)

        self.photo_buttons = {}             # 放在主視窗 ① 的工具列（進階篩選那一列）
        for key, text in (('current', '本期照片'), ('prior', '前期照片'), ('compare', '前後期照片比對'),
                          ('main', '回到缺失主頁面')):
            button = QPushButton(text)
            button.clicked.connect(lambda _=False, k=key: self.photos.emit(k))
            button.setMinimumWidth(40)
            self.photo_buttons[key] = button
        self.edit_toggle = QPushButton('修改本期')
        self.edit_toggle.setCheckable(True)
        self.edit_toggle.setToolTip('修改本期的 DERU、劣化類型、樣態、維修工法、數量、單價等（第 6 步填寫時寫入網站）')
        self.edit_toggle.toggled.connect(self.set_editing)
        heading = QHBoxLayout()
        heading.addWidget(self.title, 1)
        heading.addWidget(self.edit_toggle)

        grid = QGridLayout()
        grid.addWidget(QLabel(''), 0, 0)
        self.edit_widgets = []
        self.heads = {}
        for col, text in ((1, '前期'), (2, '本期（網站）')):
            head = QLabel(text)
            head.setStyleSheet('font-weight:bold;')
            grid.addWidget(head, 0, col)
            self.heads[col] = head
        self.cells = {}
        self.stacks = []                     # 本期欄：網站值 ⇄ 修改欄位（按「修改本期」直接在原位置換成修改介面）
        for row, (label, key) in enumerate(FIELDS, 1):
            grid.addWidget(QLabel(label), row, 0, Qt.AlignmentFlag.AlignTop)
            for col in (1, 2):
                cell = QLabel('')
                cell.setWordWrap(True)
                cell.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                cell.setFrameShape(QFrame.Shape.StyledPanel)
                cell.setMinimumHeight(22)            # 精簡：1080p 下清單與工作區可同時看到
                if col == 1:
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
            stack = QStackedWidget()
            stack.addWidget(self.cells[(label, 2)])
            stack.addWidget(deru_box if key is None else self.editors[key])
            grid.addWidget(stack, row, 2)
            self.stacks.append(stack)
        extra = len(FIELDS) + 1
        self.rule = QComboBox()
        self.rule.activated.connect(self.pick_rule)
        label = QLabel('Excel 規則')
        grid.addWidget(label, extra, 0)
        grid.addWidget(self.rule, extra, 1, 1, 2)
        self.tags = QWidget()
        self.tag_layout = QGridLayout(self.tags)
        self.tag_layout.setContentsMargins(0, 0, 0, 0)
        tag_label = QLabel('樣態')
        grid.addWidget(tag_label, extra + 1, 0)
        grid.addWidget(self.tags, extra + 1, 1, 1, 2)
        restore = QPushButton('還原網站值')
        restore.setToolTip('本期欄改回網站現值；再按「確認並下一筆」即維持網站現值、不修改')
        suggest = QPushButton('帶入程式建議')
        suggest.setToolTip('依 Excel 規則與前期資料計算的建議值（與填寫清單相同），帶入後可再調整')
        suggest.clicked.connect(self.load_proposal)
        restore.clicked.connect(self.restore)
        self.edit_status = QLabel('')
        self.edit_status.setWordWrap(True)
        self.edit_status.setMinimumWidth(90)
        buttons = QHBoxLayout()
        buttons.addWidget(self.edit_status, 1)
        self.import_button = QPushButton('導入前期')
        self.import_button.setToolTip('帶入前期的劣化類型、位置、工法、數量、單位、單價與備註；再點樣態按鈕只會換掉備註中間的敘述')
        self.import_button.clicked.connect(self.import_prior)
        buttons.addWidget(self.import_button)
        buttons.addWidget(suggest)
        buttons.addWidget(restore)            # 儲存改由「確認並下一筆」（換筆時也會自動儲存）
        button_box = QWidget()
        button_box.setLayout(buttons)
        buttons.setContentsMargins(0, 0, 0, 0)
        grid.addWidget(button_box, extra + 2, 0, 1, 3)
        self.edit_widgets += [self.rule, label, self.tags, tag_label, button_box]
        for col in (1, 2):
            grid.setColumnStretch(col, 1)
        self.editors['deterioration'].currentTextChanged.connect(self.refresh_choices)
        self.editors['remark'].textEdited.connect(self.remark_typed)
        self.editors['d_location'].textEdited.connect(self.location_typed)
        self.editors['amount'].textEdited.connect(self.amount_typed)
        self.editors['unit'].activated.connect(self.amount_typed)
        for combo in self.deru:
            combo.currentTextChanged.connect(self.mark)
        for editor in self.editors.values():
            (editor.currentTextChanged if isinstance(editor, QComboBox) else editor.textChanged).connect(self.mark)

        self.kind = QComboBox()
        self.note = ''                       # 判定說明（舊審查檔的說明照舊保留；介面不再提供輸入）
        self.status = QLabel('')
        accept = QPushButton('確認並下一筆')
        accept.setStyleSheet('font-weight:bold;')
        accept.clicked.connect(lambda: self.decided.emit(self.kind.currentText(), self.note))
        relink = QPushButton('改配對…')
        relink.setToolTip('配錯對象、或同一筆缺失前後期記在不同構件（例：河道掛淤改記在橋墩）時，手動指定對應的缺失')
        relink.clicked.connect(self.relink.emit)
        back, forward = QPushButton('上一筆'), QPushButton('下一筆')
        back.clicked.connect(lambda: self.moved.emit(-1))
        forward.clicked.connect(lambda: self.moved.emit(1))
        self.pending_only = QCheckBox('只在待處理之間移動')
        self.pending_only.setChecked(True)
        for button in (relink, back, forward, accept):
            button.setMinimumWidth(40)
        actions = QHBoxLayout()               # 判定與移動合成一列
        actions.setContentsMargins(0, 0, 0, 0)
        for widget in (QLabel('判定'), self.kind, relink, back, forward, self.pending_only):
            actions.addWidget(widget)
        actions.addStretch()
        actions.addWidget(self.status)
        actions.addWidget(accept)
        # 填寫狀態（與填寫清單同步）：需確認時在核對當下就能處理
        self.fill_banner = QFrame()
        banner = QHBoxLayout(self.fill_banner)
        banner.setContentsMargins(6, 3, 6, 3)
        self.fill_text = QLabel('')
        self.fill_text.setWordWrap(True)
        banner.addWidget(self.fill_text, 1)          # 只說明情況；要填的值直接顯示在本期欄
        self.fill_banner.hide()
        decision_box = QWidget()
        decision_box.setLayout(actions)
        grid.setVerticalSpacing(2)
        grid.setContentsMargins(0, 0, 0, 0)
        self.reasons.setWordWrap(False)       # 理由一行（完整內容看滑鼠提示），節省高度
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(3)
        layout.addLayout(heading)
        layout.addWidget(self.reasons)
        layout.addWidget(self.fill_banner)
        layout.addWidget(decision_box)     # 判定與「確認並下一筆」放最上方：1080p 修改時也不會被捲出畫面
        layout.addLayout(grid)
        layout.addStretch()
        if fill_mode:
            decision_box.hide()
            self.edit_toggle.hide()
        self.set_editing(fill_mode)

    # ------------------------------------------------------------ 顯示
    def set_rules(self, rules):
        self.rules = rules or []

    def set_editing(self, on):
        on = bool((on or self.fill_mode) and self.pair is not None and self.pair.current is not None)
        if self.edit_toggle.isChecked() != on:
            self.edit_toggle.setChecked(on)
        for widget in self.edit_widgets:
            widget.setVisible(on)
        for stack in self.stacks:              # 本期欄直接換成修改欄位（不另開一欄）
            stack.setCurrentIndex(1 if on else 0)
        self.heads[2].setText('本期（修改中：藍底＝與網站不同）' if on else '本期（將填寫；藍底＝會修改網站）')

    def show_pair(self, pair, decision, position, edits=None, planned=None):
        """planned：這一筆將填寫的值（程式依規則修正的結果＋人工修改）；本期欄直接顯示，不另外用文字說明。"""
        self.pair = pair
        any_side = pair.current or pair.prior
        component = component_label(pair)
        self.title.setText(f'{component}｜{pair.code or "—"}｜程式判定：{pair.kind}（{pair.confidence}）'
                           f'　{position}')
        self.reasons.setText('理由：' + ('；'.join(pair.reasons) if pair.reasons else '—'))
        self.reasons.setToolTip(self.reasons.text())
        planned = planned or {}
        for label, key in FIELDS:
            before, site = field_value(pair.prior, key), field_value(pair.current, key)
            if key is None:
                after = ' '.join(planned.get(k, '') or '_' for k in 'deru') if planned else site
            else:
                after = planned.get(key, site) if planned else site
            changed = pair.prior is not None and pair.current is not None and before != after
            prior_cell, current_cell = self.cells[(label, 1)], self.cells[(label, 2)]
            prior_cell.setText(before or ('（無）' if pair.prior is None else ''))
            prior_cell.setStyleSheet(f'background:{DIFFERENT};' if changed else '')
            current_cell.setText(after or ('（無）' if pair.current is None else ''))
            filled = pair.current is not None and after.replace(' ', '') != site.replace(' ', '')
            current_cell.setStyleSheet(f'background:{EDITED};' if filled else
                                       (f'background:{DIFFERENT};' if changed else ''))
            current_cell.setToolTip(f'網站現值：{site or "（空白）"}' if filled else '')
        self.heads[2].setText('本期（將填寫；藍底＝會修改網站）' if pair.current is not None else '本期')
        self.kind.clear()
        self.kind.addItems(choices_for(pair))
        self.kind.setCurrentText(decision.kind)
        self.note = decision.note
        self.status.setText(f'審查：{decision.status}' if decision.at or decision.status == PENDING else '審查：待確認')
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
        self.import_button.setEnabled(pair.prior is not None)
        for key in ('deterioration', 'unit'):
            combo = self.editors[key]
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(current.options.get(key, []))
            combo.blockSignals(False)
        self.load({**self.site_values, **(planned or {}), **(edits or {})})
        self.saved_values = self.values()
        self.planned = self.values()
        self.refresh_choices()
        self.set_editing(self.edit_toggle.isChecked())

    def type_rules(self):
        """與「修改後」劣化類型相同的規則（依樣態相似度排序）；這個類型沒有規則時用全部。"""
        kind = self.editors['deterioration'].currentText().strip()
        same = [r for r in self.component_rules if r.get('deterioration') == kind]
        return same or self.component_rules

    def refresh_choices(self, *_):
        """Excel 規則下拉與樣態按鈕：只列目前劣化類型的規則（換類型時即時更新）。"""
        if self.pair is None or self.pair.current is None:
            return
        rules = self.type_rules()
        self.rule.blockSignals(True)
        self.rule.clear()
        kind = self.editors['deterioration'].currentText().strip()
        self.rule.addItem(f'（{kind or "全部"}的 Excel 規則：帶入工法、單位、單價，並替換備註樣態）', None)
        for rule in rules:
            self.rule.addItem(f"{rule['pattern']}｜{rule['method']}｜{rule['unit']}｜{rule['price']}", rule)
        self.rule.blockSignals(False)
        while self.tag_layout.count():
            self.tag_layout.takeAt(0).widget().deleteLater()
        patterns = list(dict.fromkeys(r['pattern'] for r in rules))[:12]
        for i, pattern in enumerate(patterns):
            button = QPushButton(pattern)
            button.setToolTip('把備註中間的樣態換成這個（保留位置與尺寸），並帶入對應的工法、單位、單價')
            button.clicked.connect(lambda _=False, p=pattern: self.pick_pattern(p))
            self.tag_layout.addWidget(button, i // 6, i % 6)

    def dirty(self):
        """修改後欄位有尚未儲存的變更。"""
        return self.pair is not None and self.pair.current is not None and self.values() != getattr(self, 'saved_values', {})

    def import_prior(self):
        """導入前期並按照規則修正：
        - 本期只填範圍（例：0.5mm*2m*2）或只有數量：位置與敘述沿用前期，範圍與數量用本期（不洗掉本期填的範圍）；
        - 其他情況：帶入前期的類型、位置、工法、數量、單位、單價與備註；
        之後套用 Excel 規則（樣態、工法、單價）與位置寫法（S2G5 → 2G5、1D-2-1 → 1D2-1）。DERU 維持本期。"""
        if self.pair is None or self.pair.prior is None or self.pair.current is None:
            return
        current = self.values()
        prior = {k: str(self.pair.prior.fields.get(k, '') or '').strip() for k in TEXT_FIELDS}
        merged, why = merge_with_prior({k: current.get(k, '') for k in TEXT_FIELDS}, prior)
        values = {**current, **merged} if merged else {**current, **{k: v for k, v in prior.items() if v}}
        component, context = self.pair.current.component, dict(self.pair.current.context)
        note = why or '已導入前期資料'
        if self.rules:
            try:
                decision = automatic_plan({k: values.get(k, '') for k in TEXT_FIELDS}, component, [], self.rules, context)
            except Exception:      # 規則判斷失敗：保留導入的值
                decision = None
            if decision is not None and not decision.review:
                values.update({k: str(v or '').strip() for k, v in decision.fields.items() if k in TEXT_FIELDS})
                note += '，並依 Excel 規則修正'
            elif decision is not None and NO_RULE in decision.reason:
                adopted = adopt_prior_method({k: values.get(k, '') for k in TEXT_FIELDS}, prior)
                if adopted:
                    values.update(adopted)
                    note += '（Excel 沒有此樣態，沿用前期工法）'
        values.update(tidy_locations({k: values.get(k, '') for k in ('d_location', 'remark')}, component, context))
        self.load(values)
        self.edit_status.setText(note)

    def compose(self, pattern):
        """樣態按鈕：只換備註中間的敘述。備註空白或只有範圍時，補上位置（本期損壞位置或前期位置）與範圍
        （本期填的範圍 → 前期範圍 → 依數量），例：空白＋前期 A1-L,R 裂縫。(2m)*2 → A1-L,R 混凝土裂縫。(2m)*2。"""
        text = self.editors['remark'].text().strip()
        unit = self.editors['unit'].currentText().strip()
        prior = self.pair.prior.fields if self.pair is not None and self.pair.prior is not None else {}
        old_location, _, _ = split_remark(prior.get('remark'))
        old_size = size_part(prior.get('remark'))
        location = self.editors['d_location'].text().strip() or old_location
        found = only_size(text) if text else ('', '')
        if found is not None:
            here, size_text = found
            size = (normalize_size(size_text, unit) or size_text) if size_text else ''
            if not size:
                amount = self.editors['amount'].text().strip()
                if amount and unit in COUNT_UNITS:
                    size = f'({amount}{unit})'
                elif amount and unit == '公尺':
                    size = f'({amount}m)'
                else:
                    size = old_size
            return ' '.join(x for x in (here or location, f'{pattern}。{size}') if x)
        here, _, _ = split_remark(text)
        if not here and location:
            text = f'{location} {text}'
        return replace_pattern(text, pattern)

    # 備註與損壞位置、數量、單位雙向同步（只在使用者輸入時觸發）
    def remark_typed(self, text):
        location, description, _ = split_remark(text)
        if location:
            self.editors['d_location'].setText(location)
        size = size_part(text)
        unit = self.editors['unit'].currentText().strip()
        if size and unit and description:
            amount = quantity(description, size, unit)
            if amount is not None:
                self.editors['amount'].setText(f'{amount.normalize():f}')

    def location_typed(self, text):
        remark = self.editors['remark'].text().strip()
        location, description, tail = split_remark(remark)
        if description or tail:
            self.editors['remark'].setText(f'{text.strip()} {description}{tail}'.strip())

    def amount_typed(self, *_):
        """數量或單位改了：計數單位（處、個…）或公尺的單一範圍，備註範圍跟著改；面積、體積由備註換算數量。"""
        remark = self.editors['remark'].text().strip()
        amount, unit = self.editors['amount'].text().strip(), self.editors['unit'].currentText().strip()
        size = size_part(remark)
        if not remark or not amount:
            return
        if unit in COUNT_UNITS and (not size or re.fullmatch(r'[\(（]\s*[\d.]+\s*\D*[\)）]', size)):
            new = f'({amount}{unit})'
        elif unit == '公尺' and (not size or re.fullmatch(r'[\(（]\s*[\d.]+\s*m?\s*[\)）]', size)):
            new = f'({amount}m)'
        else:
            return
        body = remark[:len(remark) - len(size)].rstrip() if size else remark.rstrip()
        self.editors['remark'].setText((body if body.endswith('。') else body.rstrip('。') + '。') + new)

    def set_fill_status(self, item):
        """說明這一筆的填寫情況（不放按鈕）：需確認（黃）、會修改網站（綠）、已送出（藍）。要填的值在本期欄。"""
        if item is None or item.status not in (REVIEW, KEPT, READY, DONE):
            self.fill_banner.hide()
            return
        color = {REVIEW: '#fff2cc', KEPT: '#e4dfec', READY: '#e2efda', DONE: '#ddebf7'}[item.status]
        text = {REVIEW: f'需確認：{item.reason}。本期欄已帶入程式能判斷的部分，請檢查後按「確認並下一筆」',
                KEPT: f'已確認維持網站現值',
                READY: f'將修改網站：{item.reason}',
                DONE: f'已送出：{item.message}'}[item.status]
        self.fill_banner.setStyleSheet(f'QFrame{{background:{color}; border-radius:4px;}}')
        self.fill_text.setText(text)
        self.fill_banner.show()

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
        remark = self.compose(rule['pattern'])
        self.editors['remark'].setText(remark)
        # 依規則單位由備註尺寸重算數量與位置（例：前期 1 式、(10m*2m) → 平方公尺 20）
        try:
            info = parse_remark(remark, rule['pattern'], rule['unit'], aliases(rule.get('note', '')))
        except ValueError:
            return
        if info.unit == rule['unit'].replace('／', '/'):
            self.editors['amount'].setText(info.amount)
            self.editors['d_location'].setText(info.location)
            self.editors['remark'].setText(info.text)

    def pick_rule(self, index):
        rule = self.rule.itemData(index)
        if rule:
            self.apply_rule(rule)

    def pick_pattern(self, pattern):
        rule = best_rule(self.type_rules(), pattern, self.editors['unit'].currentText())
        if rule:
            self.apply_rule(rule)
            self.rule.setCurrentIndex(self.rule.findData(rule))
        else:
            self.editors['remark'].setText(self.compose(pattern))

    def load_proposal(self):
        if self.pair is None or self.pair.current is None:
            return
        try:
            decision = proposal_for(self.pair, self.rules, self.kind.currentText() == KIND_PRIOR_WRONG) if self.rules else None
        except Exception as error:
            decision, reason = None, f'程式建議失敗：{error}'
        else:
            reason = decision.reason if decision else '沒有讀到 Excel 規則表'
        if decision:
            self.load({**self.values(), **{k: str(v or '').strip() for k, v in decision.fields.items() if k in self.editors}})
        self.edit_status.setText(('程式建議（需確認）：' if decision is None or decision.review else '程式建議：') + reason)

    def restore(self):
        self.load(self.site_values)
        self.edit_status.setText('已改回網站現值；按「確認並下一筆」即維持網站現值')


COMPARE_SHARE = 0.45       # 前後期照片比對時網站區佔的比例（只顯示照片）
PAGE_WIDTH = 860            # 網站主頁面（含左側構件列表）約需的寬度（CSS 像素）
ZOOM_RANGE = (0.6, 1.0)


def fit_zoom(width):
    """依網站面板寬度自動縮放，讓第一級、第二級與主頁面不必左右拖動（螢幕解析度不同也適用）。"""
    low, high = ZOOM_RANGE
    return round(min(high, max(low, width / PAGE_WIDTH)), 2)


def site_width(total, mode):
    """網站區預設寬度：依比例分配，但每個網站最多到網頁 100% 縮放所需的寬度；多出的空間給工作區
    （4K 等大螢幕網站不會留白，工作區變寬）。"""
    count = 2 if mode == 'both' else 1
    return int(min(total * SITE_SHARE[mode], count * (PAGE_WIDTH * ZOOM_RANGE[1] + 40)))


class FitView(SiteView):
    """面板寬度改變（視窗縮放、拖動分隔線、切換單／雙網站、不同解析度）時自動調整網頁縮放。"""
    zoomed = Signal(float)

    def __init__(self, profile, parent=None):
        super().__init__(profile, parent)
        self.loadFinished.connect(lambda _ok: self.fit())

    def fit(self):
        if self.width() > 50:
            zoom = fit_zoom(self.width())
            self.setZoomFactor(zoom)
            self.zoomed.emit(zoom)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit()


def open_photos(view, url, photos_only=False):
    """開啟第二級缺失頁，載入後捲到下方的「損壞照片」；photos_only：只留照片（前後期比對用）。"""
    def scroll(ok):
        view.loadFinished.disconnect(scroll)
        if ok:
            view.probe(PHOTOS_ONLY_JS if photos_only else SCROLL_TO_PHOTOS, lambda _: None)
    view.loadFinished.connect(scroll)
    view.open(url)


class MainWindow(QMainWindow):
    def __init__(self, url=LOGIN_URL, profile_dir=PROFILE_DIR, scan_dir=SCAN_DIR, review_dir=REVIEW_DIR,
                 delay_ms=500, rewrite=None, output_dir=OUTPUT_DIR, rules_path=None, supervisor_path=None):
        super().__init__()
        self.setWindowTitle('公路橋梁填報助手')
        self.resize(1680, 960)
        self.scan_dir, self.review_dir, self.output_dir = Path(scan_dir), Path(review_dir), Path(output_dir)
        self.pairs, self.visible, self.review, self.review_file = [], [], Review(), None
        self.auto_pairs, self.order, self.prior_main, self.current_main = [], {}, '', ''
        self.current_header, self.prior_header = {}, {}
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
        lite = QPushButton('顯示完整網頁')
        lite.setToolTip('目前為精簡網頁（隱藏標題列、選單、檢測人員照片、檢測路徑、3D／Word 按鈕；構件列表收成按鈕）。'
                        '按下切換為網站原貌；只改畫面，不影響資料')
        lite.clicked.connect(self.toggle_lite)
        bar.addWidget(site)
        bar.addWidget(hide)
        bar.addWidget(lite)
        self.buttons['site_mode'], self.buttons['site'], self.buttons['lite'] = site, hide, lite
        self.lite = True
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
                                ('temp', '開啟暫存資料夾', self.open_folder),
                                ('chiefs', '單位主管設定…', self.edit_supervisors)):
            self.menu_actions[key] = self.more.addAction(text, slot)
        more = QToolButton()
        more.setText('其他功能')
        more.setMenu(self.more)
        more.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        bar.addWidget(more)
        self.summary = QLabel('登入網站後，輸入橋名按「開始」：程式自動查詢、找出本期與前期、掃描並配對。目前全程唯讀，不修改網站。')
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet('background:#eef4fb; padding:6px; border-radius:6px;')

        # ① 核對與修改：一張缺失清單（審查與填寫合併成一個狀態）＋下方工作區
        self.filter = QComboBox()
        self.filter.addItems(FILTERS)
        self.filter.currentTextChanged.connect(self.refresh_table)
        self.overall_filter = '待處理'
        self.status_buttons = {}
        cards = QHBoxLayout()
        for key in OVERALL_CARDS:
            button = QPushButton(key)
            button.setCheckable(True)
            button.setMinimumHeight(40)
            button.setMinimumWidth(40)                    # 可縮窄，不撐寬工作區
            button.setToolTip(OVERALL_TIPS[key])
            button.clicked.connect(lambda _=False, k=key: self.set_overall_filter(k))
            cards.addWidget(button)
            self.status_buttons[key] = button
        self.table = QTableWidget(0, len(TABLE_COLUMNS))
        self.table.setHorizontalHeaderLabels(TABLE_COLUMNS)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setWordWrap(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        rows = self.table.verticalHeader()          # 每列固定兩行高；完整內容看滑鼠提示或下方工作區
        rows.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        rows.setDefaultSectionSize(self.table.fontMetrics().lineSpacing() * 2 + 8)
        self.table.currentCellChanged.connect(lambda row, *_: self.select_row(row))
        self.filter.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)   # 篩選文字不省略
        self.table.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerItem)        # 捲動一次一列（分段）
        self.table.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        # 工具列：左側照片與回到缺失主頁面；右側版面、重新掃描、重新整理、進階篩選（送出都在 ② 送出資料）
        toolbar = QHBoxLayout()
        self.toolbar_left = QHBoxLayout()
        toolbar.addLayout(self.toolbar_left)
        toolbar.addStretch()
        self.side_toggle = QPushButton('工作區放右側')
        self.side_toggle.setCheckable(True)
        self.side_toggle.setMinimumWidth(40)
        self.side_toggle.setToolTip('1080p 等矮螢幕：清單與工作區左右並排，捲動清單時不必縮小工作區（設定會記住）')
        self.side_toggle.toggled.connect(self.set_work_side)
        toolbar.addWidget(self.side_toggle)
        for key, text, tip, slot in (
                ('fill_rescan', '重新掃描本期', '重新讀取網站本期缺失（前期沿用上次掃描），更新清單', lambda: self.rescan_current()),
                ('fill_plan', '重新整理', '依最新的審查與修改重新計算', lambda: self.build_fill_plan())):
            button = QPushButton(text)
            button.setToolTip(tip)
            button.setMinimumWidth(40)
            button.clicked.connect(slot)
            toolbar.addWidget(button)
            self.buttons[key] = button
        toolbar.addWidget(QLabel('進階篩選'))
        toolbar.addWidget(self.filter)
        self.card = ReviewCard()
        for button in self.card.photo_buttons.values():
            self.toolbar_left.addWidget(button)
        self.card.decided.connect(self.decide)
        self.card.moved.connect(self.move)
        self.card.relink.connect(self.relink)
        self.card.photos.connect(self.show_photos)
        self.card.edited.connect(self.save_edit)
        self.card.kept.connect(lambda: self.keep_current(self.pairs[self.index] if 0 <= self.index < len(self.pairs) else None))
        card_scroll = QScrollArea()
        card_scroll.setWidgetResizable(True)
        card_scroll.setWidget(self.card)
        check_parts = QSplitter(Qt.Orientation.Vertical)
        check_parts.addWidget(self.table)
        check_parts.addWidget(card_scroll)
        check_parts.setChildrenCollapsible(False)
        check_parts.setSizes([320, 520])
        self.check_parts = check_parts
        self.check_page = QWidget()
        check_box = QVBoxLayout(self.check_page)
        check_box.addLayout(cards)
        check_box.addLayout(toolbar)
        check_box.addWidget(check_parts, 1)
        self.rules_path = Path(rules_path) if rules_path else rules_file()
        self.migrated_rules = 0
        old = APP_DIR / 'data' / '劣化說明與維修單價1111v2.xlsx'
        if rules_path is None and self.rules_path and self.rules_path.name == RULES_NAME and old.exists():
            try:
                self.migrated_rules = migrate_auto_rules(old, self.rules_path)   # 舊檔自動新增的規則搬進 1116
            except Exception:  # 規則表開啟中等：不影響啟動
                self.migrated_rules = 0
        self.supervisor_path = Path(supervisor_path) if supervisor_path else APP_DIR / 'data' / 'supervisors.json'
        self.rules = load_rules(self.rules_path)
        self.card.set_rules(self.rules)

        # ② 送出：所有「可送出」的變更預覽、整批送出與結果（修改一律在 ①）
        fill = QWidget()
        fill_box = QVBoxLayout(fill)
        explain = QLabel('列出要填寫到網站的變更（可送出），以及送出結果。修改請在「① 核對與修改」；雙擊一列回到那一筆。'
                         '可整批送出、依構件送出，或在表格中多選（Ctrl／Shift＋點選、Ctrl＋A 全選）後送出選取的項目。'
                         '每筆送出後重新開啟核對，結果不明不會重送。')
        explain.setWordWrap(True)
        self.fill_summary = QLabel('')
        self.fill_table = QTableWidget(0, len(FILL_COLUMNS))
        self.fill_table.setHorizontalHeaderLabels(FILL_COLUMNS)
        self.fill_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.fill_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.fill_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)    # 多選：Ctrl／Shift＋點選
        self.fill_table.setWordWrap(True)
        for col, width in enumerate((70, 100, 120, 70, 300, 240)):
            self.fill_table.setColumnWidth(col, width)
        self.fill_table.setMinimumWidth(200)
        self.fill_table.horizontalHeader().setStretchLastSection(True)
        self.fill_table.cellDoubleClicked.connect(self.open_fill_item)
        self.fill_table.itemSelectionChanged.connect(self.refresh_send_buttons)
        self.fill_component = QComboBox()
        self.fill_component.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.fill_component.setToolTip('只列出這個構件；「送出此構件」送出這個構件全部「可送出」的項目')
        self.fill_component.currentIndexChanged.connect(lambda _: self.refresh_fill_table())
        actions = QHBoxLayout()
        actions.addWidget(QLabel('構件'))
        actions.addWidget(self.fill_component)
        for key, text, tip, slot in (
                ('send_component', '送出此構件', '送出目前選的構件全部「可送出」的項目（逐筆送出並回讀核對；遇到結果不明立即停止）',
                 lambda: self.fill_component_items()),
                ('send_one', '送出選取的項目', '只送出表格中選取的項目（可多選：Ctrl／Shift＋點選、Ctrl＋A 全選；'
                 '可送出、送出待確認或失敗的才會送出）', lambda: self.fill_selected()),
                ('send_all', '整批送出', '填寫全部「可送出」：逐筆送出並回讀核對；遇到結果不明立即停止', lambda: self.fill_all())):
            button = QPushButton(text)
            button.setToolTip(tip)
            button.setMinimumWidth(40)
            button.clicked.connect(slot)
            actions.addWidget(button)
            self.buttons[key] = button
        self.fill_summary.setWordWrap(True)
        fill_box.addWidget(explain)
        fill_box.addLayout(actions)
        fill_box.addWidget(self.fill_summary)
        fill_box.addWidget(self.fill_table, 1)
        self.fill_items, self.fill_visible, self.fill_log, self.filler = [], [], None, None

        # ② 填寫 → 檢測員意見（第 7 步）
        opinion = QWidget()
        box = QVBoxLayout(opinion)
        tip = QLabel('勾選即時增減意見；左下選定或填寫【】後括號消失；右側可直接修改。')
        tip.setStyleSheet('color:#666;')
        tip.setWordWrap(True)
        tip.setToolTip('6 個標準文字一律列出，符合條件自動勾選並帶入；勾選增減時保留右側的手動修改。'
                       '「寫入網站」送出到本期缺失主頁面，送出後回讀核對，結果不明不重送。')
        row = QHBoxLayout()
        for key, text, hint, slot in (('opinion_build', '重新產生', '依目前的審查與修改重新產生（會取代右側手動修改）',
                                       lambda: self.build_opinion(force=True)),
                                      ('opinion_write', '寫入網站', '把右側意見寫入本期缺失主頁面的「檢測員意見」並送出',
                                       lambda: self.write_opinion())):
            button = QPushButton(text)
            button.setToolTip(hint)
            button.clicked.connect(slot)
            row.addWidget(button)
            self.buttons[key] = button
        self.opinion_status = QLabel('')
        self.opinion_status.setWordWrap(True)
        row.addWidget(self.opinion_status, 1)
        row.addWidget(tip)
        self.opinion_list = QListWidget()
        self.opinion_list.setWordWrap(True)                 # 長文字自動換行，不用左右捲動
        self.opinion_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.opinion_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.opinion_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.opinion_list.setSpacing(2)
        self.opinion_list.itemChanged.connect(lambda _: self.opinion_changed())
        self.choice_box = QWidget()
        self.choice_form = QGridLayout(self.choice_box)
        self.choice_form.setContentsMargins(0, 0, 0, 0)
        self.choice_form.setColumnStretch(1, 1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(self.choice_box)
        left = QSplitter(Qt.Orientation.Vertical)
        for title, widget in (('條目（勾選要寫入的）', self.opinion_list), ('待選與待填（只列勾選中的條目）', scroll)):
            holder = QWidget()
            holder_box = QVBoxLayout(holder)
            holder_box.setContentsMargins(0, 0, 0, 0)
            heading = QLabel(title)
            heading.setStyleSheet('font-weight:bold;')
            holder_box.addWidget(heading)
            holder_box.addWidget(widget)
            left.addWidget(holder)
        left.setSizes([600, 220])
        self.opinion_edit = QPlainTextEdit()
        self.opinion_edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.opinion_edit.textChanged.connect(self.opinion_typed)
        self.prior_opinion = QPlainTextEdit(readOnly=True)
        self.site_opinion = QPlainTextEdit(readOnly=True)
        editor = QWidget()
        editor_box = QVBoxLayout(editor)
        editor_box.setContentsMargins(0, 0, 0, 0)
        heading = QLabel('意見（可直接修改）')
        heading.setStyleSheet('font-weight:bold;')
        advice = QGridLayout()                       # 2×2，避免工作區被撐寬
        advice.setContentsMargins(0, 0, 0, 0)
        advice.addWidget(heading, 0, 0)
        advice.addWidget(QLabel('建議：'), 1, 0, Qt.AlignmentFlag.AlignRight)
        for number, text in enumerate(ADVICE):
            button = QPushButton(text.replace('建議', ''))
            button.setToolTip(f'游標所在那一條改成「{text}」（沒有建議時補上）；游標不在編號條目上時插入在游標處')
            button.clicked.connect(lambda _=False, t=text: self.insert_advice(t))
            advice.addWidget(button, number // 2, number % 2 + 1)
        advice.setColumnStretch(3, 1)
        editor_box.addLayout(advice)
        editor_box.addWidget(self.opinion_edit)
        references = QTabWidget()                    # 參考資料收成分頁，畫面只留一個主要編輯框
        references.addTab(self.site_opinion, '網站目前的意見')
        references.addTab(self.prior_opinion, '前期意見（可複製本橋特有文字）')
        right = QSplitter(Qt.Orientation.Vertical)
        right.addWidget(editor)
        right.addWidget(references)
        right.setSizes([620, 200])
        left.setMinimumWidth(120)
        right.setMinimumWidth(200)
        parts = QSplitter(Qt.Orientation.Horizontal)
        parts.addWidget(left)
        parts.addWidget(right)
        parts.setSizes([300, 500])
        box.addWidget(tip)
        box.addLayout(row)
        box.addWidget(self.build_basic_box())
        box.addWidget(parts, 1)
        self.opinion_items, self.opinion_inputs, self.opinion_values, self.opinion_loading = [], {}, {}, False
        self.height_line = ''
        self.opinion_tab = opinion
        self.send_tab = self.fill_tab = fill          # ② 送出資料
        self.photo_panel = PhotoPanel()
        self.photo_panel.changed.connect(self.save_photo_plan)
        self.photo_panel.upload.connect(self.upload_photos)
        self.photo_panel.upload_button.setEnabled(False)
        self.inventory = ''
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
        for widget, text in zip((self.check_page, fill, self.opinion_tab, self.photo_panel, export), STEPS):
            self.tabs.addTab(widget, text)
        self.tabs.currentChanged.connect(self.tab_changed)
        log_box = QWidget()
        log_layout = QVBoxLayout(log_box)
        log_layout.setContentsMargins(0, 0, 0, 0)
        self.log_toggle = QToolButton()
        self.log_toggle.setText('紀錄')
        self.log_toggle.setCheckable(True)
        self.log_toggle.setChecked(True)
        self.log_toggle.setArrowType(Qt.ArrowType.DownArrow)
        self.log_toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.log_toggle.setAutoRaise(True)
        self.log_toggle.setToolTip('收合／展開紀錄（收合後清單與工作區更高）')
        self.log_toggle.toggled.connect(self.set_log_visible)
        log_layout.addWidget(self.log_toggle)
        log_layout.addWidget(self.output)
        self.work = QSplitter(Qt.Orientation.Vertical)     # 右側工作區：上步驟、下紀錄（可拖拉高度）
        self.work.addWidget(self.tabs)
        self.work.addWidget(log_box)
        self.work.setSizes([820, 100])

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.sites = QSplitter(Qt.Orientation.Horizontal)
        self.sites.addWidget(self.prior_view)
        self.sites.addWidget(self.view)
        self.splitter.addWidget(self.sites)
        self.splitter.addWidget(self.work)
        self.splitter.setSizes([980, 700])
        self.layout_file = self.scan_dir.parent / 'layout.json'
        self.splitter.splitterMoved.connect(self.remember_layout)
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
        # 讀取網頁的進度條（像一般瀏覽器）：左側網站與背景讀取都會顯示
        self.load_label = QLabel('')
        self.load_bar = QProgressBar()
        self.load_bar.setRange(0, 100)
        self.load_bar.setFixedWidth(180)
        self.load_bar.setTextVisible(True)
        self.zoom_label = QLabel('')
        self.zoom_label.setToolTip('網頁縮放比例：依網站面板寬度自動調整（60%～100%）')
        self.statusBar().addPermanentWidget(self.zoom_label)
        for view in (self.view, self.prior_view):
            view.zoomed.connect(lambda _z: self.show_zoom())
        self.statusBar().addPermanentWidget(self.load_label)
        self.statusBar().addPermanentWidget(self.load_bar)
        self.load_bar.hide()
        self.loading = {}
        for page, name in ((self.view.page(), '本期網站'), (self.prior_view.page(), '前期網站'),
                           (self.fetcher.page, '背景讀取')):
            self.watch_loading(page, name)
        self.set_sites(current=True)
        saved = self.layout_ratios()
        self.side_toggle.setChecked(bool(saved.get('work_side')))
        if saved.get('log_hidden'):
            self.log_toggle.setChecked(False)
        self.view.open(url)
        if self.migrated_rules:
            self.log(f'舊規則表（1111v2）中自動新增的 {self.migrated_rules} 條規則已搬入 {RULES_NAME}（舊檔移到 backup）')

    # ------------------------------------------------------------ 共用
    def tab_changed(self, _index=0):
        widget = self.tabs.currentWidget()
        if widget is self.send_tab:
            self.build_fill_plan()
        elif widget is self.opinion_tab:
            self.build_opinion()

    def log(self, message):
        self.output.appendPlainText(f'{datetime.now():%H:%M:%S} {message}')

    def set_busy(self, busy):
        self.busy = busy
        for key in ('scan', 'open'):
            self.menu_actions[key].setEnabled(not busy)
        self.buttons['export'].setEnabled(not busy and bool(self.pairs))
        for key in FILL_BUTTONS:
            self.buttons[key].setEnabled(not busy and bool(self.pairs))
        if hasattr(self, 'fill_component'):
            self.refresh_send_buttons()          # 送出選取／此構件：依選取與構件決定能不能按
        self.start_button.setEnabled(not busy)
        self.buttons['stop'].setEnabled(busy)

    def stop(self):
        self.stop_requested = True
        self.log('停止中，完成目前這一頁後停止…')

    def show_zoom(self):
        parts = [f'{name} {round(view.zoomFactor() * 100)}%' for name, view in (('本期', self.view), ('前期', self.prior_view))
                 if view.isVisibleTo(self)]
        self.zoom_label.setText('網頁縮放：' + '｜'.join(parts) if parts else '')

    def watch_loading(self, page, name):
        page.loadStarted.connect(lambda: self.show_loading(name, 0))
        page.loadProgress.connect(lambda value: self.show_loading(name, value))
        page.loadFinished.connect(lambda _ok: self.show_loading(name, None))

    def show_loading(self, name, value):
        """value=None 表示這個頁面載入完成；全部完成才隱藏進度條。"""
        if value is None or value >= 100:
            self.loading.pop(name, None)
        else:
            self.loading[name] = value
        if not self.loading:
            self.load_bar.hide()
            self.load_label.setText('')
            return
        name, value = next(reversed(self.loading.items()))
        self.load_label.setText(f'{name}載入中')
        self.load_bar.setValue(value)
        self.load_bar.show()

    def toggle_lite(self):
        """精簡網頁 ↔ 完整網頁（只改畫面，不重新載入、不影響資料）。"""
        self.lite = not self.lite
        set_lite(self.profile, self.lite, (self.view, self.prior_view))
        self.buttons['lite'].setText('顯示完整網頁' if self.lite else '精簡網頁')
        self.log('網站顯示：' + ('精簡網頁' if self.lite else '完整網頁'))

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
        self.show_zoom()
        if mode != 'none':
            total = max(self.splitter.width(), self.width(), 1000)
            saved = self.layout_ratios().get(mode)
            sites = int(total * saved) if saved else site_width(total, mode)
            self.splitter.setSizes([sites, total - sites])
            if mode == 'both':
                self.sites.setSizes([1, 1])

    # 使用者拖動網站與工作區的分隔線：記住各顯示方式的比例，下次沿用
    def layout_ratios(self):
        try:
            return json.loads(self.layout_file.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            return {}

    def remember_setting(self, key, value):
        ratios = self.layout_ratios()
        ratios[key] = value
        try:
            self.layout_file.parent.mkdir(parents=True, exist_ok=True)
            self.layout_file.write_text(json.dumps(ratios), encoding='utf-8')
        except OSError:
            pass

    def set_work_side(self, side):
        """① 的清單與工作區：上下（預設）或左右並排。"""
        self.check_parts.setOrientation(Qt.Orientation.Horizontal if side else Qt.Orientation.Vertical)
        total = sum(self.check_parts.sizes()) or 800
        self.check_parts.setSizes([int(total * (0.45 if side else 0.38)), int(total * (0.55 if side else 0.62))])
        self.side_toggle.setText('工作區放下方' if side else '工作區放右側')
        self.remember_setting('work_side', bool(side))

    def set_log_visible(self, on):
        self.output.setVisible(on)
        self.log_toggle.setArrowType(Qt.ArrowType.DownArrow if on else Qt.ArrowType.RightArrow)
        total = sum(self.work.sizes()) or 900
        bar = self.log_toggle.sizeHint().height() + 2
        self.work.setSizes([max(total - 100, 100), 100] if on else [total - bar, bar])   # 收合時空間全給清單與工作區
        self.remember_setting('log_hidden', not on)

    def remember_layout(self, *_):
        mode = self.site_mode()
        sizes = self.splitter.sizes()
        if mode == 'none' or sum(sizes) <= 0:
            return
        ratios = self.layout_ratios()
        ratios[mode] = round(sizes[0] / sum(sizes), 3)
        try:
            self.layout_file.parent.mkdir(parents=True, exist_ok=True)
            self.layout_file.write_text(json.dumps(ratios), encoding='utf-8')
        except OSError:
            pass

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
        compare = which == 'compare'
        sides = ('prior', 'current') if compare else (which,)
        for side in sides:
            defect = pair.current if side == 'current' else pair.prior
            # 比對時本期也用檢視頁（只看照片）；單看本期照片用編輯頁（可直接刪照片）
            url = defect.damage_url(edit=side == 'current' and not compare) if defect else ''
            if not url:
                QMessageBox.information(self, '照片', '這筆沒有第二級缺失頁的連結（請重新掃描）。')
                return
            self.set_sites(**{side: True})
            open_photos(self.view if side == 'current' else self.prior_view, url, photos_only=compare)
            self.log(f"開啟{'本期' if side == 'current' else '前期'}照片：{defect.component} {defect.remark.strip()}")
        if compare:                                   # 只看照片：網站區縮到約 45%，工作區按鈕不被壓縮
            total = max(self.splitter.width(), 1000)
            sites = int(total * COMPARE_SHARE)
            self.splitter.setSizes([sites, total - sites])
            self.sites.setSizes([1, 1])

    def save_edit(self, values):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        kept = self.review.set_edit(pair, values)
        self.review.save(self.review_file)
        self.log(f'修改本期：{pair.current.component}：' + (describe(kept) if kept else '還原為網站值'))
        if self.card.pair is pair:
            self.card.mark()
            self.card.saved_values = self.card.values()
        self.refresh_table()
        if self.fill_items:
            self.build_fill_plan()
        self.refresh_fill_status()
        self.update_summary()

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
        self.agency = bridge.agency or ''
        self.show_site(True)
        self.view.open(basic_urls(bridge.inventory_id)['basic'])
        latest = bridge.latest.date if bridge.latest else '—'
        note = ''
        if latest == '—' or not latest.startswith(self.this_year()):
            note = (f'\n※ 網站上還沒有 {self.this_year()} 年的檢測紀錄：確認後會先把最新一期（{latest}）當作前期掃描建檔，'
                    '再重新讀取歷史紀錄；有本年度紀錄才掃描本期，沒有就先停在這裡。')
        self.confirm_text.setText(f'請確認左側基本資料是否為要處理的橋：{bridge.name}｜橋梁系統ID {bridge.inventory_id}｜'
                                  f'{bridge.agency}｜{bridge.place}｜最新檢測 {latest}' + note)
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
            if not current.date.startswith(self.this_year()):
                current = self.prebuild_prior(bridge, history)
                if current is None:
                    self.log(f'仍沒有 {self.this_year()} 年的檢測紀錄：前期已建檔，本期先不掃描')
                    self.summary.setText(NO_CURRENT_TIP.format(bridge=bridge.name, year=self.this_year(), date=history[0].date))
                    QMessageBox.information(self, '前期已建檔', NO_CURRENT_TIP.format(
                        bridge=bridge.name, year=self.this_year(), date=history[0].date))
                    return
                self.log(f'已有本年度紀錄 {current.date}，開始掃描本期')
            if not current.edit_url:
                raise ValueError(f'最新一期 {current.date} 沒有「修改」連結，無法讀取（可能不是本公司的檢測）')
            self.log(f'本期：{current.date}（檢測系統ID {current.inspect_id}）')
            html = self.fetcher.get(current.edit_url)
            self.view.open(current.edit_url)          # 左側同步顯示本期缺失主頁面
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'中斷：{error}')
            QMessageBox.warning(self, '開始', str(error))
            return
        except Stopped as error:
            self.log(str(error))
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
        self.history_menu.addSeparator()
        clear = self.history_menu.addAction('清除歷史紀錄…', self.clear_history)
        clear.setEnabled(bool(sessions))

    def choose_history(self, sessions):
        """勾選要清除的歷史紀錄；回傳選取的項目（取消回傳空清單）。"""
        dialog = QDialog(self)
        dialog.setWindowTitle('清除歷史紀錄')
        listing = QListWidget()
        for item in sessions:
            row = QListWidgetItem(history_text(item))
            row.setFlags(row.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            row.setCheckState(Qt.CheckState.Unchecked)
            listing.addItem(row)
        def check_all(state):
            for i in range(listing.count()):
                listing.item(i).setCheckState(state)
        buttons = QDialogButtonBox()
        everything = buttons.addButton('全選', QDialogButtonBox.ButtonRole.ActionRole)
        nothing = buttons.addButton('全不選', QDialogButtonBox.ButtonRole.ActionRole)
        remove = buttons.addButton('清除勾選的紀錄', QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        everything.clicked.connect(lambda: check_all(Qt.CheckState.Checked))
        nothing.clicked.connect(lambda: check_all(Qt.CheckState.Unchecked))
        remove.clicked.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel('勾選要清除的紀錄（本期／前期掃描檔、審查與修改、填寫進度）。\n'
                                '只刪除本機紀錄，不影響網站資料；「判讀結果」資料夾的比對 Excel 不會刪除。'))
        layout.addWidget(listing)
        layout.addWidget(buttons)
        dialog.resize(760, 420)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return []
        return [sessions[i] for i in range(listing.count()) if listing.item(i).checkState() == Qt.CheckState.Checked]

    def confirm_clear(self, items):
        names = '\n'.join(f'・{history_text(i)}' for i in items[:10]) + (f'\n…另 {len(items) - 10} 筆' if len(items) > 10 else '')
        return QMessageBox.question(self, '清除歷史紀錄', f'確定清除 {len(items)} 筆紀錄？清除後無法復原：\n{names}') \
            == QMessageBox.StandardButton.Yes

    def clear_history(self):
        if self.busy:
            return
        sessions = saved_sessions(self.scan_dir, self.review_dir)
        items = self.choose_history(sessions)
        if not items or not self.confirm_clear(items):
            return
        removed = clear_sessions(items, self.scan_dir, self.review_dir)
        self.log(f'已清除 {len(items)} 筆歷史紀錄（{len(removed)} 個檔案）')
        current = (self.review.bridge, self.review.current_date)
        if any((i['bridge'], i['date']) == current for i in items):
            self.reset_session()

    def reset_session(self):
        """目前開著的紀錄被清除：清空畫面，避免之後自動存檔又把紀錄寫回來。"""
        self.pairs, self.auto_pairs, self.visible, self.fill_items = [], [], [], []
        self.review, self.review_file, self.index = Review(), None, -1
        self.refresh_table()
        self.refresh_fill_table()
        self.buttons['export'].setEnabled(False)
        self.tabs.setCurrentWidget(self.check_page)
        self.summary.setText('目前的紀錄已清除。輸入橋名按「開始」重新掃描，或從「歷史紀錄」開啟其他紀錄。')

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
        if getattr(self, 'agency', ''):
            header['管理機關'] = self.agency                # 查詢結果的管理機關（單位主管自動帶入用）
        self.stop_requested = False
        self.set_busy(True)
        started = datetime.now()
        options = dict(progress=self.progress, stopped=lambda: self.stop_requested)
        try:
            self.log(f"掃描本期：{header['橋梁名稱']} {header['檢測日期']}")
            current = scan_inspection(html, self.fetcher, **options)
            if header.get('管理機關'):
                current.header['管理機關'] = header['管理機關']
            self.save_json(header, '檢測', current.to_dict())
            record = find_prior(html, self.fetcher)
            prior = self.prebuilt_prior(record, header.get('橋梁名稱', ''))
            if prior is not None:
                self.log(f'前期：{record.date}（{record.inspectors}）沿用之前建檔的掃描，不重新掃描')
            else:
                self.log('由歷史紀錄找出前期並掃描')
                prior = scan_record(record, self.fetcher, **options).to_dict()
                self.save_json(prior['header'], '前期', prior)
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
        failed = [r for r in current.rows if not r.ok] + [r for r in prior.get('rows', []) if not r.get('ok', True)]
        for row in failed:
            row = row if isinstance(row, dict) else row.__dict__
            self.log(f"核對異常：{row.get('component')} {row.get('label')}：{row.get('check')}")
        self.log(f'掃描完成（{(datetime.now() - started).total_seconds():.0f} 秒），核對異常 {len(failed)} 列')
        self.load_pairs(prior, current.to_dict())

    # ------------------------------------------------------------ 沒有本年度紀錄：先建檔前期
    def prebuilt_path(self, date, bridge):
        name = f'{date}-{bridge}-前期.json'
        return self.scan_dir / ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in name)

    def prebuilt_prior(self, record, bridge):
        """之前「沒有本年度紀錄」時先建檔的前期掃描（同橋、同日期）；沒有回傳 None。"""
        path = self.prebuilt_path(record.date, bridge)
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            return None
        return data if data.get('prebuilt') else None

    def this_year(self):
        return str(datetime.now().year)

    def prebuild_prior(self, bridge, history):
        """最新一期不是本年度：把最新一期當作前期先掃描建檔，再重新讀取歷史紀錄看有沒有本年度紀錄。
        回傳本年度紀錄（有的話），沒有回傳 None。"""
        latest = history[0]
        year = self.this_year()
        self.log(f'找不到 {year} 年的檢測紀錄：先以最新一期 {latest.date}（{latest.unit}）作為前期掃描建檔')
        self.summary.setText(f'{bridge.name}：網站上還沒有 {year} 年的定期檢測紀錄。先把最新一期 {latest.date} 當作前期掃描建檔，'
                             '之後建立本年度紀錄時不必再掃描前期。')
        options = dict(progress=self.progress, stopped=lambda: self.stop_requested)
        if self.prebuilt_prior(latest, bridge.name) is not None:
            self.log(f'前期 {latest.date} 已建檔過，不重新掃描')
        else:
            prior = scan_record(latest, self.fetcher, **options).to_dict()
            prior['prebuilt'] = True
            self.save_json(prior['header'], '前期', prior)
            self.log(f"前期建檔完成：{latest.date}，缺失 {sum(len(r.get('defects', [])) for r in prior.get('rows', []))} 筆")
        self.log('重新讀取歷史定期檢測紀錄，確認是否已有本年度紀錄')
        history = parse_history(self.fetcher.get(history_url(bridge.inventory_id)))
        return next((r for r in history if r.date.startswith(year)), None)

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
        self.current_header, self.prior_header = dict(head), dict(prior_scan.get('header', {}))
        self.opinion_items = []
        self.inventory = current_keys.get('inventory_id', '')
        self.load_photo_plan(head)
        self.review_file = review_path(self.review_dir, head.get('橋梁名稱', ''), head.get('檢測日期', ''))
        self.review = Review.load(self.review_file) if self.review_file.exists() else Review(
            head.get('橋梁名稱', ''), head.get('檢測日期', ''), prior_scan.get('header', {}).get('檢測日期', ''))
        self.apply_review()
        self.log('配對：' + '、'.join(f'{k} {v}' for k, v in result.summary.items()))
        self.fill_items = []
        self.build_fill_plan()                 # 填寫清單同時產生，總表與填寫分頁的狀態一致
        self.overall_filter = TODO if any(self.overall(p)[0] == TODO for p in self.pairs) else '全部'
        self.index = -1
        self.refresh_table()
        self.update_summary()
        self.buttons['export'].setEnabled(True)
        self.tabs.setCurrentWidget(self.check_page)
        if self.visible:
            self.open_card(0)                      # 直接從第一筆待處理開始

    # ------------------------------------------------------------ ② 橋梁照片
    def photo_plan_path(self):
        name = f"{self.current_header.get('檢測日期', '')}-{self.current_header.get('橋梁名稱', '')}-照片.json"
        return self.review_dir / ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in name)

    def load_photo_plan(self, head):
        """依基本資料掃描檔的「是否為跨水橋」與上次存的安排，建立六格照片。"""
        name = f"{head.get('檢測日期', '')}-{head.get('橋梁名稱', '')}-基本資料.json"
        river = ''
        try:
            basic = json.loads((self.scan_dir / ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in name))
                               .read_text(encoding='utf-8'))
            river = basic.get('values', {}).get('river_cross', '')
        except (OSError, ValueError):
            pass
        plan = None
        try:
            plan = PhotoPlan.from_dict(json.loads(self.photo_plan_path().read_text(encoding='utf-8')))
        except (OSError, ValueError, TypeError):
            pass
        self.photo_panel.set_context(river, head.get('檢測日期', ''), plan)

    def save_photo_plan(self):
        if not self.current_header:
            return
        try:
            path = self.photo_plan_path()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(self.photo_panel.plan.to_dict(), ensure_ascii=False, indent=2), encoding='utf-8')
        except OSError as error:
            self.log(f'照片安排存檔失敗：{error}')

    def upload_photos(self, ask=True):
        """把放好的照片逐張上傳到網站照片簿；遇到失敗或結果不明立即停止（不重傳）。"""
        if self.busy:
            return False
        if not self.inventory:
            QMessageBox.information(self, '上傳照片', '沒有橋梁系統ID，請先掃描本期或開啟掃描檔。')
            return False
        todo = self.photo_panel.selected_slots()[:6]          # 勾選的照片（多選上傳），一次最多六張
        if not todo:
            if ask:
                QMessageBox.information(self, '上傳照片', '請先勾選要上傳的照片（每格左上角「選取」，或按「全選」）。')
            return False
        lines = '\n'.join(f'・{s.title}：{Path(s.path).name}（{s.date}）' for s in todo)
        if ask and QMessageBox.question(self, '上傳照片', f'將上傳 {len(todo)} 張照片到網站照片簿：\n{lines}\n\n'
                                        '照片簿已有同標題的照片會略過；程式不刪除照片。確定上傳？') != QMessageBox.StandardButton.Yes:
            return False
        if self.filler is None:
            self.filler = SiteFiller(self.profile, self, delay_ms=self.fetcher.delay_ms, rewrite=self.fetcher.rewrite)
            self.watch_loading(self.filler.page, '填寫')
        if self.fill_log is None:
            self.fill_log = FillLog(fill_log_path(self.review_dir, self.review.bridge, self.review.current_date))
        self.set_busy(True)
        url = photo_list_url(self.inventory)
        self.filler.start_photo_batch()
        try:
            for slot in todo:
                try:
                    message = self.filler.upload_photo(url, slot, self.fill_log)
                except FillError as error:
                    self.log(f'上傳照片「{slot.title}」失敗：{error}')
                    QMessageBox.warning(self, '上傳照片', f'「{slot.title}」：{error}\n\n已停止，其餘照片未上傳。')
                    return False
                finally:
                    self.photo_panel.refresh()
                    self.save_photo_plan()
                self.log(f'上傳照片「{slot.title}」：{message}')
        finally:
            self.set_busy(False)
        if self.view.isVisibleTo(self):
            self.view.open(url)
        return True

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

    # ------------------------------------------------------------ ② 填寫：檢測基本資料（與意見一起送出）
    def build_basic_box(self):
        group = QFrame()
        group.setFrameShape(QFrame.Shape.StyledPanel)
        grid = QGridLayout(group)
        grid.setContentsMargins(6, 4, 6, 4)
        heading = QLabel('檢測基本資料')
        heading.setStyleSheet('font-weight:bold;')
        grid.addWidget(heading, 0, 0)
        self.basic_inputs, self.basic_sources = {}, {}
        for number, name in enumerate(BASIC_NAMES):
            if name in BASIC_OPTIONS:
                widget = QComboBox()
                widget.currentTextChanged.connect(lambda _=None: self.basic_changed())
            else:
                widget = QLineEdit()
                widget.textChanged.connect(lambda _=None: self.basic_changed())
            widget.setMinimumWidth(50)                 # 不撐寬工作區
            if isinstance(widget, QComboBox):
                widget.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
                widget.setMinimumContentsLength(3)
            source = QLabel('')
            source.setStyleSheet('color:#888; font-size:11px;')
            row, col = number // 3, number % 3 * 3 + 1
            grid.addWidget(QLabel(name.replace('是否疑似鹽害橋梁', '疑似鹽害').replace('橋下有無堆放雜物', '橋下雜物')), row, col)
            grid.addWidget(widget, row, col + 1)
            widget.setToolTip(name)
            self.basic_inputs[name], self.basic_sources[name] = widget, source
        for col in (2, 5, 8):
            grid.setColumnStretch(col, 1)
        self.basic_note = QLabel('')
        self.basic_note.setStyleSheet('color:#666;')
        self.basic_note.setWordWrap(True)
        grid.addWidget(self.basic_note, 2, 1, 1, 8)
        return group

    def load_basic(self, saved=None):
        settings = supervisors(self.supervisor_path)
        values, sources = basic_defaults(self.current_header, self.prior_header, settings)
        values.update(saved or {})
        choices = basic_options(self.current_header)
        self.opinion_loading = True
        for name, widget in self.basic_inputs.items():
            if isinstance(widget, QComboBox):
                widget.clear()
                widget.addItems([''] + choices[name])
                widget.setCurrentText(values.get(name, ''))
            else:
                widget.setText(values.get(name, ''))
        self.opinion_loading = False
        self.basic_from = sources
        self.basic_changed(save=False)

    def basic_values(self):
        return {name: (w.currentText() if isinstance(w, QComboBox) else w.text()).strip()
                for name, w in self.basic_inputs.items()}

    def basic_changed(self, save=True):
        if self.opinion_loading:
            return
        values = self.basic_values()
        none = values.get('有無沖刷現象') == '無'
        self.basic_inputs['沖刷位置/備註'].setEnabled(not none)
        if none and values.get('沖刷位置/備註'):
            self.basic_inputs['沖刷位置/備註'].setText('')
            return
        diff = basic_changes(self.current_header, values)
        sources = getattr(self, 'basic_from', {})
        notes = [f'{k}：{sources[k]}' for k in BASIC_NAMES if sources.get(k) in ('沿用前期',) or
                 str(sources.get(k, '')).startswith('依管理機關')]
        missing = [k for k in BASIC_NAMES if not values.get(k) and not (k == '沖刷位置/備註' and none)]
        self.basic_note.setText(('將一併寫入：' + '、'.join(diff) if diff else '基本資料與網站相同') +
                                ('；' + '、'.join(notes) if notes else '') +
                                ('；尚未填寫：' + '、'.join(missing) if missing else ''))
        if save:
            self.save_opinion()

    def edit_supervisors(self):
        """單位主管設定：主管姓名 → 負責機關（依管理機關自動帶入單位主管）。"""
        dialog = QDialog(self)
        dialog.setWindowTitle('單位主管設定')
        dialog.resize(680, 420)
        table = QTableWidget(0, 2)
        table.setHorizontalHeaderLabels(['主管姓名', '負責機關（以頓號「、」分隔）'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        for name, agencies in supervisors(self.supervisor_path).items():
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(name))
            table.setItem(row, 1, QTableWidgetItem('、'.join(agencies)))
        add, remove = QPushButton('新增一列'), QPushButton('刪除選取列')
        add.clicked.connect(lambda: table.insertRow(table.rowCount()))
        remove.clicked.connect(lambda: [table.removeRow(r) for r in sorted({i.row() for i in table.selectedIndexes()},
                                                                          reverse=True)])
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText('儲存')
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('取消')
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        row = QHBoxLayout()
        row.addWidget(add)
        row.addWidget(remove)
        row.addStretch()
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel('依管理機關自動帶入「單位主管」。一位主管可負責多個機關，以頓號分隔（例：南投縣政府、苗栗縣政府）。'))
        layout.addWidget(table)
        layout.addLayout(row)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        return self.save_supervisors(table)

    def save_supervisors(self, table):
        values = {}
        for r in range(table.rowCount()):
            name = (table.item(r, 0).text() if table.item(r, 0) else '').strip()
            agencies = [a.strip() for a in re.split(r'[、,，]', table.item(r, 1).text() if table.item(r, 1) else '')
                        if a.strip()]
            if name or agencies:
                values[name] = agencies
        try:
            saved = supervisors(self.supervisor_path, values)
        except ValueError as error:
            QMessageBox.warning(self, '單位主管設定', f'設定格式錯誤：{error}')
            return None
        self.log('已儲存單位主管設定：' + '；'.join(f"{k}（{'、'.join(v)}）" for k, v in saved.items()))
        if self.opinion_items:
            self.load_basic()
        return saved

    # ------------------------------------------------------------ ② 填寫：檢測員意見
    def build_opinion(self, force=False):
        """產生意見草稿；已有草稿（含勾選、選項與手動修改）時接續，除非 force（重新產生）。"""
        if not self.pairs or self.busy:
            return
        if self.opinion_items and not force:
            return
        saved = {} if force else self.review.opinion
        girder = any(p.current is not None and p.current.component == '主梁' for p in self.pairs)
        if saved.get('items'):
            self.opinion_items = [OpinionItem(**item) for item in saved['items']]
            self.height_line = saved.get('height', '')
            self.opinion_values = dict(saved.get('values', {}))
        else:
            self.opinion_items = generate(self.pairs, self.review, self.current_header, self.prior_header)
            self.height_line = heights(self.current_header, self.prior_header, girder)
            self.opinion_values = {}
        for key in ('opinion_build', 'opinion_write'):
            self.buttons[key].setEnabled(True)
        self.site_opinion.setPlainText(self.current_header.get('檢測員意見', ''))
        self.prior_opinion.setPlainText(self.prior_header.get('檢測員意見', ''))
        self.fill_opinion_list()
        self.load_basic(None if force else saved.get('basic'))
        self.recompose()

    def fill_opinion_list(self):
        self.opinion_loading = True
        self.opinion_list.clear()
        for item in self.opinion_items:
            shown = item.custom or apply_choices(item.text, self.opinion_values)
            if item.kind == 'standard':
                name, _, state = item.note.partition('：')
                tag = '已帶入' if item.sources else '未偵測'
                label = f'【標準】{name}（{tag}）' + (f'\n{shown}' if item.checked else '')
            else:
                label = shown + ('（小面積，預設不寫）' if item.note.startswith('小面積') else '')
            row = QListWidgetItem(label)
            row.setFlags(row.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            row.setCheckState(Qt.CheckState.Checked if item.checked else Qt.CheckState.Unchecked)
            row.setToolTip(item.note or shown)
            if item.kind == 'standard':
                row.setBackground(QColor('#eef4fb'))
            elif item.kind == 'manual':
                row.setBackground(QColor('#fff2cc'))
            self.opinion_list.addItem(row)
        self.opinion_loading = False

    def fill_choice_panel(self):
        """勾選中的條目裡的【A／B】給下拉、【＿】給輸入框；選定或填寫後括號消失。"""
        while self.choice_form.count():
            self.choice_form.takeAt(0).widget().deleteLater()
        self.opinion_inputs = {}
        text = '\n'.join(i.text for i in self.opinion_items if i.checked and not i.custom) + '\n' + self.height_line
        row = 0
        for placeholder, options in choices(text).items():
            combo = QComboBox()
            combo.addItems(['（未選）'] + options)
            if self.opinion_values.get(placeholder) in options:
                combo.setCurrentText(self.opinion_values[placeholder])
            combo.currentTextChanged.connect(lambda value, p=placeholder: self.set_opinion_value(p, value))
            self.add_choice_row(row, placeholder, combo)
            row += 1
        for placeholder in blanks(text):
            line = QLineEdit(self.opinion_values.get(placeholder, ''))
            line.setPlaceholderText('請填寫')
            line.textChanged.connect(lambda value, p=placeholder: self.set_opinion_value(p, value))
            self.add_choice_row(row, placeholder, line)
            row += 1
        if not row:
            self.choice_form.addWidget(QLabel('（沒有待選或待填項目）'), 0, 0, 1, 2)
            row = 1
        spacer = QWidget()
        self.choice_form.addWidget(spacer, row, 0, 1, 2)
        for number in range(row + 1):
            self.choice_form.setRowStretch(number, 1 if number == row else 0)

    def add_choice_row(self, row, placeholder, widget):
        label = QLabel(placeholder)
        label.setWordWrap(True)
        label.setMaximumWidth(260)
        self.choice_form.addWidget(label, row, 0)
        self.choice_form.addWidget(widget, row, 1)
        self.opinion_inputs[placeholder] = widget

    def set_opinion_value(self, placeholder, value):
        value = value.strip()
        if value and value != '（未選）':
            self.opinion_values[placeholder] = value
        else:
            self.opinion_values.pop(placeholder, None)
        self.recompose(panel=False)

    def recompose(self, panel=True):
        """依勾選、選項與手動修改重組右側意見（勾選增減即時同步）。"""
        if panel:
            self.fill_choice_panel()
        self.opinion_loading = True
        self.opinion_edit.setPlainText(compose(self.opinion_items, apply_choices(self.height_line, self.opinion_values),
                                               self.opinion_values))
        self.opinion_loading = False
        self.opinion_state()
        self.save_opinion()

    def opinion_state(self):
        left = unresolved(self.opinion_edit.toPlainText())
        self.opinion_status.setText(f'尚有 {len(left)} 處【…】未選定或未填寫。' if left else '可以寫入網站。')

    def opinion_changed(self):
        if self.opinion_loading:
            return
        for item, index in zip(self.opinion_items, range(self.opinion_list.count())):
            item.checked = self.opinion_list.item(index).checkState() == Qt.CheckState.Checked
        self.recompose()

    def opinion_typed(self):
        """直接修改右側意見：每一條對回條目（保留修改），多出的行成為手動條目。"""
        if self.opinion_loading or not self.opinion_items:
            return
        count = len(self.opinion_items)
        checked = sum(i.checked for i in self.opinion_items)
        text = self.opinion_edit.toPlainText()
        self.height_line = absorb(self.opinion_items, text, self.opinion_values)
        self.fill_opinion_list()
        if len(self.opinion_items) != count or sum(i.checked for i in self.opinion_items) != checked:
            self.fill_choice_panel()
        if sum(i.checked for i in self.opinion_items) < checked or not numbering_ok(text):
            QTimer.singleShot(0, self.renumber_opinion)      # 刪掉條目：取消勾選並重新編號（雙向同步）
        self.opinion_state()
        self.save_opinion()

    def renumber_opinion(self):
        """意見框重新編號（保留每一條的修改），游標停在原本那一行。"""
        cursor = self.opinion_edit.textCursor()
        line = cursor.blockNumber()
        self.recompose(panel=False)
        block = self.opinion_edit.document().findBlockByNumber(min(line, self.opinion_edit.document().blockCount() - 1))
        cursor = QTextCursor(block)
        cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock)
        self.opinion_edit.setTextCursor(cursor)

    def insert_advice(self, advice):
        """游標所在的編號條目改成指定建議（沒有建議則補上）；不在條目上時插入在游標處。"""
        cursor = self.opinion_edit.textCursor()
        block = cursor.block()
        line = block.text()
        if re.match(r'^\s*\d+\s*[.．、]', line):
            body = re.sub(r'[，,]?\s*建議[^，,。]*。?$', '', line.rstrip()).rstrip('。')
            cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
            cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)
            cursor.insertText(f'{body}，{advice}。')
        else:
            cursor.insertText(advice)
        self.opinion_edit.setTextCursor(cursor)

    def save_opinion(self):
        if self.review_file is None or not self.opinion_items:
            return
        self.review.opinion = {'items': [asdict(i) for i in self.opinion_items], 'height': self.height_line,
                               'values': self.opinion_values, 'text': self.opinion_edit.toPlainText(),
                               'basic': self.basic_values()}
        self.review.save(self.review_file)

    def confirm_opinion(self, text):
        diff = basic_changes(self.current_header, self.basic_values())
        extra = ('\n\n基本資料一併寫入：' + '、'.join(f'{k}＝{v or "（空白）"}' for k, v in diff.items())) if diff else ''
        return QMessageBox.question(self, '寫入檢測員意見', f'將把下列意見寫入網站並送出：\n\n{text}{extra}\n\n確定？') \
            == QMessageBox.StandardButton.Yes

    def write_opinion(self):
        text = self.opinion_edit.toPlainText().strip()
        if self.busy or not text:
            return False
        left = unresolved(text)
        if left:
            QMessageBox.information(self, '寫入檢測員意見', '還有未選定或未填寫的項目：\n' + '\n'.join(f'【{x}】' for x in left))
            return False
        if not self.confirm_opinion(text):
            return False
        if self.filler is None:
            self.filler = SiteFiller(self.profile, self, delay_ms=self.fetcher.delay_ms, rewrite=self.fetcher.rewrite)
            self.watch_loading(self.filler.page, '填寫')
        if self.fill_log is None:
            self.fill_log = FillLog(fill_log_path(self.review_dir, self.review.bridge, self.review.current_date))
        self.set_busy(True)
        try:
            basic = self.basic_values()
            if basic.get('有無沖刷現象') == '無':
                basic.pop('沖刷位置/備註', None)           # 無沖刷時網站會停用此欄
            message = self.filler.write_opinion(self.current_main, self.current_header.get('檢測員意見', ''), text,
                                                self.fill_log, basic,
                                                {k: self.current_header.get(k, '') for k in basic})
        except FillError as error:
            self.log(f'寫入檢測員意見失敗：{error}')
            QMessageBox.warning(self, '寫入檢測員意見', str(error))
            return False
        finally:
            self.set_busy(False)
        self.current_header['檢測員意見'] = text
        self.current_header.update(basic)
        self.basic_changed(save=False)
        self.site_opinion.setPlainText(text)
        self.opinion_status.setText('已寫入網站：' + message)
        self.log(f'寫入檢測員意見：{message}')
        if self.view.isVisibleTo(self) and self.current_main:
            self.view.open(self.current_main)
        return True

    # ------------------------------------------------------------ ② 填寫
    def build_fill_plan(self):
        if not self.pairs or self.busy:
            return
        self.fill_log = FillLog(fill_log_path(self.review_dir, self.review.bridge, self.review.current_date))
        self.fill_items = build_plan(self.pairs, self.review, self.rules, self.fill_log)
        if self.add_missing_rules():
            self.fill_items = build_plan(self.pairs, self.review, self.rules, self.fill_log)
        self.refresh_fill_table()
        self.refresh_table()
        self.update_summary()

    def fill_item_for(self, pair):
        """這一筆在填寫清單中的狀態（清單尚未產生時即時計算）。"""
        if pair is None or pair.current is None:
            return None
        found = next((i for i in self.fill_items if i.key == pair.key), None)
        return found or plan_item(pair, self.review, self.rules)

    def refresh_fill_status(self):
        if self.card.pair is not None:
            self.card.set_fill_status(self.fill_item_for(self.card.pair))

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
        """② 送出：只列要送出或已送出的項目（可送出、已完成、待確認、失敗）。"""
        listed = [i for i in self.fill_items if i.status in (READY, DONE, PENDING, FAILED)]
        chosen = self.fill_component.currentData() if self.fill_component.count() else None
        names = list(dict.fromkeys(i.pair.current.component for i in listed))
        self.fill_component.blockSignals(True)
        self.fill_component.clear()
        self.fill_component.addItem(f'全部構件（可送出 {sum(i.status == READY for i in listed)}）', None)
        for name in names:
            ready = sum(i.status == READY and i.pair.current.component == name for i in listed)
            self.fill_component.addItem(f'{name}（可送出 {ready}）', name)
        index = self.fill_component.findData(chosen) if chosen else 0
        self.fill_component.setCurrentIndex(max(index, 0))
        self.fill_component.blockSignals(False)
        chosen = self.fill_component.currentData()
        self.fill_visible = [i for i in listed if chosen is None or i.pair.current.component == chosen]
        self.fill_table.setRowCount(len(self.fill_visible))
        for row, item in enumerate(self.fill_visible):
            current = item.pair.current
            values = (item.status, component_label(item.pair), current.code or '—', item.method_id, item.describe() or '—',
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
        ready = counts.get(READY, 0)
        self.buttons['send_all'].setEnabled(not self.busy and ready > 0)
        self.buttons['send_all'].setText(f'整批送出（{ready}）')
        for key in ('fill_plan', 'fill_rescan'):
            self.buttons[key].setEnabled(not self.busy and bool(self.pairs))
        self.refresh_send_buttons()

    def sendable(self, items):
        return [i for i in items if i.status in (READY, PENDING, FAILED) and i.changes]

    def selected_fill_items(self):
        rows = sorted({index.row() for index in self.fill_table.selectionModel().selectedRows()}) \
            if self.fill_table.selectionModel() else []
        return [self.fill_visible[r] for r in rows if r < len(self.fill_visible)]

    def refresh_send_buttons(self):
        """② 送出：選取幾筆、這個構件可送出幾筆，顯示在按鈕上。"""
        picked = self.sendable(self.selected_fill_items())
        self.buttons['send_one'].setText(f'送出選取的項目（{len(picked)}）')
        self.buttons['send_one'].setEnabled(not self.busy and bool(picked))
        chosen = self.fill_component.currentData()
        ready = [i for i in self.fill_visible if i.status == READY]
        self.buttons['send_component'].setText(f'送出此構件（{len(ready)}）' if chosen else '送出此構件')
        self.buttons['send_component'].setEnabled(not self.busy and chosen is not None and bool(ready))

    def open_fill_item(self, row, _col=0):
        """② 送出清單雙擊：回到 ① 那一筆修改。"""
        if 0 <= row < len(self.fill_visible):
            self.index = self.pairs.index(self.fill_visible[row].pair)
            self.show_card()

    def refresh_send_button(self):
        """① 不再有送出按鈕（送出都在 ② 送出資料）；保留給舊呼叫點。"""

    def send_current(self):
        """① 送出這一筆（下方工作區顯示的那一筆）。"""
        self.save_pending_edit()
        item = self.fill_item_for(self.pairs[self.index]) if 0 <= self.index < len(self.pairs) else None
        if item is None or item.status not in (READY, PENDING, FAILED) or not item.changes:
            QMessageBox.information(self, '送出這一筆', f'這筆是「{item.status if item else "只有前期"}」，沒有要送出的變更。')
            return
        self.run_fill([item])

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
        if self.card.pair is pair:
            self.card.edit_status.setText('已確認維持網站現值' if on else '已取消確認')
        self.build_fill_plan()
        self.refresh_fill_status()

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
        """送出表格中選取的項目（可多選）；已完成等沒有變更的略過。"""
        chosen = self.selected_fill_items()
        if not chosen:
            QMessageBox.information(self, '送出', '請先在表格選取要送出的項目（可多選：Ctrl／Shift＋點選）。')
            return
        items = self.sendable(chosen)
        if not items:
            QMessageBox.information(self, '送出', '選取的項目都是「已完成」或沒有要填寫的變更。')
            return
        if len(items) < len(chosen):
            self.log(f'送出選取的項目：{len(chosen) - len(items)} 筆已完成或沒有變更，略過')
        self.run_fill(items)

    def fill_component_items(self):
        """依構件送出：目前選的構件全部「可送出」的項目。"""
        chosen = self.fill_component.currentData()
        if chosen is None:
            QMessageBox.information(self, '送出', '請先在「構件」選一個構件。')
            return
        self.run_fill([i for i in self.fill_items if i.status == READY and i.pair.current.component == chosen])

    def fill_all(self):
        self.run_fill([i for i in self.fill_items if i.status == READY])

    def run_fill(self, items):
        if self.busy or not items or not self.confirm_fill(items):
            return
        if self.filler is None:
            self.filler = SiteFiller(self.profile, self, delay_ms=self.fetcher.delay_ms, rewrite=self.fetcher.rewrite)
            self.watch_loading(self.filler.page, '填寫')
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

    def overall(self, pair):
        """合併狀態（審查＋填寫）與說明。"""
        decision = self.review.decision(pair)
        fill = self.fill_item_for(pair) if pair.current is not None else None
        if decision.status == PENDING:
            return TODO, '待審查：' + ('；'.join(pair.reasons) or pair.kind)
        if not decision.at:                       # 程式自動判定（高信心）也要按「確認並下一筆」才算確認
            return TODO, f'待確認：程式判定「{decision.kind}」' + (f'；{fill.reason}' if fill else '')
        if fill is None:
            return NOTHING, '只有前期，網站上沒有這筆'
        if fill.status == REVIEW:
            return TODO, '填寫需確認：' + fill.reason
        if fill.status == READY:
            return SENDABLE, fill.reason
        if fill.status == DONE:
            return SENT, fill.message
        if fill.status in (PENDING, FAILED):
            return SEND_ISSUE, f'{fill.status}：{fill.message}'
        return NOTHING, fill.reason

    def set_overall_filter(self, key):
        self.overall_filter = key
        self.refresh_table()

    def update_summary(self):
        if not self.pairs:
            return
        counts = {}
        for pair in self.pairs:
            status = self.overall(pair)[0]
            counts[status] = counts.get(status, 0) + 1
        self.summary.setText(f'{self.review.bridge}　本期 {self.review.current_date}／前期 {self.review.prior_date}　'
                             f'共 {len(self.pairs)} 組：' + '、'.join(f'{k} {counts.get(k, 0)}' for k in OVERALL_CARDS[:-1])
                             + f'；已修改 {sum(1 for p in self.pairs if self.review.edit_of(p))}。審查與修改自動存檔。')

    def matches_filter(self, pair):
        chosen = self.filter.currentText()
        decision = self.review.decision(pair)
        if chosen == '全部':
            return True
        if chosen in (PENDING, CHANGED):
            return decision.status == chosen
        if chosen == '需確認':
            return pair.confidence == '需確認'
        if chosen == '填寫需確認':
            fill = self.fill_item_for(pair) if pair.current is not None else None
            return bool(fill and fill.status == REVIEW)
        if chosen == '已修改':
            return bool(self.review.edit_of(pair))
        if chosen == '人工配對':
            return pair.manual
        if chosen == '位置變更':
            return pair.moved or any('位置變更' in r for r in pair.reasons)
        return decision.kind == chosen

    def refresh_table(self, *_):
        """① 缺失清單：依數字卡片（合併狀態）與進階篩選列出；保持選取在目前那一筆。"""
        statuses = {id(p): self.overall(p) for p in self.pairs}
        counts = {}
        for status, _ in statuses.values():
            counts[status] = counts.get(status, 0) + 1
        for key, button in self.status_buttons.items():
            number = len(self.pairs) if key == '全部' else counts.get(key, 0)
            button.setText(f'{key}\n{number}')
            button.setChecked(key == self.overall_filter)
            color = OVERALL_COLORS.get(key, '#f3f3f3')
            button.setStyleSheet(f'QPushButton{{background:{color}; border:1px solid #c8c8c8; border-radius:4px;'
                                 f' font-weight:bold;}} QPushButton:checked{{border:2px solid #0078d7;}}')
        self.visible = [p for p in self.pairs if self.matches_filter(p) and
                        (self.overall_filter == '全部' or statuses[id(p)][0] == self.overall_filter)]
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.visible))
        for row, pair in enumerate(self.visible):
            decision = self.review.decision(pair)
            any_side = pair.current or pair.prior
            kind = decision.kind if decision.kind == pair.kind else f'{decision.kind}（原：{pair.kind}）'
            edits = self.review.edit_of(pair)
            remark = self.review.value(pair, 'remark') if pair.current else '—'
            status, detail = statuses[id(pair)]
            values = (component_label(pair) + (' ✎' if edits else ''), pair.code or '—',
                      remark_lines(pair.prior.remark) if pair.prior else '—',
                      ('✎ ' if 'remark' in edits else '') + remark_lines(remark), kind, pair.confidence, detail)
            color = OVERALL_COLORS.get(status)
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if color:
                    item.setBackground(QColor(color))
                item.setToolTip(f'{status}：{detail}' if col == 0 else f'{status}\n{value}')
                self.table.setItem(row, col, item)
        current = self.pairs[self.index] if 0 <= self.index < len(self.pairs) else None
        if current in self.visible:
            self.table.setCurrentCell(self.visible.index(current), 0)
        self.table.blockSignals(False)
        fit_columns(self.table, TABLE_WIDTHS)
        self.refresh_send_button()

    def select_row(self, row):
        """點選清單的一筆：下方工作區顯示（換筆前自動儲存未儲存的修改）。"""
        if 0 <= row < len(self.visible):
            pair = self.visible[row]
            if self.card.pair is not pair:
                self.save_pending_edit()
            self.index = self.pairs.index(pair)
            self.show_card(select=False)

    def open_card(self, row):
        if 0 <= row < len(self.visible):
            self.index = self.pairs.index(self.visible[row])
            self.show_card()

    def show_card(self, select=True):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        self.card.show_pair(pair, self.review.decision(pair), f'第 {self.index + 1}／{len(self.pairs)} 組',
                            self.review.edit_of(pair), self.planned_values(pair))
        self.card.set_fill_status(self.fill_item_for(pair))
        self.tabs.setCurrentWidget(self.check_page)
        if select and pair in self.visible:
            self.table.blockSignals(True)
            self.table.setCurrentCell(self.visible.index(pair), 0)
            self.table.blockSignals(False)
        self.refresh_send_button()

    def save_pending_edit(self):
        """修改後欄位有未儲存的變更：確認或換筆前自動儲存，不會默默丟掉。"""
        if self.card.dirty() and 0 <= self.index < len(self.pairs) and self.card.pair is self.pairs[self.index]:
            self.save_edit(self.card.values())
            self.log('已自動儲存修改後欄位')
            return True
        return False

    def planned_values(self, pair):
        """這一筆將填寫的值：可填寫者為填寫清單的預定值；需確認者為網站現值＋程式能判斷的部分（工作區直接顯示）。"""
        item = self.fill_item_for(pair) if pair.current is not None else None
        if item is None:
            return None
        if item.status == REVIEW:
            return {**item.site, **{k: v for k, v in (item.proposal or {}).items() if str(v or '').strip()}}
        return dict(item.target)

    def confirm_values(self, pair):
        """確認並下一筆：工作區本期欄的值就是確認的結果——與程式預定值相同不必另存；與網站相同即維持現值。"""
        if pair.current is None or self.card.pair is not pair:
            return
        values = self.card.values()
        item = self.fill_item_for(pair)
        site = {k: site_value(pair.current, k) for k in values}
        if values == site:
            if self.review.edit_of(pair):
                self.review.set_edit(pair, values)          # 清除修改
            self.review.keep(pair, True)
        elif item is None or item.status == REVIEW or values != {k: item.target.get(k, '') for k in values}:
            self.review.set_edit(pair, values)
            self.review.keep(pair, False)
        self.build_fill_plan()

    def decide(self, kind, note):
        if not 0 <= self.index < len(self.pairs):
            return
        pair = self.pairs[self.index]
        self.confirm_values(pair)
        decision = self.review.decide(pair, kind, note)
        self.review.save(self.review_file)
        self.log(f'審查：{(pair.current or pair.prior).component} {kind}（{decision.status}）')
        self.refresh_table()
        self.update_summary()
        self.move(1)

    def move(self, step):
        """移到上／下一筆。只在待處理之間移動時，到底會從頭找；全部處理完提示。"""
        if not self.pairs:
            return
        self.save_pending_edit()
        pending_only = self.card.pending_only.isChecked()
        wanted = lambda i: not pending_only or self.overall(self.pairs[i])[0] == TODO  # noqa: E731
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
            message = ('待處理全部完成，可以「整批送出」' if not any(self.overall(p)[0] == TODO for p in self.pairs)
                       else '已到最後一筆')
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
