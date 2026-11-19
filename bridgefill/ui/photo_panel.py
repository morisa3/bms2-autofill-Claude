"""② 橋梁照片：六格照片（跨水橋／陸橋），拖曳放入、格子間拖曳交換、自動產生可選拍照日期。"""
from pathlib import Path

from PySide6.QtCore import QMimeData, QPoint, Qt, Signal
from PySide6.QtGui import QDrag, QImageReader, QPixmap
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QVBoxLayout, QWidget)

from ..core.bridge_photos import PhotoPlan, bridge_kind, date_choices, exif_date, is_image, iso

SLOT_MIME = 'application/x-bridgefill-slot'
THUMB = (220, 150)


def thumbnail(path, size=THUMB):
    """縮圖（依 EXIF 方向轉正）；讀不到回傳 None。"""
    reader = QImageReader(str(path))
    reader.setAutoTransform(True)
    image = reader.read()
    if image.isNull():
        return None
    return QPixmap.fromImage(image).scaled(*size, Qt.AspectRatioMode.KeepAspectRatio,
                                           Qt.TransformationMode.SmoothTransformation)


def dropped_paths(mime):
    return [url.toLocalFile() for url in mime.urls() if url.isLocalFile() and is_image(url.toLocalFile())]


class PhotoBox(QLabel):
    """一格的照片區：接受檔案拖入；有照片時可拖到另一格交換。"""
    files = Signal(list)
    swap = Signal(int)
    clicked = Signal()

    def __init__(self, index, parent=None):
        super().__init__(parent)
        self.index = index
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(110, 90)
        self.setWordWrap(True)
        self.has_photo = False
        self.press = None
        self.highlight(False)

    def highlight(self, on):
        border = '2px solid #0078d7' if on else ('1px solid #c8c8c8' if self.has_photo else '2px dashed #9aa7b5')
        self.setStyleSheet(f'QLabel{{border:{border}; border-radius:4px; background:#fafafa; color:#666;}}')

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        if dropped_paths(mime) or (mime.hasFormat(SLOT_MIME) and int(bytes(mime.data(SLOT_MIME)).decode()) != self.index):
            event.acceptProposedAction()
            self.highlight(True)

    def dragLeaveEvent(self, event):
        self.highlight(False)

    def dropEvent(self, event):
        self.highlight(False)
        mime = event.mimeData()
        if mime.hasFormat(SLOT_MIME):
            self.swap.emit(int(bytes(mime.data(SLOT_MIME)).decode()))
        else:
            self.files.emit(dropped_paths(mime))
        event.acceptProposedAction()

    def mousePressEvent(self, event):
        self.press = event.position().toPoint()

    def mouseMoveEvent(self, event):
        if self.press is None or not self.has_photo:
            return
        if (event.position().toPoint() - self.press).manhattanLength() < 8:
            return
        mime = QMimeData()
        mime.setData(SLOT_MIME, str(self.index).encode())
        drag = QDrag(self)
        drag.setMimeData(mime)
        if self.pixmap() is not None and not self.pixmap().isNull():
            drag.setPixmap(self.pixmap().scaled(110, 75, Qt.AspectRatioMode.KeepAspectRatio))
            drag.setHotSpot(QPoint(55, 37))
        self.press = None
        drag.exec(Qt.DropAction.MoveAction)

    def mouseReleaseEvent(self, event):
        if self.press is not None:
            self.clicked.emit()
        self.press = None


class SlotWidget(QFrame):
    def __init__(self, index, panel):
        super().__init__()
        self.index, self.panel = index, panel
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.title = QLabel('')
        self.title.setStyleSheet('font-weight:bold;')
        self.box = PhotoBox(index)
        self.box.files.connect(lambda paths: panel.place(paths, index))
        self.box.swap.connect(lambda other: panel.swap(other, index))
        self.box.clicked.connect(self.browse)
        self.date = QComboBox(editable=True)
        self.date.setToolTip('拍照日期：自動列出照片拍攝日期（EXIF）、本期檢測日期與前後幾天、今天；也可直接輸入 YYYY-MM-DD')
        self.date.currentTextChanged.connect(self.changed)
        self.note = QLineEdit()
        self.note.setPlaceholderText('照片說明（選填）')
        self.note.textChanged.connect(self.changed)
        self.status = QLabel('')
        self.status.setWordWrap(True)
        clear = QPushButton('清除')
        clear.setMinimumWidth(40)
        clear.clicked.connect(lambda: panel.clear(index))
        self.pick = QCheckBox('選取')
        self.pick.setToolTip('勾選的照片會在按「上傳選取的照片」時上傳（已上傳或網站已有的不能勾選）')
        self.pick.toggled.connect(lambda on: panel.set_picked(index, on))
        head = QHBoxLayout()
        head.addWidget(self.pick)
        head.addWidget(self.title, 1)
        head.addWidget(clear)
        dates = QHBoxLayout()
        dates.addWidget(QLabel('日期'))
        dates.addWidget(self.date, 1)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(3)
        layout.addLayout(head)
        layout.addWidget(self.box, 1)
        layout.addLayout(dates)
        layout.addWidget(self.note)
        layout.addWidget(self.status)
        self.loading = False

    def browse(self):
        start = str(Path(self.panel.last_folder or Path.home()))
        paths, _ = QFileDialog.getOpenFileNames(self, f'選擇「{self.title.text()}」照片', start,
                                                '照片 (*.jpg *.jpeg *.png *.bmp *.gif)')
        if paths:
            self.panel.place(paths, self.index)

    def changed(self, *_):
        if self.loading:
            return
        slot = self.panel.plan.slots[self.index]
        slot.date, slot.note = self.date.currentText().strip(), self.note.text().strip()
        self.date.setStyleSheet('' if not slot.path or iso(slot.date) else 'background:#f8cbad;')
        self.panel.changed.emit()

    def show_slot(self, slot, choices):
        self.loading = True
        self.title.setText(slot.title)
        self.box.has_photo = bool(slot.path)
        picture = thumbnail(slot.path) if slot.path else None
        if picture is not None:
            self.box.setPixmap(picture)
            self.box.setToolTip(f'{slot.path}\n拖到其他格子可交換；點一下可換照片')
        else:
            self.box.clear()
            self.box.setText(('讀不到照片：' + Path(slot.path).name) if slot.path else '把照片拖到這裡\n（或點一下選擇）')
            self.box.setToolTip('可從檔案總管拖入；一次拖多張會依序放進空的格子')
        self.box.highlight(False)
        self.date.clear()
        own = [exif_date(slot.path)] if slot.path else []
        self.date.addItems(date_choices(self.panel.inspection_date, own + choices))
        self.date.setCurrentText(slot.date)
        self.date.setEnabled(bool(slot.path))
        self.note.setText(slot.note)
        self.status.setText(slot.status)
        done = slot.status.startswith(('已上傳', '網站已有'))
        self.status.setStyleSheet('color:#2e7d32;' if done else ('color:#c00000;' if slot.status else ''))
        usable = slot.ready and not done
        self.pick.blockSignals(True)
        self.pick.setEnabled(usable)
        self.pick.setChecked(usable and self.index in self.panel.picked)
        self.pick.blockSignals(False)
        self.setStyleSheet('SlotWidget{border:2px solid #0078d7;}' if self.pick.isChecked() else '')
        self.loading = False
        self.date.setStyleSheet('' if not slot.path or iso(slot.date) else 'background:#f8cbad;')


class PhotoPanel(QWidget):
    changed = Signal()
    upload = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.plan = PhotoPlan.create('')
        self.site_river_cross, self.inspection_date, self.last_folder, self.agency = '', '', '', ''
        self.picked = set()             # 勾選要上傳的格子（多選上傳）
        explain = QLabel('依基本資料「是否為跨水橋」產生六張照片的格子。把照片從檔案總管拖到格子（一次拖多張會依序放入空格），'
                         '格子之間拖曳可交換；拍照日期自動帶照片拍攝日期，沒有則用本期檢測日期。'
                         '上傳到網站「照片簿」，已有同標題的照片不會重複上傳，程式不刪除照片。')
        explain.setWordWrap(True)
        self.kind = QComboBox()
        self.kind.addItems(['跨水橋', '陸橋', '雲林縣'])
        self.kind.setToolTip('預設：管理機關為雲林縣政府 → 雲林縣版本；其他依基本資料「是否為跨水橋」：是 → 跨水橋，否 → 陸橋')
        self.kind.currentTextChanged.connect(self.set_kind)
        self.kind_note = QLabel('')
        self.all_date = QComboBox(editable=True)
        self.all_date.setMinimumContentsLength(10)
        self.all_date.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        apply = QPushButton('全部套用此日期')
        apply.setMinimumWidth(40)
        apply.clicked.connect(self.apply_date)
        self.upload_button = QPushButton('上傳選取的照片')
        self.upload_button.setMinimumWidth(40)
        self.upload_button.setToolTip('只上傳勾選「選取」的照片（一次最多六張），逐張：先核對照片簿 → 填入日期、標題、說明與照片 → 送出 → 重新讀照片簿核對')
        pick_all, pick_none = QPushButton('全選'), QPushButton('全不選')
        for button in (pick_all, pick_none):
            button.setMinimumWidth(40)
        pick_all.setToolTip('勾選所有可上傳的照片（已上傳、網站已有的除外）')
        pick_all.clicked.connect(lambda: self.pick_all(True))
        pick_none.clicked.connect(lambda: self.pick_all(False))
        self.upload_button.clicked.connect(self.upload.emit)
        self.kind_note.setWordWrap(True)
        top = QHBoxLayout()
        top.addWidget(QLabel('橋梁類型'))
        top.addWidget(self.kind)
        top.addWidget(self.kind_note, 1)
        actions = QHBoxLayout()
        actions.addWidget(QLabel('日期'))
        actions.addWidget(self.all_date)
        actions.addWidget(apply)
        actions.addStretch()
        actions.addWidget(pick_all)
        actions.addWidget(pick_none)
        actions.addWidget(self.upload_button)
        self.summary = QLabel('')
        self.summary.setWordWrap(True)
        grid = QGridLayout()
        grid.setSpacing(6)
        self.slots = []
        for index in range(6):
            widget = SlotWidget(index, self)
            grid.addWidget(widget, index // 3, index % 3)
            self.slots.append(widget)
        layout = QVBoxLayout(self)
        layout.addWidget(explain)
        layout.addLayout(top)
        layout.addLayout(actions)
        layout.addLayout(grid, 1)
        layout.addWidget(self.summary)
        self.setAcceptDrops(True)
        self.refresh()

    # 拖到格子以外的地方：依序放進空格
    def dragEnterEvent(self, event):
        if dropped_paths(event.mimeData()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.place(dropped_paths(event.mimeData()), 0, single_here=False)
        event.acceptProposedAction()

    def set_context(self, river_cross='', inspection_date='', plan=None, agency=''):
        """掃描或開啟掃描檔後：依管理機關（雲林縣政府）與基本資料決定版本；plan 為上次存的安排。"""
        self.site_river_cross, self.inspection_date, self.agency = river_cross, inspection_date, agency or ''
        if plan is None:
            plan = PhotoPlan.create(river_cross, agency=agency)
        elif not plan.kind:
            plan.kind = bridge_kind(plan.river_cross or river_cross, agency)
        self.plan = plan
        self.picked = {i for i in range(len(self.plan.slots)) if self.usable(i)}
        self.kind.blockSignals(True)
        self.kind.setCurrentText(self.plan.kind)
        self.kind.blockSignals(False)
        self.refresh()

    def set_kind(self, text):
        self.plan = PhotoPlan.create(self.plan.river_cross or self.site_river_cross, self.plan, kind=text)
        self.refresh()
        self.changed.emit()

    def place(self, paths, index, single_here=True):
        paths = [p for p in paths if is_image(p)]
        if not paths:
            return
        self.last_folder = str(Path(paths[0]).parent)
        if not single_here and len(paths) == 1:
            index = next((i for i, s in enumerate(self.plan.slots) if not s.path), index)
        before = {i for i, s in enumerate(self.plan.slots) if s.path}
        rest = self.plan.place(paths, self.inspection_date, index)
        self.picked |= {i for i, s in enumerate(self.plan.slots) if s.path and (i not in before or i == index)}   # 新放入的預設勾選
        self.refresh(f'有 {len(rest)} 張沒有空格可放：' + '、'.join(Path(p).name for p in rest) if rest else '')
        self.changed.emit()

    def swap(self, a, b):
        self.plan.swap(a, b)
        if (a in self.picked) != (b in self.picked):
            self.picked ^= {a, b}
        self.refresh()
        self.changed.emit()

    def clear(self, index):
        slot = self.plan.slots[index]
        slot.path, slot.date, slot.note, slot.status = '', '', '', ''
        self.refresh()
        self.changed.emit()

    def usable(self, index):
        slot = self.plan.slots[index]
        return slot.ready and not slot.status.startswith(('已上傳', '網站已有'))

    def set_picked(self, index, on):
        (self.picked.add if on else self.picked.discard)(index)
        self.refresh()

    def pick_all(self, on):
        self.picked = {i for i in range(len(self.plan.slots)) if self.usable(i)} if on else set()
        self.refresh()

    def selected_slots(self):
        """勾選且可上傳的照片（依格子順序）。"""
        return [self.plan.slots[i] for i in sorted(self.picked) if i < len(self.plan.slots) and self.usable(i)]

    def apply_date(self):
        value = iso(self.all_date.currentText())
        if not value:
            return
        for slot in self.plan.slots:
            if slot.path:
                slot.date = value
        self.refresh()
        self.changed.emit()

    def refresh(self, message=''):
        photo_dates = [exif_date(s.path) for s in self.plan.slots if s.path]
        for widget, slot in zip(self.slots, self.plan.slots):
            widget.show_slot(slot, photo_dates)
        current = self.all_date.currentText()
        self.all_date.clear()
        self.all_date.addItems(date_choices(self.inspection_date, photo_dates))
        if current:
            self.all_date.setCurrentText(current)
        site = self.site_river_cross
        notes = []
        if '雲林' in self.agency:
            notes.append(f'管理機關：{self.agency}（雲林縣版本）')
        notes.append(f'是否為跨水橋＝{site}' if site else '尚未讀到基本資料，請確認類型')
        self.kind_note.setText('（' + '；'.join(notes) + '）')
        ready = sum(s.ready for s in self.plan.slots)
        filled = sum(bool(s.path) for s in self.plan.slots)
        done = sum(s.status.startswith(('已上傳', '網站已有')) for s in self.plan.slots)
        chosen = len(self.selected_slots())
        self.summary.setText((message + '\n' if message else '') +
                             f'已放入 {filled}／6 張，可上傳 {ready} 張，已選取 {chosen} 張，已完成 {done} 張')
        self.upload_button.setText(f'上傳選取的照片（{chosen}）')
        self.upload_button.setEnabled(chosen > 0)
