"""第 3 步：唯讀掃描測試。

在內嵌網站中開到「檢測缺失主頁面」（定期檢測 → 修改），按「掃描目前檢測」：
程式以隱藏頁面依序讀取每一列的第一級清單與第二級欄位，核對「詳細資料(n)」筆數，
結果顯示在右側表格，並存成 temp/scan/*.json。全程只讀取，不送出、不修改網站資料。

用法：
    python -m bridgefill.ui.scan_test
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QHBoxLayout, QHeaderView, QLabel, QMainWindow, QMessageBox,
                               QPlainTextEdit, QPushButton, QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from ..core.local_files import APP_DIR
from ..site.pages import parse_inspection
from ..site.scan import Stopped, inventory_id, scan_basic, scan_inspection, scan_prior
from ..site.session import LOGIN_URL
from .browser import PROFILE_DIR, SHORTCUTS, SiteView, make_profile
from .fetcher import FetchError, PageFetcher

RESULT_DIR = APP_DIR / 'temp' / 'scan'
COLUMNS = ('構件', '位置／編號', 'DERU', '維修狀態', '網站筆數', '讀到', '核對')


class ScanTestWindow(QMainWindow):
    def __init__(self, url=LOGIN_URL, profile_dir=PROFILE_DIR, result_dir=RESULT_DIR, delay_ms=500, rewrite=None):
        super().__init__()
        self.setWindowTitle('公路橋梁填報助手－唯讀掃描測試')
        self.resize(1680, 940)
        self.result_dir = Path(result_dir)
        self.stop_requested = False
        self.busy = False
        self.profile = make_profile(profile_dir, self)
        self.view = SiteView(self.profile, self)
        self.fetcher = PageFetcher(self.profile, self, delay_ms=delay_ms, rewrite=rewrite)

        self.banner = QLabel('按「定期檢測查詢」搜尋橋梁，在查詢結果按「修改」開到「檢測缺失主頁面」，再按「掃描目前檢測」。'
                             '網站最上層選單也可以點一下展開。掃描只讀取資料，不會修改網站。')
        self.banner.setWordWrap(True)
        self.banner.setMargin(8)
        self.banner.setStyleSheet('background:#eef4fb; border-radius:6px; font-size:14px;')
        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.output = QPlainTextEdit(readOnly=True)
        self.output.setFont(QFont('Consolas', 9))
        self.output.setMaximumBlockCount(2000)

        self.buttons = {}
        bar = QHBoxLayout()
        for text, target in SHORTCUTS.items():
            button = QPushButton(text)
            button.setToolTip(target)
            button.clicked.connect(lambda _=False, u=target: self.view.open(u))
            bar.addWidget(button)
        bar.addSpacing(24)
        for key, text, slot in (('scan', '掃描目前檢測', lambda: self.scan(True)),
                                ('quick', '只核對筆數（較快）', lambda: self.scan(False)),
                                ('prior', '掃描前期', lambda: self.scan(True, prior=True)),
                                ('basic', '掃描基本資料', self.scan_basic),
                                ('stop', '停止', self.stop),
                                ('folder', '開啟結果資料夾', self.open_folder)):
            button = QPushButton(text)
            button.clicked.connect(slot)
            bar.addWidget(button)
            self.buttons[key] = button
        bar.addStretch()
        self.buttons['stop'].setEnabled(False)

        side = QSplitter(Qt.Orientation.Vertical)
        side.addWidget(self.table)
        side.addWidget(self.output)
        side.setSizes([620, 220])
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.view)
        splitter.addWidget(side)
        splitter.setSizes([1100, 580])
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addLayout(bar)
        layout.addWidget(self.banner)
        layout.addWidget(splitter, 1)
        self.setCentralWidget(root)
        self.view.open(url)

    # ------------------------------------------------------------ 共用
    def log(self, message):
        self.output.appendPlainText(f'{datetime.now():%H:%M:%S} {message}')

    def set_busy(self, busy):
        self.busy = busy
        for key in ('scan', 'quick', 'prior', 'basic'):
            self.buttons[key].setEnabled(not busy)
        self.buttons['stop'].setEnabled(busy)

    def stop(self):
        self.stop_requested = True
        self.log('停止中，完成目前這一筆後停止…')

    def open_folder(self):
        self.result_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.result_dir)))

    def current_html(self):
        box = {}
        self.view.page().toHtml(lambda html: box.setdefault('html', html))
        if not self.fetcher.wait_until(lambda: 'html' in box, timeout=10000, interval=100):
            raise FetchError('讀取目前頁面逾時')
        return box['html']

    def save(self, kind, header, data):
        self.result_dir.mkdir(parents=True, exist_ok=True)
        date = header.get('檢測日期') or f'{datetime.now():%Y-%m-%d}'
        name = f"{date}-{header.get('橋梁名稱', '')}-{kind}.json".replace('/', '_').replace('\\', '_')
        path = self.result_dir / name
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        return path

    # ------------------------------------------------------------ 掃描
    def scan(self, read_forms, prior=False):
        """prior=False：掃描畫面上這一期（讀編輯頁欄位，不送出）。
        prior=True：由畫面上的本期找出前期，以檢視頁掃描前期。"""
        if self.busy:
            return
        try:
            html = self.current_html()
        except FetchError as error:
            self.log(str(error))
            return
        page = parse_inspection(html)
        if not page.rows or not page.header.get('檢測日期'):
            QMessageBox.information(self, '唯讀掃描', '目前頁面不是「檢測缺失主頁面」。\n'
                                    '請在左側網站開到：檢測資料 → 定期檢測 → 查詢 → 修改。')
            return
        header = page.header
        what = '前期（由本期找出）' if prior else '目前檢測'
        self.log(f"開始掃描{what}：{header['橋梁名稱']} {header['檢測日期']}"
                 f"{'' if read_forms else '（只核對筆數）'}")
        self.table.setRowCount(0)
        self.stop_requested = False
        self.set_busy(True)
        started, before = datetime.now(), len(self.fetcher.requests)
        options = dict(read_forms=read_forms, progress=self.progress, stopped=lambda: self.stop_requested)
        try:
            if prior:
                record, result = scan_prior(html, self.fetcher, **options)
                self.log(f'前期：{record.date}（檢測系統 ID {record.inspect_id}，{record.inspectors}）')
                header = result.header
            else:
                result = scan_inspection(html, self.fetcher, **options)
        except Stopped as error:
            self.log(str(error))
            return
        except (FetchError, PermissionError, ValueError) as error:
            self.log(f'掃描中斷：{error}')
            self.dump(header, '前期' if prior else '檢測', self.fetcher.requests[before:])
            return
        finally:
            self.set_busy(False)
        # component_id 在不同構件之間會重複，須與 component_form 一起比對
        failed = {(r.keys.get('component_form'), r.keys.get('component_id')) for r in result.rows if not r.ok}
        suspects = [u for u in self.fetcher.requests[before:]
                    if any(f'component_form,{form}/component_id,{cid}/' in u for form, cid in failed if cid)]
        if suspects:
            self.dump(header, '前期' if prior else '檢測', suspects)
        for row in result.rows:
            self.add_row(row)
        seconds = (datetime.now() - started).total_seconds()
        path = self.save('前期' if prior else '檢測', header, result.to_dict())
        summary = '｜'.join(f'{k} {v}' for k, v in result.summary.items())
        self.log(f'完成（{seconds:.0f} 秒，讀取 {len(self.fetcher.requests) - before} 頁）：{summary}')
        self.log(f'結果：{path}')
        for problem in result.problems:
            self.log(f'網頁問題：{problem}')
        self.log_dialogs()

    def log_dialogs(self):
        """背景頁面遇到的網站提示視窗（已自動關閉／取消）。"""
        for kind, message in self.fetcher.page.dialogs:
            self.log(f'網站{kind}：{message}')
        self.fetcher.page.dialogs.clear()

    def dump(self, header, kind, urls, limit=40):
        """把讀不到資料的頁面原樣存檔，供診斷網站實際回傳的內容（只存在本機）。"""
        urls = list(dict.fromkeys(u for u in urls if u in self.fetcher.pages))[:limit]
        if not urls:
            return
        date = header.get('檢測日期') or f'{datetime.now():%Y-%m-%d}'
        folder = self.result_dir / 'debug' / f"{date}-{header.get('橋梁名稱', '')}-{kind}-{datetime.now():%H%M%S}"
        folder.mkdir(parents=True, exist_ok=True)
        index = []
        for number, url in enumerate(urls, 1):
            final, html = self.fetcher.pages[url]
            (folder / f'{number:02d}.html').write_text(html, encoding='utf-8')
            index.append(f'{number:02d}.html\t要求：{url}\t實際：{final}\t長度：{len(html)}')
        (folder / 'index.txt').write_text('\n'.join(index), encoding='utf-8')
        self.log(f'已存下 {len(urls)} 個讀取異常的頁面供診斷：{folder}')

    def progress(self, index, total, label):
        self.statusBar().showMessage(f'掃描中 {index}／{total}：{label}')

    def add_row(self, row):
        read = '—' if row.check.endswith('未開啟') else str(len(row.defects))
        values = (row.component, row.label, row.deru, row.repair,
                  '' if row.expected is None else str(row.expected), read, row.check or '✔')
        index = self.table.rowCount()
        self.table.insertRow(index)
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            if not row.ok:
                item.setBackground(QColor('#fde7e9'))
            self.table.setItem(index, col, item)

    def scan_basic(self):
        if self.busy:
            return
        try:
            page = parse_inspection(self.current_html())
        except FetchError as error:
            self.log(str(error))
            return
        inventory = inventory_id(page)
        if not inventory:
            QMessageBox.information(self, '掃描基本資料', '請先開到「檢測缺失主頁面」，程式會從頁面取得橋梁代號。')
            return
        self.set_busy(True)
        try:
            before = len(self.fetcher.requests)
            data = scan_basic(inventory, self.fetcher, self.fetcher.call,
                              expected_spans=int(page.header.get('總橋孔數') or 0) or None)
        except (FetchError, PermissionError) as error:
            self.log(f'基本資料掃描中斷：{error}')
            return
        finally:
            self.set_busy(False)
        path = self.save('基本資料', page.header, data)
        values = data['values']
        spans = sum(len(s['spans']) for s in data['structures'])
        girders = [len(span['girders']) for s in data['structures'] for span in s['spans']]
        self.log(f"基本資料：淨寬 {values.get('width_min')}～{values.get('width_max')}m，"
                 f"橋下淨高 {values.get('height_below')}m，跨距分配 {len(data['distribution'])} 孔，"
                 f"橋孔資料表 {spans} 孔，主梁根數 {girders}")
        self.log(f'結果：{path}')
        for problem in data['problems']:
            self.log(f'基本資料問題：{problem}')
        if data['problems']:
            self.dump(page.header, '基本資料', self.fetcher.requests[before:])

    def closeEvent(self, event):
        self.stop_requested = True
        self.fetcher.release()
        self.view.release()
        super().closeEvent(event)


def main(argv=None):
    parser = argparse.ArgumentParser(description='唯讀掃描測試')
    parser.add_argument('--url', default=LOGIN_URL)
    parser.add_argument('--delay', type=int, default=500, help='每次讀取後暫停毫秒數')
    args = parser.parse_args(argv)
    app = QApplication.instance() or QApplication(sys.argv[:1])
    window = ScanTestWindow(args.url, delay_ms=args.delay)
    window.show()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
