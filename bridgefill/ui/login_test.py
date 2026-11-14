"""第 2 步：內嵌登入測試。

確認三件事：
1. bms2 登入頁能在內嵌瀏覽器中正常顯示，並可手動完成 reCAPTCHA（或改用圖片驗證碼）登入。
2. 程式能讀到登入後的系統選單（之後掃描的前提）。
3. 關閉再開啟後仍維持登入（Cookie 保存在 browser-profile）。

程式不讀取、不填寫帳號密碼，也不操作驗證碼。結果寫在 log/login-test-*.log，不含帳密與 Cookie。

用法：
    python -m bridgefill.ui.login_test
    python -m bridgefill.ui.login_test --url 本機網頁檔.html   （離線測試）
"""
import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QTimer, QUrl, Qt
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QHBoxLayout, QLabel, QMainWindow, QMessageBox, QPlainTextEdit,
                               QPushButton, QSplitter, QVBoxLayout, QWidget)
from PySide6 import __version__ as PYSIDE_VERSION
from PySide6.QtWebEngineCore import qWebEngineChromiumVersion

from ..core.local_files import APP_DIR
from ..site.session import LOGGED_IN, LOGIN_PAGE, LOGIN_URL, PROBE_JS, classify, describe, parse_probe, redact_url
from .browser import PROFILE_DIR, SiteView, make_profile

LOG_DIR = APP_DIR / 'log'
BANNER = {
    LOGIN_PAGE: ('#fff4ce', '請在下方頁面自行輸入帳號密碼並完成驗證，再按「確定登入」。'
                            '若「我不是機器人」一直出現圖片題，可改勾「使用驗證碼登入」。'),
    LOGGED_IN: ('#dff6dd', '登入成功，程式已讀到系統選單。請關閉視窗後重新開啟，確認是否仍維持登入。'),
}


class LoginTestWindow(QMainWindow):
    def __init__(self, url=LOGIN_URL, profile_dir=PROFILE_DIR, log_dir=LOG_DIR):
        super().__init__()
        self.setWindowTitle('公路橋梁填報助手－內嵌登入測試')
        self.resize(1560, 920)
        self.start_url = url
        self.state = None
        self.loaded = 0
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = log_dir / f'login-test-{datetime.now():%Y%m%d-%H%M%S}.log'

        self.profile = make_profile(profile_dir, self)
        self.view = SiteView(self.profile, self)
        self.banner = QLabel('載入中…')
        self.banner.setWordWrap(True)
        self.banner.setMargin(10)
        self.output = QPlainTextEdit(readOnly=True)
        self.output.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.output.setFont(QFont('Consolas', 9))

        buttons = QHBoxLayout()
        for text, slot in (('回登入頁', lambda: self.view.open(self.start_url)),
                           ('重新整理', self.view.reload),
                           ('檢查登入狀態', self.check),
                           ('清除登入資料', self.clear_login),
                           ('開啟日誌資料夾', lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(log_dir))))):
            button = QPushButton(text)
            button.clicked.connect(slot)
            buttons.addWidget(button)
        buttons.addStretch()

        side = QWidget()
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.addWidget(QLabel('測試紀錄（不含帳密與 Cookie）'))
        side_layout.addWidget(self.output)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.view)
        splitter.addWidget(side)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setSizes([1220, 320])  # 網站版面約需 1150px 寬，紀錄欄可拖曳調整

        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addLayout(buttons)
        layout.addWidget(self.banner)
        layout.addWidget(splitter, 1)
        self.setCentralWidget(root)

        self.view.loadFinished.connect(self.on_loaded)
        self.view.urlChanged.connect(lambda u: u.isEmpty() or self.log(f'網址：{redact_url(u.toString())}'))
        self.log(f'PySide6 {PYSIDE_VERSION}｜Chromium {qWebEngineChromiumVersion()}｜{sys.platform}')
        self.log(f'User-Agent：{self.profile.httpUserAgent()}')
        self.log(f'設定檔：{Path(profile_dir).resolve()}')
        self.log(f'日誌：{self.log_path}')
        self.view.open(url)

    def log(self, message):
        line = f'{datetime.now():%H:%M:%S} {message}'
        self.output.appendPlainText(line)
        with open(self.log_path, 'a', encoding='utf-8') as handle:
            handle.write(line + '\n')

    def on_loaded(self, ok):
        self.loaded += 1
        if not ok:
            self.log('頁面載入失敗（網路、VPN 或網站暫停服務）')
            self.show_banner('#fde7e9', '頁面載入失敗，請確認網路後按「重新整理」。')
            return
        self.check()
        # 驗證元件與選單可能稍後才出現，延遲再判讀一次
        QTimer.singleShot(1500, self.check)

    def check(self):
        self.view.probe(PROBE_JS, self.on_probe)

    def on_probe(self, raw):
        probe = parse_probe(raw)
        state = classify(probe)
        if state != self.state:
            self.log(describe(probe))
            if state == LOGGED_IN and self.state is None:
                self.log('開啟時已是登入狀態（沿用上次登入，Cookie 保存成功）')
            elif state == LOGGED_IN:
                self.log('登入成功')
            self.state = state
        color, text = BANNER.get(state, ('#eeeeee', f'目前頁面：{describe(probe) or "無法判讀"}'))
        if state == LOGGED_IN:
            text += '\n選單：' + '、'.join(probe.get('menuItems') or [])
        self.show_banner(color, text)

    def show_banner(self, color, text):
        self.banner.setStyleSheet(f'background:{color}; border-radius:6px; font-size:14px;')
        self.banner.setText(text)

    def closeEvent(self, event):
        self.log('關閉測試視窗')
        self.view.release()
        super().closeEvent(event)

    def clear_login(self):
        answer = QMessageBox.question(self, '清除登入資料', '清除後需要重新登入。確定清除？')
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.profile.cookieStore().deleteAllCookies()
        self.state = None
        self.log('已清除登入資料（Cookie）')
        self.view.open(self.start_url)


def main(argv=None):
    parser = argparse.ArgumentParser(description='內嵌登入測試')
    parser.add_argument('--url', default=LOGIN_URL, help='開啟的網址或本機網頁檔')
    parser.add_argument('--profile', default=str(PROFILE_DIR), help='瀏覽器設定檔資料夾')
    args = parser.parse_args(argv)
    url = args.url
    if os.path.exists(url):
        url = QUrl.fromLocalFile(str(Path(url).resolve())).toString()
    app = QApplication.instance() or QApplication(sys.argv[:1])
    window = LoginTestWindow(url, args.profile)
    window.show()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
