"""單一執行檔的進入點（PyInstaller）。

- 第一次執行：把內附的規則表複製到執行檔旁的 data\\（規則表會自動新增規則，必須放在可寫入的位置）。
- 啟動失敗：錯誤寫到執行檔旁的 log\\錯誤紀錄.txt，並跳出視窗說明（不會無聲閃退）。
- --self-test：只確認內嵌瀏覽器能啟動並載入頁面（建置時自動測試用），成功回傳 0。
"""
import os
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path


def app_dir():
    return Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parents[1]


def write_error(text):
    folder = app_dir() / 'log'
    try:
        folder.mkdir(parents=True, exist_ok=True)
        with open(folder / '錯誤紀錄.txt', 'a', encoding='utf-8') as handle:
            handle.write(f'\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{text}\n')
    except OSError:
        pass
    return folder / '錯誤紀錄.txt'


def show_error(text):
    path = write_error(text)
    message = f'程式啟動發生錯誤，詳細內容已寫入：\n{path}\n\n請把這個檔案傳給維護人員。\n\n{text[-1500:]}'
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        app = QApplication.instance() or QApplication(sys.argv[:1])
        QMessageBox.critical(None, '公路橋梁填報助手', message)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, '公路橋梁填報助手', 0x10)
        except Exception:
            pass


def install_data():
    """內附資料（規則表）複製到執行檔旁；已存在的不覆蓋（保留使用者與程式自動新增的規則）。"""
    bundled = Path(getattr(sys, '_MEIPASS', app_dir())) / 'data'
    target = app_dir() / 'data'
    target.mkdir(parents=True, exist_ok=True)
    if bundled.resolve() == target.resolve():
        return
    for source in bundled.glob('*'):
        if source.is_file() and not (target / source.name).exists():
            shutil.copy2(source, target / source.name)


def self_test():
    """建置後自動測試：內嵌瀏覽器能啟動、載入頁面，且主視窗能建立。"""
    from PySide6.QtCore import QTimer, QUrl
    from PySide6.QtWebEngineWidgets import QWebEngineView      # 必須在建立 QApplication 之前匯入
    from PySide6.QtWidgets import QApplication
    import bridgefill.ui.main_window  # noqa: F401  確認所有模組都已打包
    app = QApplication.instance() or QApplication(sys.argv[:1])
    view = QWebEngineView()
    result = {'ok': False}

    def finished(ok):
        result['ok'] = ok
        app.quit()
    view.loadFinished.connect(finished)
    QTimer.singleShot(0, lambda: view.setHtml('<html><body>self-test</body></html>', QUrl('about:blank')))
    QTimer.singleShot(60000, app.quit)
    app.exec()
    print('self-test', 'ok' if result['ok'] else 'failed', flush=True)
    return 0 if result['ok'] else 1


def main():
    def hook(kind, value, trace):
        write_error(''.join(traceback.format_exception(kind, value, trace)))
        sys.__excepthook__(kind, value, trace)
    sys.excepthook = hook                      # 執行中未處理的錯誤也記錄下來
    if not getattr(sys, 'frozen', False) and str(app_dir()) not in sys.path:
        sys.path.insert(0, str(app_dir()))      # 直接以 python packaging/entry.py 執行時
    if '--self-test' in sys.argv:              # 建置測試：不跳視窗（無人可按），錯誤印出並回傳失敗
        try:
            return self_test()
        except BaseException:
            text = traceback.format_exc()
            write_error(text)
            print(text, flush=True)
            return 1
    try:
        install_data()
        os.chdir(app_dir())
        from bridgefill.ui.main_window import main as run
        return run([a for a in sys.argv[1:] if a != '--self-test'])
    except SystemExit:
        raise
    except BaseException:
        show_error(traceback.format_exc())
        return 1


if __name__ == '__main__':
    sys.exit(main())
