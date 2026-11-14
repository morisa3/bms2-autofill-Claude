"""內嵌 bms2 網站的瀏覽器元件（QtWebEngine）。

- 使用獨立、持久的設定檔，登入 Cookie 保存在程式旁的 browser-profile 資料夾，下次開啟不必重新登入。
- 帳號密碼與驗證碼一律由使用者在頁面中手動操作；程式不讀取、不填寫。
"""
from pathlib import Path

import shiboken6

from PySide6.QtCore import QUrl
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineProfile, QWebEngineScript, QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView

from ..core.local_files import APP_DIR
from ..site.session import clean_user_agent

PROFILE_DIR = APP_DIR / 'browser-profile'
PROFILE_NAME = 'bms2'


def make_profile(directory=PROFILE_DIR, parent=None):
    """建立持久設定檔：Cookie、快取與 localStorage 都寫在 directory。"""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    profile = QWebEngineProfile(PROFILE_NAME, parent)
    profile.setPersistentStoragePath(str(directory / 'storage'))
    profile.setCachePath(str(directory / 'cache'))
    profile.setPersistentCookiesPolicy(QWebEngineProfile.PersistentCookiesPolicy.ForcePersistentCookies)
    profile.setHttpAcceptLanguage('zh-TW,zh;q=0.9,en;q=0.8')
    profile.setHttpUserAgent(clean_user_agent(profile.httpUserAgent()))
    settings = profile.settings()
    settings.setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, True)
    settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanOpenWindows, True)
    settings.setAttribute(QWebEngineSettings.WebAttribute.FullScreenSupportEnabled, False)
    return profile


class SitePage(QWebEnginePage):
    """網站開新視窗（例如列印預覽）時，改在同一個檢視中開啟，不跳出外部視窗。"""

    def createWindow(self, _type):
        return self


class SiteView(QWebEngineView):
    def __init__(self, profile, parent=None):
        super().__init__(parent)
        self.setPage(SitePage(profile, self))

    def open(self, url):
        self.setUrl(QUrl.fromUserInput(str(url)))

    def release(self):
        """關閉前先釋放頁面，再釋放設定檔，確保 Cookie 完整寫入磁碟。"""
        page = self.page()
        self.setPage(QWebEnginePage(self))
        shiboken6.delete(page)

    def probe(self, script, callback):
        """在獨立的腳本環境執行唯讀腳本（看得到 DOM，但不干擾網站本身的 JavaScript），結果交給 callback。"""
        world = QWebEngineScript.ScriptWorldId.ApplicationWorld
        self.page().runJavaScript(script, world, callback)
