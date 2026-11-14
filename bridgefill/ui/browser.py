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
SITE = 'https://bms2.iot.gov.tw/bms3/'
SHORTCUTS = {
    '定期檢測查詢': SITE + 'inspect/inspect/search/',
    '基本資料搜尋': SITE + 'basic/inventory/search/',
}

# 網站最上層選單是 href="javascript:return false;"，設計為滑鼠移上去才展開；直接點會產生
# 「Illegal return statement」錯誤且沒有反應。這段腳本讓它改成點一下展開／收合（只改畫面，不送出資料）。
MENU_CLICK_JS = r"""
(() => {
  if (window.__bmsMenuClick) return;
  window.__bmsMenuClick = true;
  const close = except => document.querySelectorAll('#menu ul[data-bms-open]').forEach(ul => {
    if (except && (ul === except || ul.contains(except))) return;
    ul.style.display = ''; ul.style.visibility = ''; ul.removeAttribute('data-bms-open');
  });
  document.addEventListener('click', event => {
    const link = event.target.closest && event.target.closest('#menu a');
    if (!link) { close(); return; }
    if (!/^javascript:\s*return\s+false/i.test(link.getAttribute('href') || '')) return;
    event.preventDefault();
    const sub = link.parentElement.querySelector(':scope > ul');
    if (!sub) return;
    const open = sub.hasAttribute('data-bms-open');
    close(link.parentElement.parentElement.closest('ul[data-bms-open]'));
    if (!open) { sub.style.display = 'block'; sub.style.visibility = 'visible'; sub.setAttribute('data-bms-open', ''); }
  }, true);
})();
"""


def menu_script():
    script = QWebEngineScript()
    script.setName('bms-menu-click')
    script.setSourceCode(MENU_CLICK_JS)
    script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentReady)
    script.setWorldId(QWebEngineScript.ScriptWorldId.ApplicationWorld)
    script.setRunsOnSubFrames(False)
    return script


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
    profile.scripts().insert(menu_script())
    return profile


class SitePage(QWebEnginePage):
    """網站開新視窗（例如列印預覽）時，改在同一個檢視中開啟，不跳出外部視窗。

    quiet=True 用於背景讀取的隱藏頁面：網站的提示視窗不能停下來等人按，否則掃描會卡住。
    提示（alert）記錄後關閉；確認（confirm）與輸入（prompt）一律取消，避免誤觸刪除等動作。
    畫面上的頁面維持網站原本的提示視窗，由使用者自行判斷。
    """

    def __init__(self, profile, parent=None, quiet=False):
        super().__init__(profile, parent)
        self.quiet = quiet
        self.dialogs = []

    def createWindow(self, _type):
        return self

    def javaScriptAlert(self, origin, message):
        if not self.quiet:
            return super().javaScriptAlert(origin, message)
        self.dialogs.append(('提示', message))

    def javaScriptConfirm(self, origin, message):
        if not self.quiet:
            return super().javaScriptConfirm(origin, message)
        self.dialogs.append(('確認（已取消）', message))
        return False

    def javaScriptPrompt(self, origin, message, default):
        if not self.quiet:
            return super().javaScriptPrompt(origin, message, default)
        self.dialogs.append(('輸入（已取消）', message))
        return False, ''


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
