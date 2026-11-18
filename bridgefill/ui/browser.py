"""內嵌 bms2 網站的瀏覽器元件（QtWebEngine）。

- 使用獨立、持久的設定檔，登入 Cookie 保存在程式旁的 browser-profile 資料夾，下次開啟不必重新登入。
- 帳號密碼與驗證碼一律由使用者在頁面中手動操作；程式不讀取、不填寫。
"""
import re
from pathlib import Path

import shiboken6

from PySide6.QtCore import QUrl
from PySide6.QtWebEngineCore import (QWebEnginePage, QWebEngineProfile, QWebEngineScript, QWebEngineSettings,
                                     QWebEngineUrlRequestInfo, QWebEngineUrlRequestInterceptor)
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


# 精簡網頁：只改畫面（隱藏區塊），不改表單欄位，送出與讀取都不受影響。
# html.bf-lite 時：隱藏網站標題列與選單、檢測人員與檢測開始／結束照片、檢測路徑、3D 與輸出 Word 按鈕；
# 構件列表改成左上角的「構件列表」按鈕，點了才展開。
LITE_JS = r"""
(() => {
  const root = document.documentElement;
  root.classList.toggle('bf-lite', %s);
  if (window.__bfLite) return;
  window.__bfLite = true;
  const style = document.createElement('style');
  style.textContent = `
    html.bf-lite #topsection, html.bf-lite #menu-wrap, html.bf-lite .bf-hide,
    html.bf-lite input[value="3D動態生成"], html.bf-lite #maincontent .dropdown { display: none !important; }
    html.bf-lite #leftcolumn { display: none !important; }
    html.bf-lite #contentcolumn { margin-left: 0 !important; }
    html.bf-lite td:has(> #member_der) { width: 0 !important; padding: 0 !important; }
    html.bf-lite #member_list { display: none !important; }
    html.bf-lite.bf-members-open #member_list { display: block !important; position: fixed !important;
      left: 6px !important; top: 40px !important; max-height: 80vh; overflow: auto; z-index: 9999;
      background: #fff; box-shadow: 0 4px 16px rgba(0,0,0,.25); }
    #bf-members { display: none; }
    html.bf-lite #bf-members { display: block; position: fixed; left: 6px; top: 6px; z-index: 10000;
      background: #00bfff; color: #fff; border: 0; border-radius: 4px; padding: 4px 10px; font-weight: bold;
      cursor: pointer; box-shadow: 0 2px 6px rgba(0,0,0,.2); }`;
  (document.head || root).appendChild(style);
  const mark = () => {
    for (const th of document.querySelectorAll('th')) {
      const text = (th.textContent || '').trim();
      if (/^(檢測人員|檢測開始|檢測結束)/.test(text)) {
        const table = th.closest('table');
        if (table) table.classList.add('bf-hide');
      } else if (/^檢測路徑/.test(text)) {
        const row = th.closest('tr');
        if (row) {
          row.classList.add('bf-hide');
          let next = row.nextElementSibling;
          while (next && !next.querySelector('th')) { next.classList.add('bf-hide'); next = next.nextElementSibling; }
        }
      }
    }
    // 登入頁：只留帳號、密碼、驗證與登入按鈕（上方標題保留），左側公告列表與行動裝置下載隱藏
    for (const a of document.querySelectorAll('a.list-group-item')) {
      if (!/公告列表/.test(a.textContent || '')) continue;
      const cell = a.closest('td');
      if (cell && !cell.querySelector('form')) {
        cell.classList.add('bf-hide');
        const gap = cell.nextElementSibling;
        if (gap && !gap.querySelector('form, input')) gap.classList.add('bf-hide');
      }
    }
    for (const e of document.querySelectorAll('div, td, span, p, a, label')) {
      if (e.childElementCount > 3 || !/行動裝置下載/.test(e.textContent || '')) continue;
      let box = e;
      while (box.parentElement && !box.parentElement.querySelector('form, input[type=password]')) box = box.parentElement;
      box.classList.add('bf-hide');
    }
    if (document.getElementById('member_list') && !document.getElementById('bf-members')) {
      const button = document.createElement('button');
      button.id = 'bf-members';
      button.type = 'button';
      button.textContent = '構件列表';
      button.addEventListener('click', () => root.classList.toggle('bf-members-open'));
      document.body.appendChild(button);
      document.getElementById('member_list').addEventListener('click', event => {
        if (event.target.closest('a')) root.classList.remove('bf-members-open');
      });
    }
  };
  mark();
})();
"""


def lite_script(on):
    script = QWebEngineScript()
    script.setName('bms-lite')
    script.setSourceCode(LITE_JS % ('true' if on else 'false'))
    script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentReady)
    script.setWorldId(QWebEngineScript.ScriptWorldId.ApplicationWorld)
    script.setRunsOnSubFrames(False)
    return script


class LiteInterceptor(QWebEngineUrlRequestInterceptor):
    """精簡網頁時，缺失主頁面不下載檢測人員與檢測開始／結束照片（網站圖示、第二級的損壞照片照常載入）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.enabled = True
        self.blocked = 0

    def interceptRequest(self, info):
        if not self.enabled or info.resourceType() != QWebEngineUrlRequestInfo.ResourceType.ResourceTypeImage:
            return
        page = info.firstPartyUrl().path()
        path = info.requestUrl().path().lower()
        if '/inspect/inspect/der_content/' in page and re.search(r'\.(jpe?g|png)$', path) \
                and not re.search(r'/(img|css|images|js)/', path):
            info.block(True)
            self.blocked += 1


def set_lite(profile, on, views=()):
    """切換精簡網頁：之後開啟的頁面套用；目前開著的頁面立即切換（不重新載入）。"""
    interceptor = getattr(profile, 'lite_interceptor', None)
    if interceptor is not None:
        interceptor.enabled = on
    scripts = profile.scripts()
    for old in scripts.find('bms-lite'):
        scripts.remove(old)
    scripts.insert(lite_script(on))
    for view in views:
        view.page().runJavaScript(LITE_JS % ('true' if on else 'false'), QWebEngineScript.ScriptWorldId.ApplicationWorld)


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
    profile.scripts().insert(lite_script(True))
    profile.lite_interceptor = LiteInterceptor(profile)
    profile.setUrlRequestInterceptor(profile.lite_interceptor)
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
