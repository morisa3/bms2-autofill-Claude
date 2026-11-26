"""以隱藏的內嵌頁面讀取網站（唯讀）。

- 與畫面上的網站共用同一個設定檔，所以沿用使用者的登入狀態。
- 只開啟 is_read_only_url() 允許的網址、只呼叫 is_read_only_ajax() 允許的列表函式；從不送出表單。
- 直接開啟網址讀取（等同另開視窗），不操作畫面上的彈出視窗。
- 網頁內容一律以 toHtml 取回，不靠腳本回傳大段文字（實測在 Windows 會變成空字串）。
- 每次讀取後暫停 delay_ms 毫秒，避免對網站造成負擔。
"""
import json

import shiboken6

from PySide6.QtCore import QEventLoop, QTimer, QUrl

from ..site.pages import resolve
from ..site.scan import is_read_only_ajax, is_read_only_url
from .browser import SitePage


SEARCH_URL = 'https://bms2.iot.gov.tw/bms3/inspect/inspect/search/'


class FetchError(RuntimeError):
    pass


def absolute(url):
    """網頁上的連結可能是 /bms3/… 或 inspect/…（以 <base> 為基準）；開啟前補成完整網址。"""
    return resolve(url)


def no_images(page):
    """背景頁面只讀網頁文字：不下載圖片（照片簿列表的縮圖是原尺寸照片，每次開啟都會下載整頁照片）。"""
    from PySide6.QtWebEngineCore import QWebEngineSettings
    page.settings().setAttribute(QWebEngineSettings.WebAttribute.AutoLoadImages, False)


class PageFetcher:
    def __init__(self, profile, parent=None, delay_ms=500, timeout_ms=30000, rewrite=None):
        self.page = SitePage(profile, parent, quiet=True)
        no_images(self.page)
        self.delay_ms = delay_ms
        self.timeout_ms = timeout_ms
        self.rewrite = rewrite or (lambda url: url)   # 測試用：把正式網址導向本機
        self.requests = []
        self.pages = {}          # 網址 → (實際網址, HTML)，供讀不到資料時存檔診斷

    # ------------------------------------------------------------ 等待工具
    def wait(self, signal, timeout=None):
        """等到 signal 觸發或逾時。回傳 (是否觸發, signal 參數)。"""
        loop, box = QEventLoop(), {}
        timer = QTimer(singleShot=True)
        timer.timeout.connect(loop.quit)

        def done(*args):
            box['args'] = args
            loop.quit()
        signal.connect(done)
        timer.start(timeout or self.timeout_ms)
        loop.exec()
        try:
            signal.disconnect(done)
        except RuntimeError:     # 等待期間頁面已被釋放（例：切換分頁或關閉視窗）
            pass
        return 'args' in box, box.get('args', ())

    def wait_until(self, check, timeout=None, interval=300):
        """每 interval 毫秒檢查一次 check()；逐次執行，不會重入。"""
        waited = 0
        while waited < (timeout or self.timeout_ms):
            self.pause(interval)
            waited += interval
            if check():
                return True
        return False

    def pause(self, ms):
        loop = QEventLoop()
        QTimer.singleShot(ms, loop.quit)
        loop.exec()

    def html(self):
        box = {}
        loop = QEventLoop()

        def got(value):
            box['html'] = value
            loop.quit()
        self.page.toHtml(got)
        QTimer.singleShot(self.timeout_ms, loop.quit)
        loop.exec()
        if 'html' not in box:
            raise FetchError('讀取網頁內容逾時')
        return box['html']

    # ------------------------------------------------------------ 讀取
    def get(self, url):
        url = absolute(url)
        if not is_read_only_url(url):
            raise PermissionError(f'不在唯讀清單，拒絕開啟：{url}')
        self.requests.append(url)
        final, content = self.load(url)
        self.remember(url, content, final)
        self.pause(self.delay_ms)
        return content

    def load(self, url):
        self.page.load(QUrl(self.rewrite(url)))
        finished, args = self.wait(self.page.loadFinished)
        if self.page is None or not shiboken6.isValid(self.page):
            raise FetchError(f'讀取中途頁面已關閉：{url}')
        if not finished:
            raise FetchError(f'開啟逾時：{url}')
        if not args or not args[0]:
            raise FetchError(f'開啟失敗：{url}')
        return self.page.url().toString(), self.html()

    def remember(self, url, content, final=''):
        if len(self.pages) > 500:
            self.pages.pop(next(iter(self.pages)))
        self.pages[url] = (final or self.page.url().toString(), content)

    def release(self):
        """關閉前先釋放隱藏頁面，再釋放共用的設定檔。"""
        if self.page is not None:
            shiboken6.delete(self.page)
            self.page = None

    def __call__(self, url):
        return self.get(url)

    def search(self, name):
        """在「定期檢測 → 進階查詢」頁的橋梁名稱欄填入 name、按「開始查詢」，回傳查詢結果頁 HTML。
        這是唯一會在網站上填欄位、按按鈕的動作：只填橋梁名稱、只按同一表單內的「開始查詢」，
        結果必須是查詢結果頁（/inspect/inspect/result/）才接受。查詢不會修改資料。"""
        name = (name or '').strip()
        if not name:
            raise ValueError('請輸入橋名')
        self.get(SEARCH_URL)
        self.requests.append(f'{SEARCH_URL}#查詢:{name}')
        self.page.runJavaScript(f"""(() => {{
            const box = document.getElementById('bridge_name');
            if (!box || !box.form) return;
            const buttons = Array.from(box.form.querySelectorAll('input[type=submit], input[type=button], button'))
                .filter(b => /開始查詢/.test(b.value || b.textContent || ''));
            if (buttons.length !== 1) return;
            box.value = {json.dumps(name)};
            buttons[0].click();
        }})()""")
        finished, args = self.wait(self.page.loadFinished)
        final = self.page.url().toString()
        if not finished or not args or not args[0] or '/inspect/inspect/result/' not in final:
            raise FetchError(f'查詢「{name}」沒有出現查詢結果頁（目前：{final}）')
        content = self.html()
        self.remember(f'{SEARCH_URL}#查詢:{name}', content, final)
        self.pause(self.delay_ms)
        return content

    def call(self, page_url, action, ready_text):
        """在 page_url 頁內呼叫網站的列表函式 xajax_set_action_ajax(action)，等內容出現 ready_text 後回傳 HTML。"""
        if not is_read_only_ajax(action):
            raise PermissionError(f'不在唯讀清單，拒絕呼叫：{action}')
        page_url = absolute(page_url)
        if self.page.url().toString() != self.rewrite(page_url):
            if not is_read_only_url(page_url):
                raise PermissionError(f'不在唯讀清單，拒絕開啟：{page_url}')
            self.requests.append(page_url)
            final, content = self.load(page_url)
            self.remember(page_url, content, final)
        before = self.html()
        self.requests.append(f'{page_url}#{action}')
        self.page.runJavaScript(f'xajax_set_action_ajax({json.dumps(action)})')   # 網站自己的列表函式
        state = {}

        def ready():
            state['html'] = self.html()
            return state['html'] != before and ready_text in state['html']
        if not self.wait_until(ready):
            raise FetchError(f'載入列表逾時：{action}')
        self.remember(f'{page_url}#{action}', state['html'])
        self.pause(self.delay_ms)
        return state['html']
