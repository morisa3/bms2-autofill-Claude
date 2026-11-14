"""以隱藏的內嵌頁面讀取網站（唯讀）。

- 與畫面上的網站共用同一個設定檔，所以沿用使用者的登入狀態。
- 只開啟 is_read_only_url() 允許的網址、只呼叫 is_read_only_ajax() 允許的列表函式；從不送出表單。
- 每次讀取後暫停 delay_ms 毫秒，避免對網站造成負擔。
"""
import json

import shiboken6
from lxml import html as lxml_html

from PySide6.QtCore import QEventLoop, QTimer, QUrl
from PySide6.QtWebEngineCore import QWebEngineScript

from ..site.scan import is_read_only_ajax, is_read_only_url
from .browser import SitePage


class FetchError(RuntimeError):
    pass


SITE_ORIGIN = 'https://bms2.iot.gov.tw'


def copied_frame(page_html):
    """從 toHtml 取回的來源頁中，取出複製到隱藏文字框的框內容與網址。"""
    doc = lxml_html.fromstring(page_html)
    box = doc.get_element_by_id('__bms_copy', None)
    if box is None:
        return '', ''
    return box.get('data-url', ''), box.text or ''


def absolute(url):
    """網頁上的連結多為相對路徑（/bms3/…）；開啟前補成完整網址。"""
    url = (url or '').strip()
    if url.startswith('//'):
        return 'https:' + url
    if url.startswith('/'):
        return SITE_ORIGIN + url
    return url


class PageFetcher:
    def __init__(self, profile, parent=None, delay_ms=500, timeout_ms=30000, rewrite=None):
        self.page = SitePage(profile, parent, quiet=True)
        self.delay_ms = delay_ms
        self.timeout_ms = timeout_ms
        self.rewrite = rewrite or (lambda url: url)   # 測試用：把正式網址導向本機
        self.requests = []
        self.referer = ''        # 來源頁：設定後以頁內 iframe 讀取，比照網站彈出視窗
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
        signal.disconnect(done)
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

    def run(self, script, isolated=True):
        """執行腳本。預設在獨立環境（看得到網頁內容，但不受網站自己的 JavaScript 影響）；
        只有呼叫網站函式（例如列表函式）時才用 isolated=False 在網站環境執行。"""
        box = {}
        loop = QEventLoop()

        def got(value):
            box['value'] = value
            loop.quit()
        world = QWebEngineScript.ScriptWorldId.ApplicationWorld if isolated else QWebEngineScript.ScriptWorldId.MainWorld
        self.page.runJavaScript(script, world, got)
        QTimer.singleShot(self.timeout_ms, loop.quit)
        loop.exec()
        return box.get('value')

    # ------------------------------------------------------------ 讀取
    def get(self, url):
        """讀取唯讀網址。設定 referer（來源頁）時，比照網站彈出視窗的做法：
        隱藏頁面停在來源頁，於頁內建立看不見的 iframe 載入網址再取回內容；
        來源頁、Cookie 與載入方式都和使用者手動點開彈出視窗相同，但不碰畫面上的視窗。"""
        if not is_read_only_url(url):
            raise PermissionError(f'不在唯讀清單，拒絕開啟：{url}')
        self.requests.append(url)
        if self.referer:
            try:
                final, content = self.get_in_frame(url)
            except FetchError:
                self.compare_direct(url)
                raise
        else:
            final, content = self.load_top(url)
        self.remember(url, content, final)
        self.pause(self.delay_ms)
        return content

    def load_top(self, url):
        self.page.load(QUrl(self.rewrite(absolute(url))))
        finished, args = self.wait(self.page.loadFinished)
        if not finished:
            raise FetchError(f'開啟逾時：{url}')
        if not args or not args[0]:
            raise FetchError(f'開啟失敗：{url}')
        return self.page.url().toString(), self.html()

    def get_in_frame(self, url):
        if not is_read_only_url(self.referer):
            raise PermissionError(f'來源頁不在唯讀清單：{self.referer}')
        host = self.rewrite(absolute(self.referer))
        if self.page.url().toString() != host:
            self.load_top(self.referer)
        target = json.dumps(self.rewrite(absolute(url)))
        created = self.run(f"""(() => {{
            const old = document.getElementById('__bms_frame'); if (old) old.remove();
            const frame = document.createElement('iframe');
            frame.id = '__bms_frame'; frame.style.display = 'none';
            frame.addEventListener('load', () => frame.dataset.loaded = '1');
            frame.src = {target};
            (document.body || document.documentElement).appendChild(frame);
            return true; }})()""")
        if not created:
            raise FetchError(f'無法在來源頁建立讀取框：{url}')
        state = {}

        def copied():
            # 把框內容複製到頁面上隱藏的文字框；腳本只回傳數字（長度），大量內容改由 toHtml 取回
            state['value'] = self.run("""(() => {
                const frame = document.getElementById('__bms_frame');
                if (!frame || frame.dataset.loaded !== '1') return null;
                const doc = frame.contentDocument;
                if (!doc || !doc.documentElement) return -1;
                let box = document.getElementById('__bms_copy');
                if (!box) {
                    box = document.createElement('textarea');
                    box.id = '__bms_copy'; box.style.display = 'none';
                    (document.body || document.documentElement).appendChild(box);
                }
                const html = doc.documentElement.outerHTML;
                box.textContent = html;
                box.setAttribute('data-url', frame.contentWindow.location.href);
                return html.length;
            })()""")
            return state['value'] is not None
        if not self.wait_until(copied, interval=200):
            raise FetchError(f'開啟逾時：{url}')
        length = state['value']
        if not isinstance(length, (int, float)) or length < 0:
            self.remember(url, f'<!-- 框內容無法讀取，回傳：{length!r} -->', '')
            raise FetchError(f'框內容無法讀取（{length!r}）：{url}')
        final, content = copied_frame(self.html())
        if len(content.encode('utf-16-le')) // 2 != int(length):   # JavaScript 以 UTF-16 計算長度
            self.remember(url, f'<!-- 框內容長度不符：頁面 {int(length)}，取回 {len(content)} -->\n{content[:20000]}', final)
            raise FetchError(f'框內容長度不符（{int(length)}／{len(content)}）：{url}')
        return final, content

    def compare_direct(self, url):
        """iframe 讀取失敗時，另外直接開啟同一網址（唯讀）存下結果，供診斷比對。"""
        try:
            final, content = self.load_top(url)
        except FetchError as error:
            final, content = '', f'<!-- 直接開啟也失敗：{error} -->'
        self.remember(f'{url} (直接開啟)', content, final)

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

    def call(self, page_url, action, ready_text):
        """在 page_url 頁內呼叫網站的列表函式 xajax_set_action_ajax(action)，等內容出現 ready_text 後回傳 HTML。"""
        if not is_read_only_ajax(action):
            raise PermissionError(f'不在唯讀清單，拒絕呼叫：{action}')
        if self.page.url().toString() != self.rewrite(absolute(page_url)):
            if not is_read_only_url(page_url):
                raise PermissionError(f'不在唯讀清單，拒絕開啟：{page_url}')
            self.requests.append(page_url)
            final, content = self.load_top(page_url)   # 列表函式在這一頁本身執行，不放在 iframe
            self.remember(page_url, content, final)
        before = self.html()
        self.requests.append(f'{page_url}#{action}')
        self.run(f'xajax_set_action_ajax({json.dumps(action)})', isolated=False)
        state = {}

        def ready():
            state['html'] = self.html()
            return state['html'] != before and ready_text in state['html']
        if not self.wait_until(ready):
            raise FetchError(f'載入列表逾時：{action}')
        self.remember(f'{page_url}#{action}', state['html'])
        self.pause(self.delay_ms)
        return state['html']
