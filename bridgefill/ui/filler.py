"""第 6 步：在隱藏頁面依建議維修工法 ID 填寫第二級表單並送出（會修改網站資料）。

流程（每筆）：開啟第二級編輯頁 → 核對 ID 與網站現值 → 填入預定值 → 讀回確認 → 記錄「待確認」→
按「送出更新」→ 重新開啟回讀 → 一致才記為「已完成」。任何一步不符就停止該筆，不會自動重送。

安全限制：
- 只開啟第二級編輯頁（edit_damage）；只在按下送出的那一刻允許一次送往 update_damage 的表單送出。
- 其他導覽（刪除、新增等）一律擋下並記錄。
- 網站的確認視窗只在送出當下接受（例如「確定送出？」），其他時候一律取消；提示訊息記錄在結果中。
- 網頁內容一律以 toHtml 讀回（不靠腳本回傳值）。
"""
import json
import re

from PySide6.QtWebEngineCore import QWebEnginePage

from ..core.fill import DONE, FAILED, FILL_FIELDS, NAMES, PENDING, compare
from ..site.pages import parse_defect_form
from .browser import SitePage
from .fetcher import FetchError, PageFetcher

EDIT_PAGE = re.compile(r'/inspect/damage/edit_damage/der_id,\d+/')
MAIN_PAGE = re.compile(r'/inspect/inspect/der_content/inspect_id,\d+/')
SUBMIT_TARGET = re.compile(r'/inspect/damage/update_damage/')
MAIN_SUBMIT = re.compile(r'/inspect/inspect/update/')
OPINION_KEY = '檢測員意見'

OPINION_JS = r"""((text) => {
  const box = document.getElementById('comment');
  if (!box || !box.form) return;
  box.value = text;
  box.textContent = text;
  box.dispatchEvent(new Event('change', {bubbles: true}));
})(%s)"""

OPINION_SUBMIT_JS = r"""(() => {
  const box = document.getElementById('comment');
  if (!box || !box.form) return;
  const buttons = Array.from(box.form.querySelectorAll('input[type=submit], button[type=submit]'))
    .filter(b => /送出修改內容/.test(b.value || b.textContent || ''));
  if (buttons.length !== 1) return;
  buttons[0].click();
})()"""


def page_comment(html):
    """缺失主頁面 #comment 的內容（None：頁面沒有這個欄位）。"""
    from ..site.pages import tree
    node = next(iter(tree(html).xpath('//textarea[@id="comment"]')), None)
    return None if node is None else normalize(node.text_content())


def normalize(text):
    return '\n'.join(line.rstrip() for line in str(text or '').replace('\r\n', '\n').strip().split('\n'))

FILL_JS = r"""((values) => {
  const form = document.getElementById('form1');
  if (!form) return;
  for (const [id, value] of Object.entries(values)) {
    const e = form.querySelector('[id="' + id + '"]');
    if (!e) continue;
    if (e.tagName === 'SELECT') {
      const options = Array.from(e.options).filter(o => o.value === value || o.textContent.trim() === value);
      if (options.length !== 1) continue;             // 找不到唯一選項：不改，讀回時會發現不符
      Array.from(e.options).forEach(o => { o.selected = false; o.removeAttribute('selected'); });
      options[0].selected = true;
      options[0].setAttribute('selected', 'selected');
    } else if (e.tagName === 'TEXTAREA') {
      e.value = value;
      e.textContent = value;
    } else {
      e.value = value;
      e.setAttribute('value', value);
    }
    e.dispatchEvent(new Event('change', {bubbles: true}));
  }
})(%s)"""

SUBMIT_JS = r"""(() => {
  const form = document.getElementById('form1');
  if (!form) return;
  const buttons = Array.from(form.querySelectorAll('input[type=submit], button[type=submit]'))
    .filter(b => /送出更新/.test(b.value || b.textContent || ''));
  if (buttons.length !== 1) { document.title = '__bridgefill_no_submit__'; return; }
  buttons[0].click();
})()"""


class FillError(RuntimeError):
    pass


class FillPage(SitePage):
    """填寫用隱藏頁：只允許開啟第二級編輯頁；送出時才允許一次表單送出。"""

    def __init__(self, profile, parent=None):
        super().__init__(profile, parent, quiet=True)
        self.armed = None           # 允許下一次表單送出的目標（正規表示式）
        self.submitted = []         # 實際送出的網址
        self.blocked = []           # 被擋下的導覽

    def acceptNavigationRequest(self, url, kind, main_frame):
        address = url.toString()
        if not main_frame:
            return True
        if kind == QWebEnginePage.NavigationType.NavigationTypeFormSubmitted:
            if self.armed is not None and self.armed.search(address):
                self.armed = None
                self.submitted.append(address)
                return True
            self.blocked.append(address)
            return False
        if self.submitted or EDIT_PAGE.search(address) or MAIN_PAGE.search(address) or address.startswith('about:'):
            return True             # 送出後網站自己導回的頁面照常顯示
        self.blocked.append(address)
        return False

    def javaScriptConfirm(self, origin, message):
        if self.armed is not None:
            self.dialogs.append(('確認（已接受）', message))
            return True
        return super().javaScriptConfirm(origin, message)


class SiteFiller(PageFetcher):
    def __init__(self, profile, parent=None, delay_ms=500, timeout_ms=30000, rewrite=None):
        super().__init__(profile, parent, delay_ms, timeout_ms, rewrite)
        self.page.deleteLater()
        self.page = FillPage(profile, parent)

    def read(self, url):
        form = parse_defect_form(self.get(url))
        return form, {k: str(form.fields.get(k, '') or '').strip() for k in FILL_FIELDS if k in form.fields}

    def fill(self, item, log):
        """填寫一筆；回傳結果說明。失敗時丟出 FillError（已記入 log）。"""
        def fail(message, status=FAILED):
            log.mark(item.key, status, item.target, message)
            item.status, item.message = status, message
            raise FillError(message)

        if not item.url or not EDIT_PAGE.search(item.url):
            fail('沒有第二級編輯頁網址，請重新掃描')
        try:
            form, live = self.read(item.url)
        except (FetchError, PermissionError) as error:
            fail(f'開啟編輯頁失敗：{error}')
        if form.method_id != item.method_id:
            fail(f'建議維修工法 ID 不符（網站 {form.method_id or "讀不到"}，清單 {item.method_id}），停止，未送出')
        present = tuple(live)                    # 只核對表單上實際有的欄位
        missing = [k for k in item.changes if k not in live]
        if missing:
            fail('表單上找不到要填寫的欄位（' + '、'.join(NAMES[k] for k in missing) + '），未送出')
        previous = log.get(item.key)
        if not compare(live, item.target, present):
            log.mark(item.key, DONE, item.target, '網站已是預定值，未送出')
            item.status, item.message = DONE, '網站已是預定值（' + ('核對上次送出成功' if previous.get('status') == PENDING
                                                         else '不需送出') + '）'
            return item.message
        if previous.get('status') == PENDING:
            fail('上次送出結果不明，且網站值與預定值不同；請人工確認，不會自動重送', PENDING)
        changed = compare(live, item.site, present)
        if changed:
            fail('網站資料在掃描後有變動（' + '、'.join(NAMES[k] for k in changed) + '），請重新掃描，未送出')

        self.page.dialogs.clear()
        self.page.blocked.clear()
        self.page.submitted.clear()
        self.page.runJavaScript(FILL_JS % json.dumps({k: str(v) for k, v in {k: item.target[k] for k in item.changes}.items()},
                                                     ensure_ascii=False))
        self.pause(300)
        snapshot = parse_defect_form(self.html())
        applied = {k: str(snapshot.fields.get(k, '') or '').strip() for k in present}
        wrong = compare(applied, item.target, present)
        if wrong:
            fail('填入後網頁值與預定值不同（' + '、'.join(NAMES[k] for k in wrong) + '），可能是下拉選單沒有該選項；未送出')
        if snapshot.method_id != item.method_id:
            fail('填寫期間頁面 ID 改變，未送出')

        log.mark(item.key, PENDING, item.target, '已送出，等待回讀核對')
        self.page.armed = SUBMIT_TARGET
        self.page.runJavaScript(SUBMIT_JS)
        finished, args = self.wait(self.page.loadFinished)
        self.page.armed = None
        notes = '；'.join(f'{kind}：{text}' for kind, text in self.page.dialogs)
        if not self.page.submitted:
            log.mark(item.key, FAILED, item.target, '沒有送出：' + (notes or '找不到唯一的「送出更新」按鈕或網站檢查未通過'))
            item.status = FAILED
            item.message = '沒有送出：' + (notes or '找不到唯一的「送出更新」按鈕或網站檢查未通過')
            raise FillError(item.message)
        if not finished or not args or not args[0]:
            fail('送出後沒有收到完整回應，結果不明；保留「待確認」，不會重送' + (f'（{notes}）' if notes else ''), PENDING)
        self.pause(self.delay_ms)
        try:
            _, saved = self.read(item.url)
        except (FetchError, PermissionError) as error:
            fail(f'送出後重新開啟失敗：{error}；保留「待確認」，不會重送', PENDING)
        wrong = compare(saved, item.target, present)
        if wrong:
            fail('送出後回讀不一致（' + '、'.join(NAMES[k] for k in wrong) + '）；保留「待確認」，請人工確認，不會重送'
                 + (f'（網站訊息：{notes}）' if notes else ''), PENDING)
        log.mark(item.key, DONE, item.target, '已送出並回讀核對一致' + (f'（網站訊息：{notes}）' if notes else ''))
        item.status, item.message = DONE, '已送出並回讀核對一致'
        return item.message

    def write_opinion(self, main_url, expected, text, log):
        """把檢測員意見寫入本期缺失主頁面並送出；送出後回讀核對。expected 為掃描時網站上的意見。"""
        text = normalize(text)

        def fail(message, status=FAILED):
            log.mark(OPINION_KEY, status, {'comment': text}, message)
            raise FillError(message)
        if not MAIN_PAGE.search(main_url or ''):
            fail('沒有本期缺失主頁面網址，請重新掃描')
        try:
            live = page_comment(self.get(main_url))
        except (FetchError, PermissionError) as error:
            fail(f'開啟缺失主頁面失敗：{error}')
        if live is None:
            fail('缺失主頁面找不到「檢測員意見」欄位，未送出')
        previous = log.get(OPINION_KEY)
        if live == text:
            log.mark(OPINION_KEY, DONE, {'comment': text}, '網站已是這段意見，未送出')
            return '網站已是這段意見（未送出）'
        if previous.get('status') == PENDING:
            fail('上次送出意見的結果不明，且網站內容不同；請人工確認，不會自動重送', PENDING)
        if live != normalize(expected):
            fail('網站上的檢測員意見在掃描後有變動，請重新掃描本期後再寫入，未送出')
        self.page.dialogs.clear()
        self.page.submitted.clear()
        self.page.runJavaScript(OPINION_JS % json.dumps(text, ensure_ascii=False))
        self.pause(300)
        if page_comment(self.html()) != text:
            fail('填入後讀回的意見與預定不同，未送出')
        log.mark(OPINION_KEY, PENDING, {'comment': text}, '已送出，等待回讀核對')
        self.page.armed = MAIN_SUBMIT
        self.page.runJavaScript(OPINION_SUBMIT_JS)
        finished, args = self.wait(self.page.loadFinished)
        self.page.armed = None
        notes = '；'.join(f'{kind}：{message}' for kind, message in self.page.dialogs)
        if not self.page.submitted:
            log.mark(OPINION_KEY, FAILED, {'comment': text}, '沒有送出：' + (notes or '找不到唯一的「送出修改內容」按鈕或網站檢查未通過'))
            raise FillError('沒有送出：' + (notes or '找不到唯一的「送出修改內容」按鈕或網站檢查未通過'))
        if not finished or not args or not args[0]:
            fail('送出後沒有收到完整回應，結果不明；保留「待確認」，不會重送', PENDING)
        self.pause(self.delay_ms)
        try:
            saved = page_comment(self.get(main_url))
        except (FetchError, PermissionError) as error:
            fail(f'送出後重新開啟失敗：{error}；保留「待確認」，不會重送', PENDING)
        if saved != text:
            fail('送出後回讀的意見不一致；保留「待確認」，請人工確認，不會重送' + (f'（網站訊息：{notes}）' if notes else ''),
                 PENDING)
        log.mark(OPINION_KEY, DONE, {'comment': text}, '已送出並回讀核對一致')
        return '已送出並回讀核對一致'
