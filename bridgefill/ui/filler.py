"""第 6 步：在隱藏頁面依建議維修工法 ID 填寫第二級表單並送出（會修改網站資料）。

流程（每筆）：開啟第二級編輯頁 → 核對 ID 與網站現值 → 填入預定值 → 讀回確認 → 記錄「待確認」→
按「送出更新」→ 重新開啟回讀 → 一致才記為「已完成」。任何一步不符就停止該筆，不會自動重送。

安全限制：
- 只開啟第二級編輯頁（edit_damage）；只在按下送出的那一刻允許一次送往 update_damage 的表單送出。
- 其他導覽（刪除、新增等）一律擋下並記錄。
- 網站的確認視窗只在送出當下接受（例如「確定送出？」），其他時候一律取消；提示訊息記錄在結果中。
- 網頁內容一律以 toHtml 讀回（不靠腳本回傳值）。
"""
import base64
import json
import mimetypes
import re
from pathlib import Path
from urllib.parse import urljoin

from lxml import html as lxml_html

from PySide6.QtWebEngineCore import QWebEnginePage

from ..core.fill import DONE, FAILED, FILL_FIELDS, NAMES, PENDING, compare
from ..core.basic_info import FIELDS as BASIC_FIELDS
from ..core.bridge_photos import add_photo_url, current_page, iso, parse_photo_list, same_period
from ..site.pages import parse_defect_form, parse_inspection
from .browser import SitePage
from .fetcher import FetchError, PageFetcher

EDIT_PAGE = re.compile(r'/inspect/damage/edit_damage/der_id,\d+/')
MAIN_PAGE = re.compile(r'/inspect/inspect/der_content/inspect_id,\d+/')
SUBMIT_TARGET = re.compile(r'/inspect/damage/update_damage/')
MAIN_SUBMIT = re.compile(r'/inspect/inspect/update/')
PHOTO_PAGE = re.compile(r'/basic/inventory_photo/')          # 照片簿（列表、新增頁）
PHOTO_LIMIT = 10 * 1024 * 1024                              # 網站：單一檔案限制 10MB 以下
PHOTO_FORBIDDEN = re.compile(r'del|remove|edit|update', re.I)  # 照片簿中絕不開啟的頁面（刪除、修改）
OPINION_KEY = '檢測員意見'

OPINION_JS = r"""((text) => {
  const box = document.getElementById('comment');
  if (!box || !box.form) return;
  box.value = text;
  box.textContent = text;
  box.dispatchEvent(new Event('change', {bubbles: true}));
})(%s)"""

BASIC_JS = r"""((fields) => {
  const box = document.getElementById('comment');
  if (!box || !box.form) return;
  for (const [selector, [kind, value]] of Object.entries(fields)) {
    const e = box.form.querySelector(selector);
    if (!e) continue;
    if (kind === 'select') {
      const options = Array.from(e.options).filter(o => o.value === value || o.textContent.trim() === value);
      if (options.length !== 1) continue;
      Array.from(e.options).forEach(o => { o.selected = false; o.removeAttribute('selected'); });
      options[0].selected = true;
      options[0].setAttribute('selected', 'selected');
    } else if (kind === 'textarea') {
      e.value = value;
      e.textContent = value;
    } else {
      e.value = value;
      e.setAttribute('value', value);
    }
    e.dispatchEvent(new Event('change', {bubbles: true}));
  }
})(%s)"""

OPINION_SUBMIT_JS = r"""(() => {
  const box = document.getElementById('comment');
  if (!box || !box.form) return;
  const buttons = Array.from(box.form.querySelectorAll('input[type=submit], button[type=submit]'))
    .filter(b => /送出修改內容/.test(b.value || b.textContent || ''));
  if (buttons.length !== 1) return;
  buttons[0].click();
})()"""


# 照片簿新增表單：依欄名（拍照日期、照片標題、照片說明）找欄位填入，照片以 DataTransfer 放進檔案欄。
# 填入的欄位加上 data-bf 標記，供 toHtml 回讀核對（不靠腳本回傳值）。
PHOTO_JS = r"""((data) => {
  // 實際網頁（2026-11-17 新增照片視窗）：form#form1，表頭「上傳照片、照片預覽、拍照日期、*照片標題、照片說明」，
  // 預設 3 列 tr.add_tr（file[]、date[]、name[]、text[]），每列檔案必填。只留第 1 列填寫，其餘空列從畫面移除（不按網站的刪除鈕）。
  const forms = Array.from(document.forms).filter(f => f.querySelector('input[type=file]'));
  if (forms.length !== 1) { document.body.setAttribute('data-bf-photo', 'forms:' + forms.length); return; }
  const form = forms[0];
  const clean = s => (s || '').replace(/[\s*＊:：]/g, '');
  const rows = Array.from(form.querySelectorAll('tr')).filter(r => r.querySelector('input[type=file]') && !r.querySelector('tr'));   // 最內層的列（外層版面表格也含檔案欄）
  if (!rows.length) return;
  const head = Array.from(form.querySelectorAll('tr')).find(r => Array.from(r.children).some(c => clean(c.textContent) === '照片標題'));
  const row = rows[0];
  rows.slice(1).forEach(r => r.remove());
  const cell = (label) => {
    if (head) {
      const index = Array.from(head.children).findIndex(c => clean(c.textContent) === label);
      if (index >= 0 && row.children[index]) return row.children[index].querySelector('textarea, input:not([type=hidden]):not([type=file]):not([type=button]):not([type=submit])');
    }
    let scope = Array.from(form.querySelectorAll('th,td,label')).find(e => clean(e.textContent) === label);
    for (let i = 0; scope && i < 3; i++, scope = scope.parentElement) {
      const f = scope.querySelector('textarea, input:not([type=hidden]):not([type=file]):not([type=button]):not([type=submit])');
      if (f) return f;
    }
    return null;
  };
  for (const [label, value] of Object.entries(data.fields)) {
    const f = cell(label);
    if (!f) continue;
    f.value = value;
    if (f.tagName === 'TEXTAREA') f.textContent = value; else f.setAttribute('value', value);
    f.setAttribute('data-bf', label);
  }
  const files = row.querySelectorAll('input[type=file]');
  if (files.length !== 1) { form.setAttribute('data-bf-file', 'inputs:' + files.length); return; }
  const bytes = Uint8Array.from(atob(data.content), c => c.charCodeAt(0));
  const transfer = new DataTransfer();
  transfer.items.add(new File([bytes], data.name, {type: data.mime}));
  files[0].files = transfer.files;
  const all = form.querySelectorAll('input[type=file]');
  form.setAttribute('data-bf-file', files[0].files.length + ':' + (files[0].files[0] ? files[0].files[0].name : '') + ':rows' + all.length);
  form.setAttribute('data-bf-form', '1');
})(%s)"""

PHOTO_SUBMIT_JS = r"""(() => {
  const form = document.querySelector('form[data-bf-form]');
  if (!form) return;
  // 使用者提供（2026-11-17）：確定送出 //*[@id="form1"]/div/input
  const exact = document.evaluate('//*[@id="form1"]/div/input', document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
  if (exact.snapshotLength === 1) {
    const button = exact.snapshotItem(0);
    if (button.form === form && /確定送出/.test(button.value || '')) { button.click(); return; }
  }
  const buttons = Array.from(form.querySelectorAll('input[type=submit], button[type=submit], button:not([type])'));
  const named = buttons.filter(b => /送出|確定|儲存|上傳|新增/.test(b.value || b.textContent || ''));
  const chosen = buttons.length === 1 ? buttons : named;
  if (chosen.length !== 1) { form.setAttribute('data-bf-submit', 'buttons:' + buttons.length); return; }
  chosen[0].click();
})()"""


def photo_form(html):
    """回讀新增表單：(表單送出網址, {欄名: 值}, 檔案標記)。找不到唯一表單回傳 None。"""
    doc = lxml_html.fromstring(html or '<html></html>')
    forms = doc.xpath('//form[@data-bf-form]')
    if len(forms) != 1:
        return None
    form = forms[0]
    values = {}
    for node in form.xpath('.//*[@data-bf]'):
        values[node.get('data-bf')] = (node.text_content() if node.tag == 'textarea' else node.get('value') or '').strip()
    bridge = form.xpath('.//input[@name="inventory_id"]/@value')
    values['inventory_id'] = bridge[0].strip() if bridge else ''
    return form.get('action') or '', values, form.get('data-bf-file') or ''


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
        if (self.submitted or EDIT_PAGE.search(address) or MAIN_PAGE.search(address) or address.startswith('about:')
                or (PHOTO_PAGE.search(address) and not PHOTO_FORBIDDEN.search(address))):
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

    def main_state(self, html, keys):
        """缺失主頁面上的意見與基本資料（只取 keys）。"""
        header = parse_inspection(html).header
        state = {k: normalize(header.get(k, '')) for k in keys if k != '檢測員意見'}
        if '檢測員意見' in keys:
            state['檢測員意見'] = page_comment(html)
        return state

    def write_opinion(self, main_url, expected, text, log, basic=None, expected_basic=None):
        """把檢測員意見（及基本資料）寫入本期缺失主頁面並送出；送出後回讀核對。
        expected／expected_basic 為掃描時網站上的值：網站在掃描後被改過就停止，不覆寫。"""
        target = {'檢測員意見': normalize(text), **{k: normalize(v) for k, v in (basic or {}).items()}}
        before = {'檢測員意見': normalize(expected), **{k: normalize(v) for k, v in (expected_basic or {}).items()}}
        keys = tuple(target)

        def fail(message, status=FAILED):
            log.mark(OPINION_KEY, status, target, message)
            raise FillError(message)
        if not MAIN_PAGE.search(main_url or ''):
            fail('沒有本期缺失主頁面網址，請重新掃描')
        try:
            html = self.get(main_url)
        except (FetchError, PermissionError) as error:
            fail(f'開啟缺失主頁面失敗：{error}')
        if page_comment(html) is None:
            fail('缺失主頁面找不到「檢測員意見」欄位，未送出')
        live = self.main_state(html, keys)
        previous = log.get(OPINION_KEY)
        if live == target:
            log.mark(OPINION_KEY, DONE, target, '網站已是這些內容，未送出')
            return '網站已是這些內容（未送出）'
        if previous.get('status') == PENDING:
            fail('上次送出的結果不明，且網站內容不同；請人工確認，不會自動重送', PENDING)
        changed = [k for k in keys if k in before and live.get(k) != before[k]]
        if changed:
            fail('網站上的' + '、'.join(changed) + '在掃描後有變動，請重新掃描本期後再寫入，未送出')
        self.page.dialogs.clear()
        self.page.submitted.clear()
        self.page.runJavaScript(OPINION_JS % json.dumps(target['檢測員意見'], ensure_ascii=False))
        fields = {selector: (kind, target[name]) for name, selector, kind in BASIC_FIELDS
                  if name in target and target[name] != live.get(name)}
        if fields:
            self.page.runJavaScript(BASIC_JS % json.dumps(fields, ensure_ascii=False))
        self.pause(300)
        wrong = [k for k, v in self.main_state(self.html(), keys).items() if v != target[k]]
        if wrong:
            fail('填入後讀回的' + '、'.join(wrong) + '與預定不同（下拉選單可能沒有該選項），未送出')
        log.mark(OPINION_KEY, PENDING, target, '已送出，等待回讀核對')
        self.page.armed = MAIN_SUBMIT
        self.page.runJavaScript(OPINION_SUBMIT_JS)
        finished, args = self.wait(self.page.loadFinished)
        self.page.armed = None
        notes = '；'.join(f'{kind}：{message}' for kind, message in self.page.dialogs)
        if not self.page.submitted:
            log.mark(OPINION_KEY, FAILED, target, '沒有送出：' + (notes or '找不到唯一的「送出修改內容」按鈕或網站檢查未通過'))
            raise FillError('沒有送出：' + (notes or '找不到唯一的「送出修改內容」按鈕或網站檢查未通過'))
        if not finished or not args or not args[0]:
            fail('送出後沒有收到完整回應，結果不明；保留「待確認」，不會重送', PENDING)
        self.pause(self.delay_ms)
        try:
            saved = self.main_state(self.get(main_url), keys)
        except (FetchError, PermissionError) as error:
            fail(f'送出後重新開啟失敗：{error}；保留「待確認」，不會重送', PENDING)
        wrong = [k for k in keys if saved.get(k) != target[k]]
        if wrong:
            fail('送出後回讀的' + '、'.join(wrong) + '不一致；保留「待確認」，請人工確認，不會重送'
                 + (f'（網站訊息：{notes}）' if notes else ''), PENDING)
        log.mark(OPINION_KEY, DONE, target, '已送出並回讀核對一致')
        return '已送出並回讀核對一致'

    # ------------------------------------------------------------ 橋梁照片簿
    def photo_titles(self, list_url):
        return parse_photo_list(self.get(list_url))

    def upload_photo(self, list_url, slot, log, key=None):
        """上傳一張照片到照片簿：先核對照片簿第 1 頁沒有同標題同日期的照片（往年同標題的照片不算） → 開新增頁填入日期、標題、說明與照片 → 回讀 →
        記錄「待確認」→ 送出 → 重新讀照片簿，出現這張才記為「已完成」。不會刪除或修改既有照片，結果不明不重送。"""
        key = key or f'照片:{slot.title}'
        target = {'title': slot.title, 'date': iso(slot.date), 'path': slot.path}

        def fail(message, status=FAILED):
            log.mark(key, status, target, message)
            slot.status = message
            raise FillError(message)

        def uploaded(photos):
            return any(p['title'] == slot.title and (not target['date'] or p['date'] == target['date']) for p in photos)

        path = Path(slot.path or '')
        if not path.is_file():
            fail(f'找不到照片檔：{slot.path or "（未選擇）"}')
        if path.stat().st_size >= PHOTO_LIMIT:
            fail(f'照片超過網站限制 10MB（{path.stat().st_size / 1048576:.1f}MB），請先縮小再上傳')
        if not target['date']:
            fail('拍照日期格式不正確（YYYY-MM-DD）')
        if not PHOTO_PAGE.search(list_url or ''):
            fail('沒有照片簿網址，請先掃描本期（需要橋梁系統ID）')
        try:
            html = self.get(list_url)
        except (FetchError, PermissionError) as error:
            fail(f'開啟照片簿失敗：{error}')
        if current_page(html) != 1:
            fail('照片簿目前不在第 1 頁（網站記住了上次看的頁次），無法核對最新照片；請在網站照片簿切回第 1 頁後再上傳')
        photos = parse_photo_list(html)
        previous = log.get(key)
        if uploaded(photos):
            log.mark(key, DONE, target, '照片簿已有這張（同標題與日期），未上傳')
            slot.status = '網站已有' if previous.get('status') != PENDING else '已上傳（核對上次送出成功）'
            return slot.status
        if previous.get('status') == PENDING:
            fail('上次上傳結果不明，照片簿也沒有這張；請到網站確認，不會自動重傳', PENDING)
        near = same_period(photos, slot.title, target['date'])
        if near:
            fail(f"照片簿已有 {near[0]['date']} 的「{slot.title}」，看起來本期已上傳過；如要更換請先在網站刪除舊照片"
                 '（程式不刪除照片），或改正拍照日期')
        if 'type="file"' not in html.replace("'", '"'):
            add = add_photo_url(html)
            if not add or PHOTO_FORBIDDEN.search(add) or not PHOTO_PAGE.search(add):
                fail('照片簿找不到「新增圖片」，未上傳')
            try:
                _, html = self.load(add)
            except FetchError as error:
                fail(f'開啟新增圖片頁失敗：{error}')
            if self.page.blocked:
                fail('新增圖片頁被擋下（網址不在允許範圍）：' + self.page.blocked[-1])
        self.page.dialogs.clear()
        self.page.blocked.clear()
        self.page.submitted.clear()
        data = {'fields': {'拍照日期': target['date'], '照片標題': slot.title, '照片說明': slot.note or ''},
                'content': base64.b64encode(path.read_bytes()).decode('ascii'), 'name': path.name,
                'mime': mimetypes.guess_type(path.name)[0] or 'image/jpeg'}
        self.page.runJavaScript(PHOTO_JS % json.dumps(data, ensure_ascii=False))
        self.pause(500)
        form = photo_form(self.html())
        if form is None:
            fail('新增圖片頁找不到唯一的上傳表單（含選擇檔案），未上傳；請把新增圖片頁存成 HTML 提供給我調整')
        action, values, marker = form
        missing = [k for k in ('拍照日期', '照片標題') if values.get(k) != data['fields'][k]]
        if missing:
            fail('新增圖片頁找不到或填不進「' + '、'.join(missing) + '」欄位，未上傳')
        if not marker.startswith('1:'):
            fail(f'照片沒有放進檔案欄（{marker or "無法設定"}），未上傳')
        expected = re.search(r'inventory_id,(\d+)', list_url)
        if values.get('inventory_id') and expected and values['inventory_id'] != expected[1]:
            fail(f"新增照片視窗的橋梁系統ID（{values['inventory_id']}）與照片簿（{expected[1]}）不同，未上傳")
        if not marker.endswith(':rows1'):
            fail(f'新增照片視窗有多列上傳欄位無法整理成一列（{marker}），未上傳')
        address = urljoin(self.page.url().toString(), action) if action else self.page.url().toString()
        path_only = re.sub(r'^https?://[^/]+', '', address)
        if not PHOTO_PAGE.search(path_only) or PHOTO_FORBIDDEN.search(path_only):
            fail(f'上傳表單的送出網址不在照片簿範圍（{path_only}），未上傳')
        log.mark(key, PENDING, target, '已送出，等待回讀核對')
        self.page.armed = re.compile(re.escape(path_only.split('?')[0]))
        self.page.runJavaScript(PHOTO_SUBMIT_JS)
        finished, args = self.wait(self.page.loadFinished)
        self.page.armed = None
        notes = '；'.join(f'{kind}：{text}' for kind, text in self.page.dialogs)
        if not self.page.submitted:
            log.mark(key, FAILED, target, '沒有送出：' + (notes or '找不到唯一的送出按鈕或網站檢查未通過'))
            slot.status = '沒有送出：' + (notes or '找不到唯一的送出按鈕或網站檢查未通過')
            raise FillError(slot.status)
        if not finished or not args or not args[0]:
            fail('送出後沒有收到完整回應，結果不明；保留「待確認」，不會重傳', PENDING)
        self.pause(self.delay_ms)
        try:
            photos = parse_photo_list(self.get(list_url))
        except (FetchError, PermissionError) as error:
            fail(f'送出後重新開啟照片簿失敗：{error}；保留「待確認」，不會重傳', PENDING)
        if not uploaded(photos):
            fail('送出後照片簿沒有出現這張；保留「待確認」，請到網站確認，不會重傳' + (f'（網站訊息：{notes}）' if notes else ''),
                 PENDING)
        log.mark(key, DONE, target, '已上傳並在照片簿核對到' + (f'（網站訊息：{notes}）' if notes else ''))
        slot.status = '已上傳'
        return slot.status
