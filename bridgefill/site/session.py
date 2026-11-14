"""內嵌瀏覽器的登入狀態判讀與日誌遮罩（純邏輯，不依賴 Qt）。

程式只讀取頁面結構判斷狀態，不讀取、不填寫帳號密碼，也不操作驗證碼。
"""
import json
import re
from urllib.parse import urlsplit, urlunsplit

BASE_URL = 'https://bms2.iot.gov.tw/bms3/'
LOGIN_URL = BASE_URL + 'public/public/login/'

# 在頁面中執行，只回傳結構資訊；不讀任何輸入欄位的值。
PROBE_JS = r"""
(() => {
  const text = el => (el && el.textContent || '').replace(/\s+/g, ' ').trim();
  const menu = document.querySelector('#menu');
  const top = menu ? Array.from(menu.children).map(li => text(li.querySelector(':scope > a'))).filter(Boolean) : [];
  const frames = Array.from(document.querySelectorAll('iframe')).map(f => f.src || '');
  const forms = Array.from(document.querySelectorAll('form')).map(f => f.getAttribute('action') || '');
  return JSON.stringify({
    url: location.href,
    title: document.title,
    hasMenu: !!menu,
    menuItems: top,
    hasLoginForm: forms.some(a => /\/login\/?$/.test(a)) || !!document.querySelector('input[type=password]'),
    hasRecaptcha: frames.some(s => /recaptcha/.test(s)) ||
      !!document.querySelector('.g-recaptcha, [id*=recaptcha], script[src*="recaptcha"]'),
    hasImageCaptcha: !!document.querySelector('#captcha_code, .captcha_box'),
  });
})()
"""

LOGGED_IN = '已登入'
LOGIN_PAGE = '登入頁'
OTHER = '其他頁面'
UNKNOWN = '無法判讀'


def clean_user_agent(agent):
    """移除 QtWebEngine 字樣，保留實際 Chromium 版本，避免網站把內嵌瀏覽器當成異常用戶端。"""
    return re.sub(r'\s*QtWebEngine/\S+', '', agent or '').strip()


def redact_url(url):
    """日誌只保留網址路徑，去掉查詢參數與片段，避免記到識別資訊。"""
    if not url:
        return ''
    parts = urlsplit(str(url))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, '', ''))


def parse_probe(raw):
    """把 PROBE_JS 的回傳值轉成 dict；格式不對時回傳空 dict。"""
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def classify(probe):
    """依頁面結構判斷登入狀態。有系統選單（#menu）才算已登入。"""
    if not probe:
        return UNKNOWN
    if probe.get('hasMenu') and not probe.get('hasLoginForm'):
        return LOGGED_IN
    if probe.get('hasLoginForm'):
        return LOGIN_PAGE
    return OTHER


def describe(probe):
    """產生給使用者看、也寫入日誌的一行摘要（不含任何帳密或 Cookie）。"""
    state = classify(probe)
    parts = [state, redact_url(probe.get('url', ''))]
    if state == LOGIN_PAGE:
        found = [name for key, name in (('hasRecaptcha', 'reCAPTCHA'), ('hasImageCaptcha', '圖片驗證碼'))
                 if probe.get(key)]
        parts.append('驗證：' + ('、'.join(found) if found else '未偵測到'))
    if state == LOGGED_IN:
        items = probe.get('menuItems') or []
        parts.append(f'選單 {len(items)} 項：' + '、'.join(items))
    return '｜'.join(p for p in parts if p)
