"""橋梁照片簿：六張全景照片（2026-11-17 使用者需求）。

- 依基本資料「是否為跨水橋」（#river_cross）決定六個照片名稱：是 → 跨水橋版本，否 → 陸橋版本；
  管理機關為雲林縣政府時用雲林縣版本（2026-11-19 使用者指定）。
- 拍照日期自動產生可選清單：照片本身的拍攝日期（EXIF）、本期檢測日期與前後幾天、今天。
- 網站照片簿（/basic/inventory_photo/list/）已有相同標題的照片就不重複上傳。
本模組不依賴 Qt。
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from lxml import html as lxml_html

WATER = ('橋頭近照', '橋尾遠照', '往外上游', '往外下游', '上游側照', '下游側照')
LAND = ('橋頭近端', '橋尾遠端', '右側側照', '左側側照', '往外右側', '往外左側')
YUNLIN = ('橋頭近端', '橋尾遠端', '橋左側照', '橋右側照', '往外上游', '往外下游')    # 雲林縣政府
YUNLIN_LAND = ('橋頭近端', '橋尾遠端', '橋左側照', '橋右側照', '往外右側', '往外左側')  # 雲林縣陸橋（2026-11-27 使用者）
BUILT_IN = {'跨水橋': WATER, '陸橋': LAND, '雲林縣': YUNLIN, '雲林縣陸橋': YUNLIN_LAND}
VERSIONS = dict(BUILT_IN)          # 另含各縣市政府自訂的版本（「縣市政府－跨水橋／陸橋」，apply_custom 加入）
CUSTOM = {}                        # {縣市政府: {'跨水橋': (六個標題), '陸橋': (六個標題)}}
CUSTOM_KINDS = ('跨水橋', '陸橋')


def government(text):
    """管理機關或所在地 → 縣市政府名稱（雲林縣莿桐鄉公所、雲林縣 → 雲林縣政府）；判斷不出縣市時原樣回傳。"""
    text = str(text or '').strip()
    found = re.match(r'^(\S{2}[縣市])', text)
    return found[1].replace('台', '臺') + '政府' if found else text


def custom_kind(gov, kind):
    return f'{gov}－{kind}'


def apply_custom(custom):
    """各縣市政府自訂的照片標題（2026-11-27 使用者需求）：加入 VERSIONS；空白或不足六個的不採用。"""
    CUSTOM.clear()
    for key in [k for k in VERSIONS if k not in BUILT_IN]:
        VERSIONS.pop(key)
    for gov, sets in (custom or {}).items():
        gov = government(gov)
        clean = {kind: tuple(str(t).strip() for t in titles) for kind, titles in (sets or {}).items()
                 if kind in CUSTOM_KINDS and len([t for t in titles if str(t).strip()]) == 6}
        if gov and clean:
            CUSTOM[gov] = clean
            for kind, titles in clean.items():
                VERSIONS[custom_kind(gov, kind)] = titles
    return dict(CUSTOM)


def load_custom(path):
    import json
    try:
        return apply_custom(json.loads(Path(path).read_text(encoding='utf-8')))
    except (OSError, ValueError, AttributeError):
        return apply_custom({})


def save_custom(path, custom):
    import json
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {gov: {k: list(v) for k, v in sets.items()} for gov, sets in (custom or {}).items()}
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return apply_custom(data)
IMAGE_SUFFIXES = ('.jpg', '.jpeg', '.png', '.bmp', '.gif')
DATE = re.compile(r'(\d{4})[-/.:](\d{1,2})[-/.:](\d{1,2})')


def bridge_kind(river_cross, agency=''):
    """照片版本：管理機關所屬縣市政府有自訂標題 → 用自訂（跨水橋／陸橋）；雲林縣 → 雲林縣（陸橋用雲林縣陸橋）；
    其他依是否為跨水橋（讀不到時預設跨水橋）。"""
    land = str(river_cross or '').strip() == '否'
    sets = CUSTOM.get(government(agency), {})
    kind = '陸橋' if land else '跨水橋'
    if kind in sets:
        return custom_kind(government(agency), kind)
    if '雲林' in str(agency or ''):
        return '雲林縣陸橋' if land else '雲林縣'
    return kind


def slot_names(river_cross, agency='', kind=''):
    """六個照片名稱；kind 指定版本（跨水橋／陸橋／雲林縣）時以 kind 為準。"""
    return VERSIONS.get(kind) or VERSIONS[bridge_kind(river_cross, agency)]


def is_image(path):
    return Path(path).suffix.lower() in IMAGE_SUFFIXES


def iso(value):
    """各種日期寫法 → YYYY-MM-DD；不是日期回傳空白。"""
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    found = DATE.search(str(value or ''))
    if not found:
        return ''
    try:
        return date(int(found[1]), int(found[2]), int(found[3])).isoformat()
    except ValueError:
        return ''


def exif_date(path):
    """照片的拍攝日期（EXIF DateTimeOriginal，格式 YYYY:MM:DD HH:MM:SS）；沒有回傳空白。不需額外套件。"""
    try:
        with open(path, 'rb') as handle:
            head = handle.read(256 * 1024)
    except OSError:
        return ''
    if not head.startswith(b'\xff\xd8'):
        return ''
    found = re.search(rb'(\d{4}):(\d{2}):(\d{2}) \d{2}:\d{2}:\d{2}', head)
    return iso(found[0].decode('ascii')[:10].replace(':', '-')) if found else ''


def file_date(path):
    try:
        return datetime.fromtimestamp(Path(path).stat().st_mtime).strftime('%Y-%m-%d')
    except OSError:
        return ''


def date_choices(inspection_date='', photo_dates=(), today=None, around=3):
    """可選日期：照片拍攝日期 → 本期檢測日期 → 檢測日期前後 around 天 → 今天（不重複）。"""
    today = today or date.today()
    choices = [iso(d) for d in photo_dates if iso(d)]
    base = iso(inspection_date)
    if base:
        choices.append(base)
        day = date.fromisoformat(base)
        for offset in range(1, around + 1):
            choices += [(day - timedelta(days=offset)).isoformat(), (day + timedelta(days=offset)).isoformat()]
    choices.append(today.isoformat())
    return list(dict.fromkeys(choices))


def default_date(path, inspection_date=''):
    """預設拍照日期：照片的 EXIF 拍攝日期，沒有就用本期檢測日期。"""
    return exif_date(path) or iso(inspection_date) or file_date(path)


@dataclass
class Slot:
    title: str
    path: str = ''
    date: str = ''
    note: str = ''
    status: str = ''          # 已上傳／網站已有／失敗訊息

    @property
    def ready(self):
        return bool(self.path and self.title and iso(self.date))


@dataclass
class PhotoPlan:
    river_cross: str = ''
    slots: list = field(default_factory=list)
    kind: str = ''                 # 跨水橋／陸橋／雲林縣

    @classmethod
    def create(cls, river_cross='', previous=None, agency='', kind=''):
        """建立六格（版本依 kind，或依管理機關與是否為跨水橋）；previous 為上次的安排：
        同名的格子沿用照片與日期，名稱改了的依位置沿用（橋頭、橋尾…順序相同）。"""
        kind = kind if kind in VERSIONS else bridge_kind(river_cross, agency)
        names = VERSIONS[kind]
        before = previous.slots if previous else []
        old = {s.title: s for s in before}
        slots = []
        for index, name in enumerate(names):
            if name in old:
                slots.append(old[name])
            elif index < len(before) and before[index].title not in names:
                s = before[index]
                slots.append(Slot(name, s.path, s.date, s.note))
            else:
                slots.append(Slot(name))
        return cls(river_cross, slots, kind)

    def to_dict(self):
        return {'river_cross': self.river_cross, 'kind': self.kind,
                'slots': [dict(title=s.title, path=s.path, date=s.date, note=s.note, status=s.status) for s in self.slots]}

    @classmethod
    def from_dict(cls, data):
        slots = [Slot(**{k: s.get(k, '') for k in ('title', 'path', 'date', 'note', 'status')}) for s in data.get('slots', [])]
        kind = data.get('kind') or next((k for k, names in VERSIONS.items() if [s.title for s in slots] == list(names)), '')
        return cls(data.get('river_cross', ''), slots, kind)

    def place(self, paths, inspection_date='', start=0):
        """把多張照片依序放進空的格子（從 start 開始）；回傳沒有位置放的照片。"""
        paths = [p for p in paths if is_image(p)]
        empty = [i for i in range(start, len(self.slots)) if not self.slots[i].path] + \
                [i for i in range(0, start) if not self.slots[i].path]
        if len(paths) == 1 and start < len(self.slots):
            empty = [start]                      # 單張：放在拖入的那一格（取代原照片）
        for index, path in zip(empty, paths):
            slot = self.slots[index]
            slot.path, slot.date, slot.status = str(path), default_date(path, inspection_date), ''
        return paths[len(empty):]

    def swap(self, a, b):
        first, second = self.slots[a], self.slots[b]
        first.path, second.path = second.path, first.path
        first.date, second.date = second.date, first.date
        first.note, second.note = second.note, first.note
        first.status = second.status = ''


# ---------------------------------------------------------------- 網站照片簿

def photo_list_url(inventory_id):
    return f'https://bms2.iot.gov.tw/bms3/basic/inventory_photo/list/inventory_id,{inventory_id}/'


def text(node):
    return re.sub(r'\s+', ' ', node.text_content() if node is not None else '').strip()


def parse_photo_list(source):
    """照片簿列表：依表頭「拍照日期、照片標題、照片說明」讀出已上傳的照片 [{date, title, note}]。"""
    doc = lxml_html.fromstring(source or '<html></html>')
    for table in doc.xpath('//table'):
        rows = table.xpath('./tr|./tbody/tr|./thead/tr')
        for y, row in enumerate(rows):
            heads = [text(c) for c in row.xpath('./th|./td')]
            if '照片標題' in heads and '拍照日期' in heads:
                col = {h: i for i, h in enumerate(heads)}
                found = []
                for line in rows[y + 1:]:
                    cells = line.xpath('./th|./td')
                    if len(cells) <= max(col['照片標題'], col['拍照日期']):
                        continue
                    title = text(cells[col['照片標題']])
                    if not title:
                        continue
                    links = [a.get('href') or '' for a in line.xpath('.//a[@href]')] + \
                        [m for b in line.xpath('.//*[@onclick]')
                         for m in re.findall(r"['\"]([^'\"]*inventory_photo/edit/\d+[^'\"]*)['\"]", b.get('onclick') or '')]
                    edit = next((u for u in links if re.search(r'/inventory_photo/edit/\d+', u)), '')
                    if edit and not edit.startswith('http'):
                        edit = 'https://bms2.iot.gov.tw' + ('' if edit.startswith('/') else '/') + edit
                    ident = re.search(r'/inventory_photo/edit/(\d+)', edit)
                    found.append({'date': iso(text(cells[col['拍照日期']])), 'title': title,
                                  'note': text(cells[col['照片說明']]) if '照片說明' in col and col['照片說明'] < len(cells) else '',
                                  'id': ident[1] if ident else '', 'edit_url': edit})
                return found
    return []


def current_page(source):
    """照片簿分頁目前在第幾頁（沒有分頁回傳 1）。照片依新到舊排列，核對只看第 1 頁。"""
    doc = lxml_html.fromstring(source or '<html></html>')
    found = doc.xpath('//*[contains(concat(" ", normalize-space(@class), " "), " current ")]')
    numbers = [text(n) for n in found if text(n).isdigit()]
    return int(numbers[0]) if numbers else 1


def same_period(photos, title, day, days=60):
    """照片簿中同標題、日期相近（前後 days 天）但不同天的照片——多半是本期已上傳過；照片簿也有往年的照片，同標題不一定重複。"""
    try:
        target = date.fromisoformat(iso(day))
    except ValueError:
        return []
    near = []
    for photo in photos:
        if photo['title'] != title or not photo['date'] or photo['date'] == target.isoformat():
            continue
        if abs((date.fromisoformat(photo['date']) - target).days) <= days:
            near.append(photo)
    return near


def latest_dates(photos):
    """照片簿各標題最新的拍照日期 {標題: YYYY-MM-DD}，以及全部照片中最新的日期。"""
    latest = {}
    for photo in photos:
        if photo.get('date') and photo['date'] > latest.get(photo['title'], ''):
            latest[photo['title']] = photo['date']
    return latest, max(latest.values(), default='')


def mark_uploaded(plan, photos, year):
    """照片簿已有本年度（year）同標題照片的格子標「網站已有」（不再勾選上傳）；回傳標記的標題。
    已上傳、網站已有的格子不變。"""
    latest, _ = latest_dates(photos)
    marked = []
    for slot in plan.slots:
        day = latest.get(slot.title, '')
        if day.startswith(str(year)) and not slot.status.startswith(('已上傳', '網站已有')):
            slot.status = f'網站已有（照片簿 {day}）'
            marked.append(slot.title)
    version = uploaded_version(photos, year) or ('混用版本' if mixed_six(photos, year) else '')
    if version:                       # 本年度已依其他版本（或混用）上傳六張（例：雲林縣的橋用跨水橋標題上傳）：六格都算已有
        for slot in plan.slots:
            if not slot.status.startswith(('已上傳', '網站已有')):
                slot.status = f'網站已有（照片簿 {year} 年已有{version}格式六張）'
                marked.append(slot.title)
    return marked


HEAD_TITLES, TAIL_TITLES = ('橋頭近照', '橋頭近端'), ('橋尾遠照', '橋尾遠端')


def year_photos(photos, year):
    """本年度照片，同標題只留最新一張（照片簿新到舊）。"""
    seen, result = set(), []
    for photo in photos:
        key = title_key(photo['title'])
        if str(photo.get('date', '')).startswith(str(year)) and key not in seen:
            seen.add(key)
            result.append(photo)
    return result


def known_titles():
    return {title_key(n) for names in VERSIONS.values() for n in names}


def mixed_six(photos, year):
    """本年度有六個不同的基本照片標題（可混用各版本，例：雲林縣的橋部分用跨水橋標題），且橋頭、橋尾都有。"""
    titles = {title_key(p['title']) for p in year_photos(photos, year)} & known_titles()
    has = lambda names: any(title_key(n) in titles for n in names)
    return len(titles) >= 6 and has(HEAD_TITLES) and has(TAIL_TITLES)


def rename_plan(photos, year, target):
    """本年度照片標題不是 target 版本的（其他版本或混用）：改成 target 標題的建議對照（2026-11-26 使用者需求）。
    橋頭近照↔橋頭近端、橋尾遠照↔橋尾遠端直接對應；其餘依序補到 target 缺少的標題，使用者可在確認視窗改選。
    回傳 [{id, edit_url, date, old, new}]；不需修正回傳 []。"""
    if target not in VERSIONS:
        return []
    wanted = VERSIONS[target]
    present = year_photos(photos, year)
    have = {title_key(p['title']) for p in present}
    missing = [n for n in wanted if title_key(n) not in have]
    others = [p for p in present if title_key(p['title']) not in {title_key(n) for n in wanted}
              and title_key(p['title']) in known_titles() and p.get('id')]
    plan = []
    for group in (HEAD_TITLES, TAIL_TITLES):              # 橋頭、橋尾：意思相同直接對應
        goal = next((n for n in missing if n in group), None)
        photo = next((p for p in others if p['title'] in group), None)
        if goal and photo:
            plan.append((photo, goal))
            missing.remove(goal)
            others.remove(photo)
    plan += list(zip(others, missing))                    # 其餘依序補缺少的標題（使用者確認時可改選）
    return [{'id': p['id'], 'edit_url': p.get('edit_url', ''), 'date': p['date'], 'old': p['title'], 'new': new}
            for p, new in plan]


def parse_photo_edit(source):
    """照片「修改」頁（2026-11-26 實際網頁：form#form1 送到 /basic/inventory_photo/update/，欄位 id、inventory_id、
    date、name（照片標題）、text（照片說明），唯一的「確定送出」）。回傳 {action, id, inventory_id, date, name}；不符回傳 None。"""
    doc = lxml_html.fromstring(source or '<html></html>')
    forms = [f for f in doc.xpath('//form') if '/basic/inventory_photo/update/' in (f.get('action') or '')]
    if len(forms) != 1:
        return None
    form = forms[0]

    def field(name):
        node = form.xpath(f'.//input[@name="{name}"]')
        return (node[0].get('value') or '').strip() if len(node) == 1 else None
    submits = [b for b in form.xpath('.//input[@type="submit"]|.//button[@type="submit"]')
               if '確定送出' in ((b.get('value') or '') + text(b))]
    values = {k: field(k) for k in ('id', 'inventory_id', 'date', 'name')}
    if None in values.values() or len(submits) != 1 or form.xpath('.//input[@type="file"]'):
        return None
    return {'action': form.get('action'), **values}


def uploaded_version(photos, year):
    """本年度照片簿中標題完整的版本（跨水橋／陸橋／雲林縣）；沒有回傳空白（2026-11-26 使用者需求）。"""
    titles = {title_key(p['title']) for p in photos if str(p.get('date', '')).startswith(str(year))}
    return next((kind for kind, names in VERSIONS.items() if all(title_key(n) in titles for n in names)), '')


def title_key(text):
    """照片標題比對：去空白、全形轉半形。"""
    import unicodedata
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(text or '')))


def six_uploaded(photos, year):
    """本年度六張基本照片已上傳：任一版本六個標題都在，或六個標題混用各版本但都是基本照片標題（2026-11-26）。"""
    return bool(uploaded_version(photos, year)) or mixed_six(photos, year)


def add_photo_url(source, base='https://bms2.iot.gov.tw'):
    """照片簿「新增圖片」按鈕的網址（onclick location 或連結）；找不到回傳空白。"""
    doc = lxml_html.fromstring(source or '<html></html>')
    for node in doc.xpath('//input[@type="button" or @type="submit"]|//button|//a'):
        label = (node.get('value') or '') + text(node)
        if '新增' not in label:
            continue
        target = node.get('href') or ''
        script = node.get('onclick') or ''
        # 實際網站：onclick="easy_popup('/bms3/basic/inventory_photo/add/inventory_id,31858/_popup,1/_close,2/','1030,600')"
        found = re.search(r"['\"]([^'\"]*inventory_photo[^'\"]*)['\"]", script)
        url = found[1] if found else ('' if target.startswith(('#', 'javascript')) else target)
        if url and 'inventory_photo' in url:
            return url if url.startswith('http') else base.rstrip('/') + ('' if url.startswith('/') else '/') + url
    return ''
