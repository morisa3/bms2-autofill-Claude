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
VERSIONS = {'跨水橋': WATER, '陸橋': LAND, '雲林縣': YUNLIN}
IMAGE_SUFFIXES = ('.jpg', '.jpeg', '.png', '.bmp', '.gif')
DATE = re.compile(r'(\d{4})[-/.:](\d{1,2})[-/.:](\d{1,2})')


def bridge_kind(river_cross, agency=''):
    """照片版本：管理機關為雲林縣政府 → 雲林縣；否則依是否為跨水橋（讀不到時預設跨水橋）。"""
    if '雲林' in str(agency or ''):
        return '雲林縣'
    return '陸橋' if str(river_cross or '').strip() == '否' else '跨水橋'


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
                    found.append({'date': iso(text(cells[col['拍照日期']])), 'title': title,
                                  'note': text(cells[col['照片說明']]) if '照片說明' in col and col['照片說明'] < len(cells) else ''})
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
