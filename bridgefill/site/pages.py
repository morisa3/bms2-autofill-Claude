"""bms2 網頁解析：輸入網頁 HTML，輸出結構化資料（唯讀，不依賴 Qt）。

一律依表頭文字、元素 id 與連結參數定位，不寫死欄位或列號：
- 有無「狀態」欄會讓欄位整欄右移，所以先讀表頭再對欄。
- 每列構件以「詳細資料」連結上的 component_form／component_id／inspect_id 識別，不依列的順序。
- 每筆缺失以「建議維修工法 ID」識別。
"""
import re
from dataclasses import dataclass, field
from urllib.parse import urljoin
from decimal import Decimal, InvalidOperation

from lxml import html as lxml_html

from ..core.config import COMPONENTS, COMPONENT_SECTIONS
from ..core.deru import deru_records

REPAIR_STATES = ('已維修', '部分未維修', '未維修')
DEFECT, GOOD, UNABLE, UNKNOWN = '缺失', '良好', '無法檢測', '無法判讀'


def text(node):
    return ' '.join((node.text_content() if node is not None else '').split())


def tree(source):
    return lxml_html.fromstring(source) if isinstance(source, (str, bytes)) else source


SITE_BASE = 'https://bms2.iot.gov.tw/bms3/'


def page_base(doc):
    """網頁以 <base href> 決定相對連結的基準；實際網頁的連結常寫成 inspect/…（沒有斜線）或 /bms3/…。"""
    found = doc.xpath('//base/@href') if hasattr(doc, 'xpath') else []
    return found[0] if found else SITE_BASE


def resolve(href, base=SITE_BASE):
    """把網頁上的連結補成完整網址；javascript: 等非網址原樣保留。"""
    href = (href or '').strip()
    if not href or re.match(r'^(javascript|mailto|#)', href, re.I):
        return href
    return urljoin(base or SITE_BASE, href)


def url_params(url):
    """bms2 網址以 `key,value/` 帶參數，例：component_form,abutment/component_id,123/。"""
    return dict(re.findall(r'([a-z_]+),([^/]*)/', url or ''))


def number(value):
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None


# ---------------------------------------------------------------- 表格

def grid(table):
    """把表格展開成二維格子，處理 rowspan／colspan，只取這張表本身的列（不含巢狀表格）。"""
    rows = [tr for tr in table.iter('tr') if tr.xpath('ancestor::table[1]')[0] is table]
    cells = {}
    for y, row in enumerate(rows):
        x = 0
        for cell in row.xpath('./th|./td'):
            while (y, x) in cells:
                x += 1
            height, width = int(cell.get('rowspan') or 1), int(cell.get('colspan') or 1)
            for dy in range(height):
                for dx in range(width):
                    cells[(y + dy, x + dx)] = cell
            x += width
    width = max((x for _, x in cells), default=-1) + 1
    return rows, [[cells.get((y, x)) for x in range(width)] for y in range(len(rows))]


def header_key(value):
    return re.sub(r'\s+', '', value or '').upper()


# ---------------------------------------------------------------- 缺失主頁面

@dataclass
class ComponentRow:
    component: str            # 構件中文名稱
    table_id: str
    columns: dict             # 表頭 → 值（不含狀態、DERU、動作、維修狀態）
    deru: str
    status: str               # 缺失／良好／無法檢測／無法判讀
    count: int | None         # 詳細資料(n) 的 n
    detail_url: str
    keys: dict                # component_form、component_id、inspect_id、inventory_id
    repair: str = ''          # 已維修／部分未維修／未維修／空白
    problem: str = ''

    @property
    def identity(self):
        return (self.keys.get('component_form', self.table_id), self.keys.get('component_id', ''))

    @property
    def label(self):
        return '｜'.join(f'{k}：{v}' for k, v in self.columns.items() if re.search(r'位置|編號', k) and v)


@dataclass
class InspectionPage:
    header: dict
    rows: list = field(default_factory=list)
    problems: list = field(default_factory=list)


def deru_status(value):
    records = deru_records(value)
    if not records:
        return UNKNOWN
    d, e = records[0][0], records[0][1]
    if e == '0':
        return UNABLE
    if d == '1':
        return GOOD
    return DEFECT if d.isdigit() and e.isdigit() else UNKNOWN


def component_name(table, table_id):
    """table id 對應構件；鋼纜等共用 id 的表格以前方區段標題區分。"""
    names = [name for name, identity in COMPONENTS.items() if identity == table_id]
    if len(names) <= 1:
        return names[0] if names else table_id
    marks = table.xpath('preceding::*[starts-with(@id,"x_")][1]/@id')
    for name in names:
        if marks and marks[0] in COMPONENT_SECTIONS.get(name, ()):
            return name
    return ''


def parse_component_table(table, component, base=SITE_BASE):
    rows, cells = grid(table)
    found, problems = [], []
    headers, deru_col, action_col = [], -1, -1
    for y, row in enumerate(rows):
        line = [header_key(text(c)) if c is not None else '' for c in cells[y]]
        if any(v in ('DERU', 'DREU') for v in line):
            headers = [text(c) if c is not None else '' for c in cells[y]]
            deru_col = next(i for i, v in enumerate(line) if v in ('DERU', 'DREU'))
            action_col = next((i for i, v in enumerate(line) if v in ('動作', '操作')), -1)
            continue
        links = [a for a in row.xpath('.//a') if '詳細資料' in text(a) and a.xpath('ancestor::table[1]')[0] is table]
        if not links:
            continue
        if not headers:
            problems.append(f'{component}：找不到 DERU 表頭')
            continue
        link = links[0]
        problem = '同列有多個詳細資料連結' if len(links) > 1 else ''
        if action_col >= 0 and cells[y][action_col] is not None and link not in cells[y][action_col].iter('a'):
            problem = problem or '詳細資料不在動作欄'
        columns, repair = {}, ''
        for x, cell in enumerate(cells[y]):
            name = headers[x] if x < len(headers) else ''
            if x in (deru_col, action_col) or not name or cell is None:
                continue
            value = text(cell)
            if name == '維修狀態':
                repair = value if value in REPAIR_STATES else ''
            elif name != '狀態':
                columns[name] = value
        deru = text(cells[y][deru_col]) if deru_col < len(cells[y]) else ''
        count = re.search(r'詳細資料\s*[（(]\s*(\d+)\s*[)）]', text(link))
        url = resolve(link.get('href'), base)
        found.append(ComponentRow(component, table.get('id'), columns, deru, deru_status(deru),
                                  int(count[1]) if count else None, url,
                                  {k: v for k, v in url_params(url).items()
                                   if k in ('component_form', 'component_id', 'inspect_id', 'inventory_id')},
                                  repair, problem))
    return found, problems


def header_value(doc, label):
    for th in doc.xpath(f'//th[normalize-space(.)="{label}"]'):
        cell = th.getnext()
        if cell is None:
            continue
        field_ = cell.xpath('.//input[@type!="hidden" or not(@type)]')
        return (field_[0].get('value') or '').strip() if field_ else text(cell)
    return ''


def selected(doc, xpath):
    options = doc.xpath(xpath + '//option[@selected]')
    return text(options[0]) if options else ''


def parse_inspection(source):
    """缺失主頁面（前期或本期皆同）。"""
    doc = tree(source)
    header = {label: header_value(doc, label) for label in ('橋梁名稱', '橋梁編號', '檢測日期', '道路等級', '總橋孔數')}
    header.update({
        '單位主管': (doc.xpath('//*[@id="chief"]/@value') or [''])[0],
        '近端定義': selected(doc, '//select[@name="frontside"]'),
        '是否疑似鹽害橋梁': selected(doc, '//*[@id="salt_injury"]'),
        '有無沖刷現象': selected(doc, '//*[@id="scour_type"]'),
        '沖刷位置/備註': text(next(iter(doc.xpath('//*[@id="scour_site"]')), None)),
        '橋下有無堆放雜物': selected(doc, '//*[@id="waste"]'),
        '檢測員意見': next(iter(doc.xpath('//*[@id="comment"]')), None),
    })
    header['選項'] = {name: [text(o) for o in doc.xpath(xpath + '//option') if text(o)]
                      for name, xpath in (('近端定義', '//select[@name="frontside"]'), ('是否疑似鹽害橋梁', '//*[@id="salt_injury"]'),
                                          ('有無沖刷現象', '//*[@id="scour_type"]'), ('橋下有無堆放雜物', '//*[@id="waste"]'))}
    comment = header['檢測員意見']
    header['檢測員意見'] = comment.text_content() if comment is not None else ''
    page = InspectionPage(header)
    for table in doc.xpath('//table[@id]'):
        if table.get('id') not in set(COMPONENTS.values()):
            continue
        name = component_name(table, table.get('id'))
        if not name:
            page.problems.append(f'{table.get("id")}：共用表格找不到區段標題，未讀取')
            continue
        rows, problems = parse_component_table(table, name, page_base(doc))
        page.rows += rows
        page.problems += problems
    return page


# ---------------------------------------------------------------- 第一級：構件的缺失清單

@dataclass
class DefectSummary:
    deru: str
    deterioration: str
    method_id: str        # 建議維修工法 ID（網站唯一 ID）
    der_id: str
    view_url: str
    edit_url: str
    repaired: bool

    @property
    def status(self):
        return deru_status(self.deru)


def parse_defect_list(source):
    doc = tree(source)
    base = page_base(doc)
    found = []
    for edit in doc.xpath('//a[contains(@href,"edit_damage/")]'):
        row = edit.xpath('ancestor::tr[1]')[0]
        values = {text(h).rstrip(':：'): text(h.getnext()) for h in row.xpath('.//span[@class="der_h"]')}
        view = row.xpath('.//a[contains(@href,"view_damage/")]/@href')
        deru = ''.join(values.get(k, '_') or '_' for k in 'DERU')
        found.append(DefectSummary(deru, values.get('劣化類型', ''), values.get('建議維修工法ID', ''),
                                   url_params(edit.get('href')).get('der_id', ''),
                                   resolve(view[0], base) if view else '', resolve(edit.get('href'), base),
                                   any(text(a) == '已維修' for a in row.xpath('.//a'))))
    return found


# ---------------------------------------------------------------- 第二級：單筆缺失表單

FORM_FIELDS = ('d', 'e', 'r', 'u', 'deterioration', 'deterioration_text', 'long', 'width', 'depth', 'd_location',
               'reason', 'method', 'amount', 'unit', 'price', 'remark')


@dataclass
class DefectForm:
    method_id: str
    der_id: str
    keys: dict
    fields: dict
    options: dict          # 下拉選單選項（劣化類型、單位）


def parse_defect_form(source):
    doc = tree(source)
    form = next(iter(doc.xpath('//*[@id="form1"]')), doc)
    fields, options = {}, {}
    for name in FORM_FIELDS:
        node = next(iter(form.xpath(f'.//*[@id="{name}"]')), None)
        if node is None:
            continue
        if node.tag == 'select':
            chosen = node.xpath('.//option[@selected]')
            fields[name] = (chosen[0].get('value') or text(chosen[0])) if chosen else ''
            if name in ('deterioration', 'unit'):
                options[name] = [text(o) for o in node.xpath('.//option') if text(o)]
        elif node.tag == 'textarea':
            fields[name] = node.text_content()
        else:
            fields[name] = node.get('value') or ''
    hidden = {i.get('name'): i.get('value') or '' for i in form.xpath('.//input[@type="hidden"]')}
    return DefectForm(hidden.get('id', ''), hidden.get('der_id', ''),
                      {k: hidden.get(k, '') for k in ('component_form', 'component_id', 'inspect_id')},
                      fields, options)


VIEW_LABELS = {'劣化類型': 'deterioration', '損壞位置': 'd_location', '損壞原因': 'reason', '維修工法': 'method',
               '數量': 'amount', '單位': 'unit', '單價': 'price', '備註': 'remark'}


def parse_defect_view(source):
    """第二級「檢視」頁：純顯示的 th（欄名）／td（值）表格，沒有表單與 ID。
    識別靠網址上的 der_id（第一級清單已記下 der_id 與建議維修工法 ID 的對應）。"""
    doc = tree(source)
    fields = {}
    if doc.xpath('//form[@id="form1"]'):
        return DefectForm('', '', {}, {}, {})  # 這是編輯頁，不是檢視頁
    for th in doc.xpath('//th'):
        label = re.sub(r'\s+', '', text(th))
        cell = th.getnext()
        if cell is None or cell.tag != 'td':
            continue
        value = cell.text_content().strip() if label == '備註' else text(cell)
        if label == 'DER':
            fields.update({k.lower(): v for k, v in re.findall(r'([DER])\s*[:：]\s*([0-9_])', value)})
        elif label.startswith('急迫性'):
            fields['u'] = value
        elif label == '劣化範圍':
            for name, key in (('長', 'long'), ('寬', 'width'), ('高', 'depth')):
                found = re.search(name + r'\s*[:：]\s*([\d.]*)', value)
                fields[key] = found[1] if found else ''
        elif label in VIEW_LABELS:
            fields[VIEW_LABELS[label]] = value
    if 'deterioration' not in fields and 'remark' not in fields:
        return DefectForm('', '', {}, {}, {})
    return DefectForm('', '', {}, fields, {})


# ---------------------------------------------------------------- 歷史定期檢測紀錄

@dataclass
class InspectionRecord:
    inspect_id: str
    inventory_id: str
    date: str
    unit: str
    inspectors: str
    edit_url: str          # 缺失主頁面（der_content）
    view_url: str
    bridge: str = ''       # 橋梁名稱
    code: str = ''         # 橋梁編號（查詢結果有此欄時）
    agency: str = ''       # 管理機關
    place: str = ''        # 所在縣市＋所在區鄉


def parse_history(source):
    """歷史定期檢測紀錄，依檢測日期由新到舊排列。"""
    doc = tree(source)
    base = page_base(doc)
    for table in doc.xpath('//table'):
        rows, cells = grid(table)
        for y in range(len(rows)):
            heads = [header_key(text(c)) if c is not None else '' for c in cells[y]]
            if '檢測系統ID' in heads and '檢測日期' in heads:
                col = {h: i for i, h in enumerate(heads)}
                found = []
                for row_cells, row in zip(cells[y + 1:], rows[y + 1:]):
                    date = text(row_cells[col['檢測日期']])
                    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', date):
                        continue
                    edit = resolve(next(iter(row.xpath('.//a[contains(@href,"inspect/inspect/der_content/")]/@href')), ''), base)
                    view = resolve(next(iter(row.xpath('.//a[contains(@href,"inspect/inspect/view/")]/@href')), ''), base)
                    value = lambda name: text(row_cells[col[name]]) if name in col else ''
                    found.append(InspectionRecord(value('檢測系統ID'), value('橋梁系統ID'), date, value('檢測單位'),
                                                  value('檢測人員'), edit, view, value('橋梁名稱'), value('橋梁編號'),
                                                  value('管理機關'), value('所在縣市') + value('所在區鄉')))
                return sorted(found, key=lambda r: r.date, reverse=True)
    return []


@dataclass
class BridgeMatch:
    """查詢結果中的一座橋（以橋梁系統ID區分；龍門大橋與龍門大橋(2)是不同的橋）。"""
    inventory_id: str
    name: str
    agency: str
    place: str
    records: list          # 該橋的檢測紀錄，新到舊

    @property
    def latest(self):
        return self.records[0] if self.records else None


def group_bridges(records):
    bridges = {}
    for record in records:
        key = record.inventory_id or record.bridge
        if key not in bridges:
            bridges[key] = BridgeMatch(record.inventory_id, record.bridge, record.agency, record.place, [])
        bridges[key].records.append(record)
    for bridge in bridges.values():
        bridge.records.sort(key=lambda r: r.date, reverse=True)
    return list(bridges.values())


def needs_choice(bridges, name):
    """只有一座、且橋名與輸入完全相同時可直接使用；否則請使用者確認。"""
    exact = [b for b in bridges if b.name == name.strip()]
    return not (len(bridges) == 1 and len(exact) == 1)


def previous_inspection(history, inspect_id, date=''):
    """本期的前一期：日期早於本期的最新一筆。"""
    current = next((r for r in history if r.inspect_id == str(inspect_id)), None)
    date = current.date if current else date
    older = [r for r in history if r.inspect_id != str(inspect_id) and (not date or r.date < date)]
    return older[0] if older else None


# ---------------------------------------------------------------- 基本資料與構件明細

def parse_distribution(value):
    """`25+1@42.5+2@40+30` → [25, 42.5, 40, 40, 30]；`n@L` 表示 n 孔各長 L。長度可帶單位 M（如 `1@35M`）。"""
    spans = []
    cleaned = re.sub(r'(?<=\d)\s*(?:公尺|[mM])(?![a-zA-Z])', '', re.sub(r'\s+', '', value or ''))
    for part in re.split(r'[+＋]', cleaned):
        if not part:
            continue
        count, _, length = part.rpartition('@')
        size = number(length)
        repeat = int(count) if count.isdigit() else 1 if not count else None
        if size is None or repeat is None:
            raise ValueError(f'跨距分配無法解析：{value}')
        spans += [size] * repeat
    return spans


@dataclass
class BasicData:
    values: dict           # 欄位 id → 值
    spans: list            # 依跨距分配展開的各孔長度
    problems: list = field(default_factory=list)


BASIC_FIELDS = ('bridge_name', 'width_min', 'width_max', 'distribution', 'height_below', 'height_above', 'total_length')


def parse_basic_data(source):
    doc = tree(source)
    values = {}
    for name in BASIC_FIELDS:
        node = next(iter(doc.xpath(f'//input[@id="{name}"]')), None)
        if node is not None:
            values[name] = (node.get('value') or '').strip()
    spans, problems = [], []
    if values.get('distribution'):
        try:
            spans = parse_distribution(values['distribution'])
        except ValueError as error:
            problems.append(str(error))
    return BasicData(values, spans, problems)


@dataclass
class Span:
    code: str
    structure: str
    lanes: str
    length: Decimal | None
    full_width: Decimal | None
    girder_url: str
    span_id: str


def parse_spans(source):
    """橋孔資料表：依表頭找「橋孔編號、跨距(M)、全寬(M)」，主梁明細連結取自同列「主梁(n)」按鈕。"""
    doc = tree(source)
    for table in doc.xpath('//table'):
        rows, cells = grid(table)
        for y in range(len(rows)):
            heads = [header_key(text(c)) if c is not None else '' for c in cells[y]]
            if '橋孔編號' in heads and any(h.startswith('跨距') for h in heads):
                col = {h: i for i, h in enumerate(heads)}
                length = next(i for i, h in enumerate(heads) if h.startswith('跨距'))
                width = next((i for i, h in enumerate(heads) if h.startswith('全寬')), None)
                spans = []
                for row_cells, row in zip(cells[y + 1:], rows[y + 1:]):
                    code = text(row_cells[col['橋孔編號']])
                    if not code:
                        continue
                    girder = next(iter(row.xpath('.//a[starts-with(normalize-space(.),"主梁")]/@href')), '')
                    spans.append(Span(code, text(row_cells[col['結構型式']]) if '結構型式' in col else '',
                                      text(row_cells[col['車道數']]) if '車道數' in col else '',
                                      number(text(row_cells[length])),
                                      number(text(row_cells[width])) if width is not None else None,
                                      girder, re.search(r'span_id,(\d+)', girder)[1] if 'span_id,' in girder else ''))
                return spans
    return []


def parse_girders(source):
    """主梁明細頁：table#main 的主梁編號（例：S1G1、S1G2）；根數即筆數。"""
    doc = tree(source)
    table = next(iter(doc.xpath('//table[@id="main"]')), None)
    if table is None:
        return []
    rows, cells = grid(table)
    heads = [header_key(text(c)) for c in cells[0]] if cells else []
    if '主梁編號' not in heads:
        return []
    col = heads.index('主梁編號')
    return [text(r[col]) for r in cells[1:] if r[col] is not None and text(r[col])]


# ---------------------------------------------------------------- 淨高

def parse_clearance(comment):
    """檢測員意見中的淨高：單跨 `H=7.5m`，多跨 `P1=7.5m`。空白值（`P1=m`）回傳 None。"""
    found = {}
    for key, value in re.findall(r'(?<![A-Za-z0-9])(H|P\d+)\s*[=＝]\s*([\d.]*)\s*m', comment or '', flags=re.I):
        found[key.upper()] = number(value) if value else None
    return found


# ---------------------------------------------------------------- 筆數核對

@dataclass
class CountCheck:
    row: ComponentRow
    expected: int | None
    read: int
    ok: bool
    message: str


def check_count(row, defects):
    """比對「詳細資料(n)」與第一級實際讀到的筆數。D＝1／E＝0 的列，n＝1 是照片紀錄，不算缺失。"""
    ids = [d.method_id for d in defects if d.method_id]
    duplicate = sorted({i for i in ids if ids.count(i) > 1})
    if duplicate:
        return CountCheck(row, row.count, len(defects), False, f'讀到重複 ID：{"、".join(duplicate)}')
    if row.status in (GOOD, UNABLE) and not any(d.status == DEFECT for d in defects):
        return CountCheck(row, row.count, len(defects), True, f'{row.status}，無缺失')
    if row.count is None:
        return CountCheck(row, None, len(defects), False, '找不到詳細資料筆數')
    if len(defects) < row.count:
        return CountCheck(row, row.count, len(defects), False, f'網站 {row.count} 筆，只讀到 {len(defects)} 筆，請重掃')
    if len(defects) > row.count:
        return CountCheck(row, row.count, len(defects), False, f'網站 {row.count} 筆，讀到 {len(defects)} 筆')
    return CountCheck(row, row.count, len(defects), True, '')
