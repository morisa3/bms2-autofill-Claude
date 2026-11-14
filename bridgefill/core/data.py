"""離線資料匯入、保守比對及備註正規化。不改動原始 Word / Excel。"""
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
import re
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.escape import unescape
from .config import ALIASES, COMPONENTS, SHARED_REPAIR_SHEETS
from .locations import location_prefix,position_separator


def compact(value):
    return re.sub(r"\s+", "", str(value or ""))


def quantity(value):
    try:
        return str(Decimal(str(value).strip().replace(",", "")).normalize())
    except InvalidOperation:
        return compact(value)


def component_name(value):
    value = compact(value).replace("/", "／").replace("（","(").replace("）",")").replace("－","-").replace("–","-").replace("—","-")
    return ALIASES.get(value, value)


def normalize_remark(text):
    """調整可辨認的『位置 描述(尺寸)』，二維尺寸短邊在前，不推算數量。"""
    lines = []
    for line in str(text or "").splitlines():
        line = line.strip().rstrip('。．.').translate(str.maketrans({"＊": "*", "（": "(", "）": ")", "×": "*"}))
        line=position_separator(line)
        if not line:
            continue
        match = re.match(r"^(.+?)\s*(\([^()]+\)(?:\s*\*\s*\d+(?:\.\d+)?)?)\s*$", line)
        if not match:
            lines.append(line)
            continue
        prefix,size=match.groups()
        parts=location_prefix(prefix.strip().rstrip('。．. '))
        if not parts:
            lines.append(line);continue
        location,description=parts
        description = description.strip().rstrip("。．. ")
        size = re.sub(r"\s+", "", size)
        if "裂縫" in compact(description):
            parts_match = re.fullmatch(r"\(([^()]+)\)(\*\d+(?:\.\d+)?)?", size)
            if parts_match:
                dimensions = parts_match[1].split("*")
                widths = [x for x in dimensions if re.fullmatch(r"(?:\d+(?:\.\d+)?|\.\d+)(?:mm|cm)", x)]
                widths = [(format((Decimal(x[:-2])*Decimal(10)).normalize(),'f')+'mm') if x.endswith('cm') else x for x in widths]
                lengths = [x for x in dimensions if re.fullmatch(r"(?:\d+(?:\.\d+)?|\.\d+)m", x)]
                if len(widths) == len(lengths) == 1:
                    if len(dimensions) == 2:
                        size = f"({widths[0]}*{lengths[0]}){parts_match[2] or ''}"
                    elif len(dimensions) == 3 and re.fullmatch(r"\d+(?:\.\d+)?", dimensions[-1]):
                        count = Decimal(dimensions[-1]) * Decimal((parts_match[2] or '*1')[1:])
                        size = f"({widths[0]}*{lengths[0]})*{format(count.normalize(), 'f')}"
        else:
            dimension_match = re.fullmatch(r"\(([^()]+)\)(\*\d+(?:\.\d+)?)?", size)
            if dimension_match:
                dimensions = dimension_match[1].split("*")
                parsed = [re.fullmatch(r"(\d+(?:\.\d+)?|\.\d+)(mm|cm|m)?", value) for value in dimensions]
                if len(dimensions) == 2 and all(parsed):
                    # 單位明確時換算比較，顯示保留原數值與原單位。
                    units = [value[2] or "" for value in parsed]
                    if all(units) or not any(units):
                        factors = {"": Decimal(1), "mm": Decimal("0.001"), "cm": Decimal("0.01"), "m": Decimal(1)}
                        order = sorted(range(2), key=lambda i: Decimal(parsed[i][1]) * factors[units[i]])
                        size = f"({'*'.join(dimensions[i] for i in order)}){dimension_match[2] or ''}"
        lines.append(f"{location} {description}。{size}")
    return "\n".join(lines)


@dataclass
class History:
    component: str
    fields: dict
    source_row: int
    context: dict = field(default_factory=dict)


def read_history(path, audit=None):
    result = []
    mapping = {"劣化類型": "deterioration", "位置": "d_location", "損壞位置": "d_location",
               "維修工法": "method", "數量": "amount", "單位": "unit", "單價": "price", "金額": "price", "備註": "remark"}
    for table in Document(path).tables:
        if len(table.rows) < 3 or "檢測項目" not in table.cell(0, 0).text:
            continue
        # 報表有兩種合併儲存格：部分表格第 0、1 格都是『檢測項目』。
        titles = [compact(c.text) for c in table.rows[0].cells]
        names = [value for value in titles if value and value != "檢測項目"]
        if not names:
            continue
        name = component_name(names[0])
        headers = [compact(c.text) for c in table.rows[1].cells]
        if audit is not None and (name not in COMPONENTS or not all(h in headers for h in ('劣化類型','備註'))):
            audit.setdefault('issues',set()).add(name)
        for index, row in enumerate(table.rows[2:], 3):
            fields = {mapping[h]: row.cells[i].text.strip() for i, h in enumerate(headers)
                      if h in mapping and i < len(row.cells)}
            status={h:row.cells[i].text.strip() for i,h in enumerate(headers) if h in ('D','E','R','U') and i<len(row.cells)}
            if fields.get('deterioration') or status.get('D')=='1' or status.get('E')=='0':
                context = {h: row.cells[i].text.strip() for i,h in enumerate(headers)
                           if h in ('位置','編號','上/下層','類別','護欄材質','排水孔數量') and i < len(row.cells)}
                context.update(status)
                if status.get('D')=='1':context['_inspection_status']='前期檢測良好（D＝1）'
                elif status.get('E')=='0':context['_inspection_status']='前期無法檢測（E＝0）'
                remark = normalize_remark(fields.get('remark','')).strip()
                prefix = re.match(r'^(\S+)\s+.+', remark)
                if prefix:
                    fields['d_location'] = prefix[1]
                result.append(History(name, fields, index, context))
    if not result:
        raise ValueError("Word 中找不到含『檢測項目／劣化類型』的資料表。")
    return result


def candidates(history, component, current):
    """有填值的識別欄都必須一致；多筆候選由人決定，不取第一筆。"""
    matches = []
    for item in history:
        if item.component != component_name(component):
            continue
        if all(not current.get(key, "").strip() or
               (quantity(current[key]) == quantity(item.fields.get(key, "")) if key == "amount" else
                compact(current[key]) == compact(item.fields.get(key, "")))
               for key in ("deterioration", "d_location", "amount", "unit")):
            matches.append(item)
    return matches


def propose(current, previous=None, repair=None, replace_repair=False):
    result = dict(current)
    if previous:
        for key, value in previous.fields.items():
            if not result.get(key, "").strip():
                result[key] = value
    if result.get("remark"):
        result["remark"] = normalize_remark(result["remark"])
    if repair:
        for key in ("method", "unit", "price"):
            if replace_repair or not result.get(key, "").strip():
                result[key] = repair[key]
    return result


def repair_components(title):
    """辨識 Excel 安全分隔符、複合構件及共用工作表，不作模糊子字串比對。"""
    title = unescape(str(title)).replace('\ufeff', '').replace('\u200b', '')
    title=title.translate(str.maketrans({'（':'(', '）':')','－':'-','–':'-','—':'-'}))
    title = re.sub(r'^\s*\d+[.．、_\-]\s*', '', title).strip()
    separators = r'[/／與及或、,，|｜]'

    def key(value):
        return re.sub(separators, '', compact(value))

    for name,targets in SHARED_REPAIR_SHEETS.items():
        if key(title)==key(name):return list(targets)
    lookup = {}
    for name, identity in COMPONENTS.items():
        for alias in (name, identity):
            lookup.setdefault(key(alias), set()).add(name)
    for alias, target in ALIASES.items():
        if target in COMPONENTS:
            lookup.setdefault(key(alias), set()).add(target)

    # 先比完整名稱，保留「墩柱與帽梁」「支承／支承墊／阻尼裝置」為單一構件。
    whole = lookup.get(key(title), set())
    if len(whole) == 1:
        return list(whole)
    parts = re.split(separators, compact(title))
    if not parts or any(not part for part in parts) or len(parts) > 30:
        return []
    # 將連續片段組回已知複合名稱，再辨識共用工作表的其他構件。
    matches = {len(parts): {()}}
    for start in range(len(parts) - 1, -1, -1):
        found = set()
        for end in range(start + 1, len(parts) + 1):
            names = lookup.get(key(''.join(parts[start:end])), set())
            for name in names:
                for tail in matches.get(end, set()):
                    found.add(tuple(dict.fromkeys((name,) + tail)))
        matches[start] = found
    result = matches.get(0, set())
    return list(next(iter(result))) if len(result) == 1 else []


def read_repairs(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    result = []
    try:
        for sheet in workbook:
            rows = iter(sheet.values)
            headers = []
            # 允許表名前有標題列，最多尋找前二十列。
            required = ('劣化樣態', '維修工法', '維修單位', '單價')
            for number, row in enumerate(rows, 1):
                headers = [compact(v) for v in row]
                if all(h in headers for h in required):
                    break
                if number >= 20:
                    break
            if not all(h in headers for h in required):
                continue
            start = number + 1
            for number, row in enumerate(rows, start):
                values = [row[headers.index(h)] for h in required]
                if any(v is None or not str(v).strip() for v in values):
                    continue
                extra = {key: str(row[headers.index(header)] or '').strip() if header in headers else ''
                         for key, header in (('deterioration', '劣化類型'), ('note', '備註'), ('rule_id', '劣化ID'))}
                targets = repair_components(sheet.title)
                result.append(dict(zip(('pattern', 'method', 'unit', 'price'), (str(v).strip() for v in values)),
                                   **extra, component=targets[0] if len(targets) == 1 else '', components=targets,
                                   automatic_ready=bool(targets and extra['deterioration']),
                                   source=f'{sheet.title} 第{number}列'))
    finally:
        workbook.close()
    if not result:
        raise ValueError("Excel 需包含：劣化樣態、維修工法、維修單位、單價。")
    return result


def component_repairs(repairs, component):
    name = component_name(component)
    return [rule for rule in repairs if
            (name in rule['components'] if 'components' in rule else
             not rule.get('component') or name == rule['component'])]


def repair_candidates(repairs, current, component=None, filter_unit=True):
    text = compact(current.get('remark', '') or current.get('deterioration', ''))
    found = []
    from .conditions import keyword_score, keyword_groups, aliases
    for rule in component_repairs(repairs, component) if component else repairs:
        note = rule.get('note', '')
        score = keyword_score(note, text)
        if score is None and any(len(group) > 1 for group in keyword_groups(note)):
            continue
        pattern = compact(rule['pattern'])
        same_pattern=pattern in text or (pattern=='基礎裸露' and bool(re.search(r'基礎(?:樁帽|基樁)裸露',text)))
        if score is not None or same_pattern:
            own = max((len(compact(word)) for word in aliases(note) if compact(word) in text and compact(word) in pattern), default=0)
            strength = max(score or 0, len(pattern) if same_pattern else 0)
            found.append((rule, (score is not None, strength, own)))
    if '白華' in text and any('白華' in rule['pattern'] for rule,_ in found):
        found=[item for item in found if '白華' in item[0]['pattern'] or item[0]['pattern'] not in ('滲水','漏水')]
    if not found:
        return []
    best = max(score for _, score in found)
    return [rule for rule, score in found if score == best and
            (not filter_unit or not current.get('unit') or compact(rule['unit']) == compact(current['unit']))]


def eligible(d, e):
    # Word 報表在 D=1 或 E=0 時常將另一欄留白；先依明確排除值判斷。
    for value,target in ((d,1),(e,0)):
        try:
            if Decimal(str(value).strip())==target:return False
        except InvalidOperation:pass
    try:
        dv, ev = Decimal(str(d)), Decimal(str(e))
        if not dv.is_finite() or not ev.is_finite():
            return None
        return dv != 1 and ev != 0
    except InvalidOperation:
        return None
