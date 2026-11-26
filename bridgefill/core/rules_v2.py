"""讀取 v2 與 v3 劣化維修規則表（結構化欄位，已刪除原「備註」）。

v2：每列單一「維修單位」「單價」、各構件獨立工作表。
v3（三層整併）：
  第一層：同一列支援多單位單價欄（處單價、平方公尺單價、公尺單價、立方公尺單價、式單價、其他單位單價），
          載入時自動展開為標準規則。
  第二層：劣化類型支援多值（｜分隔）或群組參照（@群組名，查「構件劣化類型對照」工作表依構件展開）。
  第三層：支援「通用規則」基底工作表（依「適用構件」展開），各構件獨立工作表只放特有規則或覆寫。

判讀邏輯沿用 v1.00 已驗證的 core.conditions／keyword_score：由結構化欄位還原成 v1 標準寫法的備註
（synthesize_note），交給原本的判斷程式，所以新舊判斷結果 100% 一致。
"""
import re

from openpyxl import load_workbook

from .config import COMPONENTS
from .data import compact, component_name, repair_components

NUMBER_FIELDS = ('最小值', '最大值', '單獨關鍵字U值', '優先序')
RUST = re.compile('[銹鏽锈]蝕')

V3_UNIT_COLUMNS = (
    ('處單價', '處'),
    ('平方公尺單價', '平方公尺'),
    ('公尺單價', '公尺'),
    ('立方公尺單價', '立方公尺'),
    ('式單價', '式'),
)
V3_EXTRA_UNIT_COL = '其他單位單價'
COMMON_SHEET = '通用規則'
DET_MAP_SHEET = '構件劣化類型對照'


def words(value):
    return [w.strip() for w in str(value or '').split('、') if w.strip()]


def number_text(value):
    if value in (None, ''):
        return ''
    return f'{float(value):g}'


def price_text(value):
    """單價字串：整數不帶 .0，保留有效小數。"""
    if value in (None, ''):
        return ''
    text = str(value).strip()
    try:
        num = float(text.replace(',', ''))
        return f'{int(num)}' if num.is_integer() else f'{num:g}'
    except ValueError:
        return text


def synthesize_note(fields, deterioration=''):
    """結構化欄位 → v1.00 標準寫法的備註（給 core.conditions 使用）。"""
    if fields.get('處理方式') == '手動':
        return '手動處理'
    parts = []
    priority = str(fields.get('優先序') or '') in ('1', '1.0')
    any_words, required = words(fields.get('任一關鍵字')), words(fields.get('必要關鍵字'))
    # 「破損銹蝕」等含銹蝕的任一關鍵字本身就代表鋼筋外露（2026-11-20 使用者：銹蝕、鏽蝕對應鋼筋外露），
    # 有必要關鍵字時也可單獨成立，不再要求同時寫「破損」「鋼筋外露」。
    rust = [w for w in any_words if RUST.search(w)] if required else []
    any_words = [w for w in any_words if w not in rust]
    keywords = ''
    if any_words and required:
        keywords = '或'.join(f'[{w}]' for w in any_words) + '與' + '與'.join(f'[{w}]' for w in required)
    elif required:
        keywords = '與'.join(f'[{w}]' for w in required)
    elif any_words:
        keywords = '或'.join(f'[{w}]' for w in any_words)
    if rust:
        keywords += ''.join(f'或[{w}]' for w in rust)

    field, low, high, unit = (fields.get('條件欄位'), fields.get('最小值'), fields.get('最大值'),
                              fields.get('條件單位') or '')
    if field == '護欄高度':
        text = ('優先選擇，' if priority else '') + '或讀取到護欄高度' if priority else '讀取到護欄高度'
        if low not in (None, ''):
            text += f'大於等於{number_text(low)}{unit}'
        if high not in (None, ''):
            text += f'小於{number_text(high)}{unit}'
        parts.append(text + '時使用')
    elif field == '裂縫寬度':
        if low not in (None, ''):
            parts.append(f'裂縫寬度大於[{number_text(low)}{unit}]使用')
        else:
            parts.append(f'裂縫寬度小於[{number_text(high)}{unit}]使用')
    elif fields.get('前期關鍵字'):
        if keywords:
            parts.append(f'有關鍵字{keywords}即可使用')
        if priority or not keywords:
            parts.append(f"若過去資料為[{fields['前期關鍵字']}]，則優先使用")
        if fields.get('修正劣化類型為'):
            parts.append(f"但劣化類型填成[{deterioration}]，則修正成[{fields['修正劣化類型為']}]後使用")
    elif fields.get('前置判讀'):
        parts.append(('優先使用，' if priority else '') + f"先判讀是否為[{fields['前置判讀']}]，再判讀關鍵字{keywords}")
    elif fields.get('修正劣化類型為'):
        parts.append(f'若判讀到關鍵字{keywords}即可使用，但劣化類型填成[{deterioration}]，'
                     f"則修正成[{fields['修正劣化類型為']}]後使用")
    elif keywords:
        parts.append(f'有關鍵字{keywords}即可使用')

    alone, alone_u = fields.get('單獨關鍵字'), fields.get('單獨關鍵字U值')
    if alone and alone_u not in (None, ''):
        parts.append(f'或僅有關鍵字[{alone}]時，若U值為{number_text(alone_u)}優先使用')
    elif alone:
        parts.append(f'若僅寫[{alone}]則優先使用此項')
    return '，'.join(parts)


def is_v3_headers(headers):
    header_set = set(headers)
    return ({'劣化樣態', '維修工法', '劣化ID', '處理方式'} <= header_set
            and any(col in header_set for col, _ in V3_UNIT_COLUMNS))


def is_v2_headers(headers):
    return {'劣化樣態', '維修工法', '維修單位', '單價', '劣化ID', '處理方式'} <= set(headers)


def parse_unit_prices(fields):
    """從一列取出 [(unit, price_str), ...]，同時支援 v2（單一維修單位＋單價）與 v3（多單位單價欄）。"""
    if '維修單位' in fields and '單價' in fields:
        unit = str(fields.get('維修單位') or '').strip()
        price = price_text(fields.get('單價'))
        return [(unit, price)] if unit and price else []

    pairs = []
    for col, unit in V3_UNIT_COLUMNS:
        val = price_text(fields.get(col))
        if val:
            pairs.append((unit, val))
    extra = str(fields.get(V3_EXTRA_UNIT_COL) or '').strip()
    if extra:
        for chunk in re.split(r'[|｜、;\n]+', extra):
            chunk = chunk.strip()
            if not chunk or ':' not in chunk and '：' not in chunk:
                continue
            unit, price = re.split(r'[:：]', chunk, maxsplit=1)
            unit, price = unit.strip(), price_text(price)
            if unit and price:
                pairs.append((unit, price))
    return pairs


def _read_det_map(workbook):
    """讀取「構件劣化類型對照」：回傳 {(group_name, component): [det_type, ...]}。"""
    mapping = {}
    if DET_MAP_SHEET not in workbook.sheetnames:
        return mapping
    rows = list(workbook[DET_MAP_SHEET].iter_rows(values_only=True))
    if not rows:
        return mapping
    header_idx = 1 if len(rows) > 1 and len([v for v in rows[0] if v not in (None, '')]) == 1 else 0
    heads = [str(h or '').strip() for h in rows[header_idx]]
    col = {h: i for i, h in enumerate(heads) if h}
    if not {'劣化分類', '構件', '網站劣化類型'} <= set(col):
        return mapping
    for row in rows[header_idx + 1:]:
        grp = str(row[col['劣化分類']] or '').strip().lstrip('@')
        comp_raw = str(row[col['構件']] or '').strip()
        det_raw = str(row[col['網站劣化類型']] or '').strip()
        if not grp or not comp_raw or not det_raw:
            continue
        dets = [d.strip() for d in re.split(r'[|｜\n]+', det_raw) if d.strip()]
        for c in words(comp_raw):
            comp = component_name(c)
            if comp:
                bucket = mapping.setdefault((grp, comp), [])
                for d in dets:
                    if d not in bucket:
                        bucket.append(d)
    return mapping


def _resolve_targets(raw_apply, sheet_targets, is_common):
    if is_common:
        if not raw_apply or compact(raw_apply) in ('全部', '（全部）', '(全部)'):
            return list(COMPONENTS.keys())
        only = [component_name(w) for w in words(raw_apply)]
        return [c for c in dict.fromkeys(only) if c in COMPONENTS]
    only = [component_name(w) for w in words(raw_apply)]
    return [c for c in dict.fromkeys(only) if c in sheet_targets or c in COMPONENTS] if only else sheet_targets


def _expand_deteriorations(raw_det, targets, det_map):
    """展開劣化類型：
    - 直接字串或 ｜ 分隔 → 所有 targets 共用同一清單
    - @分類名 → 依「構件劣化類型對照」查出每個構件對應的網站劣化類型，回傳 [(det, [comp, ...]), ...]
    """
    raw_det = str(raw_det or '').strip()
    if not raw_det:
        return [('', list(targets))]
    if raw_det.startswith('@'):
        grp = raw_det[1:].strip()
        by_det = {}
        for comp in targets:
            for det in det_map.get((grp, comp), ()):
                by_det.setdefault(det, []).append(comp)
        return list(by_det.items())
    dets = [d.strip() for d in re.split(r'[|｜\n]+', raw_det) if d.strip()]
    return [(d, list(targets)) for d in dets]


def _rule_override_key(rule, component):
    return (component, rule['pattern'], rule['unit'], rule['deterioration'], rule['note'])


def read_rules(path):
    """讀取 v2 或 v3 規則表 → 與 v1.00 read_repairs 相同格式的規則清單（note 由結構化欄位還原）。"""
    workbook = load_workbook(path, read_only=True, data_only=True)
    common_rules = []
    specific_rules = []
    try:
        det_map = _read_det_map(workbook)
        for sheet in workbook:
            rows = sheet.iter_rows(values_only=True)
            headers = [str(h or '').strip() for h in next(rows, ())]
            if not (is_v2_headers(headers) or is_v3_headers(headers)):
                continue
            is_common = sheet.title == COMMON_SHEET
            sheet_targets = [] if is_common else repair_components(sheet.title)
            bucket = common_rules if is_common else specific_rules

            for number, row in enumerate(rows, 2):
                fields = dict(zip(headers, row))
                if not any(fields.values()) or not fields.get('劣化ID'):
                    continue
                pattern = str(fields.get('劣化樣態') or '').strip()
                method = str(fields.get('維修工法') or '').strip()
                unit_prices = parse_unit_prices(fields)
                if not pattern or not method or not unit_prices:
                    continue
                targets = _resolve_targets(fields.get('適用構件'), sheet_targets, is_common)
                if not targets:
                    continue
                base_id = str(fields['劣化ID']).strip()
                det_groups = _expand_deteriorations(fields.get('劣化類型'), targets, det_map)
                structured = {k: v for k, v in fields.items() if v not in (None, '')}

                for deterioration, comps in det_groups:
                    note = synthesize_note(fields, deterioration)
                    for unit, price in unit_prices:
                        rule_id = base_id if len(unit_prices) == 1 else f'{base_id}/{unit}'
                        bucket.append(dict(
                            pattern=pattern,
                            method=method,
                            unit=unit,
                            price=price,
                            deterioration=deterioration,
                            note=note,
                            rule_id=rule_id,
                            structured=structured,
                            component=comps[0] if len(comps) == 1 else '',
                            components=comps,
                            automatic_ready=bool(comps and deterioration),
                            source=f'{sheet.title} 第{number}列',
                        ))
    finally:
        workbook.close()

    # 若構件工作表覆寫了通用規則的同 (構件, 樣態, 單位, 劣化類型, 條件)，以構件工作表為準
    if common_rules and specific_rules:
        overridden = {_rule_override_key(r, c) for r in specific_rules for c in r['components']}
        filtered_common = []
        for rule in common_rules:
            kept = [c for c in rule['components'] if _rule_override_key(rule, c) not in overridden]
            if not kept:
                continue
            if len(kept) != len(rule['components']):
                rule = dict(rule, components=kept, component=kept[0] if len(kept) == 1 else '')
            filtered_common.append(rule)
        result = specific_rules + filtered_common
    else:
        result = specific_rules + common_rules

    if not result:
        raise ValueError('找不到 v2/v3 規則表（需有 劣化樣態、維修工法、單價、處理方式、劣化ID 欄）')
    return result
