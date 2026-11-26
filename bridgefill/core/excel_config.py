"""讀取規則表 Excel 中的 8 張設定工作表（供配對、填寫、檢測員意見使用，不依賴 Qt）。

涵蓋工作表：
- 互斥樣態（matching.score）
- 劣化類型等價（matching.score）
- 跨構件對應（matching.related）
- 構件分類（opinion.suggestion）
- 意見U值（opinion.suggestion）
- 意見排除清單（opinion.excluded）
- 意見構件名稱（opinion.display_name）
- 參數（fill.shape_remark、opinion.suggestion / water_paragraph）

讀不到工作表或尚未設定規則表時，一律自動退回內建預設值，不影響原本行為。
"""
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

from .data import compact, component_name

_path = None
_cache = {}

DEFAULT_EXCLUSIVE = (
    ('木本', '草本'),
    ('齒型', '角鋼'),
    ('AC', '混凝土'),
    ('瀝青', '混凝土'),
    ('鋼管', '鋼筋'),
    ('鋼板', '鋼管'),
)

DEFAULT_RELATED = (
    (('河道',), ('橋墩基礎', '橋台基礎', '墩柱與帽梁', '橋台', '橋墩／橋基保護設施', '護床工'),
     ('掛淤', '漂流', '淤積', '雜物', '沖刷', '垃圾')),
    (('伸縮縫',), ('橋面板', '主梁', '橫隔梁', '墩柱與帽梁', '橋台'), ('漏水', '白華')),
    (('橋梁排水設施',), ('橋面板', '主梁', '墩柱與帽梁'), ('漏水', '白華', '阻塞')),
    (('引道路堤',), ('引道道路堤保護措施', '引道路堤護欄'), ()),
    (('橋台',), ('橋台基礎', '翼牆／擋土牆'), ()),
    (('墩柱與帽梁',), ('橋墩基礎',), ()),
)

DEFAULT_MAIN = frozenset({
    '主梁', '墩柱與帽梁', '橋台', '橋台基礎', '橋墩基礎', '橋面板',
    '端牆', '中隔板', '頂板', '底板',
    '鋼纜錨碇裝置(橋塔端)', '鋼纜錨碇裝置(主梁端)', '鋼纜(吊索)-斜張鋼纜',
    '拱肋', '立柱',
})

DEFAULT_TRAFFIC = frozenset({
    '橋面板', '伸縮縫', '橋護欄', '引道路堤', '引道路堤護欄', '其他', '頂板',
})

DEFAULT_NAMES = {
    '墩柱與帽梁': '橋墩',
    '橋護欄': '護欄',
    '引道路堤護欄': '引道護欄',
    '橋梁排水設施': '排水設施',
    '引道道路堤保護措施': '引道保護措施',
    '橋墩／橋基保護設施': '橋墩保護設施',
    '其他': '',
}

DEFAULT_EXCLUDED = (
    ('伸縮縫', '漏水', ''),
    ('基礎', '混凝土剝落', ''),
    ('', '蜂窩', ''),
    ('', 'AC剝離', ''),
)

DEFAULT_ADVICE = {
    'u2_track': '建議持續追蹤',
    'u2_repair': '建議三年內派工修復',
    'u3': '建議一年內派工修復',
    'u4': '建議通知相關單位立即處置',
}

DEFAULT_PARAMS = {
    '大面積判定 E 值': {'min': 3.0, 'max': None, 'unit': ''},
    '位置含「-All」時的 E 值': {'min': 4.0, 'max': 4.0, 'unit': ''},
    '引道路堤預設長度': {'min': 5.0, 'max': 5.0, 'unit': 'm'},
    '裂縫寬度轉 cm 門檻': {'min': 10.0, 'max': None, 'unit': 'mm'},
    '基礎裸露量測誤差': {'min': 0.1, 'max': None, 'unit': 'm'},
}


def configure(path):
    """設定目前使用中的規則表路徑（由 editing.load_rules_checked 呼叫）。"""
    global _path
    _path = Path(path) if path else None


def clear_cache():
    _cache.clear()


def _words(value):
    return [w.strip() for w in str(value or '').split('、') if w.strip()]


def _u_str(value):
    if value in (None, ''):
        return ''
    try:
        return f'{float(value):g}'
    except (TypeError, ValueError):
        return str(value).strip()


def _to_float(value):
    if value in (None, ''):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _is_all(value):
    text = compact(value)
    return not text or text in ('（全部）', '(全部)', '全部', '（不限）', '(不限)', '不限')


def _sheet_rows(workbook, name):
    """找出標題列並回傳 (header_col_map, data_rows)。支援第 1 列為說明列、第 2 列為標題列。"""
    if name not in workbook.sheetnames:
        return {}, []
    rows = list(workbook[name].iter_rows(values_only=True))
    if not rows:
        return {}, []
    # 如果第 1 列只有第 1 格有長說明、第 2 列有多格標題，取第 2 列為標題
    header_idx = 0
    if len(rows) > 1:
        r0 = [v for v in rows[0] if v not in (None, '')]
        r1 = [v for v in rows[1] if v not in (None, '')]
        if len(r0) == 1 and len(r1) >= 2:
            header_idx = 1
    heads = [str(v or '').strip() for v in rows[header_idx]]
    col = {h: i for i, h in enumerate(heads) if h}
    return col, rows[header_idx + 1:]


def _cell(row, col, name):
    idx = col.get(name)
    return row[idx] if idx is not None and idx < len(row) else None


def read_config(path):
    """讀取規則表內全部設定工作表並快取。讀不到回傳空結構（呼叫端退回預設值）。"""
    path = Path(path)
    try:
        key = (str(path), path.stat().st_mtime)
    except OSError:
        return {}
    if key in _cache:
        return _cache[key]

    cfg = {}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return cfg

    try:
        # 1. 互斥樣態
        col, rows = _sheet_rows(wb, '互斥樣態')
        if '樣態A' in col and '樣態B' in col:
            pairs = []
            for row in rows:
                a = str(_cell(row, col, '樣態A') or '').strip()
                b = str(_cell(row, col, '樣態B') or '').strip()
                if a and b:
                    pairs.append((a, b))
            if pairs:
                cfg['exclusive'] = tuple(pairs)

        # 2. 劣化類型等價
        col, rows = _sheet_rows(wb, '劣化類型等價')
        if '群組' in col and '劣化類型' in col:
            groups = {}
            for row in rows:
                gid = _cell(row, col, '群組')
                dtype = str(_cell(row, col, '劣化類型') or '').strip()
                if gid in (None, '') or not dtype:
                    continue
                comp_raw = _cell(row, col, '適用構件')
                comps = frozenset() if _is_all(comp_raw) else frozenset(component_name(w) for w in _words(comp_raw))
                entry = groups.setdefault(str(gid).strip(), {'types': set(), 'components': set(), 'all': False})
                entry['types'].add(compact(dtype))
                if not comps:
                    entry['all'] = True
                else:
                    entry['components'].update(comps)
            if groups:
                cfg['det_equivalents'] = tuple(
                    (gid, frozenset(g['types']), frozenset() if g['all'] else frozenset(g['components']))
                    for gid, g in groups.items() if len(g['types']) >= 2
                )

        # 3. 跨構件對應
        col, rows = _sheet_rows(wb, '跨構件對應')
        if '構件' in col and '相關構件' in col:
            rel = []
            for row in rows:
                left = tuple(component_name(w) for w in _words(_cell(row, col, '構件')))
                right = tuple(component_name(w) for w in _words(_cell(row, col, '相關構件')))
                kw_raw = _cell(row, col, '適用樣態關鍵字')
                kws = () if _is_all(kw_raw) else tuple(_words(kw_raw))
                if left and right:
                    rel.append((left, right, kws))
            if rel:
                cfg['related'] = tuple(rel)

        # 4. 構件分類
        col, rows = _sheet_rows(wb, '構件分類')
        if '構件' in col and '主構件' in col:
            main_set, traffic_set = set(), set()
            for row in rows:
                comp = component_name(_cell(row, col, '構件'))
                if not comp:
                    continue
                if str(_cell(row, col, '主構件') or '').strip() == '是':
                    main_set.add(comp)
                if str(_cell(row, col, '可能影響行車安全') or '').strip() == '是':
                    traffic_set.add(comp)
            if main_set or traffic_set:
                cfg['classes'] = {'main': frozenset(main_set), 'traffic': frozenset(traffic_set)}

        # 5. 意見U值
        col, rows = _sheet_rows(wb, '意見U值')
        if 'U值' in col and '建議文字' in col:
            advice = dict(DEFAULT_ADVICE)
            for row in rows:
                u = _u_str(_cell(row, col, 'U值'))
                text = str(_cell(row, col, '建議文字') or '').strip().rstrip('。.')
                note = str(_cell(row, col, '說明') or '').strip()
                if not u or not text:
                    continue
                if u == '2':
                    if '追蹤' in text or '不影響' in note:
                        advice['u2_track'] = text
                    else:
                        advice['u2_repair'] = text
                elif u == '3':
                    advice['u3'] = text
                elif u == '4':
                    advice['u4'] = text
            cfg['advice'] = advice

        # 6. 意見排除清單
        col, rows = _sheet_rows(wb, '意見排除清單')
        if '構件' in col and '樣態關鍵字' in col:
            excl = []
            for row in rows:
                comp_raw = _cell(row, col, '構件')
                pat = str(_cell(row, col, '樣態關鍵字') or '').strip()
                if not pat:
                    continue
                comp = '' if _is_all(comp_raw) else str(comp_raw).strip()
                u = _u_str(_cell(row, col, '適用U值'))
                excl.append((comp, pat, u))
            if excl:
                # 向下相容 1124.xlsx 草稿（只列橋墩基礎混凝土剝落時，同步涵蓋橋台基礎）
                for comp, pat, u in list(excl):
                    if comp == '橋墩基礎' and pat == '混凝土剝落' and not any(c in ('基礎', '橋台基礎') and p == pat for c, p, _ in excl):
                        excl.append(('橋台基礎', pat, u))
                cfg['excluded'] = tuple(excl)

        # 7. 意見構件名稱
        col, rows = _sheet_rows(wb, '意見構件名稱')
        if '構件' in col and '意見中的寫法' in col:
            names = dict(DEFAULT_NAMES)
            for row in rows:
                comp = component_name(_cell(row, col, '構件'))
                written = str(_cell(row, col, '意見中的寫法') or '').strip()
                if not comp or not written:
                    continue
                if written.startswith('（省略') or written.startswith('(省略'):
                    names[comp] = ''
                elif written.startswith('依') or written.startswith('（') or written.startswith('('):
                    continue  # 屬於文字說明（如「依備註寫翼牆或擋土牆」），保留程式規則
                else:
                    names[comp] = written
            cfg['names'] = names

        # 8. 參數
        col, rows = _sheet_rows(wb, '參數')
        if '參數' in col and '最小' in col:
            params = dict(DEFAULT_PARAMS)
            for row in rows:
                pname = str(_cell(row, col, '參數') or '').strip()
                if not pname:
                    continue
                params[pname] = {
                    'min': _to_float(_cell(row, col, '最小')),
                    'max': _to_float(_cell(row, col, '最大')),
                    'unit': str(_cell(row, col, '單位') or '').strip(),
                }
            cfg['params'] = params
    finally:
        wb.close()

    _cache[key] = cfg
    return cfg


def current_config(path=None):
    target = Path(path) if path else _path
    if target is None:
        return {}
    return read_config(target)


def exclusive_pairs(path=None):
    return current_config(path).get('exclusive', DEFAULT_EXCLUSIVE)


def deterioration_equivalents(path=None):
    return current_config(path).get('det_equivalents', ())


def same_deterioration(a, b, component='', path=None):
    """兩筆劣化類型是否相同或於 Excel「劣化類型等價」屬同一群組。"""
    ca, cb = compact(a), compact(b)
    if not ca or not cb:
        return False
    if ca.replace('混凝土結構裂縫', '混凝土裂縫') == cb.replace('混凝土結構裂縫', '混凝土裂縫'):
        return True
    comp = component_name(component) if component else ''
    for _gid, types, comps in deterioration_equivalents(path):
        if comps and comp and comp not in comps:
            continue
        if ca in types and cb in types:
            return True
    return False


def related_groups(path=None):
    return current_config(path).get('related', DEFAULT_RELATED)


def component_classes(path=None):
    return current_config(path).get('classes', {'main': DEFAULT_MAIN, 'traffic': DEFAULT_TRAFFIC})


def opinion_advice(path=None):
    return current_config(path).get('advice', DEFAULT_ADVICE)


def opinion_excluded(path=None):
    return current_config(path).get('excluded', DEFAULT_EXCLUDED)


def opinion_names(path=None):
    return current_config(path).get('names', DEFAULT_NAMES)


def param_value(name, default, which='min', path=None):
    params = current_config(path).get('params', DEFAULT_PARAMS)
    entry = params.get(name)
    if not entry:
        return default
    val = entry.get(which)
    return default if val is None else val


def param_decimal(name, default, which='min', path=None):
    val = param_value(name, None, which, path)
    if val is None:
        return Decimal(str(default))
    try:
        return Decimal(str(val))
    except InvalidOperation:
        return Decimal(str(default))
