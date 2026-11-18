"""Excel「樣態同義詞」與「位置代碼對照」：填寫時把同義寫法改成標準寫法（2026-11-19 使用者指定）。

樣態同義詞（標準樣態、同義寫法、來源、適用U值、適用構件）：
- 同義寫法必須是整段樣態（前面可帶構件名稱，例：引道路堤老化），不會改到其他樣態中間的字；
- 有「適用U值」時只在該 U 值套用（植物生長：U1 → 草本植物生長、U2 → 木本植物生長）；
- 有「適用構件」時只用於列出的構件。
位置代碼對照：有填「適用構件」的列，填寫時改寫損壞位置與備註開頭（整體-ALL → 整體-L,R）。
"""
from pathlib import Path

from openpyxl import load_workbook

from .data import component_name
from .editing import split_remark

_path = None
_cache = {}


def configure(path):
    """使用中的規則表（editing.load_rules 讀取時設定）。"""
    global _path
    _path = Path(path) if path else None


def words(value):
    return [w.strip() for w in str(value or '').split('、') if w.strip()]


def read_tables(path):
    """回傳 {'patterns': [(標準, 同義, U值, 構件集合)], 'locations': [(寫法, 標準, 構件集合)]}；讀不到回傳空表。"""
    path = Path(path)
    try:
        key = (str(path), path.stat().st_mtime)
    except OSError:
        return {'patterns': [], 'locations': []}
    if key in _cache:
        return _cache[key]
    tables = {'patterns': [], 'locations': []}
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return tables
    try:
        if '樣態同義詞' in workbook.sheetnames:
            rows = list(workbook['樣態同義詞'].iter_rows(values_only=True))
            heads = [str(v or '').strip() for v in (rows[1] if len(rows) > 1 else ())]
            col = {h: i for i, h in enumerate(heads)}
            for row in rows[2:]:
                cell = lambda name: row[col[name]] if name in col and col[name] < len(row) else None  # noqa: E731
                standard = str(cell('標準樣態') or '').strip()
                if not standard:
                    continue
                raw = cell('適用U值')
                try:
                    u = f'{float(raw):g}' if raw not in (None, '') else ''
                except (TypeError, ValueError):
                    u = str(raw).strip()
                parts = {component_name(w) for w in words(cell('適用構件'))}
                for synonym in words(cell('同義寫法')):
                    tables['patterns'].append((standard, synonym, u, parts))
        if '位置代碼對照' in workbook.sheetnames:
            rows = list(workbook['位置代碼對照'].iter_rows(values_only=True))
            heads = [str(v or '').strip() for v in (rows[1] if len(rows) > 1 else ())]
            col = {h: i for i, h in enumerate(heads)}
            if '適用構件' in col:
                for row in rows[2:]:
                    parts = {component_name(w) for w in words(row[col['適用構件']] if col['適用構件'] < len(row) else '')}
                    written, standard = (str(row[col[k]] or '').strip() for k in ('寫法', '標準代碼'))
                    if parts and written and standard:
                        tables['locations'].append((written, standard, parts))
    finally:
        workbook.close()
    tables['patterns'].sort(key=lambda entry: -len(entry[1]))      # 長的寫法先比（破損鋼筋外露鏽蝕 先於 鋼筋外露鏽蝕）
    _cache[key] = tables
    return tables


def tables():
    if _path is None:
        from .editing import rules_file
        path = rules_file()
        if path is None:
            return {'patterns': [], 'locations': []}
        configure(path)
    return read_tables(_path)


def standard_pattern(remark, component, u='', table=None):
    """備註的樣態若是同義寫法，改成標準樣態（保留位置與範圍）；不是則原樣回傳。"""
    text = str(remark or '').strip()
    location, description, tail = split_remark(text)
    core = description.rstrip('。. ')
    if not core:
        return text
    name = component_name(component)
    for standard, synonym, only_u, parts in (table if table is not None else tables()['patterns']):
        if parts and name not in parts:
            continue
        if only_u and str(u).strip() != only_u:
            continue
        prefix = core[:-len(synonym)] if core.endswith(synonym) else None
        if prefix is None or (prefix and prefix not in name):
            continue
        if core == standard:
            return text
        rebuilt = standard + description[len(core):] + tail
        return f'{location} {rebuilt}' if location else rebuilt
    return text


def standard_location(fields, component, table=None):
    """位置代碼對照中有適用構件的寫法：損壞位置與備註開頭改成標準寫法（例：整體-ALL → 整體-L,R）。"""
    result = dict(fields)
    name = component_name(component)
    for written, standard, parts in (table if table is not None else tables()['locations']):
        if name not in parts:
            continue
        location = str(result.get('d_location') or '').strip()
        if location.upper().startswith(written.upper()):
            result['d_location'] = standard + location[len(written):]
        remark = str(result.get('remark') or '').strip()
        if remark.upper().startswith(written.upper()):
            result['remark'] = standard + remark[len(written):]
    return result
