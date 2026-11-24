"""Excel 沒有對應規則時，依規則格式自動新增到規則表（v2），並帶入建議值。

兩種情況：
1. 劣化類型判讀差異：同構件已有相同樣態的規則，但劣化類型不同 → 複製該規則，只改劣化類型。
2. 新的劣化樣態：同構件沒有這個樣態 → 新增一列；工法、單位、單價取自本期網站值（檢測員填的建議值），
   網站值不完整時才複製最相近的規則（同單位、同劣化類型優先）。

新增的列「處理方式」為自動、任一關鍵字為該樣態，「說明」寫明來源並標黃底請使用者確認。
寫入前先備份規則表；規則表在 Excel 中開啟時無法寫入，會回報錯誤，不影響其他功能。
"""
import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from copy import copy

from .data import component_repairs, repair_components
from .editing import split_remark
from .local_files import APP_DIR, REFERENCE_DIR
from .matching import pattern_similarity, strip_measure
from .reconcile import deterioration_key

CHECK_FILL = PatternFill('solid', fgColor='FFF2CC')
COUNT_UNITS = ('處', '個', '支', '組', '座', '塊', '根', '條', '片', '孔', '只')
NO_RULE = 'Excel 無對應值'


def remark_template(unit, pattern):
    if not pattern or not unit:
        return ''
    if unit == '公尺':
        size = '(mm*m)*1' if '裂縫' in pattern else '(m)*1'
    elif unit == '平方公尺':
        size = '(m*m)*1'
    elif unit == '立方公尺':
        size = '(m*m*m)*1'
    elif unit in COUNT_UNITS:
        size = f'(1{unit})'
    else:
        return ''
    return '{位置} {樣態}。' + size


def defect_pattern(defect):
    """本期備註的樣態（去掉位置與尺寸、結尾量測值）。"""
    _, pattern, _ = split_remark(defect.remark)
    pattern = strip_measure(pattern).strip(' 。.，,')
    return pattern


def suggest_rule(defect, rules):
    """回傳要新增的規則（dict），或 None（無法判斷樣態，或已有相同規則）。"""
    pattern = defect_pattern(defect)
    deterioration = (defect.fields.get('deterioration') or defect.deterioration or '').strip()
    if not pattern or not deterioration:
        return None
    own = component_repairs(rules, defect.component)
    if any(r['pattern'] == pattern and deterioration_key(r.get('deterioration')) == deterioration_key(deterioration)
           for r in own):
        return None
    unit = (defect.fields.get('unit') or '').strip()
    same = [r for r in own if r['pattern'] == pattern]
    site = {k: (defect.fields.get(k) or '').strip() for k in ('method', 'unit', 'price')}
    if same:
        source = next((r for r in same if r['unit'] == unit), same[0])
        kind = '劣化類型判讀差異'
    elif all(site.values()):
        source = {'method': site['method'], 'unit': site['unit'], 'price': site['price'], 'rule_id': '本期網站值'}
        kind = '新劣化樣態'
    else:
        pool = own or rules
        if not pool:
            return None

        def rank(rule):
            similarity = pattern_similarity(pattern, rule['pattern']) or 0
            return (rule['unit'] == unit, deterioration_key(rule.get('deterioration')) == deterioration_key(deterioration),
                    similarity)
        source = max(pool, key=rank)
        kind = '新劣化樣態'
    return {'劣化類型': deterioration, '劣化樣態': pattern, '維修工法': source['method'], '維修單位': source['unit'],
            '單價': source['price'], '處理方式': '自動', '任一關鍵字': pattern,
            '備註範本': remark_template(source['unit'], pattern), 'kind': kind, 'source': source.get('rule_id', ''),
            'component': defect.component}


def writable_rules_file(path):
    """規則表直接寫入使用中的檔案（data/劣化說明與維修單價1124.xlsx）。"""
    return Path(path)


def next_id(sheet, id_col):
    numbers, prefix = [], ''
    for (value,) in sheet.iter_rows(min_row=2, min_col=id_col, max_col=id_col, values_only=True):
        found = re.fullmatch(r'(\d+)-(\d+)', str(value or '').strip())
        if found:
            prefix = found[1]
            numbers.append(int(found[2]))
    return f'{prefix}-{max(numbers, default=0) + 1:02d}' if prefix else ''


def listed_components(sheet):
    """工作表「適用構件」欄列出的所有構件。"""
    headers = [c.value for c in sheet[1]]
    if '適用構件' not in headers:
        return set()
    col = headers.index('適用構件')
    return {w.strip() for row in sheet.iter_rows(min_row=2, values_only=True)
            for w in str(row[col] or '').split('、') if w.strip()}


def sheet_for(workbook, component):
    """構件所屬的規則工作表：工作表名稱含該構件，否則為「適用構件」列有該構件的工作表（例：拱肋 → 主梁或橫隔梁或橋面版）。"""
    sheets = rule_sheets(workbook)
    return (next((ws for ws in sheets if component in repair_components(ws.title)), None)
            or next((ws for ws in sheets if component in listed_components(ws)), None))


def add_rules(path, additions, note='', backup_dir=None):
    """把 additions 寫入規則表。回傳 [(劣化ID, 工作表, 新增內容)]。PermissionError：檔案開啟中。"""
    path = Path(path)
    inside = APP_DIR.resolve() in path.resolve().parents
    backup_dir = Path(backup_dir or (APP_DIR / 'backup' if inside else path.parent / 'backup'))
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup_dir / f'{path.stem}-{datetime.now():%Y%m%d-%H%M%S}{path.suffix}')
    workbook = load_workbook(path)
    added = []
    for rule in additions:
        sheet = sheet_for(workbook, rule['component'])
        if sheet is None:
            continue
        headers = [c.value for c in sheet[1]]
        id_col = headers.index('劣化ID') + 1
        rule_id = next_id(sheet, id_col)
        last = sheet.max_row
        while last > 1 and not any(c.value for c in sheet[last]):
            last -= 1
        row = last + 1
        reason = (f"{datetime.now():%Y-%m-%d} 自動新增（{rule['kind']}）：{rule.get('origin', note)}"
                  + ('工法、單位、單價取自本期網站值' if rule['source'] == '本期網站值'
                     else f"工法、單位、單價複製自 {rule['source'] or '相近規則'}") + '，請確認')
        values = dict(rule, 劣化ID=rule_id, 說明=reason, 適用構件=rule['component'])   # 只用於發生的那個構件
        for col, name in enumerate(headers, 1):
            cell = sheet.cell(row, col, values.get(name))
            above = sheet.cell(last, col)
            if above.has_style:
                cell.font, cell.border, cell.alignment = copy(above.font), copy(above.border), copy(above.alignment)
                cell.number_format = above.number_format
            if name == '說明':
                cell.fill = CHECK_FILL
        sheet.row_dimensions[row].height = sheet.row_dimensions[last].height
        for validation in sheet.data_validations.dataValidation:
            if any(cell_range.min_row <= last <= cell_range.max_row for cell_range in validation.sqref.ranges):
                validation.add(sheet.cell(row, headers.index('處理方式') + 1).coordinate)
        if sheet.auto_filter.ref:
            sheet.auto_filter.ref = re.sub(r'\d+$', str(row), sheet.auto_filter.ref)
        added.append((rule_id, sheet.title, rule))
    workbook.save(path)
    return added


def rule_sheets(workbook):
    return [ws for ws in workbook.worksheets if '劣化ID' in [c.value for c in ws[1]]]


def sync_sheets(source, target, backup_dir=None):
    """使用中的規則表（data）缺少建議版新增的工作表時（例：護床工），整張複製過去；既有工作表與自動新增的規則不動。
    回傳複製的工作表名稱。"""
    source, target = Path(source), Path(target)
    if not source.exists() or not target.exists() or source.resolve() == target.resolve():
        return []
    new = load_workbook(source)
    old = load_workbook(target)
    missing = [ws for ws in rule_sheets(new) if ws.title not in old.sheetnames]
    if not missing:
        return []
    inside = APP_DIR.resolve() in target.resolve().parents
    backup_dir = Path(backup_dir or (APP_DIR / 'backup' if inside else target.parent / 'backup'))
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(target, backup_dir / f'{target.stem}-{datetime.now():%Y%m%d-%H%M%S}{target.suffix}')
    position = len(rule_sheets(old))
    for sheet in missing:
        copy_sheet = old.create_sheet(sheet.title, position)
        position += 1
        for row in sheet.iter_rows():
            for cell in row:
                new_cell = copy_sheet.cell(cell.row, cell.column, cell.value)
                if cell.has_style:
                    new_cell.font, new_cell.border = copy(cell.font), copy(cell.border)
                    new_cell.fill, new_cell.alignment = copy(cell.fill), copy(cell.alignment)
                    new_cell.number_format = cell.number_format
                if cell.comment:
                    new_cell.comment = copy(cell.comment)
        for key, dimension in sheet.column_dimensions.items():
            copy_sheet.column_dimensions[key].width = dimension.width
        for key, dimension in sheet.row_dimensions.items():
            copy_sheet.row_dimensions[key].height = dimension.height
        copy_sheet.freeze_panes = sheet.freeze_panes
        copy_sheet.auto_filter.ref = sheet.auto_filter.ref
        for validation in sheet.data_validations.dataValidation:
            copy_sheet.add_data_validation(copy(validation))
    old.save(target)
    return [ws.title for ws in missing]


MERGED_RULES = {   # 1123 已人工合併或刪除的自動新增規則：從舊規則表搬移時略過（工作表, 劣化類型, 劣化樣態）
    ('伸縮縫', '其他損傷', '混凝土碎裂'),                       # 併入「錨固混凝土破損」（破損、碎裂）
    ('伸縮縫', '其他損傷', '木本植物生長'),                     # 改為「植物生長」
    ('主梁或橫隔梁或橋面版', '混凝土剝落、破碎、鋼筋、鋼腱或錨碇外露、銹蝕', '銹蝕蜂窩'),   # 併入「混凝土蜂窩鋼筋外露」
    ('護床工', '保護設施損壞、移動或遺失', '混凝土破損'),       # 1123v1：改為「表層混凝土破損」
}


def migrate_auto_rules(old, new, backup_dir=None):
    """舊規則表（1111v2）中程式自動新增的規則（說明含「自動新增」）搬進新規則表；同工作表、同劣化類型與樣態已有的不重複，
    MERGED_RULES 中已人工合併的也不搬。搬完把舊檔移到 backup，避免重複搬移。回傳搬入的筆數。"""
    old, new = Path(old), Path(new)
    if not old.exists() or not new.exists() or old.resolve() == new.resolve():
        return 0
    source, target = load_workbook(old), load_workbook(new)
    moved = 0
    for sheet in rule_sheets(source):
        if sheet.title not in target.sheetnames:
            continue
        headers = [c.value for c in sheet[1]]
        if '說明' not in headers:
            continue
        ws = target[sheet.title]
        names = [c.value for c in ws[1]]
        have = {(str(r[names.index('劣化類型')] or ''), str(r[names.index('劣化樣態')] or ''))
                for r in ws.iter_rows(min_row=2, values_only=True)}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            values = dict(zip(headers, row))
            if '自動新增' not in str(values.get('說明') or ''):
                continue
            key = (str(values.get('劣化類型') or ''), str(values.get('劣化樣態') or ''))
            if key in have or (sheet.title, *key) in MERGED_RULES:
                continue
            last = ws.max_row
            while last > 1 and not any(c.value for c in ws[last]):
                last -= 1
            values['劣化ID'] = next_id(ws, names.index('劣化ID') + 1)
            for col, name in enumerate(names, 1):
                cell = ws.cell(last + 1, col, values.get(name))
                above = ws.cell(last, col)
                if above.has_style:
                    cell.font, cell.border, cell.alignment = copy(above.font), copy(above.border), copy(above.alignment)
                if name == '說明':
                    cell.fill = CHECK_FILL
            have.add(key)
            moved += 1
    if moved:
        target.save(new)
    inside = APP_DIR.resolve() in old.resolve().parents
    backup_dir = Path(backup_dir or (APP_DIR / 'backup' if inside else old.parent / 'backup'))
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.move(str(old), backup_dir / f'{old.stem}-已併入新規則表-{datetime.now():%Y%m%d-%H%M%S}{old.suffix}')
    return moved
