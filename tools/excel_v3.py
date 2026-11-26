"""將 v2 規則表（data/劣化說明與維修單價1124.xlsx）三層整併為 v3 規則表（data/劣化說明與維修單價v3.xlsx）。

三層整併：
1. 同列多單位單價（處單價、平方公尺單價、公尺單價、立方公尺單價、式單價、其他單位單價）；
2. 跨構件劣化類型群組（構件劣化類型對照 + @群組名）與同構件多劣化類型（｜分隔）；
3. 通用規則基底工作表（跨 3 張以上工作表共用的規則集中維護）+ 各構件特有／覆寫工作表。

轉換完成後自動與原 1124.xlsx 逐筆比對展開結果，驗證 100% 數學等價。
"""
import sys
from collections import defaultdict
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from bridgefill.core.config import COMPONENTS  # noqa: E402
from bridgefill.core.data import component_name, repair_components  # noqa: E402
from bridgefill.core.rules_v2 import (  # noqa: E402
    COMMON_SHEET, DET_MAP_SHEET, V3_EXTRA_UNIT_COL, V3_UNIT_COLUMNS,
    is_v2_headers, price_text, read_rules, words,
)

COND_COLS = (
    '處理方式', '任一關鍵字', '必要關鍵字', '排除關鍵字', '前置判讀', '適用構件',
    '條件欄位', '最小值', '最大值', '條件單位', '單獨關鍵字', '單獨關鍵字U值',
    '前期關鍵字', '優先序', '修正劣化類型為', '備註範本', '說明', '劣化ID',
)
CORE_COND_COLS = (
    '處理方式', '任一關鍵字', '必要關鍵字', '排除關鍵字', '前置判讀',
    '條件欄位', '最小值', '最大值', '條件單位', '單獨關鍵字', '單獨關鍵字U值',
    '前期關鍵字', '優先序', '修正劣化類型為',
)
V3_HEADERS = (
    ('劣化類型', '劣化樣態', '維修工法')
    + tuple(col for col, _ in V3_UNIT_COLUMNS)
    + (V3_EXTRA_UNIT_COL,)
    + COND_COLS
)
UNIT_TO_COL = {unit: col for col, unit in V3_UNIT_COLUMNS}

HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
COMMON_FILL = PatternFill('solid', fgColor='2E75B6')
MAP_FILL = PatternFill('solid', fgColor='385723')
HEADER_FONT = Font(name='Microsoft JhengHei', size=10, bold=True, color='FFFFFF')
BODY_FONT = Font(name='Microsoft JhengHei', size=10)
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _norm(val):
    if val in (None, ''):
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def _copy_sheet(src_wb, dst_wb, title, row_offset=0):
    from openpyxl.utils.cell import range_boundaries
    src = src_wb[title]
    dst = dst_wb.create_sheet(title)
    for row in src.iter_rows():
        for cell in row:
            new_cell = dst.cell(cell.row + row_offset, cell.column, cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
            if cell.comment:
                new_cell.comment = copy(cell.comment)
    for merged in src.merged_cells.ranges:
        c1, r1, c2, r2 = range_boundaries(str(merged))
        dst.merge_cells(start_row=r1 + row_offset, start_column=c1, end_row=r2 + row_offset, end_column=c2)
    for k, dim in src.column_dimensions.items():
        if dim.width:
            dst.column_dimensions[k].width = dim.width
    for k, dim in src.row_dimensions.items():
        if dim.height:
            dst.row_dimensions[k + row_offset].height = dim.height
    if src.freeze_panes and row_offset == 0:
        dst.freeze_panes = src.freeze_panes
    if src.auto_filter.ref and row_offset == 0:
        dst.auto_filter.ref = src.auto_filter.ref
    return dst


def _style_table(ws, headers, fill=HEADER_FILL, freeze='D2'):
    ws.append(list(headers))
    for idx in range(1, len(headers) + 1):
        cell = ws.cell(1, idx)
        cell.fill = fill
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 26
    if freeze:
        ws.freeze_panes = freeze


def _finalize_table(ws, widths=None):
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center', wrap_text=True)
    if ws.max_row >= 2:
        ws.auto_filter.ref = f'A1:{get_column_letter(ws.max_column)}{ws.max_row}'
    widths = widths or {}
    for idx in range(1, ws.max_column + 1):
        name = str(ws.cell(1, idx).value or '')
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = widths.get(name, 14)


def _format_unit_cols(unit_prices):
    row = {col: None for col, _ in V3_UNIT_COLUMNS}
    extras = []
    for unit, price in unit_prices:
        val = int(price) if price.isdigit() else float(price) if _is_float(price) else price
        if unit in UNIT_TO_COL:
            row[UNIT_TO_COL[unit]] = val
        else:
            extras.append(f'{unit}:{price}')
    row[V3_EXTRA_UNIT_COL] = '｜'.join(extras) if extras else None
    return row


def _is_float(text):
    try:
        float(text)
        return True
    except ValueError:
        return False


def _core_sig(rec):
    return (rec['pattern'], rec['method']) + tuple(_norm(rec['fields'].get(c)) for c in CORE_COND_COLS)


def _assign_det_group(comp_to_dets, det_groups, pattern):
    """將 {component: tuple(dets)} 分配到相容的 @劣化分類，避免衝突並盡量共用分類名。"""
    # 檢查是否已有完全相容（無衝突）的既有分類
    for gname, gmap in det_groups.items():
        conflict = any(comp in gmap and gmap[comp] != dets for comp, dets in comp_to_dets.items())
        if not conflict:
            # 只有當重疊的構件數 >= 3 且完全一致時才共用，避免不相關分類混在一起
            overlap = sum(1 for comp, dets in comp_to_dets.items() if comp in gmap and gmap[comp] == dets)
            if overlap >= 3:
                gmap.update(comp_to_dets)
                return gname
    base = pattern
    if '鋼筋外露' in pattern or '保護層不足' in pattern:
        base = '混凝土鋼筋外露類'
    elif '裂縫' in pattern:
        base = '混凝土裂縫類'
    elif '白華' in pattern or '滲水' in pattern:
        base = '滲水白華類'
    elif any(w in pattern for w in ('剝落', '破損', '碎裂', '蜂窩', '沖蝕', '斷裂', '龜裂')):
        base = '混凝土劣化類'
    elif '鋼' in pattern:
        base = '鋼構劣化類'
    gname = base
    idx = 2
    while gname in det_groups:
        gname = f'{base}{idx}'
        idx += 1
    det_groups[gname] = dict(comp_to_dets)
    return gname


def convert(source_path, target_path, min_common_sheets=3):
    """將 v2 Excel 轉換為三層整併的 v3 Excel，回傳轉換統計資訊。"""
    source_path, target_path = Path(source_path), Path(target_path)
    src = load_workbook(source_path, data_only=True)

    # 1. 讀取所有工作表的 v2 規則列（記錄順序）
    rule_sheet_names = []
    other_sheets_before = []
    other_sheets_after = []
    raw_records = []
    seen_rule = False

    for ws in src.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        headers = [str(h or '').strip() for h in (rows[0] if rows else ())]
        if is_v2_headers(headers):
            seen_rule = True
            rule_sheet_names.append(ws.title)
            sheet_targets = repair_components(ws.title)
            for idx, row in enumerate(rows[1:], 2):
                fields = dict(zip(headers, row))
                if not any(fields.values()) or not fields.get('劣化ID'):
                    continue
                pattern = _norm(fields.get('劣化樣態'))
                method = _norm(fields.get('維修工法'))
                unit = _norm(fields.get('維修單位'))
                price = price_text(fields.get('單價'))
                if not (pattern and method and unit and price):
                    continue
                only = [component_name(w) for w in words(fields.get('適用構件'))]
                targets = tuple(c for c in dict.fromkeys(only) if c in sheet_targets or c in COMPONENTS) if only else tuple(sheet_targets)
                if not targets:
                    continue
                raw_records.append({
                    'sheet': ws.title,
                    'row': idx,
                    'targets': targets,
                    'explicit_apply': bool(only),
                    'deterioration': _norm(fields.get('劣化類型')),
                    'pattern': pattern,
                    'method': method,
                    'unit': unit,
                    'price': price,
                    'fields': fields,
                    'rule_id': _norm(fields.get('劣化ID')),
                })
        else:
            (other_sheets_after if seen_rule else other_sheets_before).append(ws.title)

    # 2. Layer 1：同工作表、同目標構件、同劣化類型、同條件 → 合併多單位單價
    l1_groups = []
    l1_index = {}
    for rec in raw_records:
        key = (rec['sheet'], rec['targets'], rec['deterioration'], _core_sig(rec))
        entry = l1_index.get(key)
        # 若同單位卻有不同單價，不可合併，開新列保留
        if entry is not None and rec['unit'] not in entry['unit_map']:
            entry['unit_map'][rec['unit']] = rec['price']
            entry['ids'].append(rec['rule_id'])
            entry['records'].append(rec)
        elif entry is not None and entry['unit_map'][rec['unit']] == rec['price']:
            entry['ids'].append(rec['rule_id'])
            entry['records'].append(rec)
        else:
            entry = {
                'sheet': rec['sheet'],
                'targets': rec['targets'],
                'explicit_apply': rec['explicit_apply'],
                'deterioration': rec['deterioration'],
                'core_sig': _core_sig(rec),
                'unit_map': {rec['unit']: rec['price']},
                'fields': dict(rec['fields']),
                'ids': [rec['rule_id']],
                'records': [rec],
            }
            l1_groups.append(entry)
            if key not in l1_index:
                l1_index[key] = entry

    # 3. Layer 3 + Layer 2：找出跨 >= min_common_sheets 工作表完全一致（除劣化類型外）的規則移入「通用規則」
    by_universal = defaultdict(list)
    for idx, entry in enumerate(l1_groups):
        # 含「修正劣化類型為」時，note 與劣化類型耦合，留在原工作表較清晰
        if _norm(entry['fields'].get('修正劣化類型為')):
            continue
        ukey = (entry['core_sig'], tuple(sorted(entry['unit_map'].items())))
        by_universal[ukey].append(idx)

    promoted_indices = set()
    common_rows = []
    det_groups = {}

    for ukey, idxs in by_universal.items():
        sheets = {l1_groups[i]['sheet'] for i in idxs}
        if len(sheets) < min_common_sheets:
            continue
        # 建立 {comp: [det, ...]} 並檢查同構件是否衝突
        comp_dets = defaultdict(list)
        for i in idxs:
            e = l1_groups[i]
            for c in e['targets']:
                if e['deterioration'] not in comp_dets[c]:
                    comp_dets[c].append(e['deterioration'])
        comp_to_dets = {c: tuple(ds) for c, ds in comp_dets.items()}
        all_dets = {d for ds in comp_to_dets.values() for d in ds}

        sample = l1_groups[idxs[0]]
        if len(all_dets) == 1:
            det_value = next(iter(all_dets))
        else:
            gname = _assign_det_group(comp_to_dets, det_groups, sample['core_sig'][0])
            det_value = f'@{gname}'

        promoted_indices.update(idxs)
        row_data = dict(sample['fields'])
        row_data['劣化類型'] = det_value
        row_data['劣化樣態'] = sample['core_sig'][0]
        row_data['維修工法'] = sample['core_sig'][1]
        row_data['適用構件'] = '、'.join(comp_to_dets.keys())
        row_data['說明'] = f'跨 {len(sheets)} 張工作表通用（整併自 {sum(len(l1_groups[i]["records"]) for i in idxs)} 列）'
        row_data['劣化ID'] = f'G-{len(common_rows) + 1:02d}'
        row_data.update(_format_unit_cols(sorted(sample['unit_map'].items())))
        common_rows.append(row_data)

    # 4. 各構件工作表保留剩餘規則，並對同工作表同條件不同劣化類型以「｜」整併
    sheet_rows = {name: [] for name in rule_sheet_names}
    per_sheet_merge = {}
    for idx, entry in enumerate(l1_groups):
        if idx in promoted_indices:
            continue
        mkey = (entry['sheet'], entry['targets'], entry['core_sig'], tuple(sorted(entry['unit_map'].items())))
        if mkey in per_sheet_merge and not _norm(entry['fields'].get('修正劣化類型為')):
            existing = per_sheet_merge[mkey]
            if entry['deterioration'] and entry['deterioration'] not in existing['_dets']:
                existing['_dets'].append(entry['deterioration'])
                existing['劣化類型'] = '｜'.join(existing['_dets'])
        else:
            row_data = dict(entry['fields'])
            row_data['_dets'] = [entry['deterioration']] if entry['deterioration'] else []
            row_data['劣化類型'] = entry['deterioration']
            row_data['劣化樣態'] = entry['core_sig'][0]
            row_data['維修工法'] = entry['core_sig'][1]
            row_data['適用構件'] = '、'.join(entry['targets']) if entry['explicit_apply'] else ''
            row_data['劣化ID'] = entry['ids'][0]
            row_data.update(_format_unit_cols(sorted(entry['unit_map'].items())))
            sheet_rows[entry['sheet']].append(row_data)
            per_sheet_merge[mkey] = row_data

    # 5. 建立 v3 工作簿
    dst = Workbook()
    dst.remove(dst.active)

    for title in other_sheets_before:
        if title == '說明':
            ws = _copy_sheet(src, dst, title, row_offset=3)
            ws['A1'] = '【v3 三層整併規則表說明】'
            ws['A1'].font = Font(name='Microsoft JhengHei', size=11, bold=True, color='1F4E78')
            ws.merge_cells('A1:D1')
            ws['A2'] = ('1. 同一列直接並列「處／平方公尺／公尺／立方公尺／式／其他單位單價」，增補單位免複製整列；'
                        '2. 跨構件通用規則集中在「通用規則」，劣化類型以 @分類名 對應「構件劣化類型對照」；'
                        '3. 各構件工作表只放該構件特有規則或覆寫（同構件同樣態同單位時，構件工作表優先於通用規則）。')
            ws['A2'].font = BODY_FONT
            ws['A2'].alignment = Alignment(wrap_text=True, vertical='center')
            ws.merge_cells('A2:D2')
            ws.row_dimensions[1].height = 24
            ws.row_dimensions[2].height = 42
        else:
            _copy_sheet(src, dst, title)

    widths = {
        '劣化類型': 26, '劣化樣態': 22, '維修工法': 26,
        '處單價': 10, '平方公尺單價': 12, '公尺單價': 10, '立方公尺單價': 12, '式單價': 10, '其他單位單價': 16,
        '處理方式': 10, '任一關鍵字': 20, '必要關鍵字': 14, '排除關鍵字': 12, '前置判讀': 14,
        '適用構件': 28, '備註範本': 26, '說明': 28, '劣化ID': 10,
    }

    # 通用規則
    ws_common = dst.create_sheet(COMMON_SHEET)
    _style_table(ws_common, V3_HEADERS, fill=COMMON_FILL)
    for row_data in common_rows:
        ws_common.append([row_data.get(h) for h in V3_HEADERS])
    _finalize_table(ws_common, widths)

    # 構件劣化類型對照
    ws_map = dst.create_sheet(DET_MAP_SHEET)
    map_headers = ('劣化分類', '構件', '網站劣化類型', '說明')
    _style_table(ws_map, map_headers, fill=MAP_FILL, freeze='A2')
    for gname, gmap in det_groups.items():
        # 同分類下相同網站劣化類型的構件排在一起，易讀
        by_det = defaultdict(list)
        for comp, dets in gmap.items():
            by_det['｜'.join(dets)].append(comp)
        for det_str, comps in by_det.items():
            ws_map.append([gname, '、'.join(comps), det_str, ''])
    _finalize_table(ws_map, {'劣化分類': 20, '構件': 34, '網站劣化類型': 38, '說明': 24})

    # 20 張構件工作表
    for name in rule_sheet_names:
        ws = dst.create_sheet(name)
        _style_table(ws, V3_HEADERS, fill=HEADER_FILL)
        for row_data in sheet_rows[name]:
            ws.append([row_data.get(h) for h in V3_HEADERS])
        _finalize_table(ws, widths)

    # 複製其餘設定工作表（樣態同義詞、互斥樣態、劣化類型等價、位置代碼對照、跨構件對應、構件分類、DERU評等、意見U值、意見排除清單、意見構件名稱、參數、修訂紀錄）
    for title in other_sheets_after:
        ws_copied = _copy_sheet(src, dst, title)
        if title == '意見排除清單':
            # 原草稿僅列「橋墩基礎」混凝土剝落，同步涵蓋橋台基礎與橋墩基礎（與 opinion.EXCLUDED 一致）
            for row in ws_copied.iter_rows(min_row=3, max_col=2):
                if str(row[0].value or '').strip() == '橋墩基礎' and str(row[1].value or '').strip() == '混凝土剝落':
                    row[0].value = '基礎'

    src.close()
    target_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(target_path)

    # 6. 嚴格等價驗證
    old_rules = read_rules(source_path)
    new_rules = read_rules(target_path)

    def expand_set(rules):
        s = set()
        for r in rules:
            for c in r['components']:
                s.add((c, r['deterioration'], r['pattern'], r['method'], r['unit'], price_text(r['price']), r['note']))
        return s

    old_set, new_set = expand_set(old_rules), expand_set(new_rules)
    missing = old_set - new_set
    extra = new_set - old_set
    if missing or extra:
        raise AssertionError(f'v3 轉換不等價：缺少 {len(missing)} 筆、多出 {len(extra)} 筆\n'
                             f'missing={list(missing)[:5]}\nextra={list(extra)[:5]}')

    v3_rows = len(common_rows) + sum(len(v) for v in sheet_rows.values())
    return {
        'v2_rows': len(raw_records),
        'v3_rows': v3_rows,
        'common_rows': len(common_rows),
        'specific_rows': sum(len(v) for v in sheet_rows.values()),
        'det_map_rows': ws_map.max_row - 1,
        'expanded_tuples': len(new_set),
    }


if __name__ == '__main__':
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'data' / '劣化說明與維修單價1124.xlsx'
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / 'data' / '劣化說明與維修單價v3.xlsx'
    stats = convert(src, dst)
    print('轉換成功並通過 100% 等價驗證：')
    for k, v in stats.items():
        print(f'  {k}: {v}')
