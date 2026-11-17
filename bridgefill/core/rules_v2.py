"""讀取 v2 劣化維修規則表（結構化欄位，已刪除原「備註」）。

判讀邏輯沿用 v1.00 已驗證的 core.conditions／keyword_score：由結構化欄位還原成 v1 標準寫法的備註
（synthesize_note），交給原本的判斷程式，所以新舊判斷結果一致（見 tests/test_rules_v2.py）。
"""
from openpyxl import load_workbook

from .config import COMPONENTS
from .data import component_name, repair_components

NUMBER_FIELDS = ('最小值', '最大值', '單獨關鍵字U值', '優先序')


def words(value):
    return [w.strip() for w in str(value or '').split('、') if w.strip()]


def number_text(value):
    if value in (None, ''):
        return ''
    return f'{float(value):g}'


def synthesize_note(fields, deterioration=''):
    """結構化欄位 → v1.00 標準寫法的備註（給 core.conditions 使用）。"""
    if fields.get('處理方式') == '手動':
        return '手動處理'
    parts = []
    priority = str(fields.get('優先序') or '') in ('1', '1.0')
    any_words, required = words(fields.get('任一關鍵字')), words(fields.get('必要關鍵字'))
    keywords = ''
    if any_words and required:
        keywords = '或'.join(f'[{w}]' for w in any_words) + '與' + '與'.join(f'[{w}]' for w in required)
    elif required:
        keywords = '與'.join(f'[{w}]' for w in required)
    elif any_words:
        keywords = '或'.join(f'[{w}]' for w in any_words)

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
        parts.append(f"若過去資料為[{fields['前期關鍵字']}]，則優先使用")
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


def read_rules(path):
    """v2 規則表 → 與 v1.00 read_repairs 相同格式的規則清單（note 由結構化欄位還原）。"""
    workbook = load_workbook(path, read_only=True, data_only=True)
    result = []
    try:
        for sheet in workbook:
            rows = sheet.iter_rows(values_only=True)
            headers = [str(h or '').strip() for h in next(rows, ())]
            if not {'劣化樣態', '維修工法', '維修單位', '單價', '劣化ID', '處理方式'} <= set(headers):
                continue
            sheet_targets = repair_components(sheet.title)
            for number, row in enumerate(rows, 2):
                fields = dict(zip(headers, row))
                if not any(fields.values()) or not fields.get('劣化ID'):
                    continue
                values = [fields.get(h) for h in ('劣化樣態', '維修工法', '維修單位', '單價')]
                if any(v is None or not str(v).strip() for v in values):
                    continue
                deterioration = str(fields.get('劣化類型') or '').strip()
                # 「適用構件」：只用於列出的構件；可列工作表名稱以外的構件（例：拱肋、斜撐、橫桿共用主梁工作表）
                only = [component_name(w) for w in words(fields.get('適用構件'))]
                targets = [c for c in dict.fromkeys(only) if c in sheet_targets or c in COMPONENTS] if only else sheet_targets
                result.append(dict(
                    zip(('pattern', 'method', 'unit', 'price'), (str(v).strip() for v in values)),
                    deterioration=deterioration, note=synthesize_note(fields, deterioration),
                    rule_id=str(fields['劣化ID']).strip(), structured={k: v for k, v in fields.items() if v not in (None, '')},
                    component=targets[0] if len(targets) == 1 else '', components=targets,
                    automatic_ready=bool(targets and deterioration), source=f'{sheet.title} 第{number}列'))
    finally:
        workbook.close()
    if not result:
        raise ValueError('找不到 v2 規則表（需有 劣化樣態、維修工法、維修單位、單價、處理方式、劣化ID 欄）')
    return result
