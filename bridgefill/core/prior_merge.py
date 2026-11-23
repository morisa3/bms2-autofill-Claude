"""本期資料不完整時，依前期補齊（在 v1.00 判斷之前執行）。

規則（使用者 2026-11-17 指定）：
1. 前期備註的尺寸換算＝本期數量（本期備註空白）：視為同一缺失，導入前期位置與敘述、前期尺寸，數量與單位用本期。
   例：前期「A2 沖蝕 (10m*2m)」1 式；本期 20 平方公尺 → 「A2 沖蝕。(10m*2m)」20 平方公尺。
2. 反過來：本期備註只有尺寸（或只有位置＋尺寸），尺寸換算＝前期數量：導入前期位置與敘述，尺寸用本期。
3. 劣化類型一致、單位一致，本期備註只留新尺寸（例：5*7），前期為 4*5：先視為與前期同一缺失——
   導入前期位置與敘述，範圍改為新尺寸並補上括號與單位（5*7 → (5m*7m)），數量依新尺寸計算。
其後仍由 v1.00 的判斷與 Excel 規則決定工法、單位、單價與備註寫法。
"""
import re
from decimal import Decimal, InvalidOperation

from .data import compact
from .editing import split_remark
from .matching import strip_measure
from .reconcile import deterioration_key
from .remarks import parse

DIMENSIONS = {'公尺': 1, '平方公尺': 2, '立方公尺': 3}
COUNT_UNITS = ('處', '個', '支', '組', '座', '塊', '根', '條', '片', '孔', '只', '式')
SIZE = re.compile(r'^[\(（]?\s*[\d.]+\s*(?:mm|cm|m|處|個|組|座|支|塊)?\s*(?:[*＊×xX]\s*[\d.]+\s*(?:mm|cm|m|處|個|組|座|支|塊)?\s*)*[\)）]?'
                  r'(?:\s*[*＊×]\s*\d+(?:\.\d+)?)?$')


def clean(value):
    return str(value or '').strip()


def number(value):
    try:
        return Decimal(clean(value))
    except InvalidOperation:
        return None


def size_part(remark):
    """備註最後的尺寸式（含括號與倍數），沒有則空白。"""
    found = re.search(r'[\(（][^()（）]+[\)）](?:\s*[*＊×]\s*\d+(?:\.\d+)?)?\s*$', clean(remark))
    return found[0].strip() if found else ''


def quantity(description, size, unit):
    """尺寸換算成 unit 的數量；無法換算回傳 None。"""
    if not size:
        return None
    try:
        info = parse(f'X {description or "缺失"}。{size}', description or '缺失', unit)
    except ValueError:
        return None
    if info.unit != unit:
        return None
    return number(info.amount)


def normalize_size(text, unit):
    """「5*7」→「(5m*7m)」；「(5m*7m)」照舊；倍數保留（5*7*2 → (5m*7m)*2）。無法判斷回傳空白。"""
    text = clean(text).translate(str.maketrans({'＊': '*', '×': '*', 'x': '*', 'X': '*', '（': '(', '）': ')'}))
    text = re.sub(r'\s+', '', text)
    if re.fullmatch(r'\([^()]+\)(\*\d+(\.\d+)?)?', text):
        return text
    parts = text.strip('()').split('*')
    if not all(re.fullmatch(r'[\d.]+(mm|cm|m|處|個|組|座|支|塊)?', p) for p in parts):
        return ''
    if len(parts) > 1 and re.fullmatch(r'[\d.]+', parts[-1]) and all(re.search(r'[a-z處個組座支塊]', p) for p in parts[:-1]):
        return '(' + '*'.join(parts[:-1]) + ')*' + parts[-1]   # 每個數字都有單位＋倍數（例：0.5mm*2m*2 → (0.5mm*2m)*2）
    if all(re.search(r'[a-z處個組座支塊]', p) for p in parts):
        return '(' + '*'.join(parts) + ')'          # 每個數字都有單位（例：0.3mm*2m）：照原樣補括號
    dims = DIMENSIONS.get(unit)
    if dims:
        if len(parts) not in (dims, dims + 1):
            return ''
        body = [p if re.search(r'[a-z]', p) else p + 'm' for p in parts[:dims]]
        extra = parts[dims:]
        return '(' + '*'.join(body) + ')' + (f'*{extra[0]}' if extra else '')
    if unit in COUNT_UNITS and len(parts) == 1:
        return f'({parts[0]}{"" if re.search(r"[^\d.]", parts[0]) else unit})'
    return ''


def size_numbers(text):
    """尺寸式中的數字（倍數不算）：「(4m*3m*1.2m)」→ ['4', '3', '1.2']。"""
    body = re.sub(r'[\)）]\s*[*＊×]\s*\d+(?:\.\d+)?\s*$', ')', clean(text))
    return re.findall(r'\d+(?:\.\d+)?', body)


SIZE_LIKE = re.compile(r'(?:約|深度?|高度?|[HD]\s*[=＝:：])?\s*\d+(?:\.\d+)?\s*m?(?:\s*[*＊×xX]\s*\d+(?:\.\d+)?\s*m?)*', re.I)


def volume_size(measure, old_size):
    """前期為「長*寬*深」三維尺寸時，本期沒寫單位的尺寸補成前期格式（2026-11-25 使用者，橋墩基礎常見）：
    - 本期也是三個數字（4*3*1.5）→ (4m*3m*1.5m)；
    - 本期只有一個數字（本期深度 1.5、H=1.5m、深1.5m）→ 長寬沿用前期：(4m*3m*1.5m)。
    無法判斷回傳空白。"""
    old = size_numbers(old_size)
    new = re.findall(r'\d+(?:\.\d+)?', clean(measure))
    if len(old) != 3 or not new or re.search(r'[\(（]', clean(measure)) and re.search(r'(?:mm|cm|m)\s*[\)）]', clean(measure)):
        return ''
    if len(new) == 3:
        return '(' + '*'.join(f'{n}m' for n in new) + ')'
    if len(new) == 1:
        return f'({old[0]}m*{old[1]}m*{new[0]}m)'
    return ''


def only_size(remark):
    """本期備註只有尺寸（或位置＋尺寸），沒有劣化敘述：回傳 (位置, 尺寸文字)；否則 None。"""
    text = clean(remark).rstrip('。')
    if not text:
        return None
    if SIZE.match(text):
        return '', text
    location, description, tail = split_remark(text)
    if location and (not description or SIZE.match(description)) and (tail or description):
        return location, (description + tail).strip()
    return None


def fitting_unit(description, size, amount, preferred=()):
    """前期尺寸換算等於本期數量的單位（本期單位空白、或前期為「式」時用）；優先試 preferred。"""
    if amount is None or not size:
        return ''
    for unit in dict.fromkeys([u for u in preferred if u] + list(DIMENSIONS) + list(COUNT_UNITS)):
        if unit == '式':
            continue
        if quantity(description, size, unit) == amount:
            return unit
    return ''


def merge_with_prior(current, prior):
    """回傳 (補齊後的本期欄位, 說明)；不適用時回傳 (None, '')。current／prior 為欄位 dict。
    規則（使用者 2026-11-18 整理）：前期劣化類型相同時——
    A. 本期備註空白：
       - 數量也空白 → 直接導入前期（位置、敘述、尺寸、數量、單位）；
       - 前期尺寸換算＝本期數量 → 導入前期，單位空白或前期為「式」時依換算結果補單位；
       - 計數單位（處、個…）→ 導入前期位置與敘述，尺寸改為本期數量。
    B. 本期備註只有尺寸（例：0.5mm*2m*2）：位置與敘述沿用前期，範圍用本期（單位空白視為與前期相同）。
    C. 本期備註有敘述但沒有位置（例：漏水。、龜裂。(12m＊12m)）：補上前期位置，尺寸沒有時依數量補。
    其後仍由 v1.00 的判斷與 Excel 規則決定工法、單位、單價與備註寫法。"""
    if not prior:
        return None, ''
    kind, old_kind = clean(current.get('deterioration')), clean(prior.get('deterioration'))
    if kind and old_kind and deterioration_key(kind) != deterioration_key(old_kind):
        return None, ''
    old_location, old_description, old_tail = split_remark(prior.get('remark'))
    old_size = size_part(prior.get('remark'))
    if not old_description:
        return None, ''
    old_description = old_description.rstrip('。. ')
    unit, amount = clean(current.get('unit')), number(current.get('amount'))
    old_unit, old_amount = clean(prior.get('unit')), number(prior.get('amount'))
    remark = clean(current.get('remark'))
    location = clean(current.get('d_location')) or old_location

    # A：本期備註空白
    if not remark:
        if not location:
            return None, ''
        if amount is None:
            if not old_size and old_unit not in COUNT_UNITS:
                return None, ''
            size = old_size or (f'({old_amount.normalize():f}{old_unit})' if old_amount is not None else '')
            fields = dict(current, remark=f'{location} {old_description}。{size}', d_location=location,
                          amount=clean(prior.get('amount')), unit=old_unit)
            return fields, '本期備註與數量空白，已導入前期資料'
        fit = fitting_unit(old_description, old_size, amount, (unit, old_unit if old_unit != '式' else ''))
        if old_size and fit and (not unit or unit == fit or unit == '式'):
            fields = dict(current, remark=f'{location} {old_description}。{old_size}', d_location=location, unit=fit)
            note = '' if unit == fit else f'（單位依換算補為 {fit}）'
            return fields, (f'前期尺寸 {old_size} 換算為 {amount.normalize():f}{fit}，與本期數量相同，'
                            f'已導入前期位置與敘述{note}')
        count_unit = unit or old_unit
        if count_unit in COUNT_UNITS and count_unit != '式':
            size = f'({amount.normalize():f}{count_unit})'
            fields = dict(current, remark=f'{location} {old_description}。{size}', d_location=location, unit=count_unit)
            return fields, f'已導入前期位置與敘述，數量改為本期 {amount.normalize():f}{count_unit}'
        return None, ''

    # B：本期備註只有尺寸
    found = only_size(remark)
    if found is not None:
        here, size_text = found
        location = here or location
        target_unit = unit or old_unit
        size = normalize_size(size_text, target_unit if target_unit != '式' else '') or volume_size(size_text, old_size)
        if not size or not location:
            return None, ''
        new_quantity = quantity(old_description, size, target_unit) if target_unit else None
        if new_quantity is None and target_unit != '式':           # 「式」維持式，不改成立方公尺等（2026-11-25）
            for candidate in ([unit] if unit else []) + list(DIMENSIONS):
                new_quantity = quantity(old_description, size, candidate)
                if new_quantity is not None:
                    target_unit = candidate
                    break
        same_unit = not unit or unit == old_unit
        equal_prior = new_quantity is not None and new_quantity == old_amount
        if not (same_unit or equal_prior):
            return None, ''
        fields = dict(current, remark=f'{location} {old_description}。{size}', d_location=location, unit=target_unit)
        if new_quantity is not None:
            fields['amount'] = f'{new_quantity.normalize():f}'
        why = (f'本期尺寸 {size} 換算與前期數量相同' if equal_prior else f'本期只填新範圍 {size}，位置與敘述沿用前期')
        return fields, why + f'，已導入前期位置「{location}」與敘述「{old_description}」'

    # D：本期敘述後直接接沒寫單位的尺寸或只寫本期深度（4*3*1.5、1.5、H=1.5m），前期為長*寬*深（2026-11-25）
    here, description, tail = split_remark(remark)
    if description and not size_part(remark):
        bare = strip_measure(description)
        measure = (description[len(bare):] if description.startswith(bare) else '') + re.sub(r'^[。.\s]+', '', tail)
        measure = measure.strip(' ，,、。')
        size = volume_size(measure, old_size) if bare and measure and SIZE_LIKE.fullmatch(measure) else ''
        if size and (here or location):
            location = here or location
            fields = dict(current, remark=f'{location} {bare.rstrip("。. ，,")}。{size}', d_location=location)
            target = unit or old_unit
            if not unit and old_unit:
                fields['unit'] = old_unit
            if target == '立方公尺':
                new_quantity = quantity(bare, size, target)
                if new_quantity is not None:
                    fields['amount'] = f'{new_quantity.normalize():f}'
            single = len(re.findall(r'\d+(?:\.\d+)?', measure)) == 1
            why = (f'本期只寫本期深度（{measure.strip()}），長寬沿用前期 {old_size}，範圍改為 {size}' if single
                   else f'本期尺寸「{measure.strip()}」沒寫單位，依前期格式補為 {size}')
            return fields, why

    # C：本期有敘述但沒有位置
    if not here and description and location:
        description = description.rstrip('。. ')
        size = size_part(remark) or normalize_size(re.sub(r'^[。.\s]+', '', tail), unit or old_unit)
        if not size and amount is not None and (unit or old_unit) in COUNT_UNITS:
            size = f'({amount.normalize():f}{unit or old_unit})'
        elif not size and amount is not None and (unit or old_unit) == '公尺':
            size = f'({amount.normalize():f}m)'                  # 長度：數量即範圍（例：伸縮縫 5 公尺 → (5m)）
        size = normalize_size(size, unit or old_unit) or size
        fields = dict(current, remark=f'{location} {description}。{size}', d_location=location)
        if not unit and old_unit and size:
            fields['unit'] = old_unit
        return fields, f'本期備註沒有位置，已補上前期位置「{location}」'
    return None, ''
