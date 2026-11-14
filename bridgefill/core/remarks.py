"""備註結構與量綱計算；只接受列舉的乘法格式，不執行文字運算式。"""
from dataclasses import dataclass
from decimal import Decimal
import re
from .data import compact, normalize_remark
from .locations import normalize_location,basic_location


@dataclass
class Remark:
    text: str
    location: str
    description: str
    amount: str
    unit: str


def number(value):
    return format(value.normalize(), 'f')


def crack_count(text):
    """有實測寬度與長度、未寫倍數時，明示原本隱含的一條。"""
    match=re.search(r'\(([^()]+)\)$',text)
    if match:
        parts=match[1].split('*')
        if len(parts)==2 and sum(bool(re.fullmatch(r'(?:\d+(?:\.\d+)?|\.\d+)mm',p)) for p in parts)==1 and sum(bool(re.fullmatch(r'(?:\d+(?:\.\d+)?|\.\d+)m',p)) for p in parts)==1:
            return text+'*1'
    return text


def parse(text, pattern, unit, aliases=()):
    text=normalize_remark(text)
    source = str(text).strip().rstrip('。．.').translate(str.maketrans({'＊': '*', '×': '*', '（': '(', '）': ')'}))
    if '\n' in source or len(source) > 3000:
        raise ValueError('多行或過長的備註需人工確認')
    match = re.fullmatch(r'(.+?)[。．.]?\s*(\([^()]+\)(?:\s*\*\s*\d+(?:\.\d+)?)?)\s*', source)
    if not match:
        raise ValueError('備註缺少可辨認的括號數量式')
    prefix, size = match.groups()
    prefix = prefix.strip().rstrip('。．. ')
    # 優先以明確空格切分；無空格時僅依唯一 Excel 樣態／備註關鍵字定位。
    split = re.match(r'^(\S+)\s+(.+)$', prefix)
    if split:
        location, description = split.groups()
        description = compact(description).rstrip('。．.')
    else:
        keys = sorted(set([compact(pattern)] + [compact(a) for a in aliases]), key=len, reverse=True)
        hits = [(prefix.find(k), k) for k in keys if k and prefix.find(k) > 0]
        if not hits:
            raise ValueError('位置與樣態間無空格，Excel 關鍵字無法辨認邊界')
        index = min(i for i, _ in hits)
        location, description = prefix[:index].strip(), compact(prefix[index:])
    if not location or not description or any(c in location for c in '()'):
        raise ValueError('無法辨認唯一缺失位置')
    # Excel 備註明確指定關鍵字替換時，採用該列的標準樣態。
    for alias in sorted(set(aliases), key=len, reverse=True):
        if alias and compact(alias) in description:
            description = compact(pattern)
    size = re.sub(r'\s+', '', size)
    normalized = normalize_remark(f'{location} {description}。{size}')
    size = normalized.split('。', 1)[-1]
    parts = re.fullmatch(r'\(([^()]+)\)(?:\*(\d+(?:\.\d+)?))?', size)
    if not parts:
        raise ValueError('數量式格式不明')
    factors = parts[1].split('*')
    if not 1 <= len(factors) <= 4:
        raise ValueError('數量式維度不明')
    values, units = [], []
    for factor in factors:
        item = re.fullmatch(r'(\d+(?:\.\d+)?|\.\d+)(mm|cm|m|公尺|平方公尺|立方公尺|組/塊|組／塊|公頃|[處座組個支面塊墩工天次式孔])?', factor)
        if not item:
            raise ValueError('數量式包含不支援的數字或單位')
        values.append(Decimal(item[1]))
        units.append(item[2] or '')
    total = Decimal(parts[2] or '1')
    if '裂縫' in description and units.count('mm') == 1 and units.count('m') == 1 and len(factors) in (2, 3):
        if len(factors) == 3 and units[-1]:
            raise ValueError('裂縫倍數不應有單位')
        for value, suffix in zip(values, units):
            if suffix != 'mm':
                total *= value
        derived = '公尺'
    elif any(suffix in ('m', 'cm', 'mm', '公尺') for suffix in units):
        dimensions = 0
        for value, suffix in zip(values, units):
            if suffix in ('m', 'cm', 'mm', '公尺'):
                dimensions += 1
                value *= {'m': Decimal(1), '公尺': Decimal(1), 'cm': Decimal('.01'), 'mm': Decimal('.001')}[suffix]
            elif suffix:
                raise ValueError('尺寸與計數單位混用')
            total *= value
        derived = {1: '公尺', 2: '平方公尺', 3: '立方公尺'}.get(dimensions)
        if not derived or any(units[i] for i in range(dimensions, len(units))) or len(factors) - dimensions > 1:
            raise ValueError('尺寸與倍數次序不明')
    elif len(factors) == 1 and units[0]:
        total *= values[0]
        derived = units[0].replace('／', '/')
    elif not any(units):
        dimensions = {'公尺': 1, '平方公尺': 2, '立方公尺': 3}.get(unit)
        if dimensions is None or len(factors) not in (dimensions, dimensions + 1):
            raise ValueError('無單位數量式無法依 Excel 唯一確認量綱')
        for value in values:
            total *= value
        derived = unit
    else:
        raise ValueError('數量式不屬於已定義的單位格式')
    return Remark(normalized, location, description, number(total), derived)


def rule_size_template(rule):
    if '裂縫' in rule.get('pattern','') or '裂縫' in rule.get('deterioration',''):
        return '(mm*m)*數量'
    unit=rule.get('unit','')
    counts={'處','座','組','個','式','支','面','塊','工','天','公頃','孔','組/塊','組／塊'}
    return {'平方公尺':'(m*m)','立方公尺':'(m*m*m)','公尺':'(m)'}.get(unit,'('+unit+')' if unit in counts else '(數量)')


def template_incomplete(text):
    text=normalize_remark(text)
    if text.startswith('損壞位置 '):return True
    match=re.search(r'\(([^()]*)\)(?:\*([^()\s]+))?$',text)
    if not match:return False
    if any(word in (match[1]+(match[2] or '')) for word in ('寬度','長度','高度','數量','處數')):return True
    units={'m','mm','cm','處','座','組','個','式','支','面','塊','工','天','公頃','孔','組/塊','組／塊'}
    return any(part in units for part in match[1].split('*'))


def prepare_fields(fields, rule=None, force_pattern=False, component='',context=None,sync_quantity=False):
    """格式及位置回填獨立於自動決策；人工指定規則可明確換成標準樣態。"""
    result = dict(fields)
    text = normalize_remark(result.get('remark',''))
    prefix=re.match(r'^(\S+)\s+(.+)$',text)
    if prefix:text=normalize_location(prefix[1],component,context)+' '+prefix[2]
    if result.get('d_location'):result['d_location']=normalize_location(result['d_location'],component,context)
    if rule:
        from .conditions import aliases
        try:
            info = parse(text, rule['pattern'], rule['unit'], aliases(rule.get('note','')))
            if force_pattern:
                info = parse(f'{info.location} {rule["pattern"]}。'+info.text.split('。',1)[-1],
                             rule['pattern'],rule['unit'])
            result.update(remark=info.text,d_location=info.location)
            if info.unit == rule['unit'].replace('／','/'):
                result['amount'] = info.amount
            elif rule['unit']=='式' and info.unit=='平方公尺' and any(word in rule['pattern'] for word in ('植物','植生')):
                result['amount']='1'
        except ValueError:
            # 空白新劣化或尚未填數字的範本，先帶入樣態及待填量式。
            size=re.search(r'\(([^()]*)\)',text)
            if force_pattern and (not text.strip() or (template_incomplete(text) and size and not re.search(r'\d',size[1]))):
                prefix=re.match(r'^(\S+)\s+',text)
                location=result.get('d_location','').strip() or (prefix[1] if prefix and prefix[1]!='損壞位置' else '') or basic_location(component,context)
                result['d_location']=location
                text=f'{location or "損壞位置"} {rule["pattern"]}。{rule_size_template(rule)}'
            # 量綱不符仍處理標點及明確位置，但不杜撰數量。
            if force_pattern:
                match = re.fullmatch(r'(\S+)\s+.+?[。．.]?\s*(\([^()]+\)(?:\s*\*\s*\d+(?:\.\d+)?)?)',text)
                if match:
                    text = normalize_remark(f'{match[1]} {rule["pattern"]}。{match[2]}')
            result['remark'] = text
    else:
        result['remark'] = text
    match = re.fullmatch(r'(\S+)\s+.+?。\([^()]+\)(?:\*\d+(?:\.\d+)?)?',result.get('remark',''))
    if match and match[1]!='損壞位置':
        result['d_location'] = match[1]
    if not rule and (sync_quantity or ('裂縫' in result.get('deterioration','') and result.get('unit')=='公尺')):
        try:
            info=parse(result['remark'],result.get('deterioration',''),result.get('unit',''))
            result.update(d_location=info.location,remark=info.text)
            if info.unit==result.get('unit','').replace('／','/'):result['amount']=info.amount
        except ValueError:pass
    if (rule or sync_quantity) and ('裂縫' in result.get('deterioration','') or (rule and '裂縫' in rule['pattern'])):
        result['remark']=crack_count(result.get('remark',''))
    return result


def crack_complete(fields):
    if '裂縫' not in fields.get('deterioration',''):return True
    text=normalize_remark(fields.get('remark',''))
    match=re.search(r'\(([^()]+)\)(?:\*(\d+(?:\.\d+)?))?$',text)
    if not match:return False
    parts=match[1].split('*')
    return (len(parts) in (2,3) and sum(bool(re.fullmatch(r'(?:\d+(?:\.\d+)?|\.\d+)mm',x)) for x in parts)==1
        and sum(bool(re.fullmatch(r'(?:\d+(?:\.\d+)?|\.\d+)m',x)) for x in parts)==1
        and (len(parts)==2 or bool(re.fullmatch(r'\d+(?:\.\d+)?',parts[-1]))))


def crack_template(fields):
    location=fields.get('d_location','')
    text=normalize_remark(fields.get('remark',''))
    match=re.match(r'^(\S+)\s+',text)
    if match:location=match[1]
    return f'{location or "損壞位置"} 混凝土裂縫。(mm*m)*數量'
