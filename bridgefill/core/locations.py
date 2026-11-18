"""只正規化已定義的構件編號與位置；保留位置尾字和範圍。"""
import re
import unicodedata

HEADS=('混凝土','鋼筋','鋼鈑','鋼管','護欄','木本','草本','植生','AC','瀝青','生銹','生鏽','銹蝕','鏽蝕','漏水','滲水','積水','堵塞','龜裂','裂縫','剝落','破損','破裂','斷裂','變形','螺栓','塗裝','防落','鋼纜')
TAILS='基礎樁帽|左右側|左側|右側|胸牆|耳牆|頰牆|基礎頂|基礎板|基座|基樁|帽梁|正背面|正背|正面|背面|正|背|上下游|上游|下游|上部|下部|梁底|頂部|底部|中間|左右|上|下|內|外|側|全長|全區'


def position_separator(value):
    # 只分隔已定義的位置尾字，不把「混凝土」等缺失敘述當作位置。
    return re.sub(r'((?:[APS]\d+(?:[GD]\d+)?|\d+[GD]\d+)(?:-[A-Za-z0-9,]+)?)[\s]*(?=(?:'+TAILS+r')(?:$|[\s,，、／/]|混凝土|鋼筋|基礎|木本|草本|生銹|漏水|滲水))',lambda m:m[1] if m[1].endswith(',') else m[1]+'-',str(value or ''),flags=re.I)


def key(value):
    value=re.sub(r'\s+','',unicodedata.normalize('NFKC',str(value or ''))).upper()
    value=position_separator(value)
    value={'橋頭':'A1','橋尾':'A2'}.get(value,value)
    return re.sub(r'^([APS])0*(\d+)',lambda m:m[1]+str(int(m[2])),value)


def location_prefix(prefix,patterns=()):
    split=re.fullmatch(r'(\S+)\s+(.+)',prefix.strip())
    if split:return split[1],split[2]
    hits=[]
    for word in (*patterns,*HEADS):
        word=re.sub(r'\s+','',str(word or ''))
        start=prefix.find(word)
        if start>0:
            position=prefix[:start]
            if re.fullmatch(r'(?:[A-Za-z0-9,，.／/~～_+\-]|'+TAILS+r')+',position):
                hits.append((start,word))
    if not hits:return None
    index=min(start for start,_ in hits)
    return prefix[:index],prefix[index:]


def spans(context):
    result=set()
    for label,value in (context or {}).items():
        if label.startswith('_') or not any(word in label for word in ('位置','編號','跨','橋孔')):continue
        for part in str(value).split('｜'):
            match=re.match(r'^S0*(\d+)(?:$|[-GD])',key(part))
            if match:result.add(int(match[1]))
    return result


def normalize_location(value,component='',context=None):
    value=position_separator(str(value or '').strip())
    if component=='主梁':
        value=re.sub(r'^S0*(\d+)-?G(\d+)',lambda m:f'{int(m[1])}G{m[2]}',value,flags=re.I)
    elif component=='橫隔梁':
        value=re.sub(r'^(\d+)D-(\d+-\d+)',r'\1D\2',value,flags=re.I)     # 1D-2-1、7D-5-1 → 1D2-1、7D5-1
        match=re.match(r'^S0*(\d+)-(?:D)?(\d+-\d+)',value,re.I)
        if match:value=f'{int(match[1])}D{match[2]}'+value[match.end():]
        elif re.match(r'^\d+-\d+',value):
            span=spans(context)
            if len(span)==1:value=f'{next(iter(span))}D'+value
    return value


def history_spans(item):
    result=spans(item.context)
    if not result:
        location=item.fields.get('d_location','') or item.fields.get('remark','').split(' ',1)[0]
        match=re.match(r'^(?:S0*(\d+)(?:$|[-GD])|(\d+)[GD])',key(location))
        if match:result.add(int(match[1] or match[2]))
    return result


def basic_location(component,context=None):
    """僅供空白新缺失的可編輯建議，不改變前期身份比對。"""
    context=context or {}
    numbers=[]
    for label,value in context.items():
        if label.startswith('_') or '編號' not in label:continue
        numbers.extend(part.strip() for part in str(value).split('｜') if part.strip())
    elements=[v for v in numbers if not re.fullmatch(r'S\d+',key(v))]
    candidates=list(dict.fromkeys(elements or numbers))
    if not candidates:
        for label,value in context.items():
            if '位置' in label and not label.startswith('_'):
                candidates.extend(part.strip() for part in str(value).split('｜') if
                    re.fullmatch(r'[APS]\d+(?:[-A-Za-z0-9,~]*)?',part.strip(),re.I) or part.strip() in ('橋頭','橋尾'))
    if not candidates:return ''
    longest=max(candidates,key=len)
    if not all(longest==candidate or (longest.startswith(candidate) and not longest[len(candidate)].isdigit()) for candidate in candidates):return ''
    value={'橋頭':'A1','橋尾':'A2'}.get(longest,longest)
    value=re.sub(r'^([APS])0*(\d+)',lambda m:m[1].upper()+str(int(m[2])),value,flags=re.I)
    span=spans(context)
    if component=='主梁' and re.fullmatch(r'G\d+',value,re.I) and len(span)==1:
        value=f'S{next(iter(span))}'+value
    return normalize_location(value,component,context)


def physical_owner(context=None, fields=None):
    """Exact abutment/pier identity; leading zeroes never imply a prefix match."""
    context=context or {};fields=fields or {}
    values=[str(v) for k,v in context.items() if any(w in k for w in ('位置','編號'))]
    family='P' if any('橋墩' in v for v in values) else 'A' if any('橋台' in v for v in values) else ''
    if not family:
        match=re.match(r'^([AP])\d+',key(fields.get('d_location','')))
        if match:family=match[1]
    owners=set()
    for value in values:
        for part in value.split('｜'):
            m=re.fullmatch(r'([AP])0*(\d+)(?:N)?',key(part))
            if m:owners.add(m[1]+str(int(m[2])))
            elif family and re.fullmatch(r'\d+',part.strip()):owners.add(family+str(int(part)))
    if len(owners)==1:return next(iter(owners))
    if owners:return '|'.join(sorted(owners))
    m=re.match(r'^([AP])0*(\d+)(?!\d)',key(fields.get('d_location','')))
    return m[1]+str(int(m[2])) if m else ''


def short_code(component, context=None):
    """清單用的簡短位置代碼（2026-11-17 使用者指定）：橋頭＝A1、橋尾＝A2，去掉欄名與重複編號。
    例：引道路堤 A1；伸縮縫「位置：橋台(A)｜編號：001」→ A1；墩柱與帽梁 P2/P02N-1；主梁 S3/G1；
    橋面板 S4；立柱 S1/R01-1H01；其他 → 設施類別（設施編號無意義）。"""
    context = context or {}
    if component == '其他':
        return next((str(v).strip() for k, v in context.items() if '設施類別' in k and str(v).strip()), '')
    owner, family, numbers, elements = '', '', [], []
    for label, value in context.items():
        if label.startswith('_') or not re.search(r'位置|編號', label):
            continue
        for part in str(value or '').split('｜'):
            part = unicodedata.normalize('NFKC', part).strip()
            if not part:
                continue
            if part in ('橋頭', '橋尾'):
                owner = owner or {'橋頭': 'A1', '橋尾': 'A2'}[part]
                continue
            kind = re.search(r'\(([APS])\)$', part) or re.fullmatch(r'(橋台|橋墩)', part)
            if kind:
                family = {'橋台': 'A', '橋墩': 'P'}.get(kind[1], kind[1])
                continue
            single = re.fullmatch(r'([APS])0*(\d+)', part, re.I)
            if single:
                code = single[1].upper() + single[2]
                if not owner:
                    owner = code
                elif code != owner:
                    elements.append(part)
                continue
            if re.fullmatch(r'\d+', part):
                numbers.append(part)
                continue
            elements.append(part)
    if family and numbers and not owner:
        owner = family + str(int(numbers.pop(0)))
    elements += numbers
    result = []
    for element in dict.fromkeys(elements):
        normal = re.sub(r'^([APS])0*(\d+)', lambda m: m[1].upper() + m[2], element, flags=re.I)
        if owner and normal == owner:
            continue
        if owner and normal.startswith(owner) and re.match(r'-?[GD]\d', normal[len(owner):], re.I):
            element = normal[len(owner):].lstrip('-')            # 主梁 S3G1 → G1
        if any(other != element and other.startswith(element) for other in elements):
            continue                                             # 拱圈 R01-1 已含在立柱 R01-1H01
        result.append(element)
    return '/'.join(([owner] if owner else []) + result)
