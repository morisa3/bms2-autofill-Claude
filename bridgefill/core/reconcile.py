"""前期完整資料與本期輸入協調；數量相同的摘要可還原上期量式。"""
import re
import unicodedata
from .data import compact, normalize_remark, quantity, component_name
from .locations import key as identity_key,spans,history_spans,normalize_location,physical_owner


def text_key(value):
    return compact(unicodedata.normalize('NFKC', str(value or '')))


def deterioration_key(value):
    return text_key(value).replace('混凝土結構裂縫','混凝土裂縫')


def numeric(value):
    value = text_key(value).replace(',', '')
    match = re.fullmatch(r'(\d+(?:\.\d+)?)(?:mm|cm|m|平方公尺|立方公尺|公尺|處|座|組|個|式)?', value)
    return quantity(match[1]) if match else None


def complete_remark(value):
    return bool(re.fullmatch(r'.+?\([^()]+\)(?:\s*\*\s*\d+(?:\.\d+)?)?', str(value or '').strip().replace('（','(').replace('）',')').replace('＊','*').replace('×','*').rstrip('。．.')))


def description(value, location=''):
    text = normalize_remark(value)
    text = re.sub(r'[。.]?\([^()]+\)(?:\*\d+(?:\.\d+)?)?$', '', text).strip().rstrip('。.')
    if location and text.startswith(location+' '):
        text = text[len(location):].strip()
    else:
        text = re.sub(r'^[A-Za-z0-9][^\s]*\s+', '', text)
    return text_key(text)


def scoped_history(history, component, context=None, deterioration='', remark='', include_status=False):
    context = context or {}
    def place_key(value):
        return identity_key(normalize_location(value,component,context))
    number_values = [part for key,value in context.items() if '編號' in key
                     for part in str(value).split('｜') if part]
    positions = [part for key,value in context.items() if '位置' in key
                 for part in str(value).split('｜') if part]
    positions = [value for value in positions if text_key(value) not in ('主','主構件','主結構','擴建','擴建結構','左側','右側')]
    if not number_values:
        number_values = [place_key(value) for value in positions if text_key(value) in ('橋頭','橋尾')]
    result = []
    for item in history:
        if item.component != component_name(component):
            continue
        if deterioration and deterioration_key(item.fields.get('deterioration')) != deterioration_key(deterioration):
            continue
        info = getattr(item, 'context', {})
        if info.get('_inspection_status') and not include_status:continue
        current_spans=spans(context)
        old_spans=history_spans(item)
        if current_spans and old_spans and not current_spans.intersection(old_spans):continue
        if current_spans and not old_spans and component in ('主梁','橫隔梁','橋面板'):continue
        owner=physical_owner(context);old_owner=physical_owner(info)
        if owner and old_owner:
            if owner!=old_owner:continue
            def members(info):
                return {identity_key(part) for k,v in info.items() if '編號' in k for part in str(v).split('｜')
                        if not re.fullmatch(r'(?:[AP])?0*\d+',identity_key(part))}
            left,right=members(context),members(info)
            if left and right and not left.intersection(right):continue
            result.append(item);continue
        element_numbers=[v for v in number_values if not re.fullmatch(r'S\d+',identity_key(v))]
        if element_numbers and info.get('編號') and not re.fullmatch(r'S\d+',identity_key(info['編號'])) and place_key(info['編號']) not in map(place_key,element_numbers):
            continue
        if positions and info.get('位置') and place_key(info['位置']) not in map(place_key, positions):
            continue
        result.append(item)
    style = description(remark) if remark else ''
    precise = [item for item in result if style and style == description(item.fields.get('remark',''), item.fields.get('d_location',''))]
    return precise or result


def type_repairs(repairs, deterioration):
    if not deterioration:
        return list(repairs)
    return [r for r in repairs if not r.get('deterioration') or
            deterioration_key(r['deterioration']) == deterioration_key(deterioration)]


def summary_of_old(current, previous):
    remark = current.get('remark', '').strip()
    old = normalize_remark(previous.get('remark', ''))
    if not remark or normalize_remark(remark) == old:
        return True
    if complete_remark(remark):
        return False
    loc, _, tail = old.partition(' ')
    description = re.sub(r'[。.]?\([^()]+\)(?:\*\d+(?:\.\d+)?)?$', '', tail).rstrip('。.')
    remaining = text_key(remark)
    for word in (loc, description, previous.get('deterioration','')):
        if word:
            remaining = remaining.replace(text_key(word), '')
    remaining = re.sub(r'數量(?:填寫|填入|填)?|單位|[。:：,，+＋]', '', remaining)
    remaining = remaining.rstrip('。．.')
    return not remaining or numeric(remaining) == numeric(previous.get('amount',''))


def reconcile(current, previous=None):
    draft = dict(current)
    if previous is None:
        return draft, '' if complete_remark(draft.get('remark')) else '本期備註不完整，且無唯一同樣態上期資料'
    old = previous.fields
    if current.get('deterioration') and deterioration_key(current['deterioration']) != deterioration_key(old.get('deterioration')):
        return draft, '本期與上期劣化類型不同，請人工確認'
    if complete_remark(current.get('remark')) and normalize_remark(current['remark']) != normalize_remark(old.get('remark','')):
        # 明確的新位置／範圍優先，保留本期再套用規則。
        return draft, ''
    if current.get('amount') and numeric(current['amount']) != numeric(old.get('amount')):
        return draft, '本期數量與上期不同，且缺少完整新範圍，請人工確認'
    if not summary_of_old(current, old):
        return draft, '本期敘述無法確認與上期相同，請人工確認'
    draft.update(old)
    draft['remark'] = normalize_remark(old.get('remark',''))
    if not complete_remark(draft['remark']):
        return draft, '上期備註也缺少完整位置與數量式，請人工確認'
    draft['d_location'] = draft['remark'].split(None,1)[0]
    return draft, ''
