"""同一本期資料的良好／缺失矛盾，只提示人工核對，不改寫 D/E。"""
import re
from decimal import Decimal,InvalidOperation
from .locations import key,physical_owner
from .data import component_name


def owner_key(record):
    context=record.get('context',{});numbers=[];positions=[]
    owner=physical_owner(context,record.get('fields',record.get('current',{})))
    if owner:return component_name(record['component']),(owner,)
    for label,value in context.items():
        if label.startswith('_'):continue
        parts=[part.strip() for part in str(value or '').split('｜') if part.strip()]
        if '編號' in label:numbers.extend(parts)
        elif '位置' in label:positions.extend(parts)
    candidates=numbers or positions
    if not candidates:
        fields=record.get('fields',record.get('current',{}));value=fields.get('d_location','')
        if value:candidates=[re.split(r'[-－]',value,maxsplit=1)[0]]
    if not candidates:return None
    def normalized(value):
        value=key(value);match=re.fullmatch(r'[AP]?0*(\d+)',value)
        return str(int(match[1])) if match else value
    return component_name(record['component']),tuple(sorted(set(map(normalized,candidates))))


def mixed_ratings(entries,good):
    keys=set()
    for entry in good:
        try:is_good=Decimal(str(entry.get('der',[''])[0]))==1
        except (InvalidOperation,IndexError):is_good=False
        if is_good and owner_key(entry):keys.add(owner_key(entry))
    return [dict(index=i,component=e['component'],location=e['fields'].get('d_location',''),detail='同構件同位置同時有 D＝1 與 D≥2，請到網站核對 D/E；程式不會自行修改評等') for i,e in enumerate(entries) if owner_key(e) in keys or e.get('context',{}).get('_mixed_d1')]
