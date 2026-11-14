"""僅供前後期對照的一對一配對；不以左欄資料覆寫本期。"""
from difflib import SequenceMatcher
import re
from .data import component_name,quantity,normalize_remark
from .reconcile import description,scoped_history,text_key
from .locations import key as location_key,physical_owner
from .reconcile import deterioration_key
from .data import repair_candidates
from .conditions import TYPE_CORRECTION,rule_deterioration


def location_conflict(record):
    fields=record.get('current',record.get('fields',{}))
    location=fields.get('remark','').strip() or fields.get('d_location','')
    actual=re.match(r'^([AP]\d+)',location_key(location))
    expected=set()
    for label,value in record.get('context',{}).items():
        if not any(word in label for word in ('位置','編號')):continue
        for part in str(value).split('｜'):
            match=re.match(r'^([AP]\d+)',location_key(part))
            if match:expected.add(match[1])
    return bool(actual and expected and actual[1] not in expected)


def pattern(fields):
    # 量式後的設計圖說不是樣態；無空格的 P6基礎裸露也須拆開。
    text=normalize_remark(fields.get('remark',''))
    text=re.split(r'\(',text,maxsplit=1)[0].strip().rstrip('。.')
    location=fields.get('d_location','')
    if location and text.startswith(location+' '):text=text[len(location):].strip()
    text=re.sub(r'^[APS]0*\d+(?:-[A-Za-z0-9]+)?[-]?','',text,flags=re.I).strip()
    text=re.sub(r'^(?:基樁|上游|下游)\s+','',text)
    text=re.sub(r'^橋墩基礎(?=混凝土|鋼筋)','',text)
    if not text:
        kind=text_key(fields.get('deterioration',''))
        text='' if any(mark in kind for mark in ('、','或')) or kind in ('其他損傷','其他','表面劣化','路堤損傷','保護設施損壞、移動或遺失') else kind
    text=re.sub(r'[、，,。．.；;：:\s/／()（）]','',text).upper()
    for a,b in (('生鏽','生銹'),('生锈','生銹'),('滲水','漏水'),('剝離','剝落'),('破損','剝落'),('木本植生','木本植物生長'),('草本植生','草本植物生長')):
        text=text.replace(a,b)
    return text


def virtual_type(fields,repairs,component):
    original=deterioration_key(fields.get('deterioration'))
    replacements={deterioration_key(rule_deterioration(r,fields.get('deterioration',''))) for r in repair_candidates(repairs,fields,component,filter_unit=False)
                  if TYPE_CORRECTION.search(text_key(r.get('note',''))) and original in tuple(deterioration_key(x) for x in TYPE_CORRECTION.search(text_key(r['note'])).groups())}
    return next(iter(replacements)) if len(replacements)==1 else original


def virtual_methods(fields,repairs,component):
    rules=repair_candidates(repairs,fields,component,filter_unit=False)
    kind=virtual_type(fields,repairs,component)
    return {text_key(r.get('method',r.get('repair',''))) for r in rules
            if deterioration_key(r.get('deterioration'))==kind} - {''}


def pair_score(current,previous,repairs=(),component=''):
    if virtual_type(current,repairs,component)!=virtual_type(previous,repairs,component):return None
    a,b=pattern(current),pattern(previous)
    for x,y in (('木本','草本'),('齒型','角鋼'),('AC','混凝土'),('瀝青','混凝土'),('鋼管','鋼筋'),('鋼板','鋼管')):
        if (x in a and y in b) or (y in a and x in b):return None
    same_type=text_key(current.get('deterioration'))==text_key(previous.get('deterioration'))
    kind=text_key(current.get('deterioration',''))
    if same_type and '白華' in kind:
        if ('白華' in a)!=('白華' in b):return None
        if '白華' in a and '白華' in b:a=b='白華'
    if same_type and '河道' in kind:
        a=re.sub(r'^河道(?=淤積|沖刷|堆積|堵塞|阻塞|變遷)','',a)
        b=re.sub(r'^河道(?=淤積|沖刷|堆積|堵塞|阻塞|變遷)','',b)
    same_amount=bool(str(current.get('amount','')).strip()) and quantity(current.get('amount'))==quantity(previous.get('amount'))
    core=lambda text:re.sub(r'防衝鋼板|防撞鋼板|混凝土|結構|表面|AC|瀝青','',text)
    ca,cb=core(a),core(b)
    similarity=max(SequenceMatcher(None,a,b,autojunk=False).ratio(),SequenceMatcher(None,ca,cb,autojunk=False).ratio()) if a and b and ca and cb else 0
    semantic=3 if a and b and (a==b or (ca and ca==cb)) else 2 if a and b and similarity>=.78 else 0
    if not semantic and same_type and same_amount and all('鋼筋' in s and any(x in s for x in ('裸露','外露')) for s in (a,b)):
        semantic=1
    if not semantic and same_type and same_amount and '混凝土' in current.get('deterioration','') and min(len(ca),len(cb))>=2 and (ca in cb or cb in ca):
        semantic=1
    if not semantic and same_type and '基礎' in current.get('deterioration','') and all('基礎' in s and any(x in s for x in ('裸露','掏空','沖刷')) for s in (a,b)):
        semantic=1
    plants=('植生','植物','木本','草本')
    if not semantic and same_type and same_amount and any(x in a for x in plants) and any(x in b for x in plants):
        # 泛稱植生可與唯一明確植生配對；不把明確木本與草本互換。
        if not (('木本' in a and '草本' in b) or ('草本' in a and '木本' in b)):semantic=1
    if not semantic and virtual_methods(current,repairs,component).intersection(virtual_methods(previous,repairs,component)):semantic=.5
    if not semantic:return None
    remark_similarity=SequenceMatcher(None,text_key(normalize_remark(current.get('remark',''))),text_key(normalize_remark(previous.get('remark',''))),autojunk=False).ratio()
    same_location=bool(current.get('d_location')) and location_key(current['d_location'])==location_key(previous.get('d_location',''))
    return int(same_location),semantic,int(same_amount),int(bool(current.get('unit')) and text_key(current.get('unit'))==text_key(previous.get('unit'))),round(remark_similarity,3)


def align_records(records,history,manual=None,repairs=(),blocked_history=(),current_statuses=()):
    manual=manual or {}
    valid=[i for i,h in enumerate(history) if not h.context.get('_inspection_status') and (h.fields.get('deterioration') or h.fields.get('remark'))]
    used=set();assigned={};candidates={};edges={};conflicts=set()
    for ci,record in enumerate(records):
        if record['id'] in manual:
            hi=manual[record['id']]
            if hi is not None:
                if hi not in valid or hi in used:raise ValueError('手動對應的前期資料不存在或已被其他本期使用')
                assigned[ci]=hi;used.add(hi)
    for ci,record in enumerate(records):
        if record['id'] in manual:continue
        if location_conflict(record):conflicts.add(ci);continue
        fields=record.get('current',record.get('fields',{}))
        scope=scoped_history([history[i] for i in valid],record['component'],record.get('context',{}))
        allowed={id(h) for h in scope}
        for hi in valid:
            if hi in used or hi in blocked_history or id(history[hi]) not in allowed:continue
            own=physical_owner(record.get("context"),fields);old=physical_owner(history[hi].context,history[hi].fields)
            if own and old and own!=old:continue
            score=pair_score(fields,history[hi].fields,repairs,record["component"])
            same_place=(own and old and own==old) or (fields.get('d_location') and location_key(fields['d_location'])==location_key(history[hi].fields.get('d_location','')))
            if score is not None and (score[1]!=.5 or same_place):edges[ci,hi]=score
        candidates[ci]=[hi for hi in valid if id(history[hi]) in allowed]
        # 空白備註不能以類型名稱當作樣態；同位置唯一同類型前期可並排參考。
        kind=text_key(fields.get('deterioration',''))
        if not str(fields.get('remark','')).strip() and kind and kind not in ('其他損傷','其他','表面劣化','路堤損傷'):
            same=[hi for hi in candidates[ci] if text_key(history[hi].fields.get('deterioration',''))==kind and
                  location_key(history[hi].fields.get('d_location',''))==location_key(fields.get('d_location',''))]
            if len(same)==1 and same[0] not in used and same[0] not in blocked_history:edges[ci,same[0]]=(0,0,0,0,0)
    # Method-only matches require a sole candidate in both directions, independent of quantity.
    method_edges=[edge for edge,score in edges.items() if score[1]==.5]
    all_edges=list(edges);method_ambiguous=set()
    for ci,hi in method_edges:
        if sum(c==ci for c,h in all_edges)>1 or sum(h==hi for c,h in all_edges)>1:
            edges.pop((ci,hi),None);method_ambiguous.add(ci)
    free=set(range(len(records)))-set(assigned)-{i for i,r in enumerate(records) if r['id'] in manual}
    while True:
        choices={}
        for ci in free:
            available=[(hi,s) for (c,hi),s in edges.items() if c==ci and hi not in used]
            if not available:continue
            best=max(s for _,s in available);top=[hi for hi,s in available if s==best]
            if len(top)==1:choices[ci]=(top[0],best)
        matches=[]
        for ci,(hi,score) in choices.items():
            contenders=[(other,s) for (other,h),s in edges.items() if h==hi and other in free]
            best=max(s for _,s in contenders)
            if score==best and sum(s==best for _,s in contenders)==1:matches.append((ci,hi))
        if not matches:break
        for ci,hi in matches:assigned[ci]=hi;used.add(hi);free.remove(ci)
    result=[]
    for ci,record in enumerate(records):
        hi=assigned.get(ci)
        ambiguous=hi is None and record['id'] not in manual and (ci in conflicts or ci in method_ambiguous or any(c==ci and h not in used for c,h in edges) or (not str(record.get('current',record.get('fields',{})).get('remark','')).strip() and any(h not in used for h in candidates.get(ci,[]))))
        result.append(dict(current_index=ci,history_index=hi,kind='matched' if hi is not None else 'ambiguous' if ambiguous else 'new',component=component_name(record['component']),problem='構件位置與備註位置不一致，請核對後手動對應' if ci in conflicts else '相同工法有多筆候選，請手動對應' if ci in method_ambiguous and hi is None else '本期備註不足或有多筆前期，請手動對應' if ambiguous else ''))
    # 前期未見於本期的列僅顯示，永遠不加入本期輸出清單。
    removed=[dict(current_index=None,history_index=hi,kind='removed',component=component_name(history[hi].component)) for hi in valid if hi not in used]
    def owner(row):
        ci=row['current_index']
        item=records[ci] if ci is not None else history[row['history_index']]
        return physical_owner(item.get('context',{}),item.get('current',item.get('fields',{}))) if isinstance(item,dict) else physical_owner(item.context,item.fields)
    rows=result+removed
    components=list(dict.fromkeys(r['component'] for r in rows))
    def order(row):
        own=owner(row);m=re.fullmatch(r'([AP])(\d+)',own)
        position=(0,0) if own=='A1' else (2,0) if own=='A2' else (1,int(m[2])) if m else (3,0)
        current_owners={owner(r) for r in result if r['component']==row['component']}
        return components.index(row['component']),own not in current_owners,position,own,row['current_index'] is None
    rows.sort(key=order)
    for row in rows:
        if row['current_index'] is not None:continue
        own=owner(row)
        active=any(r['current_index'] is not None and r['component']==row['component'] and owner(r)==own for r in rows)
        good=any(component_name(e['component'])==row['component'] and physical_owner(e.get('context'),e.get('fields'))==own and str(e.get('der',[''])[0])=='1' for e in current_statuses)
        component_active=any(r['current_index'] is not None and r['component']==row['component'] for r in rows)
        if good and not active and (own or not component_active):row['current_message']='本期改為D=1（良好）'
    return rows
