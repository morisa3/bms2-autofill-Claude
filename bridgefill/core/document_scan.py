"""生成文件輔助掃描：唯一對應自動填寫，重複樣態及歧義留給人工。"""
from collections import Counter
from .data import History,quantity,normalize_remark
from .documents import complete
from .progress import digest
from .reconcile import scoped_history,description,text_key


def owner(record):
    if record.get('owner_id'):return record['owner_id']
    return digest([record['component'],{k:v for k,v in record['context'].items() if '位置' in k or '編號' in k or k in ('類別','上/下層')}])


def style(fields):
    return description(fields.get('remark',''),fields.get('d_location','')) or text_key(fields.get('deterioration',''))


def source_compatible(entry,record):
    if entry['fields']==record['current']:return True
    source=entry.get('current',{})
    if not source or not any(str(source.get(k,'')).strip() for k in ('deterioration','remark','d_location','amount')):return False
    for key in ('deterioration','remark','amount','unit','method','price'):
        value=source.get(key,'')
        if not str(value).strip():continue
        actual=record['current'].get(key,'')
        if key in ('amount','price'):
            value,actual=quantity(value),quantity(actual)
        elif key=='remark':
            value,actual=text_key(normalize_remark(value)),text_key(normalize_remark(actual))
        else:value,actual=text_key(value),text_key(actual)
        if value!=actual:return False
    return True


def build_document_scan(manifest,job,records,components=None):
    if not str(manifest['job']).startswith('document:') and manifest['job']!=job:
        raise ValueError('生成文件不屬於目前橋梁／期別，未開始掃描')
    entries=[e for e in manifest['records'] if components is None or e['component'] in components]
    ids=[e['website_id'] for e in entries if e.get('website_id')]
    if len(ids)!=len(set(ids)):raise ValueError('文件網站缺失 ID 重複，請重新確認文件')
    current_ids=[r['website_id'] for r in records if r.get('website_id')]
    if len(current_ids)!=len(set(current_ids)):raise ValueError('本期網站缺失 ID 重複，停止掃描')
    candidates={}
    for record in records:
        same=[e for e in entries if e['component']==record['component']]
        website_id=record.get('website_id')
        exact=[e for e in same if website_id and e.get('website_id')==website_id]
        if not exact:exact=[e for e in same if manifest['job']==job and e['id']==record['id'] and not e.get('website_id')]
        if exact:candidates[record['id']]=exact;continue
        # 文件中已記錄的網站 ID 不可用內容猜成另一筆。
        same=[e for e in same if not e.get('website_id')]
        wrappers=[History(e['component'],e.get('current',e['fields']),i,e['context']) for i,e in enumerate(same)]
        scoped=scoped_history(wrappers,record['component'],record['context'])
        scoped=[same[h.source_row] for h in scoped]
        precise=[e for e in scoped if style(e.get('current',e['fields']))==style(record['current'])]
        if precise:scoped=precise
        elif len(scoped)>1:
            typed=[e for e in scoped if text_key(e.get('current',e['fields']).get('deterioration',''))==text_key(record['current'].get('deterioration',''))]
            if typed:scoped=typed
        candidates[record['id']]=scoped
    uses=Counter(e['id'] for choices in candidates.values() if len(choices)==1 for e in choices)
    repeats=Counter((owner(r),style(r['current'])) for r in records)
    planned_repeats=Counter((owner(r),style(candidates[r['id']][0]['fields'])) for r in records if len(candidates[r['id']])==1)
    result={}
    for record in records:
        choices=candidates[record['id']]
        reason=''
        if not choices:reason='生成文件找不到本筆對應，請人工核對'
        elif len(choices)!=1:reason='生成文件有多筆候選，請人工選擇'
        elif uses[choices[0]['id']]!=1:reason='多筆網站缺失對應同一文件列，請人工核對'
        elif repeats[(owner(record),style(record['current']))]>1 or planned_repeats[(owner(record),style(choices[0]['fields']))]>1:reason='同構件內有重複劣化樣態，請人工核對'
        elif str(manifest['job']).startswith('document:') and not choices[0].get('website_id') and not source_compatible(choices[0],record):reason='文件與網站原始內容無法核對，請人工確認'
        elif not complete(choices[0]['fields']):reason='文件欄位或量式不完整，請人工核對'
        elif not str(manifest['job']).startswith('document:') and choices[0]['source']!=digest([record['current'],record['der']]) and choices[0]['fields']!=record['current']:
            reason='網站原始資料在文件建立後已變動，請人工核對'
        result[record['id']]={'candidates':choices,'review':bool(reason),'reason':reason or '生成文件唯一對應，自動填入'}
    return result
