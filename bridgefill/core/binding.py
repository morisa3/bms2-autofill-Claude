"""逐筆識別與可定位的回填差異；整批核對失敗時不允許寫入。"""
from copy import deepcopy
from .data import eligible,quantity
from .config import LABELS
from .locations import key as location_key


class FillPlan(dict):
    def __init__(self):
        super().__init__();self.bindings={}


class FillMismatch(ValueError):
    def __init__(self,details):
        self.details=details
        first=details[0]
        super().__init__(f"第 {first.get('file_index','—')} 筆｜{first.get('component','')}｜{first.get('reason','無法對應')}；已列出差異及修正方式，尚未填寫。")


def structural_match(old,current):
    if old['component']!=current['component']:return False
    for key,value in old.get('context',{}).items():
        if key.startswith('_') or key in ('D','E','R','U'):continue
        available=[v for k,v in current.get('context',{}).items() if key in k]
        if value and available and location_key(value) not in [location_key(v) for v in available]:return False
    return True


def bind(manifest,job,records,components=None,overrides=None):
    expected=[e for e in manifest['records'] if components is None or e['component'] in components]
    used=set();plan=FillPlan();problems=[];overrides=overrides or {}
    for ordinal,entry in enumerate(expected,1):
        candidates=[(i,r) for i,r in enumerate(records) if i not in used and structural_match(entry,r)]
        native=entry.get('website_id')
        override=overrides.get(entry['id'])
        if native:candidates=[(i,r) for i,r in candidates if r.get('website_id')==native]
        if override:candidates=[(i,r) for i,r in candidates if r['id']==override]
        if eligible(*entry.get('der',['',''])) is False:
            good=[(i,r) for i,r in candidates if eligible(*r.get('der',['',''])) is False]
            if good:used.add(good[0][0])
            continue  # 良好列保留在文件，不要求網站必有可編輯的缺失列。
        same_status=[(i,r) for i,r in candidates if eligible(*entry.get('der',['',''])) is eligible(*r.get('der',['','']))]
        if same_status:candidates=same_status
        source=entry.get('current',{});desired=entry.get('fields',{})
        def score(record):
            value=0
            for reference in (source,desired):
                subtotal=0
                for key,weight in [('deterioration',8),('remark',12),('d_location',3),('amount',2)]:
                    a=reference.get(key,'');b=record.get('current',{}).get(key,'')
                    if a and b and (quantity(a)==quantity(b) if key=='amount' else location_key(a)==location_key(b)):
                        subtotal+=weight
                value=max(value,subtotal)
            return value
        if len(candidates)>1:
            ranked=sorted(candidates,key=lambda item:score(item[1]),reverse=True)
            if score(ranked[0][1])>=8 and score(ranked[0][1])>score(ranked[1][1]):candidates=ranked[:1]
            else:
                highest=score(ranked[0][1]);tied=[pair for pair in ranked if score(pair[1])==highest]
                same_der=[pair for pair in tied if list(map(str,entry.get('der',[])))==list(map(str,pair[1].get('der',[])))]
                if len(same_der)==1 and highest>=8:candidates=same_der
                elif all(r['current']==ranked[0][1]['current'] and r['der']==ranked[0][1]['der'] for _,r in ranked):candidates=sorted(candidates,key=lambda item:item[0])[:1]
        if len(candidates)!=1:
            problems.append(dict(file_index=ordinal,record_id=entry['id'],component=entry['component'],reason='網站 ID／位置無唯一對應' if native else '同構件、位置及樣態無唯一對應',expected=deepcopy(entry),candidates=[deepcopy(r) for _,r in candidates],actual=deepcopy(candidates[0][1]) if candidates else None,fix='核對構件及位置；在右側選取正確的網站資料。若缺失已新增、刪除或 D/E 變更，請重新匯入本期文件比對。'))
            continue
        index,current=candidates[0]
        meaningful=bool(source.get('deterioration') or source.get('remark'))
        if meaningful and not native and not override and eligible(*entry.get('der',['',''])) is not False and score(current)<8:
            problems.append(dict(file_index=ordinal,record_id=entry['id'],component=entry['component'],reason='位置相同，但原始與修訂樣態均無法確認為同筆資料',expected=deepcopy(entry),actual=deepcopy(current),candidates=[deepcopy(current)],fix='請人工選取對應資料；若本筆已刪除或移往其他構件，重新匯入完整本期文件，不會自動套到其他樣態。'))
            continue
        used.add(index)
        if list(map(str,entry.get('der',[])))!=list(map(str,current.get('der',[]))) and not (eligible(*entry.get('der',['',''])) is False and eligible(*current['der']) is False):
            problems.append(dict(file_index=ordinal,record_id=entry['id'],component=entry['component'],reason='文件與網站的 D/E 已變更',expected=deepcopy(entry),actual=deepcopy(current),candidates=[deepcopy(current)],fix='請確認本期 D/E，以網站最新檢測表重新匯入比對；不要將良好項目強制改成缺失。'))
            continue
        if eligible(*current['der']) is False:continue
        from .documents import complete
        if not complete(desired):
            from .config import FIELD_IDS
            missing=[LABELS[k] for k in FIELD_IDS if not str(desired.get(k,'')).strip()]
            problems.append(dict(file_index=ordinal,record_id=entry['id'],component=entry['component'],reason='文件待補：'+('、'.join(missing) or '備註實測量式'),expected=deepcopy(entry),actual=deepcopy(current),candidates=[deepcopy(current)],fix='按「修改此筆文件資料」，補齊黃色欄位及量式後儲存；回缺失主頁再接續填寫。'))
            continue
        invalid=[key for key,options in current.get('choices',{}).items() if desired.get(key) not in options]
        if invalid:
            problems.append(dict(file_index=ordinal,record_id=entry['id'],component=entry['component'],reason='網站沒有此選項：'+ '、'.join(LABELS.get(key,key) for key in invalid),expected=deepcopy(entry),actual=deepcopy(current),candidates=[deepcopy(current)],fix='開啟此筆修正文件的劣化類型／單位，選用網站現有選項後另存並重新導入。'))
            continue
        plan[current['id']]=deepcopy(desired)
        plan.bindings[current['id']]=dict(file_index=ordinal,record_id=entry['id'],expected=deepcopy(entry))
    extras=[r for i,r in enumerate(records) if i not in used and eligible(*r['der']) is not False]
    if extras:problems.append(dict(file_index='未列入文件',component=extras[0]['component'],record_id='',reason=f'網站另外有 {len(extras)} 筆缺失未能對應文件',expected=None,actual=deepcopy(extras[0]),candidates=deepcopy(extras),fix='重新匯入網站最新本期檢測文件，保留新增及良好列，再進行兩期比對。'))
    if problems:raise FillMismatch(problems)
    if not expected:raise ValueError('所選構件沒有文件資料')
    return plan
