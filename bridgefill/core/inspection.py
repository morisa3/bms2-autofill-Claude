"""檢測員意見草稿；不推斷未提供的現場安全判斷。"""
import json
import re
from decimal import Decimal
from pathlib import Path
from docx import Document
from .local_files import APP_DIR

DEFAULT_SUPERVISORS={'宋昌國':['台中市政府','雲林縣政府'],'顏宏宇':['南投縣政府','苗栗縣政府']}
SUFFIXES=['，建議持續追蹤。','，建議三年內派工修復。','，建議一年內派工修復。']
TEMPLATES=[
'本橋道路等級屬【道路等級】，且現況【人行道】，橋護欄高度不足110cm，建議採【加高／重作】方式改善。',
'本橋【護岸／護床工】發生【傾斜／掏空／倒塌】情形，位於【上游／下游】【左岸／右岸】約【＿m】處，並觀察流心偏向【橋台AO／橋墩PO】（【有／無】沖刷集中跡象），研判【可能影響／暫不影響】橋梁通洪及基礎安全，建議【一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。',
'本橋【Px橋墩】發生【傾斜／倒塌／裸露】情形，本期裸露深度為【本期深度，多跨以頓號區隔】，前期裸露深度為【前期深度，多跨以頓號區隔】，【Px】橋墩有【刷深／回淤】之情況，觀察主河道流心偏向【橋墩PO】，研判【可能影響／暫不影響】橋墩安全，建議【一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。',
'本橋【AO／PO】伸縮縫【間距過大／密合／具高低落差】，已造成【周邊劣化情形描述】，建議【持續觀測／一年內維修／通知相關單位立即處置】。',
'本橋橋下發現民眾堆置雜物及垃圾，具火災風險，建議通知權責單位辦理清理。',
'本橋僅供特定人員使用，其權責及通行條件尚待釐清，建議管理機關協調確認，以利後續檢測、維修及緊急應變作業。']


def supervisors(path=None, values=None):
    path=Path(path) if path else APP_DIR/'data'/'supervisors.json'
    if values is not None:
        if not isinstance(values,dict):raise ValueError('設定須為姓名對應機關清單')
        if not values or any(not str(k).strip() or not isinstance(v,list) or not v or not all(isinstance(x,str) and x.strip() for x in v) for k,v in values.items()):
            raise ValueError('主管姓名與負責機關不可空白')
        path.parent.mkdir(parents=True,exist_ok=True)
        temporary=path.with_suffix('.tmp')
        temporary.write_text(json.dumps(values,ensure_ascii=False,indent=2),encoding='utf-8')
        temporary.replace(path)
        return values
    if path.exists():
        values=json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(values,dict) or any(not isinstance(v,list) for v in values.values()):raise ValueError('主管設定格式錯誤，請檢查 data/supervisors.json')
        return values
    return {k:v[:] for k,v in DEFAULT_SUPERVISORS.items()}


HEADER_ALIASES={'負責機關':'管理機關','近端定義':'近端定義','是否疑似鹽害橋梁':'是否疑似鹽害橋梁','有無沖刷現象':'有無沖刷現象','沖刷位置/備註':'沖刷位置／備註','沖刷位置／備註':'沖刷位置／備註','檢測員意見':'檢測員意見','單位主管':'單位主管','道路等級':'道路等級','總橋孔數':'總橋孔數','橋頭里程':'橋頭里程','橋尾里程':'橋尾里程','管理機關':'管理機關','橋梁名稱':'橋梁名稱','橋梁編號':'橋梁編號','所在地標':'所在地標','所在縣市':'所在縣市','所在鄉區':'所在鄉區','路線':'路線','里程標號':'里程標號','橋梁總長':'橋梁總長','橋梁淨寬':'橋梁淨寬','總橋孔數':'總橋孔數','竣工年月':'竣工年月','橋下有無堆放雜物':'橋下有無堆放雜物','橋頭里程':'橋頭里程','橋尾里程':'橋尾里程','檢測日期':'檢測日期','檢測單位':'檢測單位'}


def metadata(path):
    if not path:return {}
    doc=Document(path)
    texts=[p.text for p in doc.paragraphs]
    pairs={}
    for table in doc.tables:
        for row in table.rows:
            cells=[];seen=set()
            for cell in row.cells:
                if cell._tc in seen:continue
                seen.add(cell._tc);cells.append(cell.text.strip())
            texts.extend(cells)
            for i,c in enumerate(cells[:-1]):
                label=re.sub(r'\s+','',c).rstrip('：:')
                if label in HEADER_ALIASES and cells[i+1]:
                    key=HEADER_ALIASES[label];value=cells[i+1]
                    if key in ('管理機關','單位主管','近端定義','是否疑似鹽害橋梁','有無沖刷現象'):value=re.sub(r'\s+','',value)
                    pairs[key]=value
    whole='\n'.join(texts)
    return dict(pairs,text=whole,mileage=bool(re.search(r'\d+(?:\.\d+)?\s*[KkＫｋ]\s*[+＋]\s*\d',whole)),sidewalk='緣石或人行道' in whole)


def canonical_choice(text):
    value=re.sub(r'\s+','',str(text or ''))
    return {'無':'否','沒有':'否','有':'是','里程樁號較小':'里程樁號較小端','竣工圖上之橋頭':'竣工圖上之橋頭端'}.get(value,value)


def supervisor_for(agency,settings):
    canonical=lambda value:re.sub(r'\s+','',str(value or '')).replace('臺','台')
    names=[name for name,agencies in settings.items() if canonical(agency) in {canonical(item) for item in agencies}]
    return names[0] if len(names)==1 else ''


def initial_values(original,meta,records,settings):
    values={key:str(original['values'].get(key,'') or '') for key in SELECTORS}
    for key,label in [('near','近端定義'),('salt','是否疑似鹽害橋梁'),('scour','有無沖刷現象')]:
        wanted=meta.get(label)
        if not wanted and key=='near':wanted='里程樁號較小端' if meta.get('mileage') else '其他'
        options=[o for o in original['choices'].get(key,[]) if canonical_choice(o['text'])==canonical_choice(wanted)]
        if len(options)==1:values[key]=options[0]['value']
    if '沖刷位置／備註' in meta:values['site']=meta['沖刷位置／備註']
    if meta.get('單位主管'):values['chief']=meta['單位主管']
    chief=supervisor_for(meta.get('管理機關'),settings)
    if chief:values['chief']=chief
    values['comment']=renumber(append_heights(values['comment'],records,meta))
    return values


def no_scour(values,choices):
    return any(option['value']==values.get('scour') and canonical_choice(option['text'])=='否' for option in choices.get('scour',[]))


def renumber(text):
    lines=str(text or '').splitlines();output=[];number=0;heights=False
    for line in lines:
        if re.match(r'^\s*(?:梁下|橋下)淨高',line):heights=True
        if heights or not line.strip():output.append(line);continue
        number+=1
        body=re.sub(r'^\s*\d+[.．、](?:\s+|(?=[^\d]))','',line).strip()
        output.append(f'{number}. {body}')
    return '\n'.join(output)


def record_fields(record):
    return record.get('fields') or record.get('current') or record.get('proposed') or {}


def deterioration_pattern(record, repairs=()):
    """樣態取自備註實際敘述或唯一 Excel 規則，不以分類欄代替。"""
    fields=record_fields(record)
    text=str(fields.get('remark','')).strip()
    if text:
        from .locations import location_prefix
        description=re.split(r'[（(]',text,1)[0].strip().rstrip('。．.')
        location=str(fields.get('d_location','')).strip()
        if location and description.startswith(location):description=description[len(location):].strip()
        else:
            parts=location_prefix(description)
            if parts:description=parts[1]
        if description and description!=location:return description.rstrip('。．.')
    candidates=[r['pattern'] for r in repairs if r.get('rule_id') and r.get('rule_id')==record.get('rule_id') and record.get('component') in r.get('components',[r.get('component')])]
    candidates=list(dict.fromkeys(candidates))
    if len(candidates)==1:return candidates[0]
    return str(record.get('pattern','')).strip() or '【劣化樣態待填】'


def pier(record):
    fields=record_fields(record)
    match=re.search(r'(?i)\bP0*(\d+)',fields.get('d_location','') or fields.get('remark',''))
    return 'P'+str(int(match[1])) if match else ''


def exposure_depth(record):
    text=record_fields(record).get('remark','').translate(str.maketrans({'（':'(', '）':')','＊':'*','×':'*'}))
    explicit=re.search(r'(?:裸露|曝露)深度(?:為|約|[:：=])?\s*(\d+(?:\.\d+)?)\s*m',text,re.I)
    if explicit:return explicit[1]
    dimensions=re.findall(r'\(\s*\d+(?:\.\d+)?m\s*\*\s*\d+(?:\.\d+)?m\s*\*\s*(\d+(?:\.\d+)?)m\s*\)',text,re.I)
    return dimensions[0] if len(dimensions)==1 else ''


def foundation_template(records, previous=()):
    selected=[r for r in records if any(c in r.get('component','') for c in ('橋墩基礎','橋墩／橋基保護設施')) and any(w in record_fields(r).get('remark','') for w in ('傾斜','倒塌','裸露'))]
    names=list(dict.fromkeys(pier(r) for r in selected if pier(r)))
    current={name:list(dict.fromkeys(exposure_depth(r) for r in selected if pier(r)==name and exposure_depth(r))) for name in names}
    old={name:list(dict.fromkeys(exposure_depth(r) for r in previous if pier(r)==name and exposure_depth(r) and any(c in r.get('component','') for c in ('基礎','橋墩／橋基保護設施')) and '裸露' in record_fields(r).get('remark',''))) for name in names}
    def depths(values,label):
        return '、'.join(name+'='+values[name][0]+'m' if len(values.get(name,[]))==1 else '【'+name+label+'深度】' for name in names)
    result=TEMPLATES[2]
    if names:
        result=result.replace('【Px橋墩】','、'.join(names)+'橋墩')
        result=result.replace('【Px】','、'.join(names))
        result=result.replace('【本期深度，多跨以頓號區隔】',depths(current,'本期')).replace('【前期深度，多跨以頓號區隔】',depths(old,'前期'))
    states=[w for w in ('傾斜','倒塌','裸露') if any(w in record_fields(r).get('remark','') for r in selected)]
    if states:result=result.replace('【傾斜／倒塌／裸露】','、'.join(states))
    deeper=[];shallower=[];stable=[]
    for name in names:
        if len(current.get(name,[]))==len(old.get(name,[]))==1:
            delta=Decimal(current[name][0])-Decimal(old[name][0])
            if delta>Decimal('0.1'):deeper.append(name)
            elif delta<Decimal('-0.1'):shallower.append(name)
            else:stable.append(name)
    changes=[]
    if deeper:changes.append('、'.join(deeper)+'橋墩有刷深之情況')
    if shallower:changes.append('、'.join(shallower)+'橋墩有回淤之情況')
    if stable:changes.append('、'.join(stable)+'橋墩裸露深度差異在±0.1m量測誤差範圍內，視為與前期一致')
    if changes:result=result.replace('、'.join(names)+'橋墩有【刷深／回淤】之情況','；'.join(changes))
    return result


def standard_templates(records,meta=None,previous=()):
    meta=meta or {}
    result=list(TEMPLATES)
    result[0]=result[0].replace('【道路等級】',meta.get('道路等級') or '【道路等級】').replace('【人行道】','設有人行道' if meta.get('sidewalk') else '無人行道')
    result[2]=foundation_template(records,previous)
    river=[r for r in records if any(c in r.get('component','') for c in ('河道','護床工'))]
    observed=' '.join(record_fields(r).get('remark','') for r in river)
    facilities=[word for word in ('護岸','護床工') if word in observed or any(word in r.get('component','') for r in river)]
    if facilities:result[1]=result[1].replace('【護岸／護床工】','、'.join(facilities))
    for placeholder,words in [('【傾斜／掏空／倒塌】',('傾斜','掏空','倒塌')),('【上游／下游】',('上游','下游')),('【左岸／右岸】',('左岸','右岸'))]:
        found=[word for word in words if word in observed]
        if found:result[1]=result[1].replace(placeholder,'、'.join(found))
    expansions=[r for r in records if r.get('component')=='伸縮縫' and any(word in record_fields(r).get('remark','') for word in ('間隙異常','間距過大','密合','高低差','高低落差'))]
    if expansions:
        positions=list(dict.fromkeys(record_fields(r).get('d_location','') for r in expansions if record_fields(r).get('d_location')))
        if positions:result[3]=result[3].replace('【AO／PO】','、'.join(positions))
        text=' '.join(record_fields(r).get('remark','') for r in expansions)
        states=[]
        for words,label in [(('間隙異常','間距過大'),'間距過大'),(('密合',),'密合'),(('高低差','高低落差'),'具高低落差')]:
            if any(word in text for word in words):states.append(label)
        if states:result[3]=result[3].replace('【間距過大／密合／具高低落差】','、'.join(states))
    flow=re.search(r'(?:主河道)?流心偏向\s*((?:[PA]0*\d+)(?:[、,，~～至及-](?:[PA]0*)?\d+)*(?:橋墩|橋台)?)',meta.get('text',''))
    if flow:
        result[1]=result[1].replace('【橋台AO／橋墩PO】',flow[1])
        result[2]=result[2].replace('【橋墩PO】',flow[1])
    return result


def draft(records, meta=None, waste=False, previous=(), repairs=(), templates=None, include_standards=True, recommendations=True):
    meta=meta or {}
    templates=templates or standard_templates(records,meta,previous)
    result=[]
    def triggered(components, words):
        return any(any(c in r.get('component','') for c in components) and any(w in (record_fields(r).get('remark','') or deterioration_pattern(r,repairs)) for w in words) for r in records)
    height_standard=include_standards and triggered(['護欄'],['高度不足'])
    if height_standard:result.append(templates[0])
    if include_standards and triggered(['河道','護床工'],['傾斜','掏空','倒塌']):result.append(templates[1])
    if include_standards and triggered(['橋墩基礎','橋墩／橋基保護設施'],['傾斜','倒塌','裸露']):result.append(templates[2])
    if include_standards and triggered(['伸縮縫'],['間隙異常','間距過大','密合','高低差','高低落差']):result.append(templates[3])
    if include_standards and waste:result.append(templates[4])
    # 第六段只供右側文字工具手動插入，不設獨立勾選項。
    groups={}
    for r in records:
        context=r.get('context',{})
        raw=context.get('U') or context.get('u') or re.search(r'U\s*[:：=]\s*(\d)',str(context))
        try:u=int(raw.group(1) if hasattr(raw,'group') else raw or 0)
        except (ValueError,TypeError):continue
        if u<2:continue
        fields=record_fields(r)
        component=r.get('component','')
        pattern=deterioration_pattern(r,repairs)
        if height_standard and '護欄' in component and '高度不足' in (fields.get('remark','')+' '+pattern):continue
        key=(component,u,pattern)
        loc=fields.get('d_location','').strip()
        if loc and loc not in groups.setdefault(key,[]):groups[key].append(loc)
    for (component,u,pattern),locations in sorted(groups.items(),key=lambda item:-item[0][1]):
        overlap=max((n for n in range(2,min(len(component),len(pattern))+1) if component.endswith(pattern[:n])),default=0)
        result.append('、'.join(locations)+component+pattern[overlap:]+(('，建議通知相關單位立即處置。' if u>=4 else SUFFIXES[2] if u==3 else SUFFIXES[0]) if recommendations else ''))
    text='\n'.join(f'{i}. {line}' for i,line in enumerate(result,1))
    return append_heights(text,records,meta)


def append_heights(text,records,meta):
    slots,note=height_slots(records,meta)
    if note or re.search(r'(?:梁下|橋下)淨高',text):return text
    title='梁下淨高' if any(r.get('component')=='主梁' for r in records) else '橋下淨高'
    lines=[s.replace(title,'',1).replace('【m】','m') for s in slots]
    return text.rstrip()+'\n\n'+title+'\n'+'\n'.join(lines)


def height_slots(records,meta):
    match=re.search(r'\d+',meta.get('總橋孔數',''))
    if not match:return [], '總橋孔數未讀取，請手動填寫淨高'
    count=int(match.group())
    if not 1<=count<=200:return [], '總橋孔數異常，請確認'
    title='梁下淨高' if any(r.get('component')=='主梁' for r in records) else '橋下淨高'
    return ([f'{title}H=【m】'] if count==1 else [f'{title}P{i}=【m】' for i in range(1,count)]),''


HEADER_METADATA={'道路等級':'xpath=//*[@id="maincontent"]/form/table/tbody/tr[1]/td/table[2]/tbody/tr[3]/td[1]', '總橋孔數':'xpath=//*[@id="maincontent"]/form/table/tbody/tr[1]/td/table[2]/tbody/tr[2]/th[3]'}

SELECTORS={'chief':'#chief','near':'xpath=//*[@id="maincontent"]/form/table/tbody/tr[2]/td/table/tbody/tr[1]/td[4]/select','salt':'#salt_injury','scour':'#scour_type','site':'#scour_site','comment':'#comment'}

def header_scope(service):
    service.require_connection()
    scopes=[s for s in service.scopes() if s.locator('#chief:visible').count()==1 and s.locator('#comment:visible').count()==1]
    if len(scopes)!=1:raise ValueError('請在專用 Chrome 開啟唯一的檢測基本資料頁，再重新讀取')
    return scopes[0]

def read_header(service):
    scope=header_scope(service)
    values={}
    choices={}
    for key,selector in SELECTORS.items():
        field=scope.locator(selector)
        if field.count()!=1:raise ValueError('找不到檢測基本資料欄位：'+key)
        values[key]=field.input_value()
        options=field.locator('option')
        if options.count():choices[key]=options.evaluate_all('(items)=>items.map(e=>({value:e.value,text:e.textContent.trim()}))')
    waste=scope.locator('#waste')
    values['waste']=waste.locator('option:checked').inner_text().strip()=='是' if waste.count()==1 else False
    page_meta={}
    for key,selector in HEADER_METADATA.items():
        field=scope.locator(selector)
        if field.count()==1:
            text=field.inner_text().strip()
            if text and text not in ('道路等級','總橋孔數'):page_meta[key]=text
            else:
                next_cell=field.locator('xpath=following-sibling::*[1]')
                if next_cell.count()==1:page_meta[key]=next_cell.inner_text().strip()
    return dict(values=values,choices=choices,url=scope.url,metadata=page_meta)


def write_header(service,payload):
    backup=APP_DIR/'temp'/'inspection-pending.json'
    if backup.exists():raise ValueError('上次送出結果待確認，請人工核對網站與 temp/inspection-pending.json；不會重送')
    scope=header_scope(service)
    before=read_header(service)
    if before!=payload['original']:raise ValueError('網頁資料已變動，請重新讀取；未送出')
    values=payload['values']
    if not values.get('comment','').strip():raise ValueError('請至少填寫檢測員意見')
    for key,selector in SELECTORS.items():
        if key=='site' and no_scour(values,before['choices']):continue
        if values[key]==before['values'][key]:continue
        service.control.check()
        field=scope.locator(selector)
        if key in before['choices']:field.select_option(value=values[key])
        else:field.fill(values[key])
    button=scope.locator('xpath=//*[@id="maincontent"]/form/table/tbody/tr[3]/td[1]/input')
    if button.count()!=1:raise ValueError('找不到唯一送出按鈕；未送出')
    # 不重試送出；回讀失敗時保留待確認記錄。
    backup=APP_DIR/'temp'/'inspection-pending.json'
    backup.parent.mkdir(parents=True,exist_ok=True)
    backup.write_text(json.dumps(payload,ensure_ascii=False),encoding='utf-8')
    with scope.page.expect_navigation(wait_until='domcontentloaded',timeout=service.timeout):
        service.control.activate(button)
    fresh=read_header(service)
    if any(fresh['values'][k]!=v for k,v in values.items() if not(k=='site' and no_scour(values,before['choices']))):raise ValueError('送出後回讀不一致，請人工確認，不會自動重送')
    backup.unlink(missing_ok=True)
    return read_header(service)
