"""網站掃描是本期缺失來源；下載文件只提供版型、基本資料及良好列。"""
from copy import deepcopy
from pathlib import Path
import json
import re
import os
import tempfile
from .config import FIELD_IDS
from .data import read_history,eligible
from .documents import export_document
from .progress import digest
from .local_files import APP_DIR,TEMP_DIR


def scan_supplement(path):
    audit={};rows=read_history(path,audit)
    if audit.get('issues'):raise ValueError('本期文件有無法辨識的構件／欄位，請先修正後導入')
    result=[]
    for i,row in enumerate(rows):
        if eligible(row.context.get('D'),row.context.get('E')) is not False:continue
        fields={k:row.fields.get(k,'') for k in FIELD_IDS}
        result.append(dict(id='supplement-'+digest([i,row.component,row.context,fields]),component=row.component,context=row.context,current=fields,der=[row.context.get('D',''),row.context.get('E','')],choices={},website_id=''))
    return result


def export_scan(source,job,scanned,output_dir=None,temp_dir=None,components=None):
    source=Path(source);records=deepcopy(scanned);seen=set()
    for number,record in enumerate(records,1):
        if eligible(*record['der']) is not True:raise ValueError('掃描資料含良好或不明 D/E，未建立本期文件')
        native=str(record.get('website_id') or '')
        if not native.isdecimal() or native in seen:raise ValueError(f'第 {number} 筆缺少或重複劣化 ID，請核對網站；未輸出文件')
        seen.add(native)
    records.extend(scan_supplement(source))
    order=list(dict.fromkeys(r['component'] for r in records));records.sort(key=lambda r:order.index(r['component']))
    directory=Path(output_dir) if output_dir else APP_DIR/'document';directory.mkdir(parents=True,exist_ok=True)
    stem=re.sub(r'-掃描資料(?:-\d+)?$','',source.stem);output=directory/(stem+'-掃描資料.docx');number=2
    while output.exists():output=directory/(stem+f'-掃描資料-{number}.docx');number+=1
    save_scan_cache(source,job,scanned,True,temp_dir)
    drafts={r['id']:dict(fields=r['current'],rule_id='',reason='網站掃描原始資料' if r.get('website_id') else '本期文件補齊良好資料') for r in records}
    export_document(source,output,job,records,[],[],drafts=drafts,allow_empty=True,scan_scope=components)
    return output


def save_scan_cache(source,job,records,complete=False,temp_dir=None):
    cache=Path(temp_dir) if temp_dir else TEMP_DIR;cache.mkdir(parents=True,exist_ok=True)
    payload=dict(version=1,job=job,source=str(Path(source).resolve()),records=records,complete=complete)
    with tempfile.NamedTemporaryFile('w',encoding='utf-8',dir=cache,delete=False) as file:
        json.dump(payload,file,ensure_ascii=False);temporary=Path(file.name)
    os.replace(temporary,cache/('scan-current-'+digest([str(Path(source).resolve()),job])+'.json'))


def load_reference(path):
    """一般下載文件不套用缺失規則；只作補充來源。"""
    from .documents import read_generated,NS
    from zipfile import ZipFile
    try:return read_generated(path,strict=False)
    except ValueError as error:
        with ZipFile(path) as archive:
            if any(name.startswith('customXml/') and NS.encode() in archive.read(name) for name in archive.namelist() if name.endswith('.xml')):raise error
    audit={};rows=read_history(path,audit)
    if audit.get('issues'):raise ValueError('本期檢測文件有無法辨識的欄位，請先修正')
    records=[dict(id='reference-'+digest([i,r.component,r.context]),component=r.component,context=r.context,current=dict(r.fields),fields=dict(r.fields),der=[r.context.get('D',''),r.context.get('E','')],website_id='') for i,r in enumerate(rows)]
    return dict(version=1,reference_only=True,job='document:'+__import__('hashlib').sha256(Path(path).read_bytes()).hexdigest(),records=records)
