"""離線逐列比對與可恢復草稿；網站識別仍由 DOCX 清單核對。"""
from copy import deepcopy
import json
import hashlib
import os
from pathlib import Path
import re
import tempfile
from .config import FIELD_IDS
from .documents import read_generated, export_document, complete
from .local_files import APP_DIR, TEMP_DIR
from .progress import digest
from .alignment import align_records
from .reconcile import scoped_history
from .remarks import prepare_fields, template_incomplete
from .data import component_repairs,component_name
from .policy import automatic_plan
from .naming import dated_bridge_name
from .conditions import rule_deterioration


def file_fingerprint(path):
    hasher=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda:stream.read(1024*1024),b''):hasher.update(block)
    return hasher.hexdigest()


def file_stamp(path):
    stat=Path(path).stat()
    return stat.st_mtime_ns,stat.st_size


def atomic_json(path, data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    fd,name=tempfile.mkstemp(dir=path.parent,suffix='.tmp')
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as stream:
            json.dump(data,stream,ensure_ascii=False);stream.flush();os.fsync(stream.fileno())
        os.replace(name,path)
    finally:
        if os.path.exists(name):os.unlink(name)


def safe_name(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',str(value)).strip(' .') or '未命名橋梁'


from .local_files import stored_path,restored_path


class ComparisonSession:
    def __init__(self, source, history, repairs, temp_dir=TEMP_DIR, output_dir=None,backup_dir=None):
        self.source=Path(source);self.history=history;self.repairs=repairs
        self.output_dir=Path(output_dir) if output_dir else APP_DIR/'document'
        self.source_stamp=file_stamp(self.source)
        self.manifest=read_generated(source,strict=False);self.inspection_info={}
        self.locked=bool(self.manifest.get('comparison',{}).get('finalized'));self.locked_path=self.source if self.locked else None
        self.ratings_reviewed=bool(self.manifest.get('comparison',{}).get('ratings_reviewed'))
        if self.manifest.get('scan_scope') is not None:self.history=[h for h in history if h.component in self.manifest['scan_scope']]
        self.source_fingerprint=file_fingerprint(self.source)
        if file_stamp(self.source)!=self.source_stamp:raise ValueError('來源文件正在寫入，請儲存完成後再導入')
        from .data import eligible
        self.retained_good=deepcopy([e for e in self.manifest['records'] if eligible(*e.get('der',['',''])) is False])
        self.entries=deepcopy([e for e in self.manifest['records'] if eligible(*e.get('der',['',''])) is not False])
        self.path=Path(temp_dir)/('compare-'+digest([stored_path(self.source),self.manifest['job']])+'.json')
        self.backup_path=(Path(backup_dir) if backup_dir else Path(temp_dir).parent/'backup')/(self.path.stem+'-'+self.source_fingerprint[:12]+'.json')
        self.index=0;self.repair_page=0;self.manual_pairs={};self.released_history=set()
        self.history_signature=digest([(h.component,h.context,h.fields) for h in self.history])
        exported_pairs=self.manifest.get('comparison',{})
        if exported_pairs.get('history_signature')==self.history_signature:
            self.manual_pairs=exported_pairs.get('manual_pairs',{})
        restored=False
        legacy=Path(temp_dir)/('compare-'+digest([str(self.source.resolve()),self.manifest['job']])+'.json')
        legacy_backup=self.backup_path.parent/(legacy.stem+'-'+self.source_fingerprint[:12]+'.json')
        saved_path=next((p for p in (self.path,self.backup_path,legacy,legacy_backup) if p.exists()),self.path)
        if saved_path.exists():
            saved=json.loads(saved_path.read_text(encoding='utf-8'))
            if saved.get('source')!=self.source_fingerprint or [r['id'] for r in saved['entries']]!=[r['id'] for r in self.entries]:
                raise ValueError('比對暫存與來源文件不一致，請保留暫存並重新選擇正確文件')
            self.locked=bool(saved.get('locked',self.locked));self.locked_path=restored_path(saved['locked_path']) if saved.get('locked_path') else self.locked_path;self.ratings_reviewed=bool(saved.get('ratings_reviewed',self.ratings_reviewed))
            self.entries=saved['entries'];self.index=min(saved.get('index',0),max(0,len(self.entries)-1))
            if saved.get('manual_pairs') and saved.get('history_signature')!=self.history_signature:
                raise ValueError('前期文件已變動，請先備份比對暫存後重新進行手動對應')
            self.manual_pairs=saved.get('manual_pairs',{});self.released_history=set(saved.get('released_history',[]))
            restored=saved.get('comparison_version')==6 and saved.get('rules_signature')==digest(repairs)
        self.rebuild_alignment()
        if not restored and not self.locked:
            for i,entry in enumerate(self.entries):
                if not entry.get('reason','').startswith('人工'):self.correct(i)
            self.rebuild_alignment()
        for i,entry in enumerate(self.entries):
            if not restored and not self.locked:
                human=entry.get('reason','').startswith('人工')
                entry['confirmed']=bool(entry.get('confirmed')) if human else entry['pair_state']=='matched' and complete(entry['fields']) and not (entry.get('needs_review') and entry.get('rule_id'))
        if self.locked:
            for entry in self.entries:entry['confirmed']=complete(entry['fields'])
        self.save()

    def rebuild_alignment(self):
        self.rows=align_records([dict(e,current=e['fields']) for e in self.entries],self.history,self.manual_pairs,repairs=self.repairs,blocked_history=self.released_history,current_statuses=self.retained_good)
        self.pairs={r['current_index']:r['history_index'] for r in self.rows if r['current_index'] is not None}
        for row in self.rows:
            if row['current_index'] is None:continue
            entry=self.entries[row['current_index']]
            entry['pair_state']=row['kind']
            entry['pair_problem']=row.get('problem','')
            entry['color']='red' if row['kind']=='new' else 'yellow' if row['kind']=='ambiguous' else 'normal'

    def histories(self, entry):
        index=next(i for i,e in enumerate(self.entries) if e is entry or e['id']==entry['id'])
        hi=self.pairs.get(index)
        return [self.history[hi]] if hi is not None else []

    def map_history(self,current_index,history_index):
        self.ensure_editable()
        entry=self.entries[current_index]
        if history_index is not None:
            if not 0<=history_index<len(self.history):raise ValueError('前期資料不存在')
            if any(i!=current_index and hi==history_index for i,hi in self.pairs.items()):raise ValueError('這筆前期已與其他本期對應，請先解除原對應')
        old=self.pairs.get(current_index)
        if old is not None and old!=history_index:self.released_history.add(old)
        if history_index is not None:self.released_history.discard(history_index)
        self.manual_pairs[entry['id']]=history_index
        self.rebuild_alignment()
        self.correct(current_index)
        self.rebuild_alignment()
        self.save()

    def correct(self,index,from_previous=False):
        """套用本期規則；僅明確按導入前期資料時複製前期欄位。"""
        self.ensure_editable()
        entry=self.entries[index]
        previous=self.histories(entry)
        if from_previous and not previous:raise ValueError('請先手動對應一筆前期資料，再按導入前期資料')
        fields=deepcopy(previous[0].fields if from_previous else entry['fields'])
        fields=prepare_fields(fields,component=entry['component'],context=entry['context'],sync_quantity=True)
        # 配對只提供唯一參考；完整本期範圍由規則保留，不以舊數量覆寫。
        plan=automatic_plan(fields,entry['component'],previous,self.rules(entry['component']),entry['context'])
        entry.update(fields=prepare_fields(plan.fields,component=entry['component'],context=entry['context'],sync_quantity=True),
            reason=plan.reason,needs_review=plan.review,confirmed=not plan.review and complete(plan.fields))
        entry['rule_id']=plan.rule_id
        return entry['fields']

    def refresh_repairs(self,repairs):
        self.repairs=repairs
        if self.locked:return
        for i,entry in enumerate(self.entries):
            if not entry.get('reason','').startswith('人工'):self.correct(i)
        self.rebuild_alignment();self.save()

    def import_previous(self):
        self.ensure_editable()
        self.correct(self.index,from_previous=True);self.rebuild_alignment();self.save()
        return self.entries[self.index]['fields']

    def editor_fields(self):
        entry=self.entries[self.index]
        return prepare_fields(entry['fields'],component=entry['component'],context=entry['context'],sync_quantity=True)

    def unpaired_history(self,current_index=None):
        used={hi for i,hi in self.pairs.items() if hi is not None and i!=current_index}
        available=[(i,h) for i,h in enumerate(self.history) if i not in used and not h.context.get('_inspection_status') and (h.fields.get('deterioration') or h.fields.get('remark'))]
        component=component_name(self.entries[current_index]['component']) if current_index is not None else ''
        return sorted(available,key=lambda pair:(component_name(pair[1].component)!=component,self.pairs.get(current_index)==pair[0]))

    def unpaired_current(self,history_index):
        component=component_name(self.history[history_index].component)
        return sorted([(i,e) for i,e in enumerate(self.entries) if self.pairs.get(i) is None],key=lambda pair:component_name(pair[1]['component'])!=component)

    def rules(self,component=None):return component_repairs(self.repairs,component or self.entries[self.index]['component'])

    def select(self,index):
        self.index=max(0,min(index,len(self.entries)-1));self.repair_page=0;self.save()

    def apply_rule(self,rule):
        self.ensure_editable()
        entry=self.entries[self.index]
        fields=prepare_fields(entry['fields'],rule,True,entry['component'],entry['context'])
        fields.update({key:rule[key] for key in ('method','unit','price')})
        fields['deterioration']=rule_deterioration(rule,entry['fields'].get('deterioration',''))
        entry.update(fields=fields,confirmed=False,reason='人工選取規則',rule_id=rule.get('rule_id',''));self.rebuild_alignment();self.save()

    def edit(self,fields):
        self.ensure_editable()
        entry=self.entries[self.index]
        fields=prepare_fields({k:str(fields.get(k,'') or '').strip() for k in FIELD_IDS},component=entry['component'],context=entry['context'],sync_quantity=True)
        entry.update(fields=fields,confirmed=complete(fields) and not template_incomplete(fields.get('remark','')),needs_review=not complete(fields),reason='人工已修改')
        self.rebuild_alignment();self.save()

    def confirm(self):
        from .config import LABELS
        fields=self.entries[self.index]['fields']
        missing=[LABELS[k] for k in FIELD_IDS if not str(fields.get(k,'')).strip()]
        if missing:raise ValueError('本期資料待補：'+'、'.join(missing))
        if not complete(fields) or template_incomplete(fields.get('remark','')):raise ValueError('本期備註的實測量式尚未完成')
        self.edit(fields)

    def save(self):
        if file_stamp(self.source)!=self.source_stamp and file_fingerprint(self.source)!=self.source_fingerprint:
            raise ValueError('來源 DOCX 在比對期間已變動，暫存保留；請先備份後重新導入')
        payload=dict(version=1,locked=self.locked,locked_path=stored_path(self.locked_path) if self.locked_path else None,ratings_reviewed=self.ratings_reviewed,comparison_version=6,released_history=sorted(self.released_history),rules_signature=digest(self.repairs),history_signature=self.history_signature,manual_pairs=self.manual_pairs,source=self.source_fingerprint,source_path=str(self.source.resolve()),index=self.index,entries=self.entries,
            history=[dict(component=h.component,context=h.context,fields=h.fields) for h in self.history],pairs={self.entries[i]['id']:hi for i,hi in self.pairs.items()})
        atomic_json(self.path,payload)
        atomic_json(self.backup_path,payload)

    def ensure_editable(self):
        if self.locked:raise ValueError('完成資料已鎖定，請先按「解除鎖定並修改」')
    def lock(self,path):
        self.locked=True;self.locked_path=Path(path);self.save()
    def unlock(self):
        self.locked=False;self.ratings_reviewed=False;self.save()
    def rating_conflicts(self):
        from .rating_review import mixed_ratings
        return mixed_ratings(self.entries,self.retained_good)

    def issues(self):
        from .config import LABELS
        problems=[]
        for i,e in enumerate(self.entries):
            missing=[LABELS[k] for k in FIELD_IDS if not str(e['fields'].get(k,'')).strip()]
            details='待補：'+'、'.join(missing) if missing else '備註實測量式未完成' if not complete(e['fields']) or template_incomplete(e['fields'].get('remark','')) else ''
            if details:problems.append(dict(index=i,component=e['component'],location=e['fields'].get('d_location',''),detail=details))
        return problems

    def finish(self,bridge):
        self.save()
        if not self.entries and not self.retained_good:raise ValueError('本期沒有需輸出的缺失；前期只作比對，不會輸出')
        records=[dict(id=e['id'],component=e['component'],context=e['context'],
            current=e.get('current',e['fields']),der=e.get('der',['2','2']),choices=e.get('choices',{}),website_id=e.get('website_id',''),owner_id=e.get('owner_id','')) for e in self.entries]
        if self.retained_good:
            by_id={r['id']:r for r in records}
            by_id.update({e['id']:dict(e,current=e.get('current',e['fields'])) for e in self.retained_good})
            records=[by_id[e['id']] for e in self.manifest['records']]
        output=self.output_dir/(dated_bridge_name(bridge)+'定期檢測表-完成檔案.docx')
        if output.resolve()==self.source.resolve() or ('完成檔案' in self.source.stem and output.exists()):
            base=output.stem+'-修正';candidate=output.with_name(base+output.suffix);count=1
            while candidate.exists() or candidate.resolve()==self.source.resolve():
                count+=1;candidate=output.with_name(base+f'-{count}'+output.suffix)
            output=candidate
        self.output_dir.mkdir(parents=True,exist_ok=True)
        fd,name=tempfile.mkstemp(prefix='.comparison-',suffix='.docx',dir=self.output_dir);os.close(fd)
        working=Path(name)
        try:
            drafts={e['id']:dict(e,reason='人工修訂（匯出）' if complete(e['fields']) and not template_incomplete(e['fields'].get('remark','')) else '待人工補齊') for e in self.entries}
            drafts.update({e['id']:e for e in self.retained_good})
            export_document(self.source,working,self.manifest['job'],records,self.history,[],drafts=drafts,inspection_info=getattr(self,'inspection_info',None),scan_scope=self.manifest.get('scan_scope'))
            # 舊版文件也保留原始網站快照，不能用人工修訂值重建來源。
            from docx import Document
            from lxml import etree
            from .documents import NS
            doc=Document(working)
            for rel in doc.part.rels.values():
                try:root=etree.fromstring(rel.target_part.blob)
                except (etree.XMLSyntaxError,AttributeError):continue
                if root.tag=='{'+NS+'}manifest':
                    manifest=json.loads(root.text)
                    for new in manifest['records']:new['source']=next(e['source'] for e in self.manifest['records'] if e['id']==new['id'])
                    manifest['comparison']=dict(history_signature=self.history_signature,manual_pairs=self.manual_pairs,finalized=not bool(self.issues()),ratings_reviewed=self.ratings_reviewed)
                    root.text=json.dumps(manifest,ensure_ascii=False);rel.target_part._blob=etree.tostring(root)
            doc.save(working)
            read_generated(working,strict=False)
            os.replace(working,output)
        finally:working.unlink(missing_ok=True)
        # 新增清單只由具完整前後期對應的比較作業輸出，掃描流程不呼叫此處。
        from .reporting import u_change
        groups=[[],[],[]]
        for i,entry in enumerate(self.entries):
            hi=self.pairs.get(i)
            delta=u_change(entry,dict(context=self.history[hi].context)) if hi is not None else None
            if entry.get('pair_state')=='new':groups[0].append(i)
            elif delta and delta>0:groups[1].append(i)
            elif delta and delta<0:groups[2].append(i)
        components=list(dict.fromkeys(e['component'] for e in self.entries))
        for indices in groups:indices.sort(key=lambda i:(components.index(self.entries[i]['component']),i))
        indices=sum(groups,[])
        self.new_defects_path=output.with_name(output.stem+'-新增劣化.docx')
        sections=[(title,[self.entries[i]['id'] for i in group]) for title,group in zip(('本期新增劣化','U值上升之劣化','U值下降之劣化'),groups)]
        export_document(self.source,self.new_defects_path,self.manifest['job'],[next(r for r in records if r['id']==self.entries[i]['id']) for i in indices],[],[],
            drafts={self.entries[i]['id']:self.entries[i] for i in indices},allow_empty=True,inspection_info=getattr(self,'inspection_info',None),scan_scope=self.manifest.get('scan_scope'),report_sections=sections)
        if not self.issues():self.path.unlink(missing_ok=True)
        return output



def import_current_document(path, history, repairs, output_dir=None, bridge=None):
    """一般本期檢測表轉為具識別的草稿；不連接網站、不推測網站缺失 ID。"""
    from .data import read_history
    from .data import eligible
    from .documents import read_generated
    try:
        read_generated(path,strict=False)
        return Path(path)
    except ValueError as exc:
        # 含有識別清單卻損壞的文件禁止降級成一般文件，避免繞過錯列檢查。
        from zipfile import ZipFile
        with ZipFile(path) as archive:
            if any(n.startswith('customXml/') and NS_TOKEN in archive.read(n) for n in archive.namelist() if n.endswith('.xml')):raise exc
    audit={};rows=read_history(path,audit)
    if audit.get('issues'):raise ValueError('本期文件有無法辨識的構件或欄位，請修正後重新匯入')
    records=[]
    for i,row in enumerate(rows):
        d,e=row.context.get('D',''),row.context.get('E','')
        # 所有列都保留；良好列僅在比對修改介面隱藏。
        if eligible(d,e) is None:raise ValueError('本期文件缺少可判讀的 D/E，請補齊後再匯入')
        fields={k:row.fields.get(k,'') for k in FIELD_IDS}
        records.append(dict(id='document-'+digest([i,row.component,row.context,fields]),component=row.component,
            context=row.context,current=fields,der=[d,e],choices={}))
    directory=Path(output_dir) if output_dir else APP_DIR/'document'
    name=dated_bridge_name(bridge or path)
    output=directory/(name+'定期檢測表-自動建立.docx')
    if output.resolve()==Path(path).resolve():raise ValueError('來源不可覆寫，請先更名本期來源文件')
    job='document:'+file_fingerprint(path)
    # 日期是輸出顯示資訊，不是工作身份；跨日再導入仍沿用同份草稿與暫存。
    if directory.exists():
        from zipfile import BadZipFile
        for candidate in sorted(directory.glob('*定期檢測表-自動建立*.docx'),key=lambda p:p.stat().st_mtime_ns,reverse=True):
            try:existing=read_generated(candidate,strict=False)
            except (ValueError,OSError,BadZipFile):continue
            if existing.get('job')==job and existing.get('alignment_version',0)>=3:return candidate
    if output.exists():
        existing=read_generated(output,strict=False)
        if existing['job']==job:
            if existing.get('alignment_version',0)>=3:return output
            # 舊版同名自動草稿可能已誤借前期；另建新版，不覆寫舊檔與手動暫存。
            output=output.with_name(output.stem+'-完整構件.docx')
            if output.exists():
                candidate=read_generated(output,strict=False)
                if candidate['job']==job and candidate.get('alignment_version',0)>=3:return output
        else:
            # 同日同橋的不同來源不能覆寫仍有暫存的舊草稿。
            output=output.with_name(output.stem+'-'+job.split(':')[-1][:10]+'.docx')
    export_document(path,output,job,records,history,repairs,review_colors=True,allow_empty=True)
    return output

NS_TOKEN=b'urn:bridgefill:generated-document:1'
