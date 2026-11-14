"""生成可編修的本期 DOCX；以非可見的識別清單防止錯位填報。"""
from copy import deepcopy
from pathlib import Path
import json
import os
import tempfile
import uuid
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from .config import FIELD_IDS
from .data import compact, component_name, component_repairs
from .policy import automatic_plan
from .reconcile import scoped_history, text_key
from .remarks import prepare_fields,crack_complete
from .progress import digest

NS = 'urn:bridgefill:generated-document:1'
HEADERS = {'劣化類型':'deterioration','損壞位置':'d_location','維修工法':'method',
           '數量':'amount','單位':'unit','金額':'price','單價':'price','備註':'remark'}


def template_component(table):
    if len(table.rows)<2 or '檢測項目' not in table.cell(0,0).text:
        return ''
    values=[compact(c.text) for c in table.rows[0].cells]
    return next((component_name(v) for v in values if v and v!='檢測項目'),'')


def basic_table(document,component):
    headers=['位置','編號','D','E','R','U','劣化類型','維修工法','數量','單位','備註']
    table=document.add_table(rows=3,cols=len(headers))
    table.style='Table Grid'
    table.cell(0,0).text='檢測項目'
    table.cell(0,1).merge(table.cell(0,len(headers)-1)).text=component
    for cell,text in zip(table.rows[1].cells,headers):cell.text=text
    return table


def add_website_id(table):
    headers=[compact(c.text) for c in table.rows[1].cells]
    if '劣化ID' in headers:return
    from docx.shared import Inches
    widths=[col.width for col in table.columns];total=sum(w or 0 for w in widths)
    width=min(int(Inches(.65)),int(total*.13)) if total else int(Inches(.65))
    table.add_column(width);table.rows[1].cells[-1].text='劣化ID'
    if total:
        for col,old in zip(table.columns,widths):col.width=int((old or 0)*(total-width)/total)
        table.columns[-1].width=width
        grid_widths=[col.width or 0 for col in table.columns]
        for row in table.rows:
            offset=0
            for cell in row._tr.tc_lst:
                span=cell.grid_span;cell.width=sum(grid_widths[offset:offset+span]);offset+=span
    table.autofit=True


def add_price(table):
    headers=[compact(c.text) for c in table.rows[1].cells]
    # 舊模板已有單價時移至單位後，不另產生第二個價格欄。
    price=next((i for i,h in enumerate(headers) if h in ('金額','單價')),None)
    unit=headers.index('單位')
    if price is not None:
        grid=table._tbl.tblGrid
        anchor=grid[unit]
        moved=grid[price]
        grid.remove(moved);anchor.addnext(moved)
        for row in table.rows[1:]:
            cell=row.cells[price]._tc
            row._tr.remove(cell)
            row.cells[unit if price>unit else unit-1]._tc.addnext(cell)
        table.rows[1].cells[unit+1 if price>unit else unit].text='金額'
        return
    grid=table._tbl.tblGrid
    original_width=sum(int(col.get(qn('w:w'),'0')) for col in grid)
    grid.insert(unit+1,deepcopy(grid[unit]))
    for i,row in enumerate(table.rows):
        if i==0:
            cell=row._tr.tc_lst[-1]
            cell.get_or_add_tcPr().get_or_add_gridSpan().val += 1
            continue
        cell=deepcopy(row.cells[unit]._tc)
        row.cells[unit]._tc.addnext(cell)
        row.cells[unit+1].text='金額' if i==1 else ''
    # 增加單價欄仍維持模板原有總寬度，避免超出列印邊界。
    total=sum(int(col.get(qn('w:w'),'0')) for col in grid)
    if total:
        widths=[round(int(col.get(qn('w:w'),'0'))*original_width/total) for col in grid]
        for col,width in zip(grid,widths):col.set(qn('w:w'),str(width))
        for row in table.rows:
            start=0
            for cell in row._tr.tc_lst:
                span=cell.grid_span
                cell.get_or_add_tcPr().get_or_add_tcW().set(qn('w:w'),str(sum(widths[start:start+span])))
                start+=span


def context_value(context,header):
    if header in context:return str(context[header])
    values=[str(v) for k,v in context.items() if header in k and not k.startswith('_')]
    if header=='編號' and not values:
        places=[str(v) for k,v in context.items() if '位置' in k]
        values=[{'橋頭':'A1','橋尾':'A2'}[v] for v in places if v in ('橋頭','橋尾')]
    return '｜'.join(dict.fromkeys(values))


def shade(row,color=None):
    for cell in row.cells:
        pr=cell._tc.get_or_add_tcPr()
        for node in list(pr.findall(qn('w:shd'))):pr.remove(node)
        if color:
            node=OxmlElement('w:shd');node.set(qn('w:fill'),color);pr.append(node)


def highlighted(row):
    return any(node.get(qn('w:fill'),'').upper() in ('FFF2CC','FFD8D8')
               for cell in row.cells for node in cell._tc.get_or_add_tcPr().findall(qn('w:shd')))


def complete(fields):
    return all(fields.get(k) is not None and str(fields.get(k,'')).strip() for k in FIELD_IDS) and crack_complete(fields)


def export_document(template_path, output_path, job, records, history, repairs, drafts=None, review_colors=False, allow_empty=False, inspection_info=None,scan_scope=None,report_sections=None):
    if not records and not allow_empty:raise ValueError('目前沒有可處理的本期缺失，未產生空白文件')
    document=Document(template_path)
    from .document_style import remove_filter_caption
    remove_filter_caption(document)
    # 以助手生成文件再輸出時替換舊識別清單，避免同名 customXml 重複。
    for rid,rel in list(document.part.rels.items()):
        if rel.reltype==RT.CUSTOM_XML:
            try:old=etree.fromstring(rel.target_part.blob)
            except etree.XMLSyntaxError:continue
            if old.tag=='{'+NS+'}manifest':document.part.drop_rel(rid)
    templates={template_component(t):deepcopy(t._tbl) for t in document.tables if template_component(t)}
    if '翼牆／擋土牆' not in templates and '橋台' in templates:
        templates['翼牆／擋土牆']=deepcopy(templates['橋台'])
    # 保留頁面、段落、頁首頁尾與各構件表格樣式；以本期順序重建資料列。
    for table in list(document.tables):
        if template_component(table):table._tbl.getparent().remove(table._tbl)
    from .inspection import metadata
    from .header_document import apply_header
    info=metadata(template_path)
    info.update(inspection_info or {})
    apply_header(document,info)
    manifest={'version':1,'alignment_version':3,'job':job,'records':[]}
    if scan_scope is not None:manifest['scan_scope']=list(scan_scope)
    from .alignment import align_records
    alignment_rows=align_records(records,history,repairs=repairs)
    aligned={r['current_index']:r['history_index'] for r in alignment_rows if r['current_index'] is not None}
    current_component=None
    rendered_fields={}
    from .reporting import document_events
    for record_index,record in document_events(records,report_sections):
        if record_index is None:
            document.add_paragraph(record);current_component=None;continue
        component=record['component']
        if component!=current_component:
            if component in templates:
                from docx.table import Table
                element=deepcopy(templates[component])
                document._body._body.insert_element_before(element,'w:sectPr')
                table=Table(element,document._body)
                if template_component(table)!=component:
                    for cell in table.rows[0].cells:
                        if compact(cell.text) and compact(cell.text)!='檢測項目':
                            cell.text=component
                            break
            else:table=basic_table(document,component)
            add_price(table)
            add_website_id(table)
            prototype=deepcopy(table.rows[2]._tr) if len(table.rows)>2 else None
            if prototype is None:prototype=deepcopy(table.add_row()._tr)
            for row in list(table.rows)[2:]:table._tbl.remove(row._tr)
            headers=[compact(c.text) for c in table.rows[1].cells]
            if not all(k in headers for k in ('劣化類型','維修工法','數量','單位','金額','備註')):
                raise ValueError(f'{component} 前期表格缺少必要欄位，無法保持版型生成')
            current_component=component
        current=record['current']
        previous_index=aligned.get(record_index)
        scoped=[history[previous_index]] if previous_index is not None else []
        from .data import eligible
        healthy=eligible(*record['der']) is False
        plan=automatic_plan(current,component,scoped,component_repairs(repairs,component),record['context'])
        if healthy:
            fields=dict(current);reason='良好／無劣化，保留原檢測資料'
        elif drafts is not None:
            fields=dict(drafts[record['id']]['fields']);reason=drafts[record['id']].get('reason','人工修訂')
        elif not plan.review:
            fields=plan.fields;reason='Excel'
        elif len(scoped)==1:
            fields=prepare_fields(plan.fields,component=component,context=record['context']);reason='待人工：'+plan.reason if plan.rule_id else '前期資料'
            for key in ('method','unit','price'):
                if not str(fields.get(key,'')).strip():fields[key]=scoped[0].fields.get(key,'')
        else:
            fields=prepare_fields(plan.fields,component=component,context=record['context']);reason='待人工：'+plan.reason
        needs_review=False if healthy else reason.startswith('待人工') or not complete(fields)
        for key,options in ({} if healthy else record.get('choices',{})).items():
            choices=[option for option in options if text_key(option)==text_key(fields.get(key,''))]
            if len(choices)==1:fields[key]=choices[0]
            else:needs_review=True
        rendered_fields[record['id']]=dict(fields)
        row_element=deepcopy(prototype);table._tbl.append(row_element)
        row=table.rows[-1]
        for header,cell in zip(headers,row.cells):
            if header=='劣化ID':value=record.get('website_id','')
            elif header in HEADERS:value=fields.get(HEADERS[header],'')
            elif header in ('D','E','R','U'):
                if header in ('D','E'):value=record['der'][('D','E').index(header)]
                else:
                    import re
                    match=re.search(rf'{header}:([^\s]+)',record['context'].get('DERU',''))
                    value=match[1] if match else record['context'].get(header,'')
            else:value=context_value(record['context'],header)
            paragraphs=cell.paragraphs
            paragraph=paragraphs[0]
            if paragraph.runs:
                paragraph.runs[0].text=str(value)
                for run in list(paragraph.runs)[1:]:paragraph._p.remove(run._r)
            else:paragraph.add_run(str(value))
            for extra in paragraphs[1:]:cell._tc.remove(extra._p)
        if '損壞位置' not in headers and not fields.get('remark','').startswith(fields.get('d_location','')+' '):
            # 備註中的損壞位置為網站欄位來源，不能只靠構件位置推算。
            needs_review=True
        color='FFF2CC' if needs_review else None
        if review_colors and not healthy:
            row_status=next(r['kind'] for r in alignment_rows if r['current_index']==record_index)
            color='FFD8D8' if row_status=='new' else 'FFF2CC' if row_status=='ambiguous' else None
            needs_review=needs_review or row_status!='matched'
        shade(row,color)
        for node in list(row._tr.iter(qn('w:bookmarkStart')))+list(row._tr.iter(qn('w:bookmarkEnd'))):
            node.getparent().remove(node)
        marker=OxmlElement('w:bookmarkStart');marker.set(qn('w:id'),str(len(manifest['records'])+1))
        marker.set(qn('w:name'),'BF_'+digest(record['id'])[:32])
        end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),marker.get(qn('w:id')))
        row.cells[0].paragraphs[0]._p.append(marker);row.cells[0].paragraphs[0]._p.append(end)
        manifest['records'].append(dict(id=record['id'],component=component,source=digest([current,record['der']]),
            table=len(document.tables)-1,row=len(table.rows)-1,context=record['context'],
            generated=digest(fields),needs_review=needs_review,reason=reason,
            d_location=fields.get('d_location',''),website_id=record.get('website_id',''),owner_id=record.get('owner_id',''),rule_id=(drafts[record['id']].get('rule_id','') if drafts is not None else plan.rule_id),current=current,der=record['der'],choices=record.get('choices',{})))
    if review_colors:
        displayed=align_records([dict(r,current=rendered_fields[r['id']]) for r in records],history,repairs=repairs)
        for paired in displayed:
            index=paired['current_index']
            if index is None:continue
            entry=manifest['records'][index]
            color='FFD8D8' if paired['kind']=='new' else 'FFF2CC' if paired['kind']=='ambiguous' else None
            if eligible(*entry['der']) is False:color=None
            shade(document.tables[entry['table']].rows[entry['row']],color)
            entry['needs_review']=entry['needs_review'] or paired['kind']!='matched'
    root=etree.Element('{'+NS+'}manifest');root.text=json.dumps(manifest,ensure_ascii=False)
    part=Part(PackURI('/customXml/bridgefill.xml'),'application/xml',etree.tostring(root),document.part.package)
    document.part.relate_to(part,RT.CUSTOM_XML)
    # Word 的資料存放區屬性，讓編輯後仍保留這份自訂 XML。
    ds='http://schemas.openxmlformats.org/officeDocument/2006/customXml'
    properties=etree.Element('{'+ds+'}datastoreItem',nsmap={'ds':ds})
    properties.set('{'+ds+'}itemID','{'+str(uuid.uuid4()).upper()+'}')
    refs=etree.SubElement(properties,'{'+ds+'}schemaRefs')
    etree.SubElement(refs,'{'+ds+'}schemaRef').set('{'+ds+'}uri',NS)
    props=Part(PackURI('/customXml/bridgefillProps.xml'),
        'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
        etree.tostring(properties),document.part.package)
    part.relate_to(props,RT.CUSTOM_XML_PROPS)
    output=Path(output_path)
    output.parent.mkdir(parents=True,exist_ok=True)
    fd,working=tempfile.mkstemp(prefix='.document-',suffix='.docx',dir=output.parent)
    os.close(fd)
    try:
        from .document_style import apply_document_font
        if report_sections is not None:
            from .document_style import format_summary
            format_summary(document)
        else:apply_document_font(document)
        document.save(working)
        os.replace(working,output)
    finally:
        if os.path.exists(working):os.unlink(working)
    return len(records),sum(r['needs_review'] for r in manifest['records'])


def read_generated(path, strict=True):
    document=Document(path)
    manifests=[]
    for rel in document.part.rels.values():
        if rel.reltype==RT.CUSTOM_XML:
            try:
                root=etree.fromstring(rel.target_part.blob)
                if root.tag=='{'+NS+'}manifest':manifests.append(json.loads(root.text))
            except (ValueError,etree.XMLSyntaxError):pass
    if len(manifests)!=1 or manifests[0].get('version')!=1:
        raise ValueError('不是助手生成的本期 DOCX，或識別資料已遺失；請重新生成')
    manifest=manifests[0]
    references=set()
    for entry in manifest['records']:
        try:
            table=document.tables[entry['table']];row=table.rows[entry['row']]
            if template_component(table)!=entry['component']:raise ValueError()
            markers=[node.get(qn('w:name')) for node in row._tr.iter(qn('w:bookmarkStart'))]
            if markers!=['BF_'+digest(entry['id'])[:32]]:
                raise ValueError()
            headers=[compact(c.text) for c in table.rows[1].cells]
            fields={HEADERS[h]:row.cells[i].text.strip() for i,h in enumerate(headers) if h in HEADERS}
        except (IndexError,ValueError):raise ValueError('文件構件、欄位或列結構已改變，請勿增刪或搬動資料列')
        if '劣化ID' in headers:
            native=row.cells[headers.index('劣化ID')].text.strip()
            if entry.get('website_id') and native!=entry['website_id']:raise ValueError('網站劣化 ID 欄已變更，請重新掃描建立文件')
            entry['website_id']=native
        fields.setdefault('d_location',entry.get('d_location',''))
        fields=prepare_fields(fields)
        from .data import eligible
        healthy=eligible(*entry.get('der',['',''])) is False
        if strict and not healthy and entry['needs_review'] and highlighted(row) and digest(fields)==entry['generated']:
            raise ValueError('文件仍有未處理的紅色／黃色列，請補齊或確認後取消底色')
        if strict and not healthy and not complete(fields):raise ValueError('文件仍有缺少欄位的資料列，請先完成手動處理（含金額單價）')
        # 構件位置及編號亦必須維持，不能移動列後沿用隱藏識別。
        for i,h in enumerate(headers):
            if h not in HEADERS and h not in ('D','E','R','U','劣化ID'):
                if row.cells[i].text.strip()!=context_value(entry['context'],h):
                    raise ValueError('文件的構件位置或編號已變動，停止避免錯位填寫')
        entry['fields']=fields
        references.add((entry['table'],entry['row']))
    actual={(ti,ri) for ti,t in enumerate(document.tables) if template_component(t)
            for ri in range(2,len(t.rows))}
    if references!=actual:raise ValueError('文件資料列數已變動，請重新生成')
    ids=[r['id'] for r in manifest['records']]
    if len(set(ids))!=len(ids):raise ValueError('生成文件有重複缺失識別')
    return manifest


def validate_generated(manifest,job,records,components=None):
    expected=manifest['records']
    if components is not None:
        available={r['component'] for r in expected}
        if not set(components)<=available:
            raise ValueError('所選構件不在生成文件中，請改選或重新生成')
        expected=[r for r in expected if r['component'] in components]
    if not expected:raise ValueError('所選構件沒有可填寫的文件資料')
    if manifest['job']!=job:raise ValueError('生成文件不屬於目前橋梁／期別，未填寫')
    if [r['id'] for r in expected]!=[r['id'] for r in records]:
        raise ValueError('本期缺失或順序已變動，請重新生成文件，未填寫')
    for old,current in zip(expected,records):
        if old.get('website_id') and old['website_id']!=current.get('website_id'):
            raise ValueError('網站缺失 ID 已變動，停止避免錯位填寫')
        if old['context']!=current['context']:
            raise ValueError('本期構件條件或 DERU 已變動，請重新生成文件，未填寫')
        if old['source']!=digest([current['current'],current['der']]) and old['fields']!=current['current']:
            raise ValueError('本期網站資料已變動，請重新生成文件，未填寫')
        for key,options in current['choices'].items():
            if old['fields'].get(key) not in options:
                raise ValueError('文件的類型或單位不在目前網站選項內，未填寫')
    return {r['id']:r['fields'] for r in expected}


def bind_document(manifest,job,records,components=None,overrides=None):
    from .binding import bind
    return bind(manifest,job,records,components,overrides)
