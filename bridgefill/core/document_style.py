"""Word 使用標楷體：完成檔 7.5pt；劣化清單內文 10pt、分類 16pt。"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def apply_document_font(document,points=7.5):
    for style in document.styles:
        if hasattr(style,'font'):
            style.font.name='標楷體';style.font.size=Pt(points)
            style.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'標楷體')
    roots=[document.element]
    roots.extend(part.element for part in document.part.related_parts.values() if hasattr(part,'element'))
    for root in roots:
        for run in root.iter(qn('w:r')):
            props=run.find(qn('w:rPr'))
            if props is None:props=OxmlElement('w:rPr');run.insert(0,props)
            fonts=props.find(qn('w:rFonts'))
            if fonts is None:fonts=OxmlElement('w:rFonts');props.append(fonts)
            for name in ('ascii','hAnsi','eastAsia','cs'):
                fonts.set(qn('w:'+name),'標楷體');fonts.attrib.pop(qn('w:'+name+'Theme'),None)
            for name in ('sz','szCs'):
                size=props.find(qn('w:'+name))
                if size is None:size=OxmlElement('w:'+name);props.append(size)
                size.set(qn('w:val'),str(round(points*2)))


def remove_filter_caption(document):
    from docx.text.paragraph import Paragraph
    for element in list(document.element.iter(qn('w:p'))):
        paragraph=Paragraph(element,document._body)
        lines=paragraph.text.splitlines()
        remaining=[line for line in lines if not ('DERU' in line.upper() and '輸出條件' in line)]
        if len(remaining)!=len(lines):
            if any(line.strip() for line in remaining):paragraph.text='\n'.join(remaining)
            elif element.getparent().tag==qn('w:tc'):paragraph.text=''
            else:element.getparent().remove(element)


def format_summary(document):
    from .documents import template_component
    from .data import compact
    from docx.shared import Twips
    for table in document.tables:
        if not template_component(table):continue
        columns=[i for i,cell in enumerate(table.rows[1].cells) if compact(cell.text) in ('單價','金額','劣化ID','網站ID')]
        grid=table._tbl.tblGrid;total=sum(int(col.get(qn('w:w'),'0')) for col in grid)
        for index in reversed(columns):
            for row in table.rows:
                offset=0
                for cell in list(row._tr.tc_lst):
                    span=cell.grid_span
                    if offset<=index<offset+span:
                        if span>1:cell.get_or_add_tcPr().get_or_add_gridSpan().val=span-1
                        else:row._tr.remove(cell)
                        break
                    offset+=span
            grid.remove(grid[index])
        width=sum(int(col.get(qn('w:w'),'0')) for col in grid)
        if width and total:
            for col in grid:col.set(qn('w:w'),str(round(int(col.get(qn('w:w'),'0'))*total/width)))
            widths=[int(col.get(qn('w:w'),'0')) for col in grid]
            for row in table.rows:
                offset=0
                for cell in row._tr.tc_lst:
                    cell.width=Twips(sum(widths[offset:offset+cell.grid_span]));offset+=cell.grid_span
    apply_document_font(document,10)
    headings={'本期新增劣化','U值上升之劣化','U值下降之劣化'}
    for paragraph in document.paragraphs:
        if paragraph.text in headings:
            for run in paragraph.runs:
                run.font.size=Pt(16)
                size=run._r.get_or_add_rPr().find(qn('w:szCs'))
                if size is not None:size.set(qn('w:val'),'32')
    # 清單供閱讀；不附帶填寫檔的隱藏單價或網站識別清單。
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from .documents import NS
    from lxml import etree
    for rid,rel in list(document.part.rels.items()):
        if rel.reltype==RT.CUSTOM_XML:
            try:is_manifest=etree.fromstring(rel.target_part.blob).tag=='{'+NS+'}manifest'
            except etree.XMLSyntaxError:is_manifest=False
            if is_manifest:document.part.drop_rel(rid)
