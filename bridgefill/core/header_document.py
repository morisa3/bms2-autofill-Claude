"""保留及更新檢測表上方基本資料，不改動缺失識別。"""
import re
from docx.shared import Pt
from .inspection import HEADER_ALIASES


def apply_header(document, info):
    info={k:str(v) for k,v in (info or {}).items() if k in set(HEADER_ALIASES.values()) and v is not None}
    found=set()
    for table in document.tables:
        for row in table.rows:
            cells=[];seen=set()
            for cell in row.cells:
                if cell._tc in seen:continue
                seen.add(cell._tc);cells.append(cell)
            for i,cell in enumerate(cells[:-1]):
                label=re.sub(r'\s+','',cell.text).rstrip('：:')
                key=HEADER_ALIASES.get(label)
                if key in info:
                    target=cells[i+1]
                    paragraph=target.paragraphs[0]
                    if paragraph.runs:
                        paragraph.runs[0].text=info[key]
                        for run in paragraph.runs[1:]:run.text=''
                    else:paragraph.add_run(info[key])
                    for extra in target.paragraphs[1:]:target._tc.remove(extra._p)
                    found.add(key)
    missing=[k for k in info if k not in found and info[k].strip()]
    if missing:
        table=document.add_table(rows=0,cols=2)
        table.style='Table Grid'
        for key in missing:
            row=table.add_row();row.cells[0].text=key;row.cells[1].text=info[key]
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:run.font.size=Pt(10)
        document._body._body.insert(0,table._tbl)
