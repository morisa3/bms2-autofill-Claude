"""前後期 U 值與報告顯示，不改寫網站評等。"""
import re
from html import escape
from .deru import deru_records


def rating(record,name):
    context=record.get('context',{})
    value=str(context.get(name,'')).strip()
    if not value:
        codes=deru_records(str(context.get('DERU','')))
        if codes and len(codes)==1:value=codes[0]['DERU'.index(name)]
    return int(value) if value in ('0','1','2','3','4') else None


def u_change(current,previous):
    a,b=rating(current,'U'),rating(previous,'U')
    return None if a is None or b is None else a-b


def rich_metadata(record,previous=None,historical=False):
    from .view_helpers import comparison_metadata
    text=comparison_metadata(record)
    if historical:text='前期｜'+'｜'.join(text.split('｜')[1:-1])
    text=escape(text)
    change=u_change(record,previous) if previous is not None else None
    if change:
        value=rating(record,'U');color='#c62828' if change>0 else '#1565c0'
        text=re.sub(rf'U:{value}(?!\d)',f'<span style="font-weight:700;color:{color}">U:{value}</span>',text)
    return text


def document_events(records,sections=None):
    if sections is None:
        yield from enumerate(records);return
    for title,identities in sections:
        yield None,title
        selected=[(i,r) for i,r in enumerate(records) if r['id'] in identities]
        if not selected:yield None,'本次無此類劣化。'
        yield from selected
