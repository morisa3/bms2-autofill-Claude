"""檢測案件以網站案件及橋梁 ID 識別；不以表單送出後的顯示路徑識別。"""
import re
from urllib.parse import urlsplit,parse_qs
from .progress import digest


def case_key(url):
    parsed=urlsplit(url);values={key:[] for key in ('inspect_id','inventory_id')}
    for part in parsed.path.split('/'):
        key,sep,value=part.partition(',')
        if sep and key in values:values[key].append(value)
    for key,items in parse_qs(parsed.query).items():
        if key in values:values[key].extend(items)
    if not all(len(v)==1 and re.fullmatch(r'\d+',v[0]) for v in values.values()):return None
    return f'case:v1:{parsed.scheme}://{parsed.netloc}/inspect/{int(values["inspect_id"][0])}/inventory/{int(values["inventory_id"][0])}'


def case_job(base,url):
    return digest([base,case_key(url) or url])


def read_case_metadata(scope):
    return scope.evaluate("""() => {
        const labels=['橋梁名稱','橋梁編號','檢測日期'], result={};
        for(const label of labels){
            const values=new Set();
            for(const cell of document.querySelectorAll('th,td')){
                if(cell.textContent.replace(/\\s/g,'')!==label)continue;
                const next=cell.nextElementSibling;if(!next)continue;
                const field=next.querySelector('input,select,textarea');
                const value=String(field?field.value:next.textContent).trim();if(value)values.add(value);
            }
            if(values.size===1)result[label]=[...values][0];
        }
        return result;
    }""")


def matching_metadata(expected,actual):
    clean=lambda v:re.sub(r'\s+','',str(v or ''))
    def date(v):
        match=re.fullmatch(r'(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?',clean(v))
        return tuple(map(int,match.groups())) if match else None
    if not clean(expected.get('橋梁編號')) or clean(expected.get('橋梁編號'))!=clean(actual.get('橋梁編號')):return False
    a,b=date(expected.get('檢測日期')),date(actual.get('檢測日期'))
    if a is None or a!=b:return False
    return not(expected.get('橋梁名稱') and actual.get('橋梁名稱')) or clean(expected['橋梁名稱'])==clean(actual['橋梁名稱'])
