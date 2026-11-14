"""以網站劣化 ID 直接填寫；不進行整批預掃描。"""
from copy import deepcopy
import re
from .data import eligible
from .config import FIELD_IDS,LABELS
from .documents import complete
from .binding import FillMismatch


def id_entries(manifest,components=None,require_complete=True,allow_empty=False):
    entries={};problems=[]
    for number,entry in enumerate(manifest.get('records',[]),1):
        if components is not None and entry['component'] not in components:continue
        if eligible(*entry.get('der',['',''])) is False:continue
        native=str(entry.get('website_id') or '').strip()
        reason=''
        if not re.fullmatch(r'\d+',native):reason='缺少有效劣化 ID，請先使用「從網站讀取本期缺失」'
        elif native in entries:reason='劣化 ID 重複：'+native
        elif require_complete and not complete(entry.get('fields',{})):
            missing=[LABELS[k] for k in FIELD_IDS if not str(entry.get('fields',{}).get(k,'')).strip()]
            reason='文件待補：'+('、'.join(missing) or '備註量式')
        if reason:problems.append(dict(file_index=number,record_id=entry.get('id',''),component=entry['component'],reason=reason,expected=deepcopy(entry),actual=None,candidates=[],fix='先掃描建立帶 ID 的本期文件，再於兩期比對完成修改並生成文件；不使用順序或樣態猜測 ID。'))
        else:entries[native]=dict(entry,website_id=native,file_index=number)
    if problems:raise FillMismatch(problems)
    if not entries and not allow_empty:raise ValueError('所選構件沒有可填寫的缺失；D＝1／E＝0 不填寫')
    return entries
