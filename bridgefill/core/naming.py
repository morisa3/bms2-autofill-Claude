"""橋名與輸出檔名；日期取使用者所在地臺灣當日。"""
from datetime import datetime, timezone, timedelta
from pathlib import Path
import re


def bridge_name(value):
    text=Path(str(value)).stem.split('定期檢測表')[0]
    groups=re.findall(r'[\[【]([^\]】]+)[\]】]',text)
    names=[s for s in groups if not re.fullmatch(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{4,8}',s)]
    if names:text=names[-1]
    else:text=re.sub(r'^\[?\d{4}[-/]\d{1,2}[-/]\d{1,2}[\] _-]*','',text)
    return text.strip('[]【】_- ') or '未命名橋梁'


def dated_bridge_name(value, now=None):
    now=now or datetime.now(timezone(timedelta(hours=8)))
    name=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',bridge_name(value)).strip(' .')
    return f'[{now:%Y-%m-%d}][{name}]'
