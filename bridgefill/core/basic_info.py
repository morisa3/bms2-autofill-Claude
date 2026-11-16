"""缺失主頁面上方的檢測基本資料（與檢測員意見同一個表單一起送出）。

預設值：本期已有的值保留；空白的沿用前期；單位主管依管理機關自動帶入（「單位主管設定」可修改對照）。
沖刷現象為「無」時，沖刷位置／備註不填。
"""
from .inspection import supervisor_for

# (欄位名稱, 表單選取器, 種類)
FIELDS = (
    ('單位主管', '#chief', 'text'),
    ('近端定義', 'select[name="frontside"]', 'select'),
    ('是否疑似鹽害橋梁', '#salt_injury', 'select'),
    ('有無沖刷現象', '#scour_type', 'select'),
    ('沖刷位置/備註', '#scour_site', 'textarea'),
    ('橋下有無堆放雜物', '#waste', 'select'),
)
NAMES = [name for name, _, _ in FIELDS]
OPTIONS = {'近端定義': ['里程樁號較小端', '竣工圖上之橋頭端', '其他'], '是否疑似鹽害橋梁': ['是', '否', '無法判斷'],
           '有無沖刷現象': ['有', '無', '無法判斷'], '橋下有無堆放雜物': ['否', '是']}


def clean(value):
    return str(value or '').strip()


def options(header):
    """下拉選項：以掃描到的網站選項為準，沒有時用預設。"""
    found = header.get('選項') or {}
    return {name: found.get(name) or OPTIONS[name] for name in OPTIONS}


def defaults(header, prior_header, settings, agency=''):
    """回傳 ({欄位: 值}, {欄位: 來源說明})。"""
    values, sources = {}, {}
    agency = agency or header.get('管理機關', '') or prior_header.get('管理機關', '')
    for name in NAMES:
        current, prior = clean(header.get(name)), clean(prior_header.get(name))
        if name == '單位主管':
            chief = supervisor_for(agency, settings) if agency else ''
            if current:
                values[name], sources[name] = current, '本期'
            elif chief:
                values[name], sources[name] = chief, f'依管理機關（{agency}）'
            elif prior:
                values[name], sources[name] = prior, '沿用前期'
            else:
                values[name], sources[name] = '', '請填寫'
            continue
        if current:
            values[name], sources[name] = current, '本期'
        elif prior:
            values[name], sources[name] = prior, '沿用前期'
        else:
            values[name], sources[name] = '', '請選擇' if name in OPTIONS else '請填寫'
    if values.get('有無沖刷現象') == '無':
        values['沖刷位置/備註'], sources['沖刷位置/備註'] = '', '無沖刷，不填'
    return values, sources


def changes(header, values):
    """與網站現值不同的欄位。"""
    return {name: clean(values.get(name)) for name in NAMES if clean(values.get(name)) != clean(header.get(name))}
