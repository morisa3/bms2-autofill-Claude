"""核對時修改本期欄位：Excel 規則載入、備註樣態替換、樣態標籤排序。

修改只存在審查檔（Review.edits），第 6 步填寫時才寫入網站。
"""
import re
from pathlib import Path

from .data import component_repairs
from .local_files import REFERENCE_DIR
from .matching import LOCATION_CODE, clean_pattern, pattern_similarity

EDIT_FIELDS = ('d', 'e', 'r', 'u', 'deterioration', 'd_location', 'method', 'amount', 'unit', 'price', 'remark')
FIELD_NAMES = {'d': 'D', 'e': 'E', 'r': 'R', 'u': 'U', 'deterioration': '劣化類型', 'd_location': '損壞位置',
               'method': '維修工法', 'amount': '數量', 'unit': '單位', 'price': '單價', 'remark': '備註'}


RULES_NAME = '劣化說明與維修單價1122.xlsx'


def rules_file():
    """使用中的規則表：data/劣化說明與維修單價1122.xlsx；沒有時找 data 中其他結構化規則表（舊版 1111v2）。"""
    main = REFERENCE_DIR / RULES_NAME
    if main.exists():
        return main
    found = sorted(p for p in REFERENCE_DIR.glob('*v2*.xlsx') if not p.name.startswith('~$')) \
        if REFERENCE_DIR.is_dir() else []
    return found[0] if found else None


def load_rules(path=None):
    """讀不到時回傳空清單（仍可手動修改，只是沒有 Excel 選項）。"""
    return load_rules_checked(path)[0]


def load_rules_checked(path=None):
    """讀取規則表並回報原因：回傳 (規則清單, 錯誤說明)；成功時錯誤說明為空白。"""
    from .rules_v2 import read_rules
    from .synonyms import configure
    path = path or rules_file()
    if not path:
        return [], f'data 資料夾中找不到規則表（{RULES_NAME}）'
    if not Path(path).is_file():
        return [], f'找不到規則表：{path}'
    configure(path)                       # 同一份規則表的「樣態同義詞」「位置代碼對照」
    try:
        rules = read_rules(path)
    except PermissionError:
        return [], f'規則表正在 Excel 中開啟，無法讀取：{Path(path).name}（請關閉後按「重新讀取」）'
    except Exception as error:  # 規則表格式錯誤不影響核對
        return [], f'規則表格式無法讀取：{error}'
    if not rules:
        return [], f'規則表中沒有讀到任何規則（工作表需有「劣化樣態、維修工法、維修單位、單價、劣化ID、處理方式」欄）：{Path(path).name}'
    return rules, ''


def split_remark(remark):
    """「A1-上下游 植生 (2處)」→ ('A1-上下游', '植生', ' (2處)')；找不到位置時位置為空。"""
    text = (remark or '').strip()
    found = re.match(r'^(\S+)\s+(.*)$', text)
    if found and (LOCATION_CODE.match(found[1]) or re.search(r'[\-－]|整體|上下游|橋頭|橋尾', found[1])):
        location, rest = found[1], found[2]
    else:
        glued = re.match(r'^([A-Za-z]*\d[A-Za-z0-9,，、~～\-_.]*)(?=[\u4e00-\u9fff])', text)   # 「P2、P3掛淤」
        location, rest = (glued[1], text[glued.end():]) if glued else ('', text)
    body = re.match(r'^(.*?)([。(（].*)?$', rest)
    return location, body[1].strip(), body[2] or ''


def replace_pattern(remark, pattern):
    """把備註中的樣態換成 pattern，保留位置與尺寸。"""
    location, _, tail = split_remark(remark)
    if tail and not tail.startswith('。'):
        tail = '。' + tail.lstrip()
    return ' '.join(x for x in (location, pattern + (tail or '。')) if x)


def ranked_rules(rules, component, fields):
    """這個構件可用的 Excel 規則，依與目前備註樣態的相似度排序。"""
    own = clean_pattern(fields)
    found = component_repairs(rules, component)
    return sorted(found, key=lambda r: -(pattern_similarity(own, r['pattern']) or 0))


def pattern_tags(rules, component, fields, limit=8):
    seen = []
    for rule in ranked_rules(rules, component, fields):
        if rule['pattern'] not in seen:
            seen.append(rule['pattern'])
    return seen[:limit]


def best_rule(rules, pattern, unit=''):
    same = [r for r in rules if r['pattern'] == pattern]
    return next((r for r in same if r['unit'] == unit), same[0] if same else None)


def describe(changes):
    return '；'.join(f'{FIELD_NAMES.get(k, k)}：{v}' for k, v in changes.items())
