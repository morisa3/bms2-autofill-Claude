"""核對時修改本期欄位：Excel 規則載入、備註樣態替換、樣態標籤排序。

修改只存在審查檔（Review.edits），第 6 步填寫時才寫入網站。
"""
import re

from .data import component_repairs
from .local_files import APP_DIR, REFERENCE_DIR
from .matching import LOCATION_CODE, clean_pattern, pattern_similarity

EDIT_FIELDS = ('d', 'e', 'r', 'u', 'deterioration', 'd_location', 'method', 'amount', 'unit', 'price', 'remark')
FIELD_NAMES = {'d': 'D', 'e': 'E', 'r': 'R', 'u': 'U', 'deterioration': '劣化類型', 'd_location': '損壞位置',
               'method': '維修工法', 'amount': '數量', 'unit': '單位', 'price': '單價', 'remark': '備註'}


def rules_file():
    """優先使用 data 資料夾中的 v2 規則表，其次是 docs/excel 的建議版。"""
    for folder in (REFERENCE_DIR, APP_DIR / 'docs' / 'excel'):
        if folder.is_dir():
            found = sorted(p for p in folder.glob('*v2*.xlsx') if not p.name.startswith('~$'))
            if found:
                return found[0]
    return None


def load_rules(path=None):
    """讀不到時回傳空清單（仍可手動修改，只是沒有 Excel 選項）。"""
    from .rules_v2 import read_rules
    path = path or rules_file()
    if not path:
        return []
    try:
        return read_rules(path)
    except Exception:  # 規則表格式錯誤不影響核對
        return []


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
