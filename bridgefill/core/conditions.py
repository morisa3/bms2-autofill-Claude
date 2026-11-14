"""Excel 備註中明確的主頁型式／高度條件；未知條件交人工。"""
from decimal import Decimal
import re
from .data import compact,normalize_remark


TYPE_CORRECTION=re.compile(r'(?:若|但)?劣化類型填成\s*[\[【「]([^\]】」]+)[\]】」][，,]?則(?:直接)?修正成\s*[\[【「]([^\]】」]+)[\]】」](?:後使用)?')


def rule_deterioration(rule,current=''):
    """只解讀明確的欄位修正格式，不執行其他備註文字。"""
    match=TYPE_CORRECTION.search(compact(rule.get('note','')))
    if match and compact(current) in (compact(match[1]),compact(match[2])):return match[2]
    return rule.get('deterioration') or current


def keyword_groups(note):
    """各組內為 AND，組間為 OR；型式、寬度門檻不當作缺失關鍵字。"""
    excluded = [match.span() for match in re.finditer(
        r'先判讀是否為\s*[\[【「][^\]】」]+[\]】」]|裂縫寬度[^，,。;；]*', note)]
    excluded.extend(match.span() for match in TYPE_CORRECTION.finditer(note))
    words = [match for match in re.finditer(r'[\[【「]([^\]】」]+)[\]】」]', note)
             if not any(start <= match.start() < end for start, end in excluded)]
    groups = []
    previous = None
    for match in words:
        choices = tuple(piece.strip() for piece in re.split(r'[|｜]', match[1]) if piece.strip())
        if not choices:
            continue
        connector = compact(note[previous.end():match.start()]) if previous else ''
        if groups and connector in ('與', '及', '且', '以及', '並且'):
            groups[-1].append(choices)
        else:
            groups.append([choices])
        previous = match
    if '裂縫寬度' in note:
        groups.append([('裂縫',)])
    return groups


def keyword_score(note, text):
    canonical=lambda s:compact(s).replace('生鏽','生銹').replace('生锈','生銹')
    text=canonical(text)
    scores = []
    for group in keyword_groups(note):
        terms = [[len(canonical(word)) for word in choices if canonical(word) in text] for choices in group]
        if terms and all(terms):
            scores.append(sum(max(term) for term in terms))
    return max(scores) if scores else None


def aliases(note):
    words=list(dict.fromkeys(word for group in keyword_groups(note) for choices in group for word in choices))
    if any(word in words for word in ('生銹','生鏽','生锈')):words=list(dict.fromkeys(words+['生銹','生鏽','生锈']))
    return words


def condition(rule, context, remark=None):
    """回傳 True/False/None（資料不足），及是否必須人工選擇。"""
    note = rule.get('note', '')
    if TYPE_CORRECTION.search(compact(note)):
        note=TYPE_CORRECTION.sub('',compact(note))
        # 「若判讀到關鍵字」已由 keyword_score 篩選，不是未知主頁條件。
        note=re.sub(r'若判讀到(?=關鍵字)','有',note)
    facts = {compact(k): compact(v) for k, v in (context or {}).items()}
    kinds = re.findall(r'先判讀是否為\s*[\[【「]([^\]】」]+)[\]】」]', note)
    for kind in kinds:
        if kind in ('齒型', '角鋼'):
            family = ('伸縮',)
        elif '支承' in kind:
            family = ('支承',)
        elif '防落' in kind or '防震' in kind:
            family = ('防落', '落橋', '止震')
        else:
            return None, True
        values = [v for k, v in facts.items() if any(word in k for word in family) and
                  any(word in k for word in ('型式', '類型', '種類', '名稱'))]
        if not values:
            values = [v for k, v in facts.items() if k in ('型式', '類型', '種類')]
        if len(values) != 1 or not values[0]:
            return None, True
        matched = compact(kind) in values[0]
        if kind == '人造橡膠支承墊':
            matched = bool(re.search(r'(?:人造|合成)橡膠支承(?:墊)?', values[0]))
        if not matched:
            return False, False

    fallback = re.search(r'或僅有(?:關鍵字)?[\[【「]([^\]】」]+)[\]】」]時[，,]?(?:則|若)?U值為([1-4])優先使用', compact(note))
    if fallback:
        primary = note.split('或僅有', 1)[0]
        # 明確草本／木本詞優先；只有泛稱植生才依目前缺失的 U 值。
        if keyword_score(primary, compact(remark or '')) is not None:
            return True, bool(re.search(r'手動選擇|手動處理', note))
        if compact(fallback[1]) not in compact(remark or ''):
            return False, False
        urgency = facts.get('U', '')
        if urgency not in ('1', '2', '3', '4'):
            return None, True
        return urgency == fallback[2], '手動選擇' in note

    width_rule = re.fullmatch(r'裂縫寬度(大於等於|小於等於|大於|小於)[\[【]?(\d+(?:\.\d+)?)(mm|cm|m)[\]】]?使用', compact(note))
    if width_rule:
        source = compact(normalize_remark(remark or '')).translate(str.maketrans({'＊':'*', '×':'*', '（':'(', '）':')'}))
        size = re.search(r'\(([^()]+)\)(?:\*\d+(?:\.\d+)?)?$', source)
        if not size:
            return None, True
        factors = size[1].split('*')
        widths = [re.fullmatch(r'(\d+(?:\.\d+)?|\.\d+)mm', factor) for factor in factors]
        widths = [match for match in widths if match]
        lengths = [factor for factor in factors if re.fullmatch(r'(?:\d+(?:\.\d+)?|\.\d+)m', factor)]
        if len(widths) != 1 or len(lengths) != 1:
            return None, True
        width = Decimal(widths[0][1])
        limit = Decimal(width_rule[2]) * {'mm':Decimal(1), 'cm':Decimal(10), 'm':Decimal(1000)}[width_rule[3]]
        valid = {'大於':width > limit, '小於':width < limit,
                 '大於等於':width >= limit, '小於等於':width <= limit}[width_rule[1]]
        # 使用者指定：0.3mm 邊界歸入原「大於 0.3mm」計價。
        if width_rule[1]=='大於' and limit==Decimal('.3'):valid=width>=limit
        return valid, '手動選擇' in note
    if '裂縫寬度' in note or 'U值' in note:
        return None, True
    comparison = r'(大於等於|小於等於|大於|小於|>=|<=|>|<)\s*(\d+(?:\.\d+)?|\.\d+)\s*(mm|cm|m)'
    bounds = re.findall(comparison, note)
    if bounds:
        clause = re.search(r'護欄高(?:度)?([^()（）,，。;；]+)', compact(note))
        if not clause:
            return None, True
        expression = re.sub(r'(?:時適用|時使用|適用|使用|時)$', '', clause[1])
        expression = re.sub(r'並且|且|以及|及', '', expression)
        if not re.fullmatch('(?:' + comparison + ')+', expression):
            return None, True
        values = [(k, v) for k, v in facts.items() if '護欄' in k and '高度' in k]
        if len(values) != 1:
            return None, True
        header, value = values[0]
        match = re.fullmatch(r'(\d+(?:\.\d+)?)(mm|cm|m)?', value)
        if not match:
            return None, True
        scale = {'mm': Decimal('.001'), 'cm': Decimal('.01'), 'm': Decimal(1)}
        header_unit = re.search(r'[（(\[](mm|cm|m)[）)\]]', header)
        height = Decimal(match[1]) * scale[match[2] or (header_unit[1] if header_unit else 'm')]
        for operator, value, unit in bounds:
            limit = Decimal(value) * scale[unit]
            valid = {'大於': height > limit, '>': height > limit, '小於': height < limit, '<': height < limit,
                     '大於等於': height >= limit, '>=': height >= limit,
                     '小於等於': height <= limit, '<=': height <= limit}[operator]
            if not valid:
                return False, False
    # Excel explicitly defines bare exposure as the foundation-exposure fallback.
    note=re.sub(r'若僅寫\[裸露\]則優先使用此項','',note)
    if re.search(r'若|當|僅|限', note) and '若過去資料為' not in note:
        return None, True
    if (('高度' in note or '護欄高' in note) and not bounds) or ('先判讀' in note and not kinds):
        return None, True
    return True, bool(re.search(r'手動選擇|手動處理', note))
