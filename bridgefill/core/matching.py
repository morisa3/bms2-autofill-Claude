"""前期與本期缺失的評分配對（第 4 步）。

1. 列對列：前期與本期以網站構件識別（component_form＋component_id）直接對上，不靠文字猜測。
2. 同一列內的缺失：依樣態、損壞位置、劣化類型、單位與數量、DERU 計分，找總分最高的一對一組合。
3. 判定：沿用、範圍加大、範圍縮小、單位數量變更、新增、已維修、疑似漏建；附信心（高／需確認）與理由。

本模組只做判讀，不修改任何資料。
"""
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from itertools import permutations

from .alignment import pattern
from .locations import key as location_key

ACCEPT = 45        # 低於此分數不配對
CONFIDENT = 75     # 達此分數為「高」信心

# 比對樣態時去掉的構件／材料字詞（例：鋼管護欄表面生銹 ≈ 生銹）
GENERIC = re.compile(r'鋼管護欄|護欄|支承墊|支承|防衝鋼板|防撞鋼板|鋼板|伸縮縫|混凝土|結構|表面|AC|瀝青|路面|基座')
# 互斥：前後期分別出現時一律不配對（與 Excel「互斥樣態」相同）
EXCLUSIVE = (('木本', '草本'), ('齒型', '角鋼'), ('鋼管', '鋼筋'), ('AC', '混凝土'), ('瀝青', '混凝土'))

# 意思相同的不同寫法（之後改由 Excel「樣態同義詞」工作表提供）
SYNONYMS = (('未設置護欄', '缺口'), ('未設置', '缺口'), ('無設置護欄', '缺口'), ('雜物堆積', '積土'))

KIND_KEEP, KIND_LARGER, KIND_SMALLER, KIND_UNIT = '沿用', '範圍加大', '範圍縮小', '單位數量變更'
KIND_NEW, KIND_REPAIRED, KIND_MISSING = '新增', '已維修', '疑似漏建'


@dataclass
class Defect:
    period: str            # 前期／本期
    row: tuple             # (component_form, component_id)
    component: str
    label: str
    deru: str
    deterioration: str
    fields: dict
    repaired: bool = False
    method_id: str = ''

    @property
    def remark(self):
        return self.fields.get('remark', '')


@dataclass
class Pair:
    kind: str
    prior: Defect | None = None
    current: Defect | None = None
    score: int = 0
    confidence: str = ''
    reasons: list = field(default_factory=list)

    @property
    def row(self):
        return (self.current or self.prior).row


def number(value):
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None


def defects_from_scan(scan, period):
    """唯讀掃描結果（ScanResult.to_dict）→ 缺失清單；只取有缺失且讀到第二級欄位的筆。"""
    found = []
    for row in scan.get('rows', []):
        key = (row['keys'].get('component_form', ''), row['keys'].get('component_id', ''))
        for item in row.get('defects', []):
            if item.get('status') != '缺失' or not item.get('fields'):
                continue
            found.append(Defect(period, key, row['component'], row.get('label', ''), item.get('deru', ''),
                                item.get('deterioration') or item['fields'].get('deterioration', ''),
                                item['fields'], bool(item.get('repaired')), item.get('method_id', '')))
    return found


# ---------------------------------------------------------------- 計分

def core(text):
    return GENERIC.sub('', text)


GENERIC_TYPES = ('其他損傷', '其他', '表面劣化')
LOCATION_CODE = re.compile(r'^([A-Za-z0-9][A-Za-z0-9,，~～\-_.]*)')


def location_code(value):
    """損壞位置欄位有時被填成「P1-1橋墩基礎裸露」；開頭是英數代碼時只取代碼。"""
    value = (value or '').strip()
    found = LOCATION_CODE.match(value)
    if found and re.search(r'[\u4e00-\u9fff]', value[found.end():]):
        return found[1].rstrip('-_')
    return value


def clean_pattern(fields):
    """v1.00 的樣態，再去掉樣態後直接接的量測值（例：「橋墩基礎裸露0.6M」→「橋墩基礎裸露」）並套用同義詞。"""
    text = pattern({**fields, 'd_location': location_code(fields.get('d_location', ''))})
    text = re.sub(r'(?:約)?[\d.]+(?:CM|MM|M)?$', '', text)
    # 損壞位置欄位空白時，v1.00 只會去掉 A1／P1／S1 開頭；其他位置代碼（如 4G1-20R）也要去掉，否則會讓兩筆看起來相像
    text = re.sub(r'^[A-Z0-9][A-Z0-9,~～\-_.]*(?=[\u4e00-\u9fff])', '', text)
    for old, new in SYNONYMS:
        text = text.replace(old, new)
    # 白華滲水、白華且滲水、滲水及白華……視為同一樣態（v1.00 已把滲水統一為漏水）
    if '白華' in text and '漏水' in text:
        text = re.sub(r'(?:漏水|白華|[且及與和])+', '漏水白華', text, count=1)
        text = re.sub(r'(?<=漏水白華)(?:漏水|白華|[且及與和])+', '', text)
    return text


def pattern_similarity(a, b):
    if not a or not b:
        return None
    ca, cb = core(a), core(b)
    ratio = max(SequenceMatcher(None, a, b, autojunk=False).ratio(),
                SequenceMatcher(None, ca, cb, autojunk=False).ratio() if ca and cb else 0)
    if min(len(ca), len(cb)) >= 2 and (ca in cb or cb in ca):
        ratio = max(ratio, 0.85)
    return ratio


def location(defect):
    """損壞位置：優先用欄位；沒有時取備註開頭。備註開頭為英數字時只取英數字段，
    避免「P1-1橋墩基礎裸露」這種位置與樣態沒有空格的寫法整串被當成位置。"""
    value = location_code(defect.fields.get('d_location') or '')
    if not value:
        remark = defect.remark.strip()
        found = LOCATION_CODE.match(remark) or re.match(r'^([^\s(（]+)', remark)
        value = found[1] if found else ''
    # 規格：備註中的 1G1 即 S1G1（主梁），1D1 同理（橫隔梁）
    value = re.sub(r'^S(\d+)(?=[GD]\d)', r'\1', value, flags=re.I)
    return location_key(value)


def score(prior, current):
    """回傳 (分數, 理由)；互斥時回傳 None。"""
    pa, pc = clean_pattern(prior.fields), clean_pattern(current.fields)
    for x, y in EXCLUSIVE:
        if (x in pa and y in pc) or (y in pa and x in pc):
            return None
    points, reasons = 0, []
    similarity = pattern_similarity(pa, pc)
    if similarity is None:
        # 本期只填了類型與 DERU（現場最低限度），或樣態無法判讀：以類型為主，交由信心判斷
        points += 25
        reasons.append('本期備註未填，以劣化類型判斷' if not current.remark.strip() else '樣態無法判讀，以劣化類型判斷')
    else:
        points += round(similarity * 50)
        if similarity >= 0.85:
            reasons.append('樣態相同' if similarity == 1 else '樣態相近')
        elif similarity < 0.5:
            points -= 10
            reasons.append('樣態不同')
    if prior.deterioration == current.deterioration:
        points += 15
    elif prior.deterioration in GENERIC_TYPES:
        points += 10    # 由「其他損傷」改為具體類型，屬常見的類型修正
        reasons.append(f'劣化類型修正（{prior.deterioration} → {current.deterioration}）')
    else:
        reasons.append(f'劣化類型不同（{prior.deterioration} → {current.deterioration}）')
    lp, lc = location(prior), location(current)
    if lp and lc:
        if lp == lc:
            points += 20
        elif lp.startswith(lc) or lc.startswith(lp):
            points += 10
            reasons.append(f'位置寫法不同（{lp}／{lc}）')
        else:
            points -= 15
            reasons.append(f'位置不同（{lp}／{lc}）')
    if prior.fields.get('unit') == current.fields.get('unit'):
        points += 10
        if number(prior.fields.get('amount')) == number(current.fields.get('amount')):
            points += 5
    elif prior.fields.get('unit') == '式':
        points += 5     # 前期以「式」暫填，改為實際單位
    if prior.deru[:1] and prior.deru[:1] == current.deru[:1]:
        points += 3
    return points, reasons


# ---------------------------------------------------------------- 配對

def best_assignment(priors, currents, scores):
    """同一列內找總分最高的一對一組合（每列筆數少，直接列舉；過多時改逐一取最高分）。"""
    if not priors or not currents:
        return []
    small, large, flip = (priors, currents, False) if len(priors) <= len(currents) else (currents, priors, True)
    if len(large) > 8:
        return greedy(scores)
    best, best_total = [], -1
    for chosen in permutations(range(len(large)), len(small)):
        picks = []
        for i, j in zip(range(len(small)), chosen):
            p, c = (j, i) if flip else (i, j)
            if scores.get((p, c)) is not None and scores[(p, c)][0] >= ACCEPT:
                picks.append((p, c))
        total = sum(scores[pc][0] for pc in picks)
        if total > best_total:
            best, best_total = picks, total
    return best


def greedy(scores):
    used_p, used_c, picks = set(), set(), []
    for (p, c), (points, _) in sorted(scores.items(), key=lambda kv: -kv[1][0]):
        if points >= ACCEPT and p not in used_p and c not in used_c:
            picks.append((p, c))
            used_p.add(p)
            used_c.add(c)
    return picks


def classify(prior, current):
    pu, cu = prior.fields.get('unit', ''), current.fields.get('unit', '')
    if pu != cu:
        return KIND_UNIT
    pa, ca = number(prior.fields.get('amount')), number(current.fields.get('amount'))
    if pa is not None and ca is not None and ca > pa:
        return KIND_LARGER
    if pa is not None and ca is not None and ca < pa:
        return KIND_SMALLER
    return KIND_KEEP


def match_rows(priors, currents):
    """同一列（同一構件）內的配對。"""
    scores = {}
    for i, p in enumerate(priors):
        for j, c in enumerate(currents):
            result = score(p, c)
            if result is not None:
                scores[(i, j)] = result
    picks = best_assignment(priors, currents, scores)
    pairs = []
    for i, j in picks:
        points, reasons = scores[(i, j)]
        # 同一筆前期或本期還有另一個也夠像的候選（分數接近且本身達高信心門檻）才算有競爭
        competing = [s for (a, b), (s, _) in scores.items() if (a == i) != (b == j)
                     and s >= max(points - 10, CONFIDENT)]
        confident = points >= CONFIDENT and not competing
        if competing:
            reasons = reasons + ['同列有其他相近候選']
        pairs.append(Pair(classify(priors[i], currents[j]), priors[i], currents[j], points,
                          '高' if confident else '需確認', reasons))
    matched_p = {i for i, _ in picks}
    matched_c = {j for _, j in picks}
    left_p = [p for i, p in enumerate(priors) if i not in matched_p]
    left_c = [c for j, c in enumerate(currents) if j not in matched_c]
    # 同一列只剩一筆前期、一筆本期，位置與單位都相同：判為同一筆、樣態或類型判斷不同（需求 2.1），一律需確認
    if len(left_p) == 1 and len(left_c) == 1:
        p, c = left_p[0], left_c[0]
        if location(p) and location(p) == location(c) and p.fields.get('unit') == c.fields.get('unit') \
                and score(p, c) is not None:
            pairs.append(Pair(classify(p, c), p, c, score(p, c)[0], '需確認',
                              ['樣態或類型不同，但位置與單位相同，判為同一筆']))
            left_p, left_c = [], []
    for c in left_c:
        reasons = ['同列有前期缺失但都不夠相近'] if priors else []
        # 同位置還有沒配上的前期缺失：可能是類型或樣態判斷不同的同一筆（需求 2.1），列出供人工判斷
        reasons += [f'可能對應前期：{p.remark.strip()}' for p in left_p if location(p) and location(p) == location(c)]
        pairs.append(Pair(KIND_NEW, current=c, confidence='高' if not priors else '需確認', reasons=reasons))
    for p in left_p:
        kind = KIND_REPAIRED if p.repaired else KIND_MISSING
        reasons = [f'可能對應本期：{c.remark.strip()}' for c in left_c if location(c) and location(c) == location(p)]
        pairs.append(Pair(kind, prior=p, confidence='高' if kind == KIND_REPAIRED else '需確認', reasons=reasons))
    return pairs


def row_code(label):
    """構件所在位置代碼：「橋孔編號：S003｜橋面板編號：S3」→ S3，「橋墩編號：P001」→ P1，「編號：A1」→ A1。"""
    for token in re.findall(r'[SPA]\d+', label or ''):
        return token[0] + str(int(token[1:]))
    return ''


def remark_code(defect):
    """備註位置開頭的構件代碼：S4-L → S4，4G1-20R → S4（主梁編號前的數字是孔號），P2-1 → P2。"""
    value = location(defect)
    found = re.match(r'^(\d+)[GD]\d', value) or re.match(r'^S(\d+)', value)
    if found:
        return 'S' + str(int(found[1]))
    found = re.match(r'^([PA])(\d+)', value)
    return found[1] + str(int(found[2])) if found else ''


def check_location(pair):
    """本期備註寫的位置和所在構件不符（例：S3 那列寫成 S4）：加註並改為需確認。"""
    if not pair.current:
        return pair
    own, written = row_code(pair.current.label), remark_code(pair.current)
    if own and written and own[0] == written[0] and own != written:
        pair.reasons = pair.reasons + [f'本期備註位置 {written} 與所在構件 {own} 不符']
        pair.confidence = '需確認'
    return pair


@dataclass
class MatchResult:
    pairs: list

    @property
    def summary(self):
        counts = {}
        for pair in self.pairs:
            counts[pair.kind] = counts.get(pair.kind, 0) + 1
        counts['需確認'] = sum(1 for p in self.pairs if p.confidence == '需確認')
        return counts


def match(prior_defects, current_defects):
    rows = {}
    for d in prior_defects:
        rows.setdefault(d.row, ([], []))[0].append(d)
    for d in current_defects:
        rows.setdefault(d.row, ([], []))[1].append(d)
    pairs = []
    for priors, currents in rows.values():
        pairs += [check_location(pair) for pair in match_rows(priors, currents)]
    return MatchResult(pairs)


def match_scans(prior_scan, current_scan):
    return match(defects_from_scan(prior_scan, '前期'), defects_from_scan(current_scan, '本期'))
