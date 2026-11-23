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
from .locations import key as location_key, short_code

ACCEPT = 45        # 低於此分數不配對
CONFIDENT = 75     # 達此分數為「高」信心

# 比對樣態時去掉的構件／材料字詞（例：鋼管護欄表面生銹 ≈ 生銹）
GENERIC = re.compile(r'鋼管護欄|護欄|支承墊|支承|防衝鋼板|防撞鋼板|鋼板|伸縮縫|混凝土|結構|表面|AC|瀝青|路面|基座')
# 互斥：前後期分別出現時一律不配對（與 Excel「互斥樣態」相同）
EXCLUSIVE = (('木本', '草本'), ('齒型', '角鋼'), ('鋼管', '鋼筋'), ('AC', '混凝土'), ('瀝青', '混凝土'))

# 意思相同的不同寫法（之後改由 Excel「樣態同義詞」工作表提供）
SYNONYMS = (('未設置護欄', '缺口'), ('未設置', '缺口'), ('無設置護欄', '缺口'), ('雜物堆積', '積土'))

# 跨構件對應：同一筆缺失前後期可能記在不同構件（例：前期記在河道的掛淤，本期拆到橋墩基礎）。
# (構件群 A, 構件群 B, 樣態關鍵字；空白代表不限)。只用來提示「可能對應」，不自動配對。
# 與 Excel「跨構件對應」工作表相同（草稿，待使用者確認）。
RELATED = (
    (('河道',), ('橋墩基礎', '橋台基礎', '墩柱與帽梁', '橋台', '橋墩／橋基保護設施', '護床工'),
     ('掛淤', '漂流', '淤積', '雜物', '沖刷', '垃圾')),
    (('伸縮縫',), ('橋面板', '主梁', '橫隔梁', '墩柱與帽梁', '橋台'), ('漏水', '白華')),
    (('橋梁排水設施',), ('橋面板', '主梁', '墩柱與帽梁'), ('漏水', '白華', '阻塞')),
    (('引道路堤',), ('引道道路堤保護措施', '引道路堤護欄'), ()),
    (('橋台',), ('橋台基礎', '翼牆／擋土牆'), ()),
    (('墩柱與帽梁',), ('橋墩基礎',), ()),
)
CROSS_SIMILARITY = 0.85

KIND_KEEP, KIND_LARGER, KIND_SMALLER, KIND_UNIT = '沿用', '範圍加大', '範圍縮小', '單位數量變更'
KIND_NEW, KIND_REPAIRED, KIND_MISSING = '新增', '已維修', '疑似漏建'
KIND_MERGED = '與其他缺失重複合併'
# 視為同一缺失的劣化（前期分兩筆記、本期合成一筆）：前期多出的那筆判為「與其他缺失重複合併」，本期不另建
MERGE_GROUPS = (('保護層不足', '鋼筋外露'),)


def merged_into(prior, currents):
    """前期缺失與同列本期某一筆屬於同一組（例：保護層不足／鋼筋外露）時，回傳那一筆本期。"""
    mine = clean_pattern(prior.fields)
    for group in MERGE_GROUPS:
        if not any(w in mine for w in group):
            continue
        for current in currents:
            theirs = clean_pattern(current.fields)
            if any(w in theirs for w in group) and location(current) == location(prior) or \
                    (any(w in theirs for w in group) and not location(prior)):
                if any(w in theirs and w not in mine for w in group) or all(w in theirs for w in group):
                    return current
    return None


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
    der_id: str = ''
    inspect_id: str = ''
    options: dict = field(default_factory=dict)     # 第二級下拉選單（劣化類型、單位）
    context: dict = field(default_factory=dict)     # 構件列資料（型式、護欄高度等）＋U，供 Excel 條件判斷

    @property
    def remark(self):
        return self.fields.get('remark', '')

    @property
    def display(self):
        """畫面顯示的構件名稱：「其他」加上設施類別（例：其他（鋪面））。"""
        kind = next((str(v).strip() for k, v in self.context.items() if '設施類別' in k and str(v).strip()), '')
        return f'{self.component}（{kind}）' if self.component == '其他' and kind else self.component

    @property
    def code(self):
        """清單用的簡短位置代碼（A1、P2/P02N-1、S3/G1、鋪面…）；判斷不出來時用原本的位置文字。"""
        return short_code(self.component, self.context) or self.label

    def damage_url(self, edit=False):
        """第二級缺失頁（下方有損壞照片）。本期用編輯頁（可刪照片），前期用檢視頁。"""
        if not (self.der_id and self.inspect_id):
            return ''
        page = 'edit_damage' if edit else 'view_damage'
        return (f'https://bms2.iot.gov.tw/bms3/inspect/damage/{page}/der_id,{self.der_id}/inspect_id,{self.inspect_id}/'
                f'component_form,{self.row[0]}/component_id,{self.row[1]}/_popup,1/')


@dataclass
class Pair:
    kind: str
    prior: Defect | None = None
    current: Defect | None = None
    score: int = 0
    confidence: str = ''
    reasons: list = field(default_factory=list)
    manual: bool = False        # 人工配對

    @property
    def row(self):
        return (self.current or self.prior).row

    @property
    def moved(self):
        """前後期記在不同構件（位置變更）。"""
        return bool(self.prior and self.current and self.prior.component != self.current.component)

    @property
    def display(self):
        if self.moved:
            return f'{self.prior.display} → {self.current.display}'
        return (self.current or self.prior).display

    @property
    def code(self):
        if self.moved and self.prior.code != self.current.code:
            return f'{self.prior.code} → {self.current.code}'
        return (self.current or self.prior).code

    @property
    def key(self):
        """審查紀錄用的固定識別：本期以建議維修工法 ID，只有前期時以前期的 der_id（皆為網站產生的唯一值）。"""
        if self.current:
            return f'本期:{defect_id(self.current)}'
        return f'前期:{defect_id(self.prior)}'


def defect_id(defect):
    """一筆缺失的固定識別：本期用建議維修工法 ID，前期用 der_id（網站產生的唯一值）。"""
    if defect is None:
        return ''
    if defect.period == '本期':
        return defect.method_id or defect.der_id
    return defect.der_id or defect.method_id


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
                                item['fields'], bool(item.get('repaired')), item.get('method_id', ''),
                                item.get('der_id', ''), row['keys'].get('inspect_id', ''), item.get('options') or {},
                                {**(row.get('columns') or {}), 'U': item['fields'].get('u') or item.get('deru', '')[3:4]}))
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


UNIT = r'(?:cm|mm|m|公尺|公分|米)?'
NUMBER = r'\d+(?:\.\d+)?\s*' + UNIT
MEASURE = re.compile(r'[\s，,。、;；:：]*(?:約|深度?|高度?|寬度?|長度?|厚度?|[HDLWT]\s*[=＝:：])?\s*' + NUMBER +
                     r'(?:\s*[*xX×＊]\s*' + NUMBER + r')*\s*(?:立方公尺|平方公尺|m3|m2|m³|m²)?$', re.I)


def strip_measure(text):
    """去掉樣態結尾的量測值（2026-11-25：沒寫單位的體積「4*3*1.2」、只寫本期深度「1.5」、H=1.2m、深1.5m 等），
    避免被當成新的劣化樣態（自動新增規則）或降低前後期配對分數。"""
    text = str(text or '')
    while True:
        stripped = MEASURE.sub('', text).rstrip(' 。.，,、')
        if stripped == text or not stripped:
            return stripped or text
        text = stripped


def clean_pattern(fields):
    """v1.00 的樣態，再去掉樣態後直接接的量測值（例：「橋墩基礎裸露0.6M」→「橋墩基礎裸露」）並套用同義詞。"""
    text = pattern({**fields, 'd_location': location_code(fields.get('d_location', ''))})
    text = strip_measure(text)
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
    if pu != cu and cu.strip():                  # 本期單位空白＝沒填（填寫時沿用前期），不算單位變更
        return KIND_UNIT
    pa, ca = number(prior.fields.get('amount')), number(current.fields.get('amount'))
    if pa is not None and ca is not None and ca > pa:
        return KIND_LARGER
    if pa is not None and ca is not None and ca < pa:
        return KIND_SMALLER
    return KIND_KEEP


def match_rows(priors, currents):
    """同一列（同一構件）內的配對。本期備註未填的先依劣化類型配對（blank_pairs），其餘依分數配對。"""
    first = blank_pairs(priors, currents)
    if first:
        used = {id(x) for pair in first for x in (pair.prior, pair.current)}
        return first + match_rows([p for p in priors if id(p) not in used], [c for c in currents if id(c) not in used])
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
        target = merged_into(p, currents) if kind == KIND_MISSING else None
        if target is not None:
            kind = KIND_MERGED
            reasons = [f'與本期「{target.remark.strip()}」視為同一缺失（' + '／'.join(
                next(g for g in MERGE_GROUPS if any(w in clean_pattern(p.fields) for w in g))) + '），本期不另建']
        pairs.append(Pair(kind, prior=p, confidence='高' if kind == KIND_REPAIRED else '需確認', reasons=reasons))
    return pairs


def no_description(defect):
    """本期備註沒有敘述（空白，或只有尺寸／位置＋尺寸）。"""
    from .prior_merge import only_size
    remark = defect.remark.strip()
    return not remark or only_size(remark) is not None


def blank_pairs(priors, currents):
    """本期只填 DERU、劣化類型（數量）而備註未填（2026-11-19 實際案例）：與同一列、同劣化類型的前期配對。
    雙方都只有一筆 → 同一筆缺失（高）；各有多筆 → 依 DERU、數量一對一配對（需確認）。"""
    from .reconcile import deterioration_key
    pairs = []
    blanks = [c for c in currents if no_description(c)]
    for kind in dict.fromkeys(deterioration_key(c.deterioration) for c in blanks):
        mine = [c for c in blanks if deterioration_key(c.deterioration) == kind]
        theirs = [p for p in priors if deterioration_key(p.deterioration) == kind]
        if not theirs:
            continue
        if len(mine) == 1 and len(theirs) == 1:
            pairs.append(Pair(classify(theirs[0], mine[0]), theirs[0], mine[0], CONFIDENT, '高',
                              ['本期備註未填，同構件同劣化類型只有一筆，視為同一缺失']))
            continue
        left = list(theirs)
        for c in mine:                          # 多筆：DERU 相同優先，再看數量相同，最後依順序
            if not left:
                break
            best = max(left, key=lambda p: (p.deru == c.deru, number(p.fields.get('amount')) == number(c.fields.get('amount'))
                                            and number(c.fields.get('amount')) is not None, -left.index(p)))
            left.remove(best)
            pairs.append(Pair(classify(best, c), best, c, ACCEPT, '需確認',
                              [f'本期備註未填，同構件同劣化類型有 {len(theirs)} 筆前期，依 DERU／數量配對，請確認']))
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


def codes(defect):
    """缺失提到的構件位置代碼（P2、A1、S3）；「P2、P3掛淤」→ {P2, P3}。"""
    text = f"{defect.fields.get('d_location', '')} {defect.remark}"
    return {m[0].upper() + str(int(m[1])) for m in re.findall(r'(?<![A-Za-z])([PASpas])0*(\d+)', text)}


def related(a, b, text_a='', text_b=''):
    """兩個構件是否屬於 RELATED 中的同一組，且樣態符合該組關鍵字。"""
    for left, right, words in RELATED:
        if (a in left and b in right) or (a in right and b in left):
            if not words or any(w in text_a and w in text_b for w in words):
                return True
    return False


def cross_hints(pairs):
    """配對後剩下的「只有前期」與「只有本期」，若在相關構件、樣態相同、位置代碼吻合：加註可能對應（位置變更）。"""
    priors = [p for p in pairs if p.prior and not p.current]
    currents = [p for p in pairs if p.current and not p.prior]
    for left in priors:
        for right in currents:
            p, c = left.prior, right.current
            if p.component == c.component:
                continue
            pa, pc = clean_pattern(p.fields), clean_pattern(c.fields)
            if not related(p.component, c.component, pa + p.deterioration, pc + c.deterioration):
                continue
            similarity = pattern_similarity(pa, pc)
            if similarity is None or similarity < CROSS_SIMILARITY:
                continue
            cp, cc = codes(p), codes(c)
            if cp and cc and not cp & cc:
                continue
            left.reasons = left.reasons + [f'可能對應本期：{c.component} {c.remark.strip()}（位置變更）']
            right.reasons = right.reasons + [f'可能對應前期：{p.component} {p.remark.strip()}（位置變更）']
    return pairs


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
    return MatchResult(cross_hints(pairs))


def match_scans(prior_scan, current_scan):
    return match(defects_from_scan(prior_scan, '前期'), defects_from_scan(current_scan, '本期'))
