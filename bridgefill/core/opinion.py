"""第 7 步：檢測員意見草稿（不依賴 Qt）。

依《規格》第 7 節與竹子坑橋歸納的修正規則：
- U 決定建議：U4 立即處置、U3 一年內派工修復；U2 看是否影響行車安全（橋面）或主構件且 E≧3 → 三年內派工修復，否則持續追蹤。
- 位置只寫到主構件（A1-下游 → A1，1G1-0,13,26 → S1）；構件名稱用實際名稱；樣態已含名稱不重複；「整體-」刪除。
- 相同樣態、相同建議的條目合併（A1、A2橋台、P2橋墩混凝土剝落鋼筋外露）；連號寫成 S1~S3。
- 不重要的樣態不寫入（伸縮縫漏水、基礎混凝土剝落、蜂窩、AC剝離）；小面積的混凝土剝落、沖蝕、小破裂（U2、E≦2）預設不勾選。
- 標準文字：護欄高度不足、河道護岸與基礎（**護岸與橋台／橋墩基礎同時掏空、裸露時合併成一段，建議只寫一次**）、伸縮縫間距、橋下雜物。
  已被標準文字敘明的缺失不再逐條重複。
- 【A／B】為待選項目，由使用者在畫面上選定後才寫入。
- 淨高：本期有新值則保留，沒有沿用前期意見中的值。
"""
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

from .editing import split_remark

SUFFIX = {2: '建議持續追蹤', 3: '建議一年內派工修復', 4: '建議通知相關單位立即處置'}
THREE_YEARS = '建議三年內派工修復'
# 構件分類（Excel「構件分類」草稿）：主構件、可能影響行車安全
MAIN = {'主梁', '墩柱與帽梁', '橋台', '橋台基礎', '橋墩基礎', '橋面板', '端牆', '中隔板', '頂板', '底板',
        '鋼纜錨碇裝置(橋塔端)', '鋼纜錨碇裝置(主梁端)', '鋼纜(吊索)-斜張鋼纜', '拱肋', '立柱'}
TRAFFIC = {'橋面板', '伸縮縫', '橋護欄', '引道路堤', '引道路堤護欄', '其他', '頂板'}
NAMES = {'墩柱與帽梁': '橋墩', '橋護欄': '護欄', '引道路堤護欄': '引道護欄', '橋梁排水設施': '排水設施',
         '引道道路堤保護措施': '引道保護措施', '橋墩／橋基保護設施': '橋墩保護設施', '其他': ''}
# (構件關鍵字, 樣態關鍵字)：不寫入意見；構件空白代表全部
EXCLUDED = (('伸縮縫', '漏水'), ('基礎', '混凝土剝落'), ('', '蜂窩'), ('', 'AC剝離'))
MINOR = ('混凝土剝落', '沖蝕', '小破裂')

RIVER = ('河道', '護床工')
RIVER_WORDS = ('傾斜', '掏空', '倒塌')
FOUNDATION = ('橋墩基礎', '橋台基礎', '橋墩／橋基保護設施', '墩柱與帽梁', '橋台')
FOUNDATION_WORDS = ('傾斜', '倒塌', '裸露', '掏空')
JOINT_WORDS = ('間隙異常', '間距過大', '密合', '高低差', '高低落差')
WASTE_WORDS = ('雜物', '垃圾')
WATER_TAIL = ('，觀察主河道流心偏向【橋台AO／橋墩PO】（【有／無】沖刷集中跡象），研判【可能影響／暫不影響】橋梁通洪及基礎安全，'
              '建議【列入颱洪期間應注意橋梁／一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。')
# v1.00 的 6 個標準文字（未偵測到時原樣列出、不勾選，需要時自行勾選）
TEMPLATES = (
    '本橋道路等級屬【道路等級】，且現況【有／無】人行道，橋護欄高度不足110cm，建議採【加高／重作】方式改善。',
    '本橋【護岸／護床工】發生【傾斜／掏空／倒塌】情形，位於【上游／下游】【左岸／右岸】約【＿m】處，並觀察流心偏向【橋台AO／橋墩PO】'
    '（【有／無】沖刷集中跡象），研判【可能影響／暫不影響】橋梁通洪及基礎安全，'
    '建議【列入颱洪期間應注意橋梁／一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。',
    '本橋【Px橋墩】發生【傾斜／倒塌／裸露】情形，本期裸露深度為【本期深度】，前期裸露深度為【前期深度】，【Px】橋墩有【刷深／回淤】之情況，'
    '觀察主河道流心偏向【橋墩PO】，研判【可能影響／暫不影響】橋墩安全，'
    '建議【列入颱洪期間應注意橋梁／一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。',
    '本橋【AO／PO】伸縮縫【間距過大／密合／具高低落差】，已造成【周邊劣化情形描述】，建議【持續觀測／一年內維修／通知相關單位立即處置】。',
    '本橋橋下發現民眾堆置雜物及垃圾，具火災風險，建議通知權責單位辦理清理。',
    '本橋僅供特定人員使用，其權責及通行條件尚待釐清，建議管理機關協調確認，以利後續檢測、維修及緊急應變作業。',
)
STANDARD_NAMES = ('護欄高度不足', '護岸／護床工', '橋墩基礎', '伸縮縫', '橋下雜物', '特定人員使用')
FOUNDATION_TAIL = ('，觀察主河道流心偏向【橋墩PO】，研判【可能影響／暫不影響】橋墩安全，'
                   '建議【列入颱洪期間應注意橋梁／一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】。')
CHOICE = re.compile(r'【([^】]*／[^】]*)】')
BLANK = re.compile(r'【[^】／]*】')
STYLE_DETAILED, STYLE_COMPACT = 'detailed', 'compact'
HIGH_E_OPTIONS = '持續追蹤／三年內派工修復／一年內派工修復'


@dataclass
class OpinionItem:
    text: str
    kind: str                      # standard／defect／manual
    checked: bool = True
    sources: list = field(default_factory=list)     # Pair.key
    note: str = ''
    custom: str = ''               # 使用者在意見框直接改過的這一條（空白：用 text 套用選項）


@dataclass
class Entry:
    """一筆要寫入意見的本期缺失（已套用逐筆審查的修改）。"""
    key: str
    component: str
    fields: dict
    d: int
    e: int
    u: int
    prior_fields: dict = field(default_factory=dict)
    context: dict = field(default_factory=dict)      # 構件列資料（「其他」的設施類別等）

    @property
    def remark(self):
        return self.fields.get('remark', '')


def number(value):
    try:
        return int(str(value).strip())
    except ValueError:
        return 0


def entries_from(pairs, review, planned=None):
    """意見依據的本期值：planned(pair) 有值時用它（工作區本期欄「將填寫的值」：程式依規則補齊＋人工修改），
    否則用網站現值＋人工修改。不必先送出、重新掃描，核對完就能產生正確的意見。"""
    result = []
    keys = ('d', 'e', 'r', 'u', 'deterioration', 'd_location', 'remark', 'amount', 'unit')
    for pair in pairs:
        if pair.current is None:
            continue
        values = {k: review.value(pair, k) for k in keys}
        filled = planned(pair) if planned else None
        if filled:
            values.update({k: str(filled[k]).strip() for k in keys if str(filled.get(k, '') or '').strip()})
        result.append(Entry(pair.key, pair.current.component, values, number(values['d']), number(values['e']),
                            number(values['u']), dict(pair.prior.fields) if pair.prior else {},
                            dict(pair.current.context)))
    return result


# ---------------------------------------------------------------- 位置、名稱、樣態

def locations(entry):
    """主構件位置代碼（可能多個）：A1-下游 → A1，1G1-0,13,26 → S1，P2、P3 → P2、P3，整體 → 無。"""
    text = (entry.fields.get('d_location') or split_remark(entry.remark)[0] or '').strip()
    if not text or text.startswith('整體') or re.search(r'(?i)-all', text):
        return []
    found = []
    for match in re.finditer(r'(?i)(?<![A-Z0-9])(?:S?(\d+)[GD]\d+|([SPA])0*(\d+))', text):
        code = f'S{int(match[1])}' if match[1] else f'{match[2].upper()}{int(match[3])}'
        if code not in found:
            found.append(code)
    return found


def display_name(component, pattern, codes=(), context=None):
    from .excel_config import opinion_names
    name = opinion_names().get(component, component)
    if component == '其他':
        name = facility(context)            # 「其他」用設施類別（例：鋪面），不寫「其他」
    if '／' in name:
        parts = name.split('／')
        name = next((p for p in parts if p in pattern), parts[0])
    if component == '墩柱與帽梁' and codes and all(c.startswith('A') for c in codes):
        name = '橋台'
    if name and (name in pattern or any(part in pattern for part in re.split(r'與|或', name) if len(part) >= 2)):
        return ''              # 樣態已含構件名稱，不重複
    return name


def facility(context):
    """「其他」構件的設施類別（網站欄名「設施類別」）。"""
    return next((str(v).strip() for k, v in (context or {}).items() if '設施類別' in k and str(v).strip()), '')


def pattern_of(entry):
    _, pattern, _ = split_remark(entry.remark)
    pattern = re.sub(r'(?:約)?[\d.]+\s*(?:CM|MM|M|cm|mm|m)?$', '', pattern).strip(' 。.，,')
    pattern = re.sub(r'(?:滲水|漏水)(?:及|且|與|和|、)?白華|白華(?:及|且|與|和|、)?(?:滲水|漏水)', '滲水白華', pattern)
    return pattern or entry.fields.get('deterioration', '') or '【劣化樣態待填】'


def suggestion(entry):
    from .excel_config import component_classes, opinion_advice, param_value
    advice = opinion_advice()
    if entry.u >= 4:
        return advice['u4']
    if entry.u == 3:
        return advice['u3']
    classes = component_classes()
    large_e = int(param_value('大面積判定 E 值', 3))
    if entry.component in classes['traffic']:
        return advice['u2_repair']
    if entry.component in classes['main'] and entry.e >= large_e:
        return advice['u2_repair']
    return advice['u2_track']


def excluded(entry, pattern):
    from .excel_config import opinion_excluded
    u_text = str(entry.u)
    return any((not c or c in entry.component) and w in pattern and (not only_u or u_text == only_u)
               for c, w, only_u in opinion_excluded())


def minor(entry, pattern):
    return entry.u == 2 and entry.e <= 2 and any(w in pattern for w in MINOR) and '鋼筋' not in pattern


def compress(codes):
    """S1、S2、S3 → S1~S3；其他照原順序。"""
    result, run = [], []

    def flush():
        if len(run) >= 3:
            result.append(f'{run[0]}~{run[-1]}')
        else:
            result.extend(run)
        run.clear()
    for code in codes:
        if run and code[0] == run[-1][0] and int(code[1:]) == int(run[-1][1:]) + 1:
            run.append(code)
        else:
            flush()
            run.append(code)
    flush()
    return '、'.join(result)


# ---------------------------------------------------------------- 標準文字

def has(entry, components, words):
    text = entry.remark + entry.fields.get('deterioration', '')
    return any(c in entry.component for c in components) and any(w in text for w in words)


def exposure_depth(fields):
    text = str(fields.get('remark', '')).translate(str.maketrans({'（': '(', '）': ')', '＊': '*', '×': '*'}))
    explicit = re.search(r'(?:裸露|曝露)深度(?:為|約|[:：=])?\s*(\d+(?:\.\d+)?)\s*m', text, re.I)
    if explicit:
        return explicit[1]
    sizes = re.findall(r'\(\s*\d+(?:\.\d+)?m\s*\*\s*\d+(?:\.\d+)?m\s*\*\s*(\d+(?:\.\d+)?)m\s*\)', text, re.I)
    return sizes[0] if len(sizes) == 1 else ''


def water_paragraph(river, foundations):
    """護岸／護床工與基礎合併成一段，研判與建議只寫一次。"""
    parts = []
    if river:
        text = ' '.join(e.remark for e in river)
        facilities = [w for w in ('護岸', '護床工') if w in text or any(w in e.component for e in river)] or ['護岸']
        states = [w for w in RIVER_WORDS if w in text] or ['傾斜／掏空／倒塌']
        side = '、'.join(w for w in ('上游', '下游') if w in text) or '上游／下游'
        bank = '、'.join(w for w in ('左岸', '右岸') if w in text) or '左岸／右岸'
        parts.append(f"{'、'.join(facilities)}發生{'、'.join(states) if len(states) > 1 or '／' not in states[0] else '【' + states[0] + '】'}"
                     f"情形，位於【{side}】【{bank}】約【＿m】處".replace('【上游】', '上游').replace('【下游】', '下游')
                     .replace('【左岸】', '左岸').replace('【右岸】', '右岸'))
    if foundations:
        codes = []
        for e in foundations:
            codes += [c for c in locations(e) if c[0] in 'AP' and c not in codes]
        names = '、'.join(codes) or '【Px】'
        text = ' '.join(e.remark for e in foundations)
        states = '、'.join(w for w in FOUNDATION_WORDS if w in text)
        now = {c: exposure_depth(e.fields) for e in foundations for c in locations(e) if exposure_depth(e.fields)}
        old = {c: exposure_depth(e.prior_fields) for e in foundations for c in locations(e) if e.prior_fields
               and exposure_depth(e.prior_fields)}
        depth = lambda values, label: '、'.join(f'{c}={values[c]}m' if c in values else f'【{c}{label}深度】'  # noqa: E731
                                              for c in codes) or f'【{label}深度】'
        sentence = f'{names}基礎發生{states}情形，本期裸露深度為{depth(now, "本期")}，前期裸露深度為{depth(old, "前期")}'
        from .excel_config import param_decimal
        tolerance = param_decimal('基礎裸露量測誤差', '0.1')
        changes = []
        for c in codes:
            if c in now and c in old:
                try:
                    delta = Decimal(now[c]) - Decimal(old[c])
                except InvalidOperation:
                    continue
                changes.append(f'{c}有刷深之情況' if delta > tolerance else f'{c}有回淤之情況'
                               if delta < -tolerance else f'{c}裸露深度差異在±{tolerance:g}m量測誤差範圍內，視為與前期一致')
        parts.append(sentence + ('，' + '、'.join(changes) if changes else '，【Px】有【刷深／回淤】之情況'))
    return '本橋' + '；'.join(parts) + (WATER_TAIL if river else FOUNDATION_TAIL)


def standard_items(entries, header, everything=()):
    """6 個標準文字一律列出：符合條件的自動帶入並勾選；不符合的原樣列出、不勾選。
    護岸與基礎同時符合時合併成一段（研判、建議只寫一次），第 3 項不再另列。"""
    def item(index, text, found, merged=''):
        note = STANDARD_NAMES[index] + ('：' + merged if merged else '：符合條件，已自動帶入' if found
                                        else '：未偵測到，需要時勾選')
        return OpinionItem(text, 'standard', bool(found), [e.key for e in found], note)
    items = []
    guard = [e for e in entries if '護欄' in e.component and '高度不足' in e.remark]
    grade = header.get('道路等級')
    items.append(item(0, TEMPLATES[0].replace('【道路等級】', grade) if grade else TEMPLATES[0], guard))
    river = [e for e in entries if has(e, RIVER, RIVER_WORDS) or ('護岸' in e.remark and has(e, ('',), RIVER_WORDS))]
    foundations = [e for e in entries if e not in river and has(e, FOUNDATION, FOUNDATION_WORDS)
                   and ('基礎' in e.component or '基礎' in e.remark or '保護設施' in e.component)]
    if river and foundations:
        items.append(item(1, water_paragraph(river, foundations), river + foundations, '護岸與基礎都符合，合併成一段'))
    else:
        items.append(item(1, water_paragraph(river, []) if river else TEMPLATES[1], river))
        items.append(item(2, water_paragraph([], foundations) if foundations else TEMPLATES[2], foundations))
    joints = [e for e in entries if has(e, ('伸縮縫',), JOINT_WORDS)]
    text = TEMPLATES[3]
    # 伸縮縫位置自動判讀：有間距、密合、高低差的那幾道；沒有時列出本期有缺失的伸縮縫；再沒有就依橋孔數給選項
    codes = []
    for e in joints or [e for e in everything if '伸縮縫' in e.component]:
        codes += [c for c in locations(e) if c not in codes and c[0] in 'AP']
    if not codes:
        spans = number(header.get('總橋孔數'))
        options = ['A1', 'A2'] + [f'P{i}' for i in range(1, spans)] if spans else []
        if options:
            text = text.replace('【AO／PO】', '【' + '／'.join(options) + '】')
    if joints:
        remarks = ' '.join(e.remark for e in joints)
        states = [label for words, label in ((('間隙異常', '間距過大'), '間距過大'), (('密合',), '密合'),
                                             (('高低差', '高低落差'), '具高低落差')) if any(w in remarks for w in words)]
        if states:
            text = text.replace('【間距過大／密合／具高低落差】', '、'.join(states))
    if codes:
        text = text.replace('【AO／PO】', '、'.join(codes))
    items.append(item(3, text, joints))
    waste = [e for e in entries if any(w in e.remark for w in WASTE_WORDS)]
    if not waste and str(header.get('橋下有無堆放雜物', '')).strip() in ('是', '有'):
        waste = [Entry('header', '', {'remark': '橋下有堆放雜物'}, 0, 0, 0)]
    items.append(item(4, TEMPLATES[4], waste))
    items.append(item(5, TEMPLATES[5], []))
    return items


# ---------------------------------------------------------------- 產生

def generate(pairs, review, header=None, prior_header=None, planned=None, style=STYLE_DETAILED):
    """產生意見條目。
    style='detailed'（預設）：依構件位置列出並壓縮連號（S1~S3主梁…）。
    style='compact'（精簡版）：U=2 且 E<3 不列位置、跨構件合併同樣態（橫隔梁、橋面板、主梁混凝土剝落鋼筋外露，建議持續追蹤。）；
    U=2 且 E≧3（不含伸縮縫漏水等排除項）獨立成條，供人工判斷是否需維修。"""
    header, prior_header = header or {}, prior_header or {}
    everything = entries_from(pairs, review, planned)
    entries = [e for e in everything if e.u >= 2]
    items = standard_items(entries, header, everything)
    covered = {key for item in items if item.checked for key in item.sources}
    from .excel_config import opinion_advice, param_value
    compact = style == STYLE_COMPACT
    large_e = int(param_value('大面積判定 E 值', 3))
    track_advice = opinion_advice()['u2_track']
    groups, high_e = {}, {}
    for entry in entries:
        if entry.key in covered:
            continue
        pattern = pattern_of(entry)
        if excluded(entry, pattern):
            continue
        codes = locations(entry)
        name = display_name(entry.component, pattern, codes, entry.context)
        if compact and entry.u == 2 and entry.e >= large_e:
            bucket = high_e.setdefault((name, pattern), {'codes': [], 'keys': [], 'e': 0})
            bucket['codes'] += [c for c in codes if c not in bucket['codes']]
            bucket['keys'].append(entry.key)
            bucket['e'] = max(bucket['e'], entry.e)
            continue
        advice = track_advice if (compact and entry.u == 2) else suggestion(entry)
        group = groups.setdefault((pattern, advice), {'parts': {}, 'keys': [], 'minor': True, 'u': 0})
        part = group['parts'].setdefault(name, [])
        if not (compact and entry.u == 2):
            part += [c for c in codes if c not in part]
        group['keys'].append(entry.key)
        group['minor'] &= minor(entry, pattern)
        group['u'] = max(group['u'], entry.u)

    urgent = [(k, g) for k, g in groups.items() if g['u'] >= 3]
    normal = [(k, g) for k, g in groups.items() if g['u'] < 3]

    def format_subject(parts):
        pieces = []
        for name, codes in parts.items():
            label = compress(codes) + name
            if label and label not in pieces:
                pieces.append(label)
        return '、'.join(pieces)

    for (pattern, advice), group in sorted(urgent, key=lambda kv: -kv[1]['u']):
        subject = format_subject(group['parts'])
        items.append(OpinionItem(f'{subject}{pattern}，{advice}。', 'defect', not group['minor'], group['keys'],
                                 '小面積、U2：預設不寫入' if group['minor'] else ''))
    for (name, pattern), group in high_e.items():
        title = f'{name}{pattern}'
        where = compress(group['codes'])
        note = f'E≧3（最大 E={group["e"]}，位置：{where or "—"}），請人工判斷是否需維修'
        items.append(OpinionItem(f'{title}，建議【{title}｜{HIGH_E_OPTIONS}】。', 'review_e', True, group['keys'], note))
    for (pattern, advice), group in normal:
        subject = format_subject(group['parts'])
        items.append(OpinionItem(f'{subject}{pattern}，{advice}。', 'defect', not group['minor'], group['keys'],
                                 '小面積、U2：預設不寫入' if group['minor'] else ''))
    return items


def heights(header, prior_header, has_girder, spans=None):
    """淨高：本期有新值則保留，沒有沿用前期。"""
    def parse(text):
        return {m[1].upper(): m[2] for m in re.finditer(r'(?i)\b(P\d+|H)\s*[=＝]\s*(\d+(?:\.\d+)?)', text or '')}
    values = {**parse(prior_header.get('檢測員意見')), **parse(header.get('檢測員意見'))}
    count = spans or number(header.get('總橋孔數'))
    keys = ['H'] if count == 1 else [f'P{i}' for i in range(1, count)] if count > 1 else sorted(values)
    title = '梁下淨高' if has_girder else '橋下淨高'
    return f"{title}：" + '、'.join(f'{k}={values[k]}m' if k in values else f'{k}=【m】' for k in keys)


def height_text(opinion):
    """意見中的淨高那一行（例：梁下淨高：P1=3.35m、P2=5.15m）；沒有回傳空白。"""
    lines, after = [], False
    for line in (line.strip() for line in str(opinion or '').splitlines()):
        value = bool(re.match(r'(?i)^[PH]\s*\d*\s*[=＝:：]', line))     # 「淨高：」下一行起的 P1=3.3m、H=2.2m
        if '淨高' in line or re.search(r'(?i)\bH\s*[=＝]', line) or (after and value):
            lines.append(line)
            after = True
        elif line:
            after = False
    return '\n'.join(lines)                           # 分行照原意見（範本：橋下淨高：換行 P1=3.2m…）


def choices(text):
    """待選項目：【A／B／C】→ ['A', 'B', 'C']；支援標籤前綴【標籤｜A／B／C】→ ['A', 'B', 'C']。"""
    result = {}
    for m in dict.fromkeys(CHOICE.findall(text)):
        options = m.split('｜', 1)[1] if '｜' in m else m
        result[f'【{m}】'] = options.split('／')
    return result


def choice_label(placeholder):
    """待選項目顯示名稱：【主梁混凝土裂縫｜持續追蹤／…】→ 「主梁混凝土裂縫」；一般則原樣回傳。"""
    inner = str(placeholder or '').strip('【】')
    return inner.split('｜', 1)[0] if '｜' in inner else placeholder


# 舊版存檔的建議選項（4.5.000.0 前少「列入颱洪期間應注意橋梁」，4.5.000.0 順序不同）：載入草稿時換成目前的選項
ADVICE = '【列入颱洪期間應注意橋梁／一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】'
OLD_ADVICE = ('【一年內維修並於維修前列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】',
              '【一年內維修並於維修前列入颱洪期間應注意橋梁／列入颱洪期間應注意橋梁／辦理詳細檢測／定期觀測河床變化】')


def upgrade_text(text):
    for old in OLD_ADVICE:
        text = text.replace(old, ADVICE)
    return text


def upgrade_saved(items, values):
    """舊草稿的意見文字與已選值改用目前的建議選項（items：dict 清單；values：{待選項目: 值}）。"""
    for item in items:
        for key in ('text', 'custom'):
            if item.get(key):
                item[key] = upgrade_text(item[key])
    return [dict(item) for item in items], {upgrade_text(k): v for k, v in (values or {}).items()}


def apply_choices(text, chosen):
    """chosen：{'【A／B】': 'A', '【＿m】': '15m'}；選定或填寫後括號消失。"""
    for placeholder, value in (chosen or {}).items():
        if value:
            text = text.replace(placeholder, value)
    return text


def compose(items, height_line='', chosen=None):
    lines = [item.custom or apply_choices(item.text, chosen) for item in items if item.checked]
    text = '\n'.join(f'{i}. {line}' for i, line in enumerate(lines, 1))
    return (text + ('\n' + height_line if height_line else '')).strip()


def blanks(text):
    """待填項目：【＿m】、【周邊劣化情形描述】等（沒有「／」的【…】）。"""
    return list(dict.fromkeys(BLANK.findall(text)))


NUMBERED = re.compile(r'^\s*\d+\s*[.．、]\s*')


def absorb(items, text, chosen=None):
    """使用者直接修改意見框後：依內容把每一行對回勾選中的條目（順序不變，取最相近的；改過的記為 custom），
    刪掉的行對應的條目取消勾選，多出的行成為手動條目；沒有編號的行（淨高）回傳。勾選增減時就能保留手動修改並同步。"""
    from difflib import SequenceMatcher
    lines = [line for line in text.splitlines() if line.strip()]
    numbered = [NUMBERED.sub('', line).strip() for line in lines if NUMBERED.match(line)]
    numbered = [line for line in numbered if line]          # 剛打好編號、還沒有內容的行：先不處理
    rest = '\n'.join(line for line in lines if not NUMBERED.match(line))
    checked = [item for item in items if item.checked]
    shown = [item.custom or apply_choices(item.text, chosen) for item in checked]
    pointer, matched, added = 0, set(), []
    for position, line in enumerate(numbered):
        best, score = None, 0.0
        for index in range(pointer, len(checked)):
            ratio = 1.0 if line == shown[index] else SequenceMatcher(None, line, shown[index]).ratio()
            if ratio > score:
                best, score = index, ratio
            if ratio == 1.0:
                break
        if (best is None or score < 0.5) and pointer < len(checked) and \
                len(numbered) - position >= len(checked) - pointer:
            best, score = pointer, 1.0                     # 行數夠對應剩下的條目：視為這一條整句改寫
        if best is not None and score >= 0.5:
            item = checked[best]
            item.custom = '' if line in (item.text, apply_choices(item.text, chosen)) else line
            matched.add(best)
            pointer = best + 1
        else:
            added.append(line)
    for index, item in enumerate(checked):
        if index not in matched:
            item.checked, item.custom = False, ''      # 意見框中刪掉的條目：取消勾選
    for line in added:
        items.append(OpinionItem(line, 'manual', True, note='手動新增'))
    return rest


def numbering_ok(text):
    """意見框的編號是否為 1、2、3…連號。"""
    numbers = [int(re.match(r'^\s*(\d+)', line)[1]) for line in text.splitlines() if NUMBERED.match(line)]
    return numbers == list(range(1, len(numbers) + 1))


def unresolved(text):
    """還沒選定或填寫的【…】。"""
    return list(dict.fromkeys(CHOICE.findall(text) + [m.strip('【】') for m in BLANK.findall(text)]))
