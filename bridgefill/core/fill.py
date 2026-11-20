"""第 6 步：依建議維修工法 ID 填寫網站——填寫清單與進度紀錄（不依賴 Qt）。

每筆本期缺失的「預定值」＝網站現值 → 套用 Excel 規則的程式建議（v1.00 automatic_plan，前期資料以配對結果為準）
→ 再套用使用者在逐筆審查的修改（優先）。與網站現值不同的欄位才填寫。

保護（沿用 v1.00 的防錯原則）：
- 以建議維修工法 ID 鎖定：開啟的表單 ID 必須與清單一致，否則停止。
- 網站現值與掃描時不同（有人改過）：停止，請重新掃描，不覆寫。
- 送出前先記錄「待確認」；送出後重新開啟回讀，一致才記為「已完成」。
- 上次送出結果不明（待確認）：重新開啟核對，網站值已是預定值就記為完成，否則交人工，**絕不自動重送**。
"""
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .data import History, eligible
from .prior_merge import merge_with_prior
from .review import KIND_PRIOR_WRONG, site_value
from .rule_add import NO_RULE

FILL_FIELDS = ('d', 'e', 'r', 'u', 'deterioration', 'd_location', 'method', 'amount', 'unit', 'price', 'remark')
TEXT_FIELDS = ('deterioration', 'd_location', 'method', 'amount', 'unit', 'price', 'remark')   # v1.00 填寫的 7 欄
NAMES = {'d': 'D', 'e': 'E', 'r': 'R', 'u': 'U', 'deterioration': '劣化類型', 'd_location': '損壞位置',
         'method': '維修工法', 'amount': '數量', 'unit': '單位', 'price': '單價', 'remark': '備註'}
READY, REVIEW, SAME, SKIP, DONE, PENDING, FAILED = '可填寫', '需確認', '無需修改', '不填寫', '已完成', '待確認', '失敗'
KEPT = '已確認'          # 需確認的項目經人工確認維持網站現值


def clean(value):
    return str(value if value is not None else '').strip()


@dataclass
class FillItem:
    pair: object
    site: dict                       # 掃描時的網站值
    target: dict                     # 預定值
    status: str
    reason: str = ''
    proposal: dict | None = None     # 程式建議（需確認時供人工帶入）
    rule_id: str = ''
    message: str = ''                # 填寫結果說明

    @property
    def key(self):
        return self.pair.key

    @property
    def method_id(self):
        return self.pair.current.method_id

    @property
    def url(self):
        return self.pair.current.damage_url(edit=True)

    @property
    def changes(self):
        return {k: (self.site.get(k, ''), self.target.get(k, '')) for k in FILL_FIELDS
                if clean(self.site.get(k)) != clean(self.target.get(k))}

    def describe(self):
        return '；'.join(f'{NAMES[k]}：{a or "（空）"} → {b or "（空）"}' for k, (a, b) in self.changes.items())


def site_fields(defect):
    return {k: site_value(defect, k) for k in FILL_FIELDS}


def proposal_for(pair, rules, prior_wrong=False):
    """v1.00 的自動判斷：本期現值＋配對到的前期＋Excel 規則 → Decision。
    跨構件配對、或審查判定「前期選擇或填寫錯誤」時不帶前期（只依本期與 Excel）。"""
    from .policy import automatic_plan
    current = pair.current
    base = {k: site_value(current, k) for k in TEXT_FIELDS}
    history, why = [], ''
    if pair.prior is not None and not pair.moved and not prior_wrong:
        prior = {k: clean(pair.prior.fields.get(k)) for k in TEXT_FIELDS}
        merged, why = merge_with_prior(base, prior)      # 尺寸換算一致、只留新尺寸等：先導入前期位置與敘述
        if merged is not None:
            base = merged
        else:
            history = [History(pair.prior.component, prior, 0, dict(pair.prior.context))]
    base = standardize(base, current.component, site_value(current, 'u'))      # Excel 樣態同義詞
    base = shape_remark(base, current.component, current.context, rules)        # 補單位、位置、尺寸範本
    decision = automatic_plan(base, current.component, history, rules, dict(current.context))
    if why and decision.review and NO_RULE in decision.reason:
        adopted = adopt_prior_method(decision.fields, pair.prior.fields)   # 規則 6：Excel 沒有此樣態，沿用前期工法
        if adopted:
            decision.fields, decision.review = adopted, False
            decision.reason = '前期劣化類型與樣態相同，Excel 沒有此樣態，沿用前期工法、單位、單價'
    if why:
        decision.reason = f'{why}；{decision.reason}'
    if decision.review and re.search(r'\(m(\*m)*\)', clean(decision.fields.get('remark'))):
        decision.reason = '已補上位置與尺寸範本，請在本期欄把 m 改成實際尺寸；' + decision.reason
    decision.fields = tidy_locations(decision.fields, current.component, current.context)
    return decision


def adopt_prior_method(fields, prior):
    """Excel 無對應值、但已導入前期位置與敘述（同類型、同樣態）：工法、單位、單價沿用前期。回傳新欄位或 None。"""
    from .editing import split_remark
    _, description, _ = split_remark(fields.get('remark'))
    _, old_description, _ = split_remark(prior.get('remark'))
    if not description or description.rstrip('。. ') != old_description.rstrip('。. '):
        return None
    result = dict(fields)
    for key in ('method', 'unit', 'price'):
        if not clean(result.get(key)) or key != 'unit':
            result[key] = clean(prior.get(key)) or clean(result.get(key))
    if not all(clean(result.get(k)) for k in ('method', 'unit', 'price', 'amount', 'remark')):
        return None
    location = clean(result.get('remark')).split(' ', 1)[0]
    result['d_location'] = clean(result.get('d_location')) or location
    return result


def standardize(fields, component, u=''):
    """Excel「樣態同義詞」：備註的同義寫法改成標準樣態（老化 → AC剝離、植物生長 U1 → 草本植物生長）。"""
    from .synonyms import standard_pattern
    result = dict(fields)
    if clean(result.get('remark')):
        result['remark'] = standard_pattern(result['remark'], component, u)
    return result


BARE_SIZE = re.compile(r'^(?P<head>.*?[^\d\s.(（])[\s。.]*[(（]?(?P<size>\d+(?:\.\d+)?(?:mm|cm|m)?'
                       r'(?:\s*[*＊×xX]\s*\d+(?:\.\d+)?(?:mm|cm|m)?){1,3})[)）]?\s*$')
SIZE_TEMPLATE = {'平方公尺': '(m*m)', '立方公尺': '(m*m*m)', '公尺': '(m)'}
BANK_LENGTH = 5          # Excel 參數「引道路堤預設長度」：引道路堤 AC 缺尺寸時以 最小淨寬 × 5m 估算


def shape_remark(fields, component, context=None, rules=()):
    """備註寫法補齊（2026-11-20 使用者需求）：
    1. 尺寸沒寫單位與括號：「5*7」→「(5m*7m)」（平方公尺）、「1*2*3」→「(1m*2m*3m)」（立方公尺）；
    2. 只寫樣態沒有位置：補上構件位置（前期位置已由 merge_with_prior 補上；沒有前期時用清單位置代碼，例 A1）；
    3. 有樣態沒有尺寸：依 Excel 單位補尺寸範本「(m*m)」供修改；引道路堤 AC 缺失依 最小淨寬×5m 自動估算。"""
    from .editing import split_remark
    from .locations import short_code
    from .prior_merge import COUNT_UNITS, normalize_size
    result = dict(fields)
    remark = clean(result.get('remark'))
    if not remark:
        return result
    unit = clean(result.get('unit'))
    bare = BARE_SIZE.match(remark)
    if bare and not re.search(r'[a-z]', bare['size']):
        parts = re.split(r'\s*[*＊×xX]\s*', bare['size'])
        size_unit = unit if unit in SIZE_TEMPLATE else {2: '平方公尺', 3: '立方公尺'}.get(len(parts), '')
        size = normalize_size('*'.join(parts), size_unit) if size_unit and len(parts) > 1 else ''
        if size:
            remark = f"{bare['head'].rstrip('。. ')}。{size}"
            if not unit:
                result['unit'] = unit = size_unit
    location, description, tail = split_remark(remark)
    description = description.rstrip('。. ')
    if not description or re.fullmatch(r'[\d.*()（）mc\s]+', description):
        return dict(result, remark=remark)
    if not location:
        location = clean(result.get('d_location')) or short_code(component, context or {})
        if location and not re.search(r'[\u4e00-\u9fff]{2,}', location):
            remark = f'{location} {description}{tail}'
    if '(' not in remark and '（' not in remark:
        size = ''
        candidates = rule_units(rules, dict(result, remark=remark), component)
        size_unit = unit if unit in SIZE_TEMPLATE else (candidates[0] if len(candidates) == 1 else '')
        width = clean((context or {}).get('_最小淨寬'))
        if (component == '引道路堤' and size_unit == '平方公尺' and description.upper().startswith('AC')
                and re.fullmatch(r'\d+(\.\d+)?', width) and float(width) > 0):
            size = f'({width}m*{BANK_LENGTH}m)'
            result['unit'] = size_unit
        elif size_unit in SIZE_TEMPLATE:
            size = SIZE_TEMPLATE[size_unit]
        elif unit in COUNT_UNITS and unit != '式' and clean(result.get('amount')):
            size = f"({clean(result['amount'])}{unit})"
        if size:
            remark = f"{remark.rstrip('。. ')}。{size}"
    result['remark'] = remark
    return result


def rule_units(rules, fields, component):
    """備註樣態可用的 Excel 規則單位（不重複）。"""
    if not rules:
        return []
    from .data import repair_candidates
    from .reconcile import type_repairs
    try:
        found = repair_candidates(type_repairs(rules, clean(fields.get('deterioration'))), fields, component,
                                  filter_unit=False)
    except Exception:  # 規則格式異常時不補範本
        return []
    return list(dict.fromkeys(clean(r.get('unit')) for r in found))


def tidy_locations(fields, component, context=None):
    """主梁、橫隔梁的位置寫法（S2G5 → 2G5、1D-2-1 → 1D2-1），損壞位置與備註開頭一起修正。"""
    from .locations import normalize_location
    from .synonyms import standard_location
    result = standard_location(dict(fields), component)          # Excel 位置代碼對照（整體-ALL → 整體-L,R）
    location = clean(result.get('d_location'))
    if location:
        result['d_location'] = normalize_location(location, component, context)
    remark = clean(result.get('remark'))
    if remark and ' ' in remark:
        from .data import normalize_remark
        head, rest = remark.split(' ', 1)
        head = re.sub(r'(?<=[0-9A-Z])\.(?=\d*[LR]\b)', ',', normalize_location(head, component, context))   # A1-0R.0L → A1-0R,0L
        rest = rest.replace(' ', '')                            # 敘述中間不留空白（支撐座 混凝土裂縫）
        tidy = normalize_remark(f'{head} {rest}') if re.search(r'[\(（][^()（）]*[\)）]', rest) else f'{head} {rest}'
        result['remark'] = tidy
        if clean(result.get('d_location')) and clean(result.get('d_location')) == remark.split(' ', 1)[0]:
            result['d_location'] = tidy.split(' ', 1)[0]
    return result


def plan_item(pair, review, rules):
    current = pair.current
    site = site_fields(current)
    edits = {k: clean(v) for k, v in review.edit_of(pair).items() if k in FILL_FIELDS}
    decision, problem = None, ''
    if rules:
        try:
            decision = proposal_for(pair, rules, review.decision(pair).kind == KIND_PRIOR_WRONG)
        except Exception as error:   # 規則判斷失敗只影響這一筆
            problem = f'程式建議失敗：{error}'
    else:
        problem = '沒有讀到 Excel 規則表，只能填寫人工修改'
    proposal = {k: clean(decision.fields.get(k, site[k])) for k in TEXT_FIELDS} if decision else None
    automatic = decision is not None and not decision.review
    target = dict(site)
    if automatic:
        target.update(proposal)
    target.update(edits)
    item = FillItem(pair, site, target, READY, decision.reason if decision else problem, proposal,
                    getattr(decision, 'rule_id', '') if decision else '')
    if eligible(target['d'], target['e']) is False and not edits:
        item.status, item.reason = SKIP, 'D＝1 或 E＝0（良好或無此項），不填寫'
        item.target = dict(site)
        return item
    if not automatic and not edits:
        item.target = dict(site)
        original = decision.reason if decision else problem
        if pair.key in review.kept:
            item.status = KEPT
            item.reason = f"已人工確認維持網站現值（{review.kept[pair.key].replace('T', ' ')}）；程式原判：{original}"
        else:
            item.status = REVIEW
            item.reason = original + '；請在本期欄修改（或還原網站值）後按「確認並下一筆」'
        return item
    for name in ('deterioration', 'unit'):
        choices = current.options.get(name) or []
        if choices and clean(target[name]) and clean(target[name]) not in choices:
            item.status, item.reason = REVIEW, f'網站「{NAMES[name]}」下拉選單沒有「{target[name]}」，請人工修改'
            item.target = dict(site)
            return item
    if not item.changes:
        item.status, item.reason = SAME, '網站值已符合規則與修改，不需填寫'
    elif edits and not automatic:
        item.reason = '只填寫人工修改（程式建議需確認：' + (decision.reason if decision else problem) + '）'
    elif edits:
        item.reason = '程式建議＋人工修改（人工優先）：' + item.reason
    return item


def build_plan(pairs, review, rules, log=None):
    """本期有的缺失才填寫（只有前期的漏建、已維修不在網站上）。"""
    items = [plan_item(pair, review, rules) for pair in pairs if pair.current is not None]
    if log is not None:
        for item in items:
            entry = log.get(item.key)
            if entry.get('status') == DONE and entry.get('target') == item.target:
                item.status, item.message = DONE, f"已於 {entry.get('at', '')} 填寫並核對"
            elif entry.get('status') == PENDING:
                item.status, item.message = PENDING, '上次送出結果不明，填寫時會先重新開啟核對，不會重送'
            elif entry.get('status') == FAILED and item.status == READY:
                item.message = f"上次失敗：{entry.get('message', '')}"
    return items


class FillLog:
    """填寫進度檔（每筆以 Pair.key 記錄狀態、預定值、時間）。先寫暫存檔再取代，避免中斷時損毀。"""

    def __init__(self, path):
        self.path = Path(path)
        self.entries = {}
        if self.path.exists():
            self.entries = json.loads(self.path.read_text(encoding='utf-8')).get('entries', {})

    def get(self, key):
        return self.entries.get(key, {})

    def mark(self, key, status, target=None, message=''):
        entry = dict(self.entries.get(key, {}))
        entry.update(status=status, message=message, at=datetime.now().isoformat(timespec='seconds'))
        if target is not None:
            entry['target'] = dict(target)
        self.entries[key] = entry
        self.save()
        return entry

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix('.tmp')
        temporary.write_text(json.dumps({'version': 1, 'entries': self.entries}, ensure_ascii=False, indent=2),
                             encoding='utf-8')
        temporary.replace(self.path)


def fill_log_path(folder, bridge, current_date):
    safe = ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in f'{current_date}-{bridge}')
    return Path(folder) / f'{safe}-填寫.json'


def compare(actual, expected, keys=FILL_FIELDS):
    """回傳不一致的欄位名稱。"""
    return [k for k in keys if clean(actual.get(k)) != clean(expected.get(k))]
