"""人工改配對：使用者指定前期與本期的對應（可一對多、可跨構件），或拆開程式的配對。

- links：前期缺失 ID → 本期缺失 ID 清單（一筆前期可對多筆本期，例：前期河道「P2、P3掛淤」→ 本期橋墩 P2、P3）。
- splits：被使用者拆開的 (前期 ID, 本期 ID)。
人工指定一律優先於自動配對；被拆散留下的另一筆改為單獨一筆、需確認。
本模組只做判讀，不修改網站。
"""
from dataclasses import dataclass

from .matching import (KIND_KEEP, KIND_LARGER, KIND_MISSING, KIND_NEW, KIND_REPAIRED, KIND_SMALLER, KIND_UNIT,
                       Pair, clean_pattern, codes, defect_id, number, pattern_similarity, score)

MANUAL = '人工配對'
SPLIT = '人工拆開'
LEFT = '原配對已改為人工配對'
TIERS = {1: '同一構件列', 2: '同構件其他列', 3: '其他構件（未配對）', 4: '已配給其他缺失'}


def classify_group(prior, currents):
    """一筆前期對多筆本期：單位相同時以本期數量合計比較。"""
    unit = prior.fields.get('unit', '')
    if any(c.fields.get('unit', '') != unit for c in currents):
        return KIND_UNIT
    before = number(prior.fields.get('amount'))
    amounts = [number(c.fields.get('amount')) for c in currents]
    if before is None or any(a is None for a in amounts):
        return KIND_KEEP
    total = sum(amounts)
    return KIND_LARGER if total > before else KIND_SMALLER if total < before else KIND_KEEP


def alone(defect, reason, confidence):
    if defect.period == '本期':
        return Pair(KIND_NEW, current=defect, confidence=confidence, reasons=[reason])
    return Pair(KIND_REPAIRED if defect.repaired else KIND_MISSING, prior=defect, confidence=confidence,
                reasons=[reason])


def linked_pairs(prior, currents):
    kind = classify_group(prior, currents)
    reasons = [MANUAL]
    moved = sorted({c.component for c in currents if c.component != prior.component})
    if moved:
        reasons.append(f"位置變更（{prior.component} → {'、'.join(moved)}）")
    if len(currents) > 1:
        total = sum((number(c.fields.get('amount')) or 0) for c in currents)
        reasons.append(f"一對多：本期共 {len(currents)} 筆，合計 {total.normalize():f} {prior.fields.get('unit', '')}"
                       f"（前期 {prior.fields.get('amount', '')}）")
    result = []
    for current in currents:
        points = score(prior, current)
        result.append(Pair(kind, prior, current, points[0] if points else 0, '高', list(reasons), manual=True))
    return result


def apply_manual(pairs, links=None, splits=None):
    """在自動配對結果上套用人工指定，回傳新的配對清單（不改動原清單）。"""
    priors = {defect_id(p.prior): p.prior for p in pairs if p.prior}
    currents = {defect_id(p.current): p.current for p in pairs if p.current}
    valid = {}
    for pid, cids in (links or {}).items():
        chosen = [c for c in dict.fromkeys(cids) if c in currents]
        if pid in priors and chosen:
            valid[pid] = chosen
    linked_c = {c for cids in valid.values() for c in cids}
    splits = {tuple(s) for s in splits or ()}
    result = []
    for pair in pairs:
        pid, cid = defect_id(pair.prior), defect_id(pair.current)
        split = bool(pair.prior and pair.current and (pid, cid) in splits)
        touched_p, touched_c = pid in valid, cid in linked_c
        if not (split or touched_p or touched_c):
            result.append(pair)
            continue
        if pair.prior and not touched_p:
            result.append(alone(pair.prior, SPLIT if split else LEFT, '高' if split else '需確認'))
        if pair.current and not touched_c:
            result.append(alone(pair.current, SPLIT if split else LEFT, '高' if split else '需確認'))
    for pid, cids in valid.items():
        result += linked_pairs(priors[pid], [currents[c] for c in cids])
    return result


@dataclass
class Candidate:
    tier: int
    defect: object
    similarity: float
    partner: str = ''          # 目前配給誰（空白：未配對）
    selected: bool = False     # 目前就是和選取的這筆配對

    @property
    def tier_name(self):
        return TIERS[self.tier]


def candidates(pairs, target):
    """改配對的候選清單。選取的配對有前期時列本期，否則列前期。
    排序：① 同一構件列（未配對在前）② 同構件其他列的未配對 ③ 其他構件的未配對（依樣態相似度）④ 已配給其他缺失。"""
    anchor = target.prior or target.current
    want_current = target.prior is not None
    own = {defect_id(p.current) for p in pairs if want_current and p.prior is target.prior and p.current}
    if not want_current and target.current:
        own = set()
    found, seen = [], set()
    anchor_pattern, anchor_codes = clean_pattern(anchor.fields), codes(anchor)
    for index, pair in enumerate(pairs):
        defect = pair.current if want_current else pair.prior
        if defect is None or defect_id(defect) in seen:
            continue
        seen.add(defect_id(defect))
        other = pair.prior if want_current else pair.current
        selected = defect_id(defect) in own
        partner = '' if other is None or selected else f'{other.period} {other.component} {other.remark.strip()}'
        if defect.row == anchor.row:
            tier = 1
        elif partner:
            tier = 4
        elif defect.component == anchor.component:
            tier = 2
        else:
            tier = 3
        similarity = pattern_similarity(anchor_pattern, clean_pattern(defect.fields)) or 0
        if anchor_codes & codes(defect):
            similarity += 0.1
        found.append((tier, not selected, bool(partner), -similarity, index,
                      Candidate(tier, defect, round(similarity, 2), partner, selected)))
    return [item[-1] for item in sorted(found, key=lambda item: item[:5])]
