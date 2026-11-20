"""依需求的三種表單狀態建立建議；任何歧義或量綱衝突暫停自動化。"""
from dataclasses import dataclass,replace
import re
from decimal import Decimal
from .data import repair_candidates
from .remarks import parse, prepare_fields,crack_count
from .conditions import TYPE_CORRECTION, aliases, condition, rule_deterioration
from .data import compact
from .reconcile import reconcile, scoped_history, type_repairs, complete_remark,numeric


@dataclass
class Decision:
    fields: dict
    reason: str
    review: bool = False
    rule_id: str = ""


def same_outcome(parsed, deterioration):
    """結果完全相同的候選（樣態、工法、單位、單價、修正後劣化類型都一樣）只留一筆，劣化類型與本期相同的優先。
    例：1120 的 1-03（路堤損傷 AC龜裂）與 1-12（其他損傷 AC龜裂，修正為路堤損傷）。"""
    kept = {}
    for entry in sorted(parsed, key=lambda e: e[0].get('deterioration') != (deterioration or e[0].get('deterioration'))
                        or bool(TYPE_CORRECTION.search(compact(e[0].get('note', ''))))):
        rule = entry[0]
        fixed = TYPE_CORRECTION.search(compact(rule.get('note', '')))
        key = (rule['pattern'], rule['method'], rule['unit'], str(rule['price']), entry[1].text,
               fixed[2] if fixed else rule.get('deterioration', ''))
        kept.setdefault(key, entry)
    return list(kept.values()) if len(kept) < len(parsed) else parsed


def automatic_plan(current, component, history, repairs, context=None, apply_rules=True):
    matches = scoped_history(history, component, context, current.get('deterioration',''), current.get('remark',''))
    # 同一構件／編號／劣化樣態只有一筆，才可還原摘要或修正誤填單位。
    previous = matches[0] if len(matches) == 1 and current.get('deterioration') else None
    current = prepare_fields(current,component=component,context=context)
    draft, problem = reconcile(current, previous)
    draft = prepare_fields(draft,component=component,context=context)
    if problem:
        return Decision(draft, problem, True)
    if not apply_rules:
        from .data import normalize_remark
        draft['remark'] = normalize_remark(draft.get('remark',''))
        if not complete_remark(draft.get('remark')) or ' ' not in draft['remark']:
            return Decision(draft, '未套用 Excel，備註仍無法辨識完整位置及數量式', True)
        draft['d_location'] = draft['remark'].split(None,1)[0]
        if not all(draft.get(key,'').strip() for key in ('deterioration','d_location','method','amount','unit','price','remark')):
            return Decision(draft, '未套用 Excel，仍有缺少欄位，請人工確認', True)
        return Decision(draft, '還原唯一上期完整資料，不套用 Excel 規則')
    repairs = type_repairs(repairs, draft.get('deterioration',''))
    rules = repair_candidates([r for r in repairs if r.get('automatic_ready', True)], draft, component, filter_unit=False)
    if not rules:
        return Decision(draft, 'Excel 無對應值，暫停此筆自動處理，人工完成送出後繼續', True)
    parsed = []
    manual_matches=[]
    unknown = False
    for rule in rules:
        eligible, manual = condition(rule, context, draft.get('remark', ''))
        if eligible is False:
            continue
        if eligible is None:
            unknown = True
            continue
        if manual:manual_matches.append(rule)
        try:
            text=draft.get('remark','')
            if rule['pattern']=='基礎裸露':
                text=re.sub(r'^([AP]\d+)(?=(?:橋墩|橋台)基礎)',r'\1 ',text)
            info = parse(text, rule['pattern'], rule['unit'], aliases(rule.get('note', '')))
            plant=any(word in rule['pattern'] for word in ('植物','植生'))
            count=numeric(draft.get('amount',''))
            if plant and rule['unit']=='式' and info.unit=='平方公尺':
                info=replace(info,unit='式',amount='1')
            elif not plant and rule['unit']=='式' and draft.get('unit')=='式' and count is not None and Decimal(count)>0:
                # 固定「式」沿用本期明填的式數，尺寸僅描述範圍，不換算成式。
                info=replace(info,unit='式',amount=count)
            if '裂縫' in info.description:info=replace(info,text=crack_count(info.text))
            if info.unit == rule['unit'].replace('／', '/'):
                parsed.append((rule, info, manual))
        except ValueError:
            continue
    if unknown:
        ids='、'.join(r.get('rule_id','未編號') for r in rules)
        return Decision(draft, f'候選規則 {ids}：使用條件缺少主頁資料或無法確認，尚未套用單價', True)
    preferred = [entry for entry in parsed if any(word in entry[0].get('note', '') for word in ('優先選擇', '優先使用'))]
    if preferred:
        parsed = preferred
    parsed = same_outcome(parsed, draft.get('deterioration', ''))
    if len(parsed) != 1:
        if len(rules)==1 and len(manual_matches)==1:
            rule=manual_matches[0]
            # 人工規則可先提供明確工法／單價，尺寸不換算成「式」數量。
            draft.update(method=rule['method'],unit=rule['unit'],price=rule['price'],
                         deterioration=rule_deterioration(rule,draft.get('deterioration','')))
            return Decision(draft,'Excel 指定手動處理：已帶入工法、單位與單價；保留本期數量／備註，請核對後確認',True,rule.get('rule_id',''))
        ids='、'.join(r.get('rule_id','未編號') for r in rules)
        reason='多個可用方案' if len(parsed)>1 else '備註量式無法解析或與規則單位不符'
        return Decision(draft,f'候選規則 {ids}：{reason}，尚未套用單價',True)
    rule, info, manual = parsed[0]
    if manual:
        draft.update(method=rule['method'],unit=rule['unit'],price=rule['price'],
                     deterioration=rule_deterioration(rule,draft.get('deterioration','')))
        return Decision(draft, 'Excel 指定手動選擇：已帶入工法、單位與單價，請人工確認方案', True,rule.get('rule_id',''))
    # 新版表單明確授權檢核及替換工法、單位、單價與劣化類型。
    draft.update(remark=info.text, d_location=info.location, amount=info.amount,
                 method=rule['method'], unit=rule['unit'], price=rule['price'])
    draft['deterioration'] = rule_deterioration(rule,draft.get('deterioration',''))
    return Decision(draft, '依備註回填位置與計算數量，並按 Excel 檢核類型、工法、單位及單價',rule_id=rule.get('rule_id',''))
