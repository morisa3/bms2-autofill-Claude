"""以 v1.00 比對檔中使用者確認過的配對當標準答案，量測新配對引擎的準確率。

用法：
    python tools/evaluate_matching.py 比對未修改.json 比對修改完成.json
「未修改」提供原始欄位，「修改完成」的 pairs（本期 → 前期第幾筆，None 表示新增）為標準答案。
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bridgefill.core.locations import key as location_key   # noqa: E402
from bridgefill.core.matching import Defect, match          # noqa: E402


def row_key(component, context):
    """v1.00 沒有網站構件 ID，以構件＋第一個構件代碼當作列。
    前後期的欄位名稱不同（位置／橋孔編號／所屬橋台編號），寫法也不同（S001／S1、橋頭／A001），
    所以只取值裡的第一個代碼並統一寫法；沒有代碼的構件（如河道）整個構件為一列。"""
    for label, value in (context or {}).items():
        if label.startswith('_') or not re.search(r'位置|編號', label):
            continue
        text = str(value).strip()
        if text in ('橋頭', '橋尾'):
            return (component, 'A1' if text == '橋頭' else 'A2')
        found = re.search(r'([APS])0*(\d+)', text)
        if found:
            return (component, found[1] + found[2])
    return (component, '')


def deru(context):
    if context.get('DERU'):
        return ''.join(re.findall(r'[DERU]:(\S)', context['DERU']))
    return ''.join(str(context.get(k, '_')) for k in 'DERU')


# v1.00 比對檔中經使用者重新確認、更正的答案（2026-11-15）：(構件, 本期備註) → 前期備註（None 表示新增）
CORRECTIONS = {
    ('伸縮縫', 'P2 雜物堆積。(9m)'): 'P2 雜物堆積 (9m)',          # 沿用，非新增
    ('墩柱與帽梁', 'P1-正 混凝土蜂窩。(1m*1m)'): 'P1正 混凝土蜂窩',   # 對應前期 P1正 混凝土蜂窩
    ('墩柱與帽梁', 'P1-2 混凝土剝落。(1處)'): None,                  # 因此本筆為新增
}


def load(original, confirmed):
    before = json.loads(Path(original).read_text(encoding='utf-8'))
    after = json.loads(Path(confirmed).read_text(encoding='utf-8'))
    priors = [Defect('前期', row_key(h['component'], h.get('context')), h['component'], '', deru(h.get('context', {})),
                     h['fields'].get('deterioration', ''), h['fields']) for h in before['history']]
    currents, truth = [], {}
    for entry in before['entries']:
        fields = entry.get('fields') or entry.get('current') or {}
        defect = Defect('本期', row_key(entry['component'], entry.get('context')), entry['component'], '',
                        deru(entry.get('context', {})), fields.get('deterioration', ''), fields,
                        method_id=entry['id'])
        currents.append(defect)
        truth[entry['id']] = after['pairs'].get(entry['id'])
        fixed = (entry['component'], fields.get('remark', '').strip())
        if fixed in CORRECTIONS:
            answer = CORRECTIONS[fixed]
            truth[entry['id']] = None if answer is None else next(
                i for i, p in enumerate(priors) if p.component == entry['component'] and p.remark.strip() == answer)
    return priors, currents, truth


def evaluate(original, confirmed, show=True):
    priors, currents, truth = load(original, confirmed)
    result = match(priors, currents)
    predicted = {p.current.method_id: (priors.index(p.prior) if p.prior else None)
                 for p in result.pairs if p.current}
    right = wrong = 0
    details = []
    for entry_id, answer in truth.items():
        guess = predicted.get(entry_id)
        ok = guess == answer
        right += ok
        wrong += not ok
        if not ok:
            current = next(c for c in currents if c.method_id == entry_id)
            details.append((current, None if answer is None else priors[answer],
                            None if guess is None else priors[guess]))
    cross_row = sum(1 for e, a in truth.items() if a is not None and
                    priors[a].row != next(c for c in currents if c.method_id == e).row)
    if show:
        print(f'本期 {len(currents)} 筆：正確 {right}、錯誤 {wrong}（準確率 {right / len(currents):.1%}）')
        print(f'標準答案中跨列的配對：{cross_row} 筆（v1.00 資料沒有構件 ID，以位置欄位推列）')
        for current, answer, guess in details:
            print(f'  ✗ {current.component}｜{current.remark}')
            print(f'      答案：{answer.remark if answer else "新增"}')
            print(f'      程式：{guess.remark if guess else "新增"}')
    return right, wrong, details


if __name__ == '__main__':
    evaluate(*sys.argv[1:3])
