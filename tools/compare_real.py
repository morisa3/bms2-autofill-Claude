"""實際案例比對：未修改的掃描 → 程式自動建議，與使用者最終送出的值逐欄比較，量測自動化程度。

資料（不上傳）：fixtures/private/real-2026-11-18/
  未修改__{日期}-{橋名}-檢測.json、未修改__{前期日期}-{橋名}-前期.json：新建本期時的網站狀態（多數只有 DERU、類型、數量）
  修改完成__{日期}-{橋名}-填寫.json、修改完成__{日期}-{橋名}-審查.json：最終送出值與審查判定
用法：
  python tools/compare_real.py            # 不帶審查（全自動）
  python tools/compare_real.py review     # 帶入使用者的審查判定與改配對
  python tools/compare_real.py review v   # 列出每一筆不同的欄位
"""
import collections
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bridgefill.core.editing import load_rules, rules_file  # noqa: E402
from bridgefill.core.fill import TEXT_FIELDS, build_plan  # noqa: E402
from bridgefill.core.manual import apply_manual  # noqa: E402
from bridgefill.core.matching import match_scans  # noqa: E402
from bridgefill.core.review import Review  # noqa: E402

DATA = Path(__file__).resolve().parents[1] / 'fixtures' / 'private' / 'real-2026-11-18'


def compare(folder=DATA, use_review=False, verbose=False):
    rules = load_rules(rules_file())
    result = {}
    for current_path in sorted(folder.glob('未修改__*-檢測.json')):
        bridge = current_path.name.split('__')[1][len('2026-11-18-'):-len('-檢測.json')]
        prior_path = next(folder.glob(f'未修改__*-{bridge}-前期.json'))
        final_path = next(folder.glob(f'修改完成__*-{bridge}-填寫.json'), None)
        if final_path is None:
            continue
        current = json.loads(current_path.read_text(encoding='utf-8'))
        prior = json.loads(prior_path.read_text(encoding='utf-8'))
        final = json.loads(final_path.read_text(encoding='utf-8'))['entries']
        pairs = match_scans(prior, current).pairs
        review = Review()
        if use_review:
            review = Review.load(next(folder.glob(f'修改完成__*-{bridge}-審查.json')))
            pairs = apply_manual(pairs, review.links, review.splits)
            review.sync(pairs)
        items = build_plan(pairs, review, rules)
        same, differ, fields, prefilled = 0, [], collections.Counter(), 0
        for item in items:
            target = final.get(item.key, {}).get('target')
            if not target:
                continue
            shown = {**item.target, **{k: v for k, v in (item.proposal or {}).items() if v}} \
                if item.status == '需確認' else item.target          # 工作區本期欄預填的值
            bad = [k for k in TEXT_FIELDS + ('d', 'e', 'r', 'u')
                   if str(item.target.get(k, '')).strip() != str(target.get(k, '')).strip()]
            prefilled += not [k for k in TEXT_FIELDS + ('d', 'e', 'r', 'u')
                              if str(shown.get(k, '')).strip() != str(target.get(k, '')).strip()]
            if bad:
                differ.append((item, target, bad))
                fields.update(bad)
            else:
                same += 1
        status = collections.Counter(i.status for i in items)
        print(f'{bridge}：程式狀態 {dict(status)}；與最終相同 {same}、不同 {len(differ)}'
              f'（工作區預填後相同 {prefilled}）；不同欄位 {dict(fields)}')
        if verbose:
            for item, target, bad in differ:
                c = item.pair.current
                print(f'  [{item.status}] {c.component} {c.code}  前期：{item.pair.prior.remark if item.pair.prior else "—"}')
                for k in bad:
                    print(f'      {k}：程式＝{item.target.get(k)!r}  最終＝{target.get(k)!r}')
        result[bridge] = (same, len(differ))
    return result


if __name__ == '__main__':
    compare(use_review='review' in sys.argv[1:], verbose='v' in sys.argv[1:])
