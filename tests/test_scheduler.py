"""Phase 4 單元測試：橋梁總表目錄讀取、重複橋名防誤判分流、監看時段輪詢與穩定緩衝。"""
import tempfile
import unittest
from datetime import datetime, time, timedelta
from pathlib import Path

from openpyxl import Workbook

from bridgefill.core.scheduler import (
    STATUS_AMBIGUOUS, STATUS_DONE, STATUS_STABILIZING, STATUS_WAITING,
    WatchItem, WatchQueue, disambiguate_bridge, load_list_bridges,
)
from bridgefill.site.pages import BridgeMatch, InspectionRecord


def _bridge(inv_id, name, place='雲林縣北港鎮', years=()):
    records = [
        InspectionRecord(inspect_id=f'{inv_id}{y}', inventory_id=inv_id, date=f'{y}-05-01',
                         unit='雲林縣政府', inspectors='王小明', edit_url=f'/edit/{inv_id}/{y}',
                         view_url=f'/view/{inv_id}/{y}', bridge=name, code='', agency='雲林縣政府', place=place)
        for y in sorted(years, reverse=True)
    ]
    return BridgeMatch(inv_id, name, '雲林縣政府', place, records)


class DisambiguationTests(unittest.TestCase):
    """驗證重複橋名分流優先順序與防改建新舊橋誤判。"""

    def test_exact_id_wins_immediately(self):
        cands = [_bridge('1001', '平和橋', '雲林縣北港鎮', [2024]),
                 _bridge('2002', '平和橋', '雲林縣斗六市', [2024])]
        res = disambiguate_bridge(cands, wanted_name='平和橋', inventory_id='2002', township='北港鎮', current_year=2026)
        self.assertFalse(res.ambiguous)
        self.assertEqual(res.method, 'id')
        self.assertEqual(res.bridge.inventory_id, '2002')

    def test_id_matching_demolished_bridge_warns_and_blocks(self):
        cands = [_bridge('1001', '平和橋(已拆除)', '雲林縣北港鎮', [2020]),
                 _bridge('2002', '平和橋', '雲林縣北港鎮', [2024])]
        res = disambiguate_bridge(cands, wanted_name='平和橋', inventory_id='1001', current_year=2026)
        self.assertTrue(res.ambiguous)
        self.assertIsNone(res.bridge)
        self.assertIn('已拆除', res.warning)

    def test_excludes_demolished_before_other_checks(self):
        cands = [_bridge('1001', '仁愛橋(已拆除)', '雲林縣北港鎮', [2022]),
                 _bridge('2002', '仁愛橋', '雲林縣北港鎮', [2024])]
        res = disambiguate_bridge(cands, wanted_name='仁愛橋', current_year=2026)
        self.assertFalse(res.ambiguous)
        self.assertEqual(res.method, 'exclude_demolished')
        self.assertEqual(res.bridge.inventory_id, '2002')

    def test_township_matches_place(self):
        cands = [_bridge('1001', '中正橋', '雲林縣北港鎮', [2024]),
                 _bridge('2002', '中正橋', '雲林縣虎尾鎮', [2024])]
        res = disambiguate_bridge(cands, wanted_name='中正橋', township='虎尾鎮(115)', current_year=2026)
        self.assertFalse(res.ambiguous)
        self.assertEqual(res.method, 'township')
        self.assertEqual(res.bridge.inventory_id, '2002')

    def test_cycle_date_fallback_picks_active_over_retired_2020(self):
        # 2026 年檢測：一座有 2024/2022（活躍），另一座最後停在 2020（已無使用）
        cands = [_bridge('1001', '成功橋', '雲林縣北港鎮', [2020, 2018]),
                 _bridge('2002', '成功橋', '雲林縣北港鎮', [2024, 2022])]
        res = disambiguate_bridge(cands, wanted_name='成功橋', township='北港鎮', current_year=2026)
        self.assertFalse(res.ambiguous)
        self.assertEqual(res.method, 'cycle_date')
        self.assertEqual(res.bridge.inventory_id, '2002')

    def test_rebuilt_new_bridge_vs_unmarked_old_bridge_is_ambiguous(self):
        # 改建後新橋（僅 2026 或無前期紀錄）與未標註「已拆除」的舊橋（2024 有檢）並存 → 絕不自動亂猜
        cands = [_bridge('1001', '復興橋', '雲林縣北港鎮', [2024, 2022]),
                 _bridge('3003', '復興橋', '雲林縣北港鎮', [2026])]
        res = disambiguate_bridge(cands, wanted_name='復興橋', township='北港鎮', current_year=2026)
        self.assertTrue(res.ambiguous)
        self.assertIsNone(res.bridge)
        self.assertIn('改建', res.warning)


class WatchQueueTests(unittest.TestCase):
    """驗證監看時段輪詢、上傳穩定緩衝與提早收工立即掃描。"""

    def test_stability_buffer_waits_for_unchanged_defect_count(self):
        q = WatchQueue(start_time=time(13, 0), end_time=time(18, 0), interval_minutes=10, stabilize_minutes=5, enabled=True)
        q.add_items([{'sheet': '北港鎮', 'name': '平和橋', 'inventory_id': '12345'}])
        item = q.items[0]
        t0 = datetime(2026, 6, 1, 14, 0)

        self.assertTrue(q.is_due(t0))
        # 1. 尚未上傳本期 → 繼續等待
        self.assertFalse(q.evaluate_probe(item, has_current=False, now=t0))
        self.assertEqual(item.status, STATUS_WAITING)

        # 2. 14:10 首次看到本期（3 筆缺失）→ 進入等待穩定
        t1 = t0 + timedelta(minutes=10)
        self.assertFalse(q.evaluate_probe(item, has_current=True, defect_count=3, now=t1))
        self.assertEqual(item.status, STATUS_STABILIZING)

        # 3. 14:14 App 又補傳了 2 筆（變成 5 筆）→ 重新計時
        t2 = t1 + timedelta(minutes=4)
        self.assertFalse(q.evaluate_probe(item, has_current=True, defect_count=5, now=t2))
        self.assertEqual(item.seen_count, 5)
        self.assertIn('3 → 5', item.message)

        # 4. 14:17 僅過 3 分鐘（未滿 5 分鐘）→ 仍等待
        t3 = t2 + timedelta(minutes=3)
        self.assertFalse(q.evaluate_probe(item, has_current=True, defect_count=5, now=t3))
        self.assertEqual(item.status, STATUS_STABILIZING)

        # 5. 14:19 滿 5 分鐘且筆數維持 5 → 放行掃描！
        t4 = t2 + timedelta(minutes=5)
        self.assertTrue(q.evaluate_probe(item, has_current=True, defect_count=5, now=t4))
        q.mark_done(item, now=t4)
        self.assertEqual(item.status, STATUS_DONE)
        self.assertIn('已掃完 1', q.status_summary())

    def test_early_return_force_immediate_bypasses_stability_buffer(self):
        q = WatchQueue(start_time=time(16, 30), end_time=time(18, 0), stabilize_minutes=5, enabled=True)
        q.add_items([{'sheet': '斗六市', 'name': '雲林橋', 'inventory_id': '8888'}])
        item = q.items[0]
        # 16:20 提早收工：不在原定視窗內，但手動按「立即掃描剩餘排程」可直接掃
        early = datetime(2026, 6, 1, 16, 20)
        self.assertFalse(q.is_within_window(early))
        self.assertTrue(q.evaluate_probe(item, has_current=True, defect_count=4, now=early, force_immediate=True))

    def test_ambiguous_item_excluded_from_pending(self):
        q = WatchQueue(enabled=True)
        q.add_items([WatchItem(name='重複橋')])
        q.mark_ambiguous(q.items[0], '需人工確認')
        self.assertEqual(q.pending_items(), [])
        self.assertIn('同名需人工確認 1', q.status_summary())


class LoadListBridgesTests(unittest.TestCase):
    """驗證從本機 Excel 總表與 Google 試算表載入各鄉鎮橋梁目錄。"""

    def test_load_from_local_excel(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = '北港鎮'
        ws1.append(['序號', '二代系統ID', '橋梁名稱', '系統填寫完成'])
        ws1.append([1, 10001, '平和橋', '2026-06-01'])
        ws1.append([2, '10002', '仁愛橋', None])
        ws2 = wb.create_sheet('斗六市')
        ws2.append(['序號', '二代系統ID', '橋梁名稱', '系統填寫完成'])
        ws2.append([1, 20001, '雲林橋', ''])

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'master.xlsx'
            wb.save(path)
            catalog = load_list_bridges(local_path=path)
            self.assertEqual(catalog['sheets'], ['北港鎮', '斗六市'])
            self.assertEqual(len(catalog['bridges']), 3)
            self.assertTrue(catalog['bridges'][0]['done'])
            self.assertEqual(catalog['bridges'][0]['inventory_id'], '10001')
            self.assertFalse(catalog['bridges'][1]['done'])

    def test_load_from_google_client(self):
        class FakeClient:
            def list_bridges(self):
                return {'sheets': [{'name': '虎尾鎮', 'bridges': [
                    {'name': '虎尾橋', 'id': '30001', 'done': False, 'row': 2},
                    {'name': '平和橋', 'id': '30002', 'done': True, 'row': 3},
                ]}]}

        catalog = load_list_bridges(google_client=FakeClient())
        self.assertEqual(catalog['sheets'], ['虎尾鎮'])
        self.assertEqual(len(catalog['bridges']), 2)
        self.assertEqual(catalog['bridges'][0]['inventory_id'], '30001')
        self.assertTrue(catalog['bridges'][1]['done'])


if __name__ == '__main__':
    unittest.main()
