"""測試歷史紀錄依檢測日期分組與依日期清除。"""
import json
import tempfile
import unittest
from pathlib import Path

from bridgefill.core.review import (
    clear_sessions, group_sessions_by_date, saved_sessions,
    sessions_before_date, sessions_for_dates,
)


def _write_scan(scan_dir, date, bridge):
    p = Path(scan_dir) / f'{date}-{bridge}-檢測.json'
    p.write_text(json.dumps({'header': {'橋梁名稱': bridge, '檢測日期': date}, 'rows': []},
                            ensure_ascii=False), encoding='utf-8')
    return p


class HistoryDateGroupingTests(unittest.TestCase):
    def test_group_sessions_by_latest_date_descending(self):
        with tempfile.TemporaryDirectory() as folder:
            scan_dir = Path(folder) / 'scan'
            review_dir = Path(folder) / 'review'
            scan_dir.mkdir()
            _write_scan(scan_dir, '2026-05-10', '甲橋')
            _write_scan(scan_dir, '2026-06-02', '乙橋')
            _write_scan(scan_dir, '2026-06-02', '丙橋')
            _write_scan(scan_dir, '2026-05-20', '丁橋')
            _write_scan(scan_dir, '', '無日期橋')

            sessions = saved_sessions(scan_dir, review_dir)
            grouped = group_sessions_by_date(sessions)
            self.assertEqual([d for d, _ in grouped], ['2026-06-02', '2026-05-20', '2026-05-10', '未知日期'])
            self.assertEqual(len(grouped[0][1]), 2)

            # 依日期篩選與刪除（只刪 2026-05-20，其餘保留）
            target = sessions_for_dates(sessions, ['2026-05-20'])
            self.assertEqual([s['bridge'] for s in target], ['丁橋'])
            clear_sessions(target, scan_dir, review_dir)

            # 刪除早於 2026-06-02 的紀錄（應刪甲橋，保留乙橋、丙橋與無日期橋）
            after = saved_sessions(scan_dir, review_dir)
            older = sessions_before_date(after, '2026-06-02')
            self.assertEqual([s['bridge'] for s in older], ['甲橋'])
            clear_sessions(older, scan_dir, review_dir)

            remaining = saved_sessions(scan_dir, review_dir)
            self.assertEqual({s['bridge'] for s in remaining}, {'乙橋', '丙橋', '無日期橋'})


if __name__ == '__main__':
    unittest.main()
