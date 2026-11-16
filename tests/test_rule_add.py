"""Excel 沒有對應規則時自動新增（新劣化樣態、劣化類型判讀差異），新增後填寫清單可直接套用。"""
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from bridgefill.core.editing import rules_file
from bridgefill.core.fill import READY, SAME, plan_item
from bridgefill.core.matching import Defect, Pair
from bridgefill.core.review import Review
from bridgefill.core.rule_add import NO_RULE, add_rules, suggest_rule, writable_rules_file
from bridgefill.core.rules_v2 import read_rules


def pier_guard(remark, kind, method='', unit='平方公尺', price=''):
    fields = {'deterioration': kind, 'd_location': 'P1-R', 'method': method, 'amount': '4', 'unit': unit,
              'price': price, 'remark': remark, 'd': '2', 'e': '2', 'r': '2', 'u': '2'}
    d = Defect('本期', ('pier_protect', '1'), '橋墩／橋基保護設施', '', '2222', kind, fields, method_id='9100',
               der_id='700', inspect_id='9', options={'deterioration': [kind], 'unit': ['平方公尺', '處']})
    return Pair('新增', current=d, confidence='高')


class RuleAddTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.path = Path(self.folder.name) / '規則.xlsx'
        shutil.copy(rules_file(), self.path)
        self.rules = read_rules(self.path)

    def tearDown(self):
        self.folder.cleanup()

    def test_type_difference_copies_same_pattern_rule(self):
        pair = pier_guard('P1-R 防衝鋼板損毀。(2m*2m)', '防撞鋼板損傷')     # 規則表只有「保護設施損壞」的損毀
        self.assertTrue(plan_item(pair, Review(), self.rules).reason.startswith(NO_RULE))
        rule = suggest_rule(pair.current, self.rules)
        source = next(r for r in self.rules if r['rule_id'] == rule['source'])
        self.assertEqual((rule['kind'], rule['劣化樣態'], rule['劣化類型']), ('劣化類型判讀差異', '防衝鋼板損毀', '防撞鋼板損傷'))
        self.assertEqual((rule['維修工法'], rule['單價']), (source['method'], source['price']))
        added = add_rules(self.path, [rule], backup_dir=Path(self.folder.name) / 'bk')
        rule_id, sheet, _ = added[0]
        self.assertTrue(list((Path(self.folder.name) / 'bk').glob('*.xlsx')))     # 寫入前備份
        rules = read_rules(self.path)
        new = next(r for r in rules if r['rule_id'] == rule_id)
        self.assertEqual((new['deterioration'], new['pattern']), ('防撞鋼板損傷', '防衝鋼板損毀'))
        item = plan_item(pair, Review(), rules)                      # 新增後即可依規則填寫
        self.assertEqual(item.status, READY)
        self.assertEqual(item.target['price'], source['price'])
        ws = load_workbook(self.path)[sheet]
        headers = [c.value for c in ws[1]]
        row = next(r for r in ws.iter_rows(min_row=2) if r[headers.index('劣化ID')].value == rule_id)
        note = row[headers.index('說明')]
        self.assertIn('自動新增（劣化類型判讀差異）', note.value)
        self.assertEqual(note.fill.fgColor.rgb[-6:], 'FFF2CC')
        self.assertIsNone(suggest_rule(pair.current, rules))         # 已有就不重複新增

    def test_new_pattern_uses_site_values(self):
        pair = pier_guard('P1-R 鋼板塗鴉。(2m*2m)', '防撞鋼板損傷', method='高壓水槍清洗', price='350')
        rule = suggest_rule(pair.current, self.rules)
        self.assertEqual((rule['kind'], rule['劣化樣態'], rule['維修工法'], rule['維修單位'], rule['單價']),
                         ('新劣化樣態', '鋼板塗鴉', '高壓水槍清洗', '平方公尺', '350'))
        add_rules(self.path, [rule], backup_dir=Path(self.folder.name) / 'bk')
        self.assertEqual(plan_item(pair, Review(), read_rules(self.path)).status, SAME)

    def test_sync_adds_missing_sheets_only(self):
        from openpyxl import load_workbook as load
        from bridgefill.core.rule_add import sync_sheets
        old = Path(self.folder.name) / '使用中.xlsx'
        workbook = load(self.path)
        del workbook['護床工']
        workbook.save(old)
        rule = suggest_rule(pier_guard('P1-R 鋼板塗鴉。(2m*2m)', '防撞鋼板損傷', method='高壓水槍清洗', price='350').current,
                            read_rules(old))
        add_rules(old, [rule], backup_dir=Path(self.folder.name) / 'bk')          # 使用中的檔案有自動新增的規則
        self.assertEqual(sync_sheets(self.path, old, Path(self.folder.name) / 'bk'), ['護床工'])
        rules = read_rules(old)
        expected = len([r for r in read_rules(self.path) if '護床工' in r['components']])
        self.assertEqual(len([r for r in rules if '護床工' in r['components']]), expected)
        self.assertTrue([r for r in rules if r['pattern'] == '鋼板塗鴉'])                # 自動新增的保留
        self.assertEqual(sync_sheets(self.path, old, Path(self.folder.name) / 'bk'), [])  # 已有：不重複

    def test_migrate_auto_rules_from_old_file(self):
        from bridgefill.core.rule_add import migrate_auto_rules
        old = Path(self.folder.name) / '劣化說明與維修單價1111v2.xlsx'
        shutil.copy(self.path, old)
        rule = suggest_rule(pier_guard('P1-R 鋼板塗鴉。(2m*2m)', '防撞鋼板損傷', method='高壓水槍清洗', price='350').current,
                            read_rules(old))
        add_rules(old, [rule], backup_dir=Path(self.folder.name) / 'bk')
        new = Path(self.folder.name) / '劣化說明與維修單價1116.xlsx'
        shutil.copy(self.path, new)
        self.assertEqual(migrate_auto_rules(old, new, Path(self.folder.name) / 'bk'), 1)
        self.assertFalse(old.exists())                                    # 舊檔移到 backup
        self.assertTrue([r for r in read_rules(new) if r['pattern'] == '鋼板塗鴉'])
        self.assertEqual(migrate_auto_rules(old, new), 0)

    def test_suggestion_file_is_copied_to_data_folder_only(self):
        self.assertEqual(writable_rules_file(self.path), self.path)


if __name__ == '__main__':
    unittest.main()
