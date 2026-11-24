"""拱橋構件（拱肋、立柱、斜撐、橫桿）與清單用的簡短編號。"""
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from bridgefill.core.config import COMPONENTS
from bridgefill.core.data import component_name, component_repairs
from bridgefill.core.locations import short_code
from bridgefill.core.rule_add import add_rules, sheet_for
from bridgefill.core.rules_v2 import read_rules
from bridgefill.site.pages import parse_inspection

ROOT = Path(__file__).resolve().parents[1]
RULES = ROOT / 'data' / '劣化說明與維修單價1124.xlsx'


class ArchTests(unittest.TestCase):
    def test_components_and_tables(self):
        self.assertEqual([COMPONENTS[c] for c in ('拱肋', '立柱', '斜撐', '橫桿')],
                         ['arch', 'arch_column', 'lateral_bracing', 'crossbeam'])
        self.assertEqual(component_name('拱肋（拱圈）'), '拱肋')
        page = """<html><body><table><tr><th>橋梁名稱</th><td>拱橋</td></tr><tr><th>檢測日期</th><td>2026-11-04</td></tr></table>
<table id="arch_column"><tr><th>狀態</th><th>橋孔編號</th><th>拱圈編號</th><th>立柱編號</th><th>DERU</th><th>動作</th><th>維修狀態</th></tr>
<tr><td></td><td>S001</td><td>R01-1</td><td>R01-1H01</td><td>2 4 1 1</td>
<td><a href="/bms3/inspect/inspect/der_list/component_form,arch_column/component_id,5/inspect_id,9/">詳細資料(1)</a></td><td>未維修</td></tr></table></body></html>"""
        rows = parse_inspection(page).rows
        self.assertEqual((rows[0].component, short_code(rows[0].component, rows[0].columns)), ('立柱', 'S1/R01-1H01'))

    def test_rules_shared_with_existing_sheets(self):
        rules = read_rules(RULES)
        kinds = {c: {r['deterioration'] for r in component_repairs(rules, c)} for c in ('拱肋', '立柱', '斜撐', '橫桿', '主梁')}
        self.assertIn('混凝土剝落、破碎、鋼筋外露、銹蝕', kinds['拱肋'])           # 與橫隔梁、橋面板相同的選單文字
        self.assertNotIn('混凝土剝落、破碎、鋼筋、鋼腱或錨碇外露、銹蝕', kinds['拱肋'])   # 主梁專用
        self.assertIn('墩柱傾斜、沉陷', kinds['立柱'])
        self.assertIn('銲接處損傷', kinds['立柱'])                                  # 立柱獨有，加在墩柱與帽梁
        self.assertNotIn('銲接處損傷', {r['deterioration'] for r in component_repairs(rules, '墩柱與帽梁')})
        self.assertEqual(kinds['斜撐'], kinds['橫桿'])
        self.assertTrue(component_repairs(rules, '主梁'))

    def test_auto_added_rule_goes_to_shared_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'rules.xlsx'
            shutil.copy(RULES, path)
            self.assertEqual(sheet_for(load_workbook(path), '斜撐').title, '主梁或橫隔梁或橋面版')
            self.assertEqual(sheet_for(load_workbook(path), '立柱').title, '墩柱與帽梁')
            added = add_rules(path, [{'劣化類型': '其他損傷', '劣化樣態': '鋼構件變形', '維修工法': '整修', '維修單位': '處',
                                      '單價': 1000, '處理方式': '自動', '任一關鍵字': '鋼構件變形', 'kind': '新劣化樣態',
                                      'source': '本期網站值', 'component': '斜撐'}], backup_dir=Path(folder) / 'backup')
            self.assertEqual(added[0][1], '主梁或橫隔梁或橋面版')
            rule = next(r for r in read_rules(path) if r['pattern'] == '鋼構件變形')
            self.assertEqual(rule['components'], ['斜撐'])


class CodeTests(unittest.TestCase):
    def test_short_codes(self):
        cases = [('引道路堤', {'引道路堤位置': '橋頭'}, 'A1'), ('引道路堤', {'引道路堤位置': '橋尾'}, 'A2'),
                 ('伸縮縫', {'位置': '橋台(A)', '編號': '001'}, 'A1'), ('伸縮縫', {'位置': '橋墩(P)', '編號': '003'}, 'P3'),
                 ('支承／支承墊／阻尼裝置', {'位置': '橋墩(P)', '編號': '2'}, 'P2'),
                 ('墩柱與帽梁', {'橋墩編號': 'P002', '編號': 'P02N-1'}, 'P2/P02N-1'),
                 ('主梁', {'橋孔位置': 'S003', '編號': 'S3G1'}, 'S3/G1'), ('橋面板', {'橋孔編號': 'S004', '橋面板編號': 'S4'}, 'S4'),
                 ('其他', {'設施類別': '鋪面', '設施編號': '08'}, '鋪面'), ('拱肋', {'橋孔編號': 'S003', '拱圈編號': 'R01-1'}, 'S3/R01-1'),
                 ('立柱', {'橋孔編號': 'S001', '拱圈編號': 'R01-1', '立柱編號': 'R01-1H01'}, 'S1/R01-1H01'),
                 ('斜撐', {'橋孔編號': 'S001'}, 'S1'), ('橫桿', {'橋孔編號': 'S001', '橫桿編號': 'R01-B01'}, 'S1/R01-B01'),
                 ('橋台', {'橋台位置': '橋頭', '編號': 'A1'}, 'A1'), ('橋護欄', {}, '')]
        for component, context, expected in cases:
            self.assertEqual(short_code(component, context), expected, (component, context))


if __name__ == '__main__':
    unittest.main()
