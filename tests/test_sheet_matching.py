"""採 v2 的實際工作表名稱建立不含使用者單價的測試活頁簿。"""
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook
from bridgefill.core.data import repair_components, read_repairs, component_repairs

V2_SHEETS = {
    '引道路堤': ['引道路堤'],
    '引道路堤護欄或橋護欄': ['引道路堤護欄', '橋護欄'],
    '引道路堤保護措施': ['引道道路堤保護措施'],
    '河道': ['河道'],
    '橋台基礎或橋墩基礎': ['橋台基礎', '橋墩基礎'],
    '橋台': ['橋台'],
    '_x0009_伸縮縫': ['伸縮縫'],
    '橋梁排水設施': ['橋梁排水設施'],
    '支承或支承墊或阻尼裝置': ['支承／支承墊／阻尼裝置'],
    '_x0009_防落設施': ['防落設施'],
    '橋墩或橋基保護設施': ['橋墩／橋基保護設施'],
    '墩柱與帽梁': ['墩柱與帽梁'],
    '主梁或橫隔梁或橋面版': ['主梁', '橫隔梁', '橋面板'],
    '其他': ['其他'],
}


class SheetMatchingTests(unittest.TestCase):
    def test_every_v2_sheet_maps_to_expected_components(self):
        for title, expected in V2_SHEETS.items():
            with self.subTest(sheet=title):
                self.assertEqual(repair_components(title), expected)

    def test_separator_variants_preserve_composite_and_shared_components(self):
        for separator in ('/', '／', '或', '與', '及', '、', '|', '｜'):
            with self.subTest(separator=separator):
                self.assertEqual(repair_components(separator.join(['支承','支承墊','阻尼裝置'])), ['支承／支承墊／阻尼裝置'])
                self.assertEqual(repair_components(separator.join(['橋台基礎','橋墩基礎'])), ['橋台基礎','橋墩基礎'])
                self.assertEqual(repair_components(separator.join(['主梁','橫隔梁','橋面版'])), ['主梁','橫隔梁','橋面板'])
        self.assertEqual(repair_components('橋台或支承或支承墊或阻尼裝置'), ['橋台','支承／支承墊／阻尼裝置'])
        self.assertEqual(repair_components('翼牆或擋土牆'), ['翼牆／擋土牆'])
        self.assertEqual(repair_components('主梁或未完成命名'), [])
        self.assertEqual(repair_components('非橋台構件'), [])

    def test_hidden_whitespace_number_prefix_and_ids(self):
        self.assertEqual(repair_components('_x0009_06.伸縮縫'), ['伸縮縫'])
        self.assertEqual(repair_components('\ufeff\u200b\t防落設施'), ['防落設施'])
        self.assertEqual(repair_components('\t橋面版 '), ['橋面板'])
        self.assertEqual(repair_components('bearing'), ['支承／支承墊／阻尼裝置'])

    def test_import_v2_names_and_scope_without_cross_component_rules(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'v2.xlsx'
            workbook = Workbook()
            workbook.remove(workbook.active)
            for index, title in enumerate(V2_SHEETS):
                sheet = workbook.create_sheet(title)
                sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
                sheet.append(['其他損傷','測試樣態',f'測試工法{index}','處',index,None])
            workbook.save(path)
            rules = read_repairs(path)
            self.assertEqual(len(rules), 14)
            self.assertTrue(all(rule['automatic_ready'] for rule in rules))
            for index, (title, expected) in enumerate(V2_SHEETS.items()):
                with self.subTest(sheet=title):
                    self.assertEqual(rules[index]['components'], expected)
                    for component in expected:
                        self.assertIn(rules[index], component_repairs(rules, component))
            slab = component_repairs(rules, '橋面板')
            self.assertEqual([rule['method'] for rule in slab], ['測試工法12'])
            bearing = component_repairs(rules, '支承／支承墊／阻尼裝置')
            self.assertEqual([rule['method'] for rule in bearing], ['測試工法8'])
