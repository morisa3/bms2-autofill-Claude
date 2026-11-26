"""v3 三層整併規則表與 Excel 8 張設定工作表接線驗證。"""
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from bridgefill.core.data import component_repairs
from bridgefill.core.editing import load_rules_checked, rules_file
from bridgefill.core.excel_config import (
    clear_cache, component_classes, configure as configure_excel,
    deterioration_equivalents, exclusive_pairs, opinion_advice,
    opinion_excluded, opinion_names, param_value, related_groups, same_deterioration,
)
from bridgefill.core.fill import shape_remark
from bridgefill.core.matching import Defect, Pair, related, score
from bridgefill.core.opinion import generate
from bridgefill.core.policy import automatic_plan
from bridgefill.core.review import Review
from bridgefill.core.rule_add import add_rules
from bridgefill.core.rules_v2 import price_text, read_rules

ROOT = Path(__file__).resolve().parent.parent
V3_PATH = ROOT / 'data' / '劣化說明與維修單價v3.xlsx'


def _defect(component, remark, det='其他損傷', unit='處', deru='2222', location=''):
    fields = {'deterioration': det, 'd_location': location, 'remark': remark,
              'amount': '1', 'unit': unit, 'd': deru[0], 'e': deru[1], 'r': deru[2], 'u': deru[3]}
    return Defect('本期', (component, '1'), component, '', deru, det, fields, method_id='1001')


class V3RuleEquivalenceTests(unittest.TestCase):
    """確認 data/劣化說明與維修單價v3.xlsx 為預設規則表。"""

    def test_v3_file_exists_and_is_default(self):
        self.assertTrue(V3_PATH.exists())
        self.assertEqual(rules_file().name, '劣化說明與維修單價v3.xlsx')


class V3OverrideAndAddTests(unittest.TestCase):
    """確認構件工作表覆寫通用規則，且 rule_add.add_rules 可寫入 v3 工作表。"""

    def test_component_sheet_overrides_common_sheet(self):
        wb = Workbook()
        ws_common = wb.active
        ws_common.title = '通用規則'
        headers = ['劣化類型', '劣化樣態', '維修工法', '處單價', '平方公尺單價', '公尺單價', '立方公尺單價',
                   '式單價', '其他單位單價', '處理方式', '任一關鍵字', '適用構件', '劣化ID']
        ws_common.append(headers)
        ws_common.append(['其他損傷', '鋼筋未切', '切除塗防銹漆', 500, None, None, None, None, None,
                          '自動', '鋼筋未切', '主梁、橋台', 'G-01'])
        ws_beam = wb.create_sheet('主梁或橫隔梁或橋面版')
        ws_beam.append(headers)
        ws_beam.append(['其他損傷', '鋼筋未切', '高空作業切除塗防銹漆', 1200, None, None, None, None, None,
                        '自動', '鋼筋未切', '主梁', '1-01'])

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'override_v3.xlsx'
            wb.save(path)
            rules = read_rules(path)
            beam = component_repairs(rules, '主梁')
            abut = component_repairs(rules, '橋台')
            self.assertEqual(len(beam), 1)
            self.assertEqual(beam[0]['price'], '1200')
            self.assertEqual(beam[0]['method'], '高空作業切除塗防銹漆')
            self.assertEqual(len(abut), 1)
            self.assertEqual(abut[0]['price'], '500')

    def test_add_rules_writes_v3_unit_column(self):
        with tempfile.TemporaryDirectory() as folder:
            dst = Path(folder) / 'test_v3.xlsx'
            shutil.copy2(V3_PATH, dst)
            added = add_rules(dst, [{
                'component': '橋台', '劣化類型': '其他損傷', '劣化樣態': '測試新樣態',
                '維修工法': '測試工法', '維修單位': '平方公尺', '單價': '3200',
                '處理方式': '自動', '任一關鍵字': '測試新樣態', 'kind': '新劣化樣態', 'source': '本期網站值',
            }], backup_dir=Path(folder) / 'bak')
            self.assertEqual(len(added), 1)
            reloaded = read_rules(dst)
            found = [r for r in component_repairs(reloaded, '橋台') if r['pattern'] == '測試新樣態']
            self.assertEqual(len(found), 1)
            self.assertEqual(found[0]['unit'], '平方公尺')
            self.assertEqual(found[0]['price'], '3200')


class ExcelConfigSheetsTests(unittest.TestCase):
    """驗證 8 張原本未讀取的 Excel 設定工作表皆已接線生效。"""

    def tearDown(self):
        configure_excel(None)
        clear_cache()

    def test_loads_all_eight_config_sheets_from_v3(self):
        rules, err = load_rules_checked(V3_PATH)
        self.assertEqual(err, '')
        self.assertTrue(rules)
        self.assertIn(('鋼板', '鋼管'), exclusive_pairs())
        self.assertTrue(deterioration_equivalents())
        self.assertTrue(same_deterioration('防撞鋼板損傷', '保護設施損壞、移動或遺失', '橋墩／橋基保護設施'))
        self.assertTrue(same_deterioration('混凝土結構裂縫', '混凝土裂縫', '主梁'))
        self.assertTrue(related_groups())
        self.assertIn('主梁', component_classes()['main'])
        self.assertIn('橋面板', component_classes()['traffic'])
        self.assertEqual(opinion_advice()['u3'], '建議一年內派工修復')
        self.assertTrue(any(p == '漏水' for _, p, _ in opinion_excluded()))
        self.assertEqual(opinion_names().get('其他'), '')
        self.assertEqual(param_value('引道路堤預設長度', 99), 5.0)
        self.assertEqual(param_value('大面積判定 E 值', 99), 3.0)

    def test_custom_config_sheets_affect_matching_fill_and_opinion(self):
        wb = Workbook()
        ws = wb.active
        ws.title = '互斥樣態'
        ws.append(['說明'])
        ws.append(['樣態A', '樣態B', '說明'])
        ws.append(['木本', '草本', ''])
        ws.append(['焦油', '水泥', '自訂互斥'])

        ws_eq = wb.create_sheet('劣化類型等價')
        ws_eq.append(['說明'])
        ws_eq.append(['群組', '劣化類型', '適用構件', '說明'])
        ws_eq.append([1, '型態甲', '主梁', ''])
        ws_eq.append([1, '型態乙', '主梁', ''])

        ws_rel = wb.create_sheet('跨構件對應')
        ws_rel.append(['說明'])
        ws_rel.append(['構件', '相關構件', '適用樣態關鍵字', '說明'])
        ws_rel.append(['主梁', '支承／支承墊／阻尼裝置', '位移', '自訂跨構件'])

        ws_param = wb.create_sheet('參數')
        ws_param.append(['說明'])
        ws_param.append(['參數', '最小', '最大', '單位', '說明'])
        ws_param.append(['引道路堤預設長度', 8, 8, 'm', ''])
        ws_param.append(['大面積判定 E 值', 4, None, None, ''])

        ws_ex = wb.create_sheet('意見排除清單')
        ws_ex.append(['說明'])
        ws_ex.append(['構件', '樣態關鍵字', '適用U值', '說明'])
        ws_ex.append(['主梁', '微細痕跡', 2, ''])

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'custom_cfg.xlsx'
            wb.save(path)
            configure_excel(path)
            clear_cache()

            # 1. 自訂互斥樣態生效
            p1 = _defect('主梁', 'S1 焦油脫落。(1處)')
            c1 = _defect('主梁', 'S1 水泥脫落。(1處)')
            self.assertIsNone(score(p1, c1))

            # 2. 自訂劣化類型等價生效（獲得等價加分與理由）
            p2 = _defect('主梁', 'S1 混凝土裂縫。(1m)', det='型態甲', unit='公尺')
            c2 = _defect('主梁', 'S1 混凝土裂縫。(1m)', det='型態乙', unit='公尺')
            pts, reasons = score(p2, c2)
            self.assertGreaterEqual(pts, 75)
            self.assertTrue(any('等價' in r for r in reasons), reasons)

            # 3. 自訂跨構件對應生效
            self.assertTrue(related('主梁', '支承／支承墊／阻尼裝置', '位移', '位移'))

            # 4. 自訂引道路堤預設長度=8m 生效
            shaped = shape_remark({'remark': 'A1 AC龜裂', 'unit': '平方公尺'}, '引道路堤', {'_最小淨寬': '6'})
            self.assertIn('(6m*8m)', shaped['remark'])

            # 5. 自訂意見排除清單 + 大面積 E=4 生效
            pair_excl = Pair('新增', None, _defect('主梁', 'S1 微細痕跡。(1處)', deru='2222'), confidence='高')
            pair_e3 = Pair('新增', None, _defect('主梁', 'S1 混凝土裂縫。(1處)', deru='3322'), confidence='高')
            items = generate([pair_excl, pair_e3], Review().sync([pair_excl, pair_e3]), {}, {})
            defect_texts = [i.text for i in items if i.kind == 'defect']
            self.assertFalse(any('微細痕跡' in t for t in defect_texts))
            self.assertIn('S1主梁混凝土裂縫，建議持續追蹤。', defect_texts)   # 大面積門檻改 4，故 E=3 仍持續追蹤


if __name__ == '__main__':
    unittest.main()
