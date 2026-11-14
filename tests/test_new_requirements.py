import tempfile
from pathlib import Path
import unittest
from openpyxl import Workbook
from bridgefill.core.data import read_repairs, History
from bridgefill.core.policy import automatic_plan
from bridgefill.core.remarks import parse


class RemarkTests(unittest.TestCase):
    def test_dimensions_counts_and_cracks(self):
        cases = [
            ('A1 AC剝離 （5m＊5m）', 'AC剝離', '平方公尺', '25', 'A1 AC剝離。(5m*5m)'),
            ('A1AC剝離。 （5*5）', 'AC剝離', '平方公尺', '25', 'A1 AC剝離。(5*5)'),
            ('P1 混凝土裂縫 。 （5m＊0.3mm） ＊ 4', '混凝土裂縫', '公尺', '20', 'P1 混凝土裂縫。(0.3mm*5m)*4'),
            ('P1 裂縫。(0.3mm*5m*4)', '裂縫', '公尺', '20', 'P1 裂縫。(0.3mm*5m)*4'),
            ('P1 裂縫。(5m*0.3mm*4)*2', '裂縫', '公尺', '40', 'P1 裂縫。(0.3mm*5m)*8'),
            ('A1 掏空。(2m*3m*4m)*2', '掏空', '立方公尺', '48', 'A1 掏空。(2m*3m*4m)*2'),
            ('A1 剝落。(50cm*2m)*3', '剝落', '平方公尺', '3', 'A1 剝落。(50cm*2m)*3'),
            ('A1 植生。(1處)', '植生', '處', '1', 'A1 植生。(1處)'),
            ('A1 排水。(2式)', '排水', '式', '2', 'A1 排水。(2式)'),
            ('A1 脫落。(5m)*4', '脫落', '公尺', '20', 'A1 脫落。(5m)*4'),
        ]
        for text, pattern, unit, amount, expected in cases:
            with self.subTest(text=text):
                result = parse(text, pattern, unit)
                self.assertEqual((result.text, result.amount, result.unit), (expected, amount, unit))

    def test_unknown_formats_do_not_compute(self):
        for text in ('A1 無法判讀', 'A1缺失(5m)', 'A1 損傷。(1+2)', 'A1 損傷。(__import__(os))',
                     'A1 損傷。(5m*2*3)', 'A1 損傷。(1m*2處)', 'A1 損傷。(2m)\nA2 損傷。(3m)'):
            with self.subTest(text=text):
                with self.assertRaises(ValueError):
                    parse(text, '損傷', '公尺')


class NewPolicyTests(unittest.TestCase):
    def setUp(self):
        self.rule = dict(pattern='AC剝離', method='新工法', unit='平方公尺', price='350',
                         deterioration='路堤損傷', component='引道路堤', note='若過去資料為[AC老化]，則優先填寫AC剝離')
        self.current = dict(deterioration='路堤損傷', d_location='錯誤位置', amount='1', unit='處',
                            method='舊工法', price='300', remark='A1AC老化 （5m＊5m）')

    def test_new_remark_overwrites_location_amount_type_and_repair(self):
        plan = automatic_plan(self.current, '引道路堤', [], [self.rule])
        self.assertFalse(plan.review)
        self.assertEqual(plan.fields, dict(deterioration='路堤損傷', d_location='A1', amount='25', unit='平方公尺',
                         method='新工法', price='350', remark='A1 AC剝離。(5m*5m)'))

    def test_area_and_count_choose_rule_by_expression_not_old_unit(self):
        rule = dict(self.rule, pattern='積土植生', deterioration='植物生長', note='有[植生]即可使用')
        rules = [rule, dict(rule, unit='處', price='100')]
        for remark, expected in [('A1 積土植生。(2處)', '處'), ('A1 植生。(2m*3m)', '平方公尺')]:
            plan = automatic_plan(dict(self.current, remark=remark, deterioration='植物生長'), '引道路堤', [], rules)
            self.assertFalse(plan.review)
            self.assertEqual(plan.fields['unit'], expected)
            self.assertEqual(plan.fields['remark'].count('積土植生'), 1)

    def test_history_requires_same_type_quantity_and_unique_position(self):
        old = History('引道路堤', dict(self.current, deterioration='路堤損傷'), 3)
        current = dict(deterioration='路堤損傷', amount='1')
        plan = automatic_plan(current, '引道路堤', [old], [self.rule])
        self.assertFalse(plan.review)
        self.assertEqual(plan.fields['amount'], '25')
        for fields in (dict(current, amount='2'), dict(current, deterioration='植物生長')):
            self.assertTrue(automatic_plan(fields, '引道路堤', [old], [self.rule]).review)
        self.assertTrue(automatic_plan(current, '引道路堤', [old, old], [self.rule]).review)

    def test_generic_keyword_replaces_whole_old_pattern_without_duplicate_prefix(self):
        rule = dict(self.rule, pattern='AC表面破損', note='有[破損]字樣即可替換')
        plan = automatic_plan(dict(self.current, remark='A1 路面破損。(2m*3m)'), '引道路堤', [], [rule])
        self.assertFalse(plan.review)
        self.assertEqual(plan.fields['remark'], 'A1 AC表面破損。(2m*3m)')

    def test_extra_work_method_condition_is_not_assumed_from_keyword(self):
        rule = dict(self.rule, pattern='路面坑洞', unit='立方公尺',
                    note='若傷及混凝土層時使用，關鍵字[坑洞]')
        plan = automatic_plan(dict(self.current, remark='A1 路面坑洞。(2m*3m*0.2m)'), '引道路堤', [], [rule])
        self.assertTrue(plan.review)
        self.assertIn('使用條件', plan.reason)

    def test_rules_cannot_cross_component_or_use_incomplete_sheet(self):
        self.assertTrue(automatic_plan(self.current, '橋台', [], [self.rule]).review)
        self.assertTrue(automatic_plan(self.current, '引道路堤', [], [dict(self.rule, automatic_ready=False)]).review)

    def test_invalid_dimension_or_conflicting_repairs_pause(self):
        self.assertTrue(automatic_plan(dict(self.current, remark='A1 AC剝離。(1處)'), '引道路堤', [], [self.rule]).review)
        self.assertTrue(automatic_plan(self.current, '引道路堤', [], [self.rule, dict(self.rule, price='400')]).review)

    def test_import_six_columns_skips_missing_values_and_keeps_scope(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'rules.xlsx'
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '引道路堤'
            sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
            sheet.append(['路堤損傷','AC剝離','修復 ','平方公尺',350,'[AC老化]優先替換'])
            sheet.append(['路堤損傷','AC車轍','修復','平方公尺',None,None])
            other = workbook.copy_worksheet(sheet)
            other.title = '橋台'
            workbook.save(path)
            rules = read_repairs(path)
            self.assertEqual(len(rules), 2)
            self.assertEqual(rules[0]['component'], '引道路堤')
            self.assertEqual(rules[0]['deterioration'], '路堤損傷')
            self.assertEqual(rules[0]['price'], '350')
            self.assertEqual(rules[0]['method'], '修復')
            self.assertTrue(rules[0]['automatic_ready'])
            self.assertTrue(rules[1]['automatic_ready'])


class AllComponentImportTests(unittest.TestCase):
    def test_every_registered_component_uses_same_excel_schema_and_policy(self):
        from bridgefill.core.config import COMPONENTS
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'all.xlsx'
            workbook = Workbook()
            workbook.remove(workbook.active)
            for index, component in enumerate(COMPONENTS):
                sheet = workbook.create_sheet(component)
                sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
                sheet.append(['其他損傷','缺失','工法'+str(index),'處',index,None])
            workbook.save(path)
            rules = read_repairs(path)
            self.assertEqual(len(rules), len(COMPONENTS))
            for index, component in enumerate(COMPONENTS):
                with self.subTest(component=component):
                    plan = automatic_plan({'remark':'A1 缺失。(1處)'}, component, [], rules)
                    self.assertFalse(plan.review)
                    self.assertEqual(plan.fields['method'], '工法'+str(index))
                    self.assertEqual(plan.fields['price'], str(index))

    def test_shared_sheet_number_prefix_ids_aliases_and_unknown_scope(self):
        from bridgefill.core.data import repair_components
        self.assertEqual(repair_components('01.引道路堤護欄與橋護欄'), ['引道路堤護欄','橋護欄'])
        self.assertEqual(repair_components('expansion'), ['伸縮縫'])
        self.assertEqual(repair_components('引道路堤保護設施'), ['引道道路堤保護措施'])
        self.assertEqual(repair_components('橋墩 / 帽梁'), ['墩柱與帽梁'])
        self.assertEqual(repair_components('引道路堤與未定構件'), [])

    def test_shared_sheet_is_scoped_to_both_components_only(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'shared.xlsx'
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '引道路堤護欄與橋護欄'
            sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
            sheet.append(['其他損傷','破損','修復','處',500,None])
            workbook.save(path)
            rules = read_repairs(path)
            for component in ('引道路堤護欄', '橋護欄'):
                self.assertFalse(automatic_plan({'remark':'A1 破損。(1處)'}, component, [], rules).review)
            self.assertTrue(automatic_plan({'remark':'A1 破損。(1處)'}, '橋台', [], rules).review)

    def test_missing_type_or_unknown_sheet_does_not_become_global_rule(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'incomplete.xlsx'
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '橋台'
            sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
            sheet.append([None,'破損','修復','處',500,None])
            other = workbook.copy_worksheet(sheet)
            other.title = '尚未命名'
            other.cell(2,1).value = '其他損傷'
            workbook.save(path)
            rules = read_repairs(path)
            self.assertEqual(len(rules), 2)
            self.assertTrue(all(not rule['automatic_ready'] for rule in rules))
            self.assertTrue(automatic_plan({'remark':'A1 破損。(1處)'}, '橋台', [], rules).review)
