import tempfile
from pathlib import Path
import unittest
from openpyxl import Workbook
from bridgefill.core.data import read_repairs, repair_components, repair_candidates
from bridgefill.core.policy import automatic_plan
from bridgefill.core.conditions import condition


def plant_rules():
    return [dict(pattern='草本植物生長', deterioration='其他損傷', method='清除植生', unit='處', price='101',
                 components=['橋台','翼牆／擋土牆'],
                 note='有關鍵字[草本植物]或[草本植生]即可使用，或僅有[植生]時，則U值為1優先使用'),
            dict(pattern='木本植物生長', deterioration='其他損傷', method='清除植生', unit='處', price='202',
                 components=['橋台','翼牆／擋土牆'],
                 note='有關鍵字[木本植物]或[木本植生]即可使用，或僅有[植生]時，若U值為2優先使用')]


class V3RuleTests(unittest.TestCase):
    def test_shared_abutment_wingwall_sheet_roundtrip(self):
        title = '橋台或翼牆或擋土牆'
        self.assertEqual(repair_components(title), ['橋台','翼牆／擋土牆'])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'v3.xlsx'
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = title
            sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註'])
            for rule in plant_rules():
                sheet.append([rule['deterioration'],rule['pattern'],rule['method'],rule['unit'],int(rule['price']),rule['note']])
            workbook.save(path)
            rules = read_repairs(path)
            for component in ('橋台','翼牆／擋土牆'):
                for urgency, price in [('1','101'),('2','202')]:
                    result = automatic_plan({'remark':'A1 植生。(2處)'}, component, [], rules, {'U':urgency})
                    self.assertFalse(result.review)
                    self.assertEqual(result.fields['price'], price)
            self.assertTrue(automatic_plan({'remark':'A1 植生。(2處)'}, '橋墩基礎', [], rules, {'U':'1'}).review)

    def test_explicit_plant_keyword_overrides_u_fallback_and_missing_u_is_manual(self):
        rules = plant_rules()
        for remark, urgency, price in [('草本植生','2','101'),('木本植物','1','202')]:
            result = automatic_plan({'remark':f'A1 {remark}。(2處)'}, '翼牆／擋土牆', [], rules, {'U':urgency})
            self.assertFalse(result.review)
            self.assertEqual(result.fields['price'], price)
        for context in ({}, {'U':'3'}, {'U':''}):
            self.assertTrue(automatic_plan({'remark':'A1 植生。(2處)'}, '翼牆／擋土牆', [], rules, context).review)

    def test_and_keywords_require_both_and_outrank_generic_paint_rule(self):
        generic = dict(pattern='塗裝剝落', method='一般塗裝', unit='平方公尺', price='10',note='有關鍵字[塗裝剝落]即可使用')
        bolts = dict(generic, pattern='螺栓塗裝剝落', method='螺栓塗裝', note='有關鍵字[螺栓]與[塗裝剝落]即可使用')
        self.assertEqual(repair_candidates([bolts], {'remark':'A1 螺栓。(1處)'}), [])
        self.assertEqual(repair_candidates([bolts], {'remark':'A1 塗裝剝落。(1m*2m)'}), [])
        result = automatic_plan({'remark':'A1 螺栓處塗裝剝落。(1m*2m)'}, '主梁', [], [generic,bolts])
        self.assertFalse(result.review)
        self.assertEqual(result.fields['method'], '螺栓塗裝')

    def test_fall_prevention_type_and_rust_spelling(self):
        rules = [dict(pattern=f'{name}生銹', method=name, unit='處',price='10',note=f'先判讀是否為[{name}]，再判讀關鍵字[生銹]或[生鏽]')
                 for name in ('防震拉桿','防落托架')]
        for name in ('防震拉桿','防落托架'):
            result = automatic_plan({'remark':'P1 生鏽。(2處)'}, '防落設施', [], rules, {'防落設施型式':name})
            self.assertFalse(result.review)
            self.assertEqual(result.fields['method'], name)
        self.assertTrue(automatic_plan({'remark':'P1 生鏽。(2處)'}, '防落設施', [], rules, {}).review)

    def test_crack_width_selects_rule_and_boundary_uses_upper_price(self):
        rules = [dict(pattern='混凝土裂縫', method='寬裂縫',unit='公尺',price='20', note='裂縫寬度大於[0.3mm]使用'),
                 dict(pattern='混凝土裂縫', method='細裂縫',unit='公尺',price='10', note='裂縫寬度小於[0.3mm]使用')]
        for width, expected in [('0.5','寬裂縫'),('0.3','寬裂縫'),('0.2','細裂縫')]:
            result = automatic_plan({'remark':f'A1 混凝土裂縫。(5m*{width}mm)*2'}, '翼牆／擋土牆', [], rules)
            self.assertFalse(result.review)
            self.assertEqual(result.fields['method'], expected)
            self.assertEqual(result.fields['amount'], '10')
        for size in ('5m','0.2mm','1m*2m'):
            self.assertTrue(automatic_plan({'remark':f'A1 混凝土裂縫。({size})'}, '翼牆／擋土牆', [], rules).review)

    def test_bearing_type_priority_selects_matching_pattern_not_cross_alias(self):
        rules = [dict(pattern='支承墊橡膠'+damage, method=method,unit='組',price='10',
                      note=prefix+'先判讀是否為[人造橡膠支承墊]，再判讀關鍵字[變形]或[開裂]')
                 for method,prefix in [('一般',''),('優先方案','優先使用，')] for damage in ('變形','開裂')]
        result = automatic_plan({'remark':'P1 支承表面橡膠開裂。(2組)'}, '支承／支承墊／阻尼裝置', [], rules, {'支承型式':'合成橡膠支承'})
        self.assertFalse(result.review)
        self.assertEqual(result.fields['method'], '優先方案')
        self.assertIn('開裂', result.fields['remark'])
        self.assertEqual(condition(rules[0], {'支承型式':'鋼支承'}, 'P1 變形。(2組)'), (False,False))

    def test_guardrail_height_abbreviated_note(self):
        rule = dict(note='優先選擇(護欄高大於0.85m小於1.1m適用)')
        self.assertEqual(condition(rule, {'橋護欄高度':'0.95'}), (True,False))
        self.assertEqual(condition(rule, {'橋護欄高度':'1.15'}), (False,False))
        self.assertEqual(condition(dict(note='護欄高≧0.85m適用'), {'橋護欄高度':'0.95'}), (None,True))
