import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'tools'))

from openpyxl import load_workbook  # noqa: E402
from excel_v2 import (ADDITIONS, BOX_CULVERT_TYPES, EDITS, NEW_SHEETS, add_new_sheets, apply_edits,  # noqa: E402
                      convert, parse_note, template)
from bridgefill.core.conditions import condition, keyword_score  # noqa: E402
from bridgefill.core.remarks import prepare_fields  # noqa: E402
from bridgefill.core.data import read_repairs  # noqa: E402
from bridgefill.core.rules_v2 import read_rules  # noqa: E402

SOURCE = ROOT / 'data' / '劣化說明與維修單價1111v1.xlsx'


class ParseNoteTests(unittest.TestCase):
    def parse(self, note, unit='處', pattern='樣態'):
        fields, review, remark = parse_note(note, unit, pattern)
        return {k: v for k, v in fields.items() if v not in ('', None)}, review, remark

    def test_any_and_required_keywords(self):
        fields, review, _ = self.parse('有關鍵字[碎裂]或[破碎]與[鋼筋外露]即可使用')
        self.assertEqual(fields['任一關鍵字'], '碎裂、破碎')
        self.assertEqual(fields['必要關鍵字'], '鋼筋外露')
        self.assertFalse(review)

    def test_all_required_keywords(self):
        fields, _, _ = self.parse('有關鍵字[破損]與[鋼筋外露]即可使用')
        self.assertEqual(fields['必要關鍵字'], '破損、鋼筋外露')
        self.assertNotIn('任一關鍵字', fields)

    def test_crack_width_boundary_follows_v0976(self):
        wide, _, remark = self.parse('裂縫寬度大於[0.3mm]使用', '公尺', '混凝土裂縫')
        narrow, _, _ = self.parse('裂縫寬度小於[0.3mm]使用', '公尺', '混凝土裂縫')
        self.assertEqual((wide['條件欄位'], wide['最小值'], wide['條件單位']), ('裂縫寬度', 0.3, 'mm'))
        self.assertEqual(narrow['最大值'], 0.3)
        self.assertIn('0.3mm 歸入', remark)
        self.assertEqual(wide['備註範本'], '{位置} {樣態}。(mm*m)*1')

    def test_alone_keyword_with_u_value(self):
        fields, _, _ = self.parse('有關鍵字[木本植物]或[木本植生]即可使用，或僅有關鍵字[植生]時，若U值為2優先使用')
        self.assertEqual((fields['單獨關鍵字'], fields['單獨關鍵字U值']), ('植生', 2))

    def test_type_correction_and_precheck(self):
        fields, _, _ = self.parse('若判讀到關鍵字[生銹]或[生鏽]即可使用，但劣化類型填成[其他損傷]，則修正成[防撞鋼板損傷]後使用')
        self.assertEqual(fields['修正劣化類型為'], '防撞鋼板損傷')
        fields, _, _ = self.parse('優先使用，先判讀是否為[人造橡膠支承墊]，再判讀關鍵字[變形]或[開裂]')
        self.assertEqual((fields['前置判讀'], fields['任一關鍵字'], fields['優先序']), ('人造橡膠支承墊', '變形、開裂', 1))

    def test_unparseable_conditions_need_review(self):
        _, review, remark = self.parse('若傷及混凝土層時使用，有關鍵字[坑洞]即可使用')
        self.assertTrue(review)
        self.assertIn('傷及混凝土層', remark)
        _, review, _ = self.parse('手動選擇，或讀取到護欄高度大於0.60小於0.85時使用', '公尺', '護欄高度不足')
        self.assertTrue(review)

    def test_templates_by_unit(self):
        self.assertEqual(template('平方公尺', 'AC剝離'), '{位置} {樣態}。(m*m)*1')
        self.assertEqual(template('立方公尺', '路面坑洞'), '{位置} {樣態}。(m*m*m)*1')
        self.assertEqual(template('公尺', '漏水'), '{位置} {樣態}。(m)*1')
        self.assertEqual(template('式', '基礎裸露'), '')


GUARD_NOTES = {
    '2-55': '讀取到護欄高度小於0.30m時使用',
    '2-40': '優先選擇，或讀取到護欄高度大於等於0.85m小於1.10m時使用',
    '2-41': '讀取到護欄高度大於等於0.60m小於0.85m時使用',
    '2-42': '讀取到護欄高度大於等於0.30m小於0.60m時使用',
}


class ConvertTests(unittest.TestCase):
    def test_v2_keeps_the_confirmed_edits_and_drops_the_old_note(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            stats = convert(SOURCE, target)
            self.assertEqual((stats['rows'], stats['review']), (593, 17))   # 含箱涵 63、護床工 17、主梁劣化類型複製 18 筆
            headers = [c.value for c in load_workbook(target)['橋台或翼牆或擋土牆'][1]]
            self.assertNotIn('備註', headers)
            self.assertNotIn('轉換狀態', headers)
            self.assertEqual(headers[-1], '劣化ID')
            ids = {r['rule_id']: r for r in read_rules(target)}
            self.assertEqual(ids['1-05']['pattern'], '路面混凝土層破洞')
            for rule_id in ('2-54', '2-55'):
                self.assertEqual((ids[rule_id]['method'], ids[rule_id]['unit'], ids[rule_id]['price']),
                                 ('紐澤西護欄施作＋單管護欄加高', '公尺', '9900'))

    def test_old_pothole_remark_is_rewritten_to_new_pattern(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            convert(SOURCE, target)
            rule = {r['rule_id']: r for r in read_rules(target)}['1-05']
        self.assertIsNotNone(keyword_score(rule['note'], 'A1 路面坑洞。(1m*1m*0.1m)'))
        fields = dict(deterioration='路堤損傷', d_location='A1', method='', amount='0.1', unit='立方公尺',
                      price='', remark='A1 路面坑洞。(1m*1m*0.1m)')
        self.assertEqual(prepare_fields(fields, rule, force_pattern=True, component='引道路堤')['remark'],
                         'A1 路面混凝土層破洞。(1m*1m*0.1m)')

    def test_no_guardrail_does_not_collide_with_2_28(self):
        self.assertIsNone(keyword_score('有關鍵字[未設置]即可使用', 'S1-L 無設置護欄。(10m)'))
        self.assertIsNone(keyword_score('有關鍵字[無設置護欄]或[未設置護欄]或[無護欄]即可使用', 'S1-L 護欄未設置。(10m)'))

    def test_guardrail_edits_match_user_ranges(self):
        notes = {rule_id: changes['備註'] for _, rule_id, changes, _ in EDITS if rule_id in GUARD_NOTES}
        notes.update({values[6]: values[5] for _, _, values, _, _ in ADDITIONS if values[6] in GUARD_NOTES})
        self.assertEqual(notes, GUARD_NOTES)
        expected = {'0.10': '2-55', '0.29': '2-55', '0.30': '2-42', '0.59': '2-42', '0.60': '2-41', '0.84': '2-41',
                    '0.85': '2-40', '1.09': '2-40', '1.10': None}
        for height, rule_id in expected.items():
            chosen = [k for k, note in GUARD_NOTES.items()
                      if condition({'note': note}, {'護欄高度(m)': height})[0]]
            self.assertEqual(chosen, [rule_id] if rule_id else [], height)
        for note in GUARD_NOTES.values():
            self.assertEqual(condition({'note': note}, {}), (None, True))  # 讀不到高度轉人工
            fields, review, _ = parse_note(note, '公尺', '護欄高度不足')
            self.assertFalse(review)
            self.assertEqual(fields['條件欄位'], '護欄高度')


class BoxCulvertTests(unittest.TestCase):
    def test_box_culvert_rules_use_site_types_and_copied_prices(self):
        wb = load_workbook(SOURCE)
        add_new_sheets(wb)
        for _, title, prefix, rows in NEW_SHEETS:
            sheet = [r for r in wb[title].iter_rows(min_row=2, values_only=True) if r[6]]
            self.assertEqual(len(sheet), len(rows))
            if title in BOX_CULVERT_TYPES:
                self.assertTrue(all(r[0] in BOX_CULVERT_TYPES[title] for r in sheet), title)
            self.assertTrue(all(str(r[6]).startswith(prefix + '-') for r in sheet))
        honeycomb = next(r for r in wb['端牆或中隔板'].iter_rows(min_row=2, values_only=True) if r[6] == '15-11')
        self.assertEqual((honeycomb[0], honeycomb[3], honeycomb[4]), ('混凝土空洞(或蜂窩)', '平方公尺', 3200))  # 截圖 3200

    def test_girder_types_split_by_component(self):
        from bridgefill.core.data import component_repairs
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            convert(SOURCE, target)
            rules = read_rules(target)
        concrete = lambda c: {r['deterioration'] for r in component_repairs(rules, c) if '剝落' in r['deterioration']}  # noqa: E731
        self.assertEqual(concrete('主梁'), {'混凝土剝落、破碎、鋼筋、鋼腱或錨碇外露、銹蝕', '混凝土剝落、鋼筋外露、銹蝕'})
        self.assertEqual(concrete('橫隔梁'), {'混凝土剝落、破碎、鋼筋外露、銹蝕'})
        self.assertEqual(concrete('橋面板'), {'混凝土剝落、破碎、鋼筋外露、銹蝕'})
        main = [r for r in component_repairs(rules, '主梁') if r['deterioration'] == '混凝土剝落、鋼筋外露、銹蝕']
        self.assertEqual(len(main), 17)
        self.assertEqual({r['method'] for r in main if r['pattern'] == '混凝土剝落鋼筋外露'},
                         {'混凝土修復，鋼筋銹蝕處理，水刀鑿除'})

    def test_bed_protection_sheet(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            convert(SOURCE, target)
            rules = [r for r in read_rules(target) if '護床工' in r['components']]
        self.assertEqual(len(rules), 17)
        ids = {r['rule_id']: r for r in rules}
        self.assertEqual((ids['18-07']['pattern'], ids['18-07']['unit']), ('護床工掏空', '立方公尺'))

    def test_reader_maps_box_culvert_sheets_to_components(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / 'v2.xlsx'
            convert(SOURCE, target)
            rules = read_rules(target)
        self.assertTrue({'端牆', '中隔板', '頂板', '底板'} <= {c for r in rules for c in r['components']})
        ids = {r['rule_id']: r for r in rules}
        self.assertEqual((ids['17-11']['pattern'], ids['17-19']['pattern']), ('底板沖刷裸露', '底板沖刷裸露'))
        self.assertIn('石籠', ids['17-19']['method'])
        self.assertIn('灌漿回填', ids['17-12']['method'])
        self.assertEqual(ids['17-13']['price'], '3200')


if __name__ == '__main__':
    unittest.main()
