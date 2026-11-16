"""第 7 步檢測員意見：建議文字、合併、排除、標準文字（護岸與基礎合併一段）、淨高。"""
import unittest

from bridgefill.core.matching import Defect, Pair
from bridgefill.core.opinion import (TEMPLATES, absorb, apply_choices, blanks, choices, compose, compress, generate,
                                     heights, locations, unresolved)
from bridgefill.core.review import Review

_count = [0]


def make(component, remark, deru='2222', location='', prior_remark=None, kind='其他損傷'):
    _count[0] += 1
    fields = {'deterioration': kind, 'd_location': location, 'remark': remark, 'amount': '1', 'unit': '處',
              'd': deru[0], 'e': deru[1], 'r': deru[2], 'u': deru[3]}
    current = Defect('本期', (component, str(_count[0])), component, '', deru, kind, fields,
                     method_id=str(9000 + _count[0]))
    prior = None
    if prior_remark:
        prior = Defect('前期', (component, str(_count[0])), component, '', deru, kind,
                       dict(fields, remark=prior_remark), der_id=str(500 + _count[0]))
    return Pair('沿用' if prior else '新增', prior, current, confidence='高')


def texts(pairs, **header):
    return [(i.text, i.checked) for i in generate(pairs, Review().sync(pairs), header, {})]


class OpinionTests(unittest.TestCase):
    def test_suggestion_by_u_traffic_and_main_component(self):
        result = dict(texts([make('橋台', 'A1-下游 混凝土剝落鋼筋外露。(1處)', '3223'),
                             make('橋面板', 'S1-L 混凝土剝落鋼筋外露。(1處)', '3222'),
                             make('主梁', '1G1-0,13 混凝土剝落鋼筋外露。(1處)', '3322'),
                             make('橋台', 'A2 混凝土剝落鋼筋外露。(1處)', '3222')]))
        self.assertIn('A1橋台混凝土剝落鋼筋外露，建議一年內派工修復。', result)
        self.assertIn('S1橋面板、S1主梁混凝土剝落鋼筋外露，建議三年內派工修復。', result)   # 橋面、主構件 E3
        self.assertIn('A2橋台混凝土剝落鋼筋外露，建議持續追蹤。', result)

    def test_merge_compress_exclude_and_minor(self):
        result = texts([make('主梁', f'{n}G1-1 混凝土裂縫。(1處)', location=f'{n}G1') for n in (1, 2, 3)]
                       + [make('伸縮縫', 'A1 漏水。(9m)'), make('橋台基礎', 'A1 混凝土剝落。(1處)'),
                          make('橋台', 'A1 混凝土剝落。(1處)', '2222'), make('橋墩', 'P1 滲水及白華(1處)'),
                          make('橋墩', 'P2 滲水白華(1處)'), make('橋面板', 'S1 混凝土剝落。(1處)', '2221')])
        lines = dict(result)
        self.assertIn('S1~S3主梁混凝土裂縫，建議持續追蹤。', lines)
        self.assertFalse([t for t in lines if '漏水' in t and '伸縮縫' in t])               # 伸縮縫漏水不寫
        self.assertIn('P1、P2橋墩滲水白華，建議持續追蹤。', lines)
        self.assertFalse(lines['A1橋台混凝土剝落，建議持續追蹤。'])                       # 小面積 U2 預設不勾
        self.assertFalse([t for t in lines if t.startswith('S1橋面板')])                  # U1 不寫

    def test_river_and_foundation_share_one_suggestion(self):
        pairs = [make('河道', '上游 左岸護岸掏空。(10m)', '3323'),
                 make('橋墩基礎', 'P1 橋墩基礎裸露。(2m*3m*0.8m)', '3323', prior_remark='P1 橋墩基礎裸露。(2m*3m*0.5m)'),
                 make('橋台基礎', 'A1 橋台基礎掏空。(1處)', '3323'),
                 make('主梁', 'S1G1 混凝土裂縫。(1處)')]
        items = generate(pairs, Review().sync(pairs), {}, {})
        water = [i.text for i in items if i.checked and ('基礎' in i.text or '護岸' in i.text)]
        self.assertIn('合併', next(i.note for i in items if i.text == water[0]))
        self.assertEqual(len([i for i in items if i.kind == 'standard']), 5)          # 2、3 合併
        self.assertEqual(len(water), 1)
        text = water[0]
        self.assertTrue(text.startswith('本橋護岸發生掏空情形，位於上游左岸約【＿m】處；P1、A1基礎發生裸露、掏空情形'))
        self.assertIn('本期裸露深度為P1=0.8m、【A1本期深度】', text)
        self.assertIn('前期裸露深度為P1=0.5m', text)
        self.assertIn('P1有刷深之情況', text)
        self.assertEqual(text.count('建議'), 1)
        self.assertTrue(text.endswith('辦理詳細檢測／定期觀測河床變化】。'))
        self.assertFalse([i for i in items if i.kind == 'defect' and ('裸露' in i.text or '掏空' in i.text)])

    def test_choices_and_compose(self):
        pairs = [make('橋護欄', 'S1-L,R 護欄高度不足。(45m)*2')]
        items = generate(pairs, Review().sync(pairs), {'道路等級': '鄉道'}, {})
        text = compose(items, '橋下淨高：H=3m')
        self.assertTrue(text.startswith('1. 本橋道路等級屬鄉道，且現況【有／無】人行道，橋護欄高度不足110cm'))
        self.assertFalse([t for t in text.splitlines() if '護欄高度不足。' in t])   # 已由標準文字敘明
        options = choices(text)
        self.assertEqual(options['【加高／重作】'], ['加高', '重作'])
        done = apply_choices(text, {'【有／無】': '無', '【加高／重作】': '加高'})
        self.assertIn('現況無人行道', done)
        self.assertEqual(unresolved(done), [])
        self.assertTrue(done.endswith('橋下淨高：H=3m'))

    def test_six_standard_texts_always_listed(self):
        pairs = [make('主梁', 'S1G1 混凝土裂縫。(1處)')]
        items = generate(pairs, Review().sync(pairs), {}, {})
        standard = [i for i in items if i.kind == 'standard']
        self.assertEqual([i.text for i in standard], list(TEMPLATES))
        self.assertFalse(any(i.checked for i in standard))
        self.assertTrue(all('未偵測到' in i.note for i in standard))
        joint = generate([make('伸縮縫', 'A1 伸縮縫高低差。(1處)', '3333')], Review(), {}, {})[3]
        self.assertTrue(joint.checked)
        self.assertTrue(joint.text.startswith('本橋A1伸縮縫具高低落差'))

    def test_joint_locations_detected_or_offered(self):
        def joint(pairs, **header):
            return generate(pairs, Review().sync(pairs), header, {})[3].text
        self.assertTrue(joint([make('伸縮縫', 'A2 伸縮縫間隙異常。(1處)', '3333'), make('伸縮縫', 'A1 漏水。(9m)')])
                        .startswith('本橋A2伸縮縫間距過大'))                     # 有狀況的那一道
        self.assertTrue(joint([make('伸縮縫', 'P1 漏水。(9m)', '2221')]).startswith('本橋P1伸縮縫【'))   # 有缺失的伸縮縫
        self.assertTrue(joint([make('主梁', 'S1G1 裂縫。(1處)')], 總橋孔數='3')
                        .startswith('本橋【A1／A2／P1／P2】伸縮縫'))               # 依橋孔數給選項

    def test_other_uses_facility_category(self):
        pair = make('其他', 'S1-L,R 植生積土。(16m*0.2m)', '2122')
        pair.current.context = {'設施類別': '鋪面', '設施編號': '08'}
        lines = [t for t, _ in texts([pair])]
        self.assertIn('S1鋪面植生積土，建議三年內派工修復。', lines)
        self.assertEqual(pair.display, '其他（鋪面）')

    def test_blanks_fill_and_brackets_disappear(self):
        text = TEMPLATES[1]
        self.assertEqual(blanks(text), ['【＿m】'])
        done = apply_choices(text, {'【＿m】': '15m', '【護岸／護床工】': '護岸'})
        self.assertIn('本橋護岸發生', done)
        self.assertIn('約15m處', done)

    def test_manual_edits_survive_toggling(self):
        pairs = [make('主梁', 'S1G1 混凝土裂縫。(1處)'), make('橋台', 'A1 混凝土剝落鋼筋外露。(1處)')]
        items = generate(pairs, Review().sync(pairs), {}, {})
        text = compose(items, '梁下淨高：H=3m')
        lines = text.splitlines()
        lines[0] = '1. 手改的第一條'
        rest = absorb(items, '\n'.join(lines + ['3. 新增的一條']))
        self.assertEqual(rest, '梁下淨高：H=3m')
        defects = [i for i in items if i.checked]
        self.assertEqual(defects[0].custom, '手改的第一條')
        self.assertEqual(defects[-1].kind, 'manual')
        defects[1].checked = False                                        # 取消勾選：同步刪除，手改保留
        again = compose(items, rest)
        self.assertEqual(again, '1. 手改的第一條\n2. 新增的一條\n梁下淨高：H=3m')

    def test_heights_keep_current_then_prior(self):
        line = heights({'檢測員意見': 'P1=3.35', '總橋孔數': '3'}, {'檢測員意見': '橋下淨高：P1=3.85m、P2=5.15m'}, True)
        self.assertEqual(line, '梁下淨高：P1=3.35m、P2=5.15m')
        self.assertEqual(heights({'總橋孔數': '1'}, {}, False), '橋下淨高：H=【m】')

    def test_locations(self):
        pair = make('主梁', '1G1-0,13,26 混凝土剝落。(1處)')
        self.assertEqual(locations(Review().sync([pair]) and type('E', (), {'fields': pair.current.fields,
                                                                             'remark': pair.current.remark})()),
                         ['S1'])
        self.assertEqual(compress(['S1', 'S2', 'S3', 'P1', 'P3']), 'S1~S3、P1、P3')


if __name__ == '__main__':
    unittest.main()
