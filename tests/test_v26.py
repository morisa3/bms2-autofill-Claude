"""2.6.000.0（2026-11-20 使用者需求）：規則表 1120、鋼筋外露的銹蝕關鍵字、數量不用科學記號、備註補單位／位置／尺寸範本、
引道路堤 AC 依淨寬估算、多組帳號密碼、歷史紀錄縣市、照片簿最新日期。"""
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from bridgefill.core.bridge_photos import PhotoPlan, latest_dates, mark_uploaded
from bridgefill.core.data import quantity
from bridgefill.core.editing import load_rules
from bridgefill.core.fill import proposal_for, shape_remark
from bridgefill.core.matching import Defect, Pair
from bridgefill.core.review import county_of
from bridgefill.core.rules_v2 import synthesize_note
from bridgefill.core.conditions import keyword_groups
from bridgefill.ui import credentials

RULES_PATH = Path(__file__).resolve().parents[1] / 'data' / '劣化說明與維修單價1120.xlsx'
ABUTMENT = '混凝土剝落、破碎、鋼筋外露、銹蝕'


def rules():
    found = load_rules(RULES_PATH)
    return found[0] if isinstance(found, tuple) else found


def decide(remark, kind, component='橋台', prior=None, context=None, unit='', amount=''):
    fields = {'deterioration': kind, 'd_location': '', 'method': '', 'amount': amount, 'unit': unit, 'price': '',
              'remark': remark, 'd': '2', 'e': '2', 'r': '2', 'u': '2'}
    current = Defect('本期', ('x', '1'), component, '', '2222', kind, fields, method_id='1', der_id='1',
                     inspect_id='1', context=dict(context or {'位置': '橋頭'}))
    old = None
    if prior:
        old = Defect('前期', ('x', '1'), component, '', '2222', kind, dict(fields, **prior), context=dict(current.context))
    return proposal_for(Pair('相同' if old else '新增', prior=old, current=current, confidence='高'), RULES)


RULES = rules()


class RustKeywordTests(unittest.TestCase):
    def test_screenshot_case_keeps_exposed_rebar(self):
        """前期「A1-下游背牆 混凝土剝落，鋼筋外露 (1處)」、本期空白 → 混凝土剝落鋼筋外露 800（不再判成混凝土剝落 500）。"""
        prior = dict(remark='A1-下游背牆 混凝土剝落，鋼筋外露 (1處)', method='混凝土修復，鋼筋銹蝕處理，水刀鑿除',
                     unit='處', amount='1', price='800', d_location='A1-下游背牆')
        result = decide('', ABUTMENT, prior=prior)
        self.assertFalse(result.review)
        self.assertEqual((result.fields['remark'], result.fields['price']), ('A1-下游背牆 混凝土剝落鋼筋外露。(1處)', '800'))

    def test_rust_words_mean_exposed_rebar_on_every_component(self):
        for component, remark, expected in (('橋台', 'A1 混凝土破損銹蝕。(1處)', '混凝土破損鋼筋外露'),
                                            ('橋台', 'A1 混凝土蜂窩鏽蝕。(1處)', '混凝土蜂窩鋼筋外露'),
                                            ('墩柱與帽梁', 'P1 混凝土剝落銹蝕。(1處)', '混凝土剝落鋼筋外露'),
                                            ('主梁', 'S1/G1 混凝土剝落鏽蝕。(1處)', '混凝土剝落鋼筋外露')):
            with self.subTest(component=component, remark=remark):
                kind = ABUTMENT if component in ('橋台', '墩柱與帽梁') else '混凝土剝落、鋼筋外露、銹蝕'
                result = decide(remark, kind, component)
                self.assertIn(expected, result.fields['remark'])
                self.assertEqual(result.fields['price'], '800')

    def test_rust_word_stands_alone_with_required_words(self):
        note = synthesize_note({'任一關鍵字': '破損銹蝕、破損鏽蝕', '必要關鍵字': '破損、鋼筋外露'})
        self.assertEqual(keyword_groups(note), [[('破損',), ('鋼筋外露',)], [('破損銹蝕',)], [('破損鏽蝕',)]])

    def test_prior_keyword_rows_keep_their_keywords(self):
        self.assertEqual(synthesize_note({'任一關鍵字': '車轍', '前期關鍵字': 'AC老化', '優先序': 2}), '有關鍵字[車轍]即可使用')
        self.assertIn('優先使用', synthesize_note({'任一關鍵字': '剝離', '前期關鍵字': 'AC老化', '優先序': 1}))

    def test_equivalent_rules_are_not_ambiguous(self):
        """1-03（路堤損傷 AC龜裂）與 1-12（其他損傷 AC龜裂，修正為路堤損傷）結果相同，不算多個方案。"""
        result = decide('A1 AC龜裂。(8m*5m)', '', '引道路堤')
        self.assertFalse(result.review, result.reason)
        self.assertEqual((result.fields['amount'], result.fields['price'], result.fields['deterioration']),
                         ('40', '350', '路堤損傷'))


class NumberTests(unittest.TestCase):
    def test_quantity_never_uses_scientific_notation(self):
        self.assertEqual([quantity(v) for v in ('400', '440', '1,200', '0.50', '1E+2')], ['400', '440', '1200', '0.5', '100'])

    def test_count_unit_keeps_plain_number(self):
        result = decide('A1 植生。(1式)', '其他損傷', '橋台', unit='式', amount='400')
        self.assertNotIn('E+', json.dumps(result.fields, ensure_ascii=False))


class RemarkShapeTests(unittest.TestCase):
    def test_bare_sizes_get_units_and_brackets(self):
        for remark, unit, expected in (('A1 AC龜裂 5*7', '', 'A1 AC龜裂。(5m*7m)'),
                                       ('A1 AC龜裂 (5*7)', '平方公尺', 'A1 AC龜裂。(5m*7m)'),
                                       ('A1 路面混凝土層破洞 1*2*3', '', 'A1 路面混凝土層破洞。(1m*2m*3m)')):
            with self.subTest(remark=remark):
                self.assertEqual(shape_remark({'remark': remark, 'unit': unit}, '引道路堤')['remark'], expected)
        result = decide('A1 路面混凝土層破洞 1*2*3', '路堤損傷', '引道路堤')
        self.assertEqual((result.fields['unit'], result.fields['amount']), ('立方公尺', '6'))

    def test_pattern_only_gets_location_and_size_template(self):
        result = decide('AC龜裂', '路堤損傷', '引道路堤')
        self.assertEqual(result.fields['remark'], 'A1 AC龜裂。(m*m)')
        self.assertTrue(result.review)
        self.assertIn('尺寸範本', result.reason)

    def test_prior_location_is_used(self):
        prior = dict(remark='A1-L AC老化。(5m*7m)', unit='平方公尺', amount='35', method='x', price='350',
                     deterioration='路堤損傷')
        result = decide('AC龜裂', '路堤損傷', '引道路堤', prior=prior)
        self.assertTrue(result.fields['remark'].startswith('A1-L AC龜裂。'))

    def test_bank_ac_uses_width_times_five(self):
        result = decide('AC龜裂', '路堤損傷', '引道路堤', context={'位置': '橋頭', '_最小淨寬': '7.5'})
        self.assertFalse(result.review, result.reason)
        self.assertEqual((result.fields['remark'], result.fields['amount'], result.fields['unit']),
                         ('A1 AC龜裂。(5m*7.5m)', '37.5', '平方公尺'))


class AccountTests(unittest.TestCase):
    def setUp(self):
        self.vault = credentials.MemoryVault()
        credentials.use(self.vault)

    def tearDown(self):
        credentials.use(None)

    def test_many_accounts_with_labels(self):
        credentials.save('yunlin', 'a', '雲林縣')
        credentials.save('chiayi', 'b', '嘉義縣')
        self.assertEqual(credentials.accounts(), [('嘉義縣', 'chiayi'), ('雲林縣', 'yunlin')])
        self.assertEqual(credentials.default_label(), '雲林縣')            # 第一組為預設
        self.assertEqual(credentials.load(), ('yunlin', 'a'))
        credentials.set_default('嘉義縣')
        self.assertEqual(credentials.load(), ('chiayi', 'b'))
        credentials.save('yunlin', 'new', '雲林縣')                        # 同標籤更新
        self.assertEqual(credentials.load('雲林縣'), ('yunlin', 'new'))
        credentials.clear('嘉義縣')
        self.assertEqual(credentials.accounts(), [('雲林縣', 'yunlin')])
        self.assertEqual(credentials.load(), ('yunlin', 'new'))           # 預設被刪除改用第一組
        credentials.clear()
        self.assertEqual((credentials.accounts(), credentials.load()), ([], None))

    def test_old_single_account_is_kept(self):
        self.vault.write(credentials.LEGACY, 'old', 'secret')
        self.assertEqual(credentials.accounts(), [('old', 'old')])
        self.assertEqual(credentials.load(), ('old', 'secret'))
        self.assertIsNone(self.vault.read(credentials.LEGACY))

    def test_label_rules(self):
        self.assertEqual(credentials.save('user1', 'x'), 'user1')           # 標籤空白用帳號
        with self.assertRaises(ValueError):
            credentials.save('user2', 'x', 'a:b')


class HistoryAndPhotoTests(unittest.TestCase):
    def test_county(self):
        self.assertEqual(county_of({'所在地': '雲林縣莿桐鄉', '管理機關': '交通部公路局'}), '雲林縣')
        self.assertEqual(county_of({'管理機關': '台中市政府'}), '臺中市')
        self.assertEqual(county_of({'管理機關': '交通部公路局'}), '交通部公路局')

    def test_photo_book_marks_this_year(self):
        plan = PhotoPlan.create('是')
        photos = [{'title': '橋頭近照', 'date': '2026-03-02', 'note': ''}, {'title': '橋頭近照', 'date': '2025-02-01', 'note': ''},
                  {'title': '橋尾遠照', 'date': '2025-02-01', 'note': ''}]
        latest, newest = latest_dates(photos)
        self.assertEqual((latest['橋頭近照'], newest), ('2026-03-02', '2026-03-02'))
        self.assertEqual(mark_uploaded(plan, photos, 2026), ['橋頭近照'])
        self.assertTrue(plan.slots[0].status.startswith('網站已有'))
        self.assertEqual(plan.slots[1].status, '')                          # 只有往年的照片：仍要上傳


class PriorReuseTests(unittest.TestCase):
    def test_any_saved_prior_is_reused(self):
        from bridgefill.ui.main_window import MainWindow
        with tempfile.TemporaryDirectory() as folder:
            window = SimpleNamespace(scan_dir=Path(folder))
            window.prebuilt_path = lambda date, bridge: MainWindow.prebuilt_path(window, date, bridge)
            path = window.prebuilt_path('2024-05-11', '測試橋')
            path.write_text(json.dumps({'header': {'檢測日期': '2024-05-11'}, 'rows': []}), encoding='utf-8')
            record = SimpleNamespace(date='2024-05-11')
            self.assertIsNotNone(MainWindow.prebuilt_prior(window, record, '測試橋'))       # 一般掃描的前期也沿用
            self.assertIsNone(MainWindow.prebuilt_prior(window, SimpleNamespace(date='2023-01-01'), '測試橋'))


if __name__ == '__main__':
    unittest.main()


class EditCardTests(unittest.TestCase):
    """修改本期（2026-11-20 使用者選 A+B）：滾輪不會改到沒點過的下拉選單；修改時各列不擴大。"""

    @classmethod
    def setUpClass(cls):
        from PySide6.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def wheel(self, widget):
        from PySide6.QtCore import QPoint, QPointF, Qt
        from PySide6.QtGui import QWheelEvent
        center = QPointF(widget.rect().center())
        event = QWheelEvent(center, QPointF(widget.mapToGlobal(QPoint(5, 5))), QPoint(0, 0), QPoint(0, -120),
                            Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier, Qt.ScrollPhase.NoScrollPhase, False)
        self.app.sendEvent(widget, event)
        return event

    def test_wheel_needs_click_first(self):
        from bridgefill.ui.main_window import WheelSafeCombo
        combo = WheelSafeCombo()
        combo.addItems(['1', '2', '3'])
        combo.show()
        event = self.wheel(combo)
        self.assertEqual(combo.currentIndex(), 0)
        self.assertFalse(event.isAccepted())          # 交給外層捲動
        combo.setFocus()
        self.app.processEvents()
        if combo.hasFocus():                           # offscreen 平台不一定給焦點
            self.wheel(combo)
            self.assertEqual(combo.currentIndex(), 1)
        combo.close()

    def test_editing_keeps_row_heights(self):
        from bridgefill.ui.main_window import ReviewCard
        card = ReviewCard()
        heights = {card.editors[k].height() for k in ('deterioration', 'unit', 'remark')} | {c.height() for c in card.deru}
        self.assertEqual(len({card.editors[k].maximumHeight() for k in card.editors}), 1)
        self.assertTrue(heights)
        self.assertEqual(card.tag_menu.actions(), [])


class FillerNetworkTests(unittest.TestCase):
    """方案 A：只讀取的步驟連線失敗自動重試；方案 C：核對待確認依網站值改狀態（不送出）。"""

    def filler(self):
        from bridgefill.ui.filler import SiteFiller
        filler = SiteFiller.__new__(SiteFiller)
        filler.retries, filler.waits = 0, []
        filler.pause = filler.waits.append
        return filler

    def test_read_only_steps_retry(self):
        from bridgefill.ui.fetcher import FetchError
        from bridgefill.ui.filler import RETRY_TIMES
        filler, tries = self.filler(), []

        def flaky(url):
            tries.append(url)
            if len(tries) < 3:
                raise FetchError('開啟逾時')
            return 'ok'
        self.assertEqual(filler.patient(flaky, 'u'), 'ok')
        self.assertEqual((len(tries), filler.retries, filler.waits), (3, 2, [3000, 6000]))
        with self.assertRaises(FetchError):
            filler.patient(lambda url: (_ for _ in ()).throw(FetchError('斷線')), 'u')
        self.assertEqual(RETRY_TIMES, 2)

    def test_verify_pending(self):
        from bridgefill.core.fill import FillLog
        from bridgefill.ui.filler import FillError
        url = 'https://bms2.iot.gov.tw/bms3/inspect/damage/edit_damage/der_id,1/inspect_id,9/'
        with tempfile.TemporaryDirectory() as folder:
            log = FillLog(Path(folder) / 'log.json')
            for live, expected in (({'price': '800'}, '已完成'), ({'price': '500'}, '失敗'), ({'price': '9'}, '待確認')):
                item = SimpleNamespace(url=url, key='k', target={'price': '800'}, site={'price': '500'}, status='待確認',
                                       message='')
                filler = self.filler()
                filler.read = lambda u, live=live: (None, live)
                filler.verify(item, log)
                self.assertEqual(log.get('k')['status'], expected)
            filler = self.filler()
            from bridgefill.ui.fetcher import FetchError
            filler.read = lambda u: (_ for _ in ()).throw(FetchError('斷線'))
            with self.assertRaises(FillError) as caught:
                filler.verify(item, log)
            self.assertTrue(caught.exception.network)
