import unittest
from bridgefill.core.data import History
from bridgefill.core.remarks import prepare_fields
from bridgefill.core.reconcile import scoped_history
from bridgefill.core.policy import automatic_plan
from ui_helpers import ui


class ManualRepairTests(unittest.TestCase):
    def setUp(self):
        self.fields=dict(deterioration='其他損傷',d_location='',method='持續追蹤',amount='1',unit='式',price='',remark='P13-2正背面 混凝土龜裂 （8m＊5m）＊2')
        self.old=History('墩柱與帽梁',self.fields,3,{'編號':'P013N-2'})
        self.rule=dict(deterioration='其他損傷',pattern='混凝土表面龜裂',method='表面處理',unit='平方公尺',price='300',note='有關鍵字[龜裂]即可使用')

    def test_manual_selection_replaces_style_quantity_and_location_despite_mismatched_current_amount(self):
        app,controller,page=ui();controller.busy=True
        app.show_review(dict(component='墩柱與帽梁',target='墩柱',current=dict(self.fields,remark='',amount='80'),
            matches=[self.old],repairs=[self.rule],reason='數量不同',suggested=None,
            choices={'deterioration':['其他損傷'],'unit':['式','平方公尺']},context={'編號':'P013N-2'}))
        app.select_history(self.old);app.select_repair(self.rule)
        self.assertEqual(app.fields['remark'].value,'P13-2-正背面 混凝土表面龜裂。(5m*8m)*2')
        self.assertEqual(app.fields['d_location'].value,'P13-2-正背面')
        self.assertEqual(app.fields['amount'].value,'80')

    def test_keyword_alias_replaces_even_when_partial_pattern_already_present(self):
        fields=dict(self.fields,remark='P1 混凝土表面龜裂鋼筋外露 （2m＊3m）')
        rule=dict(self.rule,note='有關鍵字[龜裂]即可使用')
        result=automatic_plan(fields,'墩柱與帽梁',[],[rule])
        self.assertFalse(result.review,result.reason)
        self.assertEqual(result.fields['remark'],'P1 混凝土表面龜裂。(2m*3m)')

    def test_basic_format_runs_even_when_excel_or_history_matching_fails(self):
        fields=dict(self.fields,remark='A1 混凝土裂縫（5m＊0.2mm＊4）')
        result=automatic_plan(fields,'橋台',[],[])
        self.assertTrue(result.review)
        self.assertEqual(result.fields['remark'],'A1 混凝土裂縫。(0.2mm*5m)*4')
        self.assertEqual(result.fields['d_location'],'A1')

    def test_head_tail_alias_applies_only_to_identity_not_defect_location(self):
        a1=History('橋台',self.fields,3,{'位置':'橋頭','編號':'A1'})
        a2=History('橋台',self.fields,4,{'位置':'橋尾','編號':'A2'})
        self.assertEqual(scoped_history([a1,a2],'橋台',{'位置':'A1'}),[a1])
        self.assertEqual(scoped_history([a1,a2],'橋台',{'位置':'橋尾'}),[a2])
