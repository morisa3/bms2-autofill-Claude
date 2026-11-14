import tempfile
from pathlib import Path
import unittest
from docx import Document
from bridgefill.core.config import FIELD_IDS
from bridgefill.core.data import History, read_history
from bridgefill.core.policy import automatic_plan
from bridgefill.core.reconcile import scoped_history, reconcile
from ui_helpers import ui


class ReconcileTests(unittest.TestCase):
    def setUp(self):
        self.old=History('橋台',dict(deterioration='混凝土剝落、破碎、鋼筋外露、銹蝕',
            d_location='A1-L',method='舊工法',amount='1.5',unit='平方公尺',price='100',
            remark='A1-L 混凝土剝落鋼筋外露。(0.5m*3m)'),3,{'位置':'橋頭','編號':'A1'})
        self.rules=[dict(deterioration=self.old.fields['deterioration'],pattern='混凝土剝落鋼筋外露',
            components=['橋台'],method='新工法',unit='平方公尺',price='200')]
        self.context={'橋台位置':'橋頭','編號':'A1','DERU':'D:2 E:2 R:1 U:1'}

    def current(self, **values):
        fields={key:'' for key in FIELD_IDS}
        fields['deterioration']=self.old.fields['deterioration']
        fields.update(values)
        return fields

    def test_summary_variants_restore_complete_formula_and_wrong_unit(self):
        for remark in ('','A1-L 混凝土剝落鋼筋外露','A1-L 1.5m','A1-L 數量填1.5m','1.5','1.5 平方公尺'):
            with self.subTest(remark=remark):
                result=automatic_plan(self.current(amount='1.5m',unit='公尺',remark=remark), '橋台',[self.old],self.rules,self.context)
                self.assertFalse(result.review,result.reason)
                self.assertEqual(result.fields['remark'],self.old.fields['remark'])
                self.assertEqual(result.fields['d_location'],'A1-L')
                self.assertEqual(result.fields['amount'],'1.5')
                self.assertEqual(result.fields['unit'],'平方公尺')
                self.assertEqual(result.fields['price'],'200')

    def test_new_expanded_extent_preserved_and_excel_applied(self):
        remark='A1-L,R 混凝土剝落鋼筋外露。(0.5m*3m)*2'
        result=automatic_plan(self.current(amount='3',remark=remark), '橋台',[self.old],self.rules,self.context)
        self.assertFalse(result.review)
        self.assertEqual(result.fields['remark'],remark)
        self.assertEqual(result.fields['d_location'],'A1-L,R')
        self.assertEqual(result.fields['amount'],'3')
        self.assertEqual(result.fields['method'],'新工法')

    def test_ambiguous_same_style_or_different_incomplete_amount_needs_review(self):
        second=History('橋台',dict(self.old.fields,d_location='A1-R',remark=self.old.fields['remark'].replace('A1-L','A1-R')),4,self.old.context)
        self.assertTrue(automatic_plan(self.current(amount='1.5',unit='公尺'), '橋台',[self.old,second],self.rules,self.context).review)
        self.assertTrue(automatic_plan(self.current(amount='2'), '橋台',[self.old],self.rules,self.context).review)

    def test_number_type_and_specific_style_scope(self):
        wrong=History('橋台',dict(self.old.fields),4,{'位置':'橋尾','編號':'A2'})
        unrelated=History('橋台',dict(self.old.fields,deterioration='混凝土結構裂縫'),5,self.old.context)
        self.assertEqual(scoped_history([self.old,wrong,unrelated],'橋台',self.context,self.old.fields['deterioration']),[self.old])
        self.assertEqual(scoped_history([self.old],'橋台',{'編號':'A2'},self.old.fields['deterioration']),[])

    def test_unique_same_style_within_broad_type_can_restore(self):
        other=History('橋台',dict(self.old.fields,remark='A1-L 混凝土破損。(1m*2m)'),4,self.old.context)
        result=automatic_plan(self.current(amount='1.5',remark='A1-L 混凝土剝落鋼筋外露'), '橋台',[other,self.old],self.rules,self.context)
        self.assertFalse(result.review)
        self.assertEqual(result.fields['remark'],self.old.fields['remark'])

    def test_without_excel_restores_all_old_fields(self):
        result=automatic_plan(self.current(amount='1.5',unit='公尺'), '橋台',[self.old],[],self.context,False)
        self.assertFalse(result.review)
        self.assertEqual(result.fields,self.old.fields)

    def test_different_type_never_uses_other_type_excel_rule(self):
        current=self.current(deterioration='混凝土結構裂縫',remark=self.old.fields['remark'])
        result=automatic_plan(current,'橋台',[self.old],self.rules,self.context)
        self.assertTrue(result.review)
        self.assertEqual(result.fields['deterioration'],'混凝土結構裂縫')

    def test_explicit_type_change_filters_lists_and_restores_selected_history(self):
        app,controller,page=ui()
        controller.busy=True
        controller.history=[self.old,History('橋台',self.old.fields,4,{'編號':'A2'})]
        controller.repairs=self.rules
        app.show_review(dict(component='橋台',target='橋台 A1',current=self.current(deterioration='混凝土結構裂縫'),
            matches=[],repairs=[],reason='人工確認類型',suggested=None,
            choices={'deterioration':['混凝土結構裂縫',self.old.fields['deterioration']],'unit':['平方公尺']},context=self.context))
        app.fields['deterioration'].value=self.old.fields['deterioration']
        app.type_changed(None)
        self.assertEqual(app.payload['matches'],[self.old])
        self.assertEqual(sum(c.content is not None for c in app.repair_list.controls),1)
        self.assertEqual(app.fields['remark'].value,self.old.fields['remark'])
        self.assertEqual(app.history_fields['d_location'].value,'A1-L')

    def test_word_keeps_identity_separate_from_damage_location(self):
        with tempfile.TemporaryDirectory() as directory:
            document=Document()
            table=document.add_table(rows=3,cols=6)
            for row,values in zip(table.rows,[['檢測項目','橋台','','','',''],
                ['位置','編號','劣化類型','數量','單位','備註'],
                ['橋頭','A1',self.old.fields['deterioration'],'1.5','平方公尺',self.old.fields['remark']]]):
                for cell,value in zip(row.cells,values):cell.text=value
            path=Path(directory)/'history.docx'
            document.save(path)
            record=read_history(path)[0]
            self.assertEqual(record.context,{'位置':'橋頭','編號':'A1'})
            self.assertEqual(record.fields['d_location'],'A1-L')

    def test_explicit_manual_rule_populates_method_unit_price(self):
        app,controller,page=ui()
        controller.busy=True
        rule=dict(self.rules[0],note='手動選擇')
        current=self.current(amount='2',remark='數量範圍待人工確認')
        app.show_review(dict(component='橋台',target='橋台 A1',current=current,
            matches=[self.old],repairs=[rule],reason='人工確認',suggested=None,
            choices={'deterioration':[self.old.fields['deterioration']],'unit':['平方公尺']},context=self.context))
        app.select_repair(rule)
        self.assertEqual(app.fields['method'].value,'新工法')
        self.assertEqual(app.fields['price'].value,'200')
        self.assertEqual(app.fields['amount'].value,'2')
        self.assertEqual(app.fields['remark'].value,'數量範圍待人工確認')

    def test_integrated_flet_review_old_and_new_panels_and_submit(self):
        app,controller,page=ui()
        controller.busy=True
        current=self.current(amount='1.5',unit='公尺')
        suggestion=automatic_plan(current,'橋台',[self.old],self.rules,self.context)
        app.show_review(dict(component='橋台',target='構件：橋台 ｜ 編號：A1',current=current,
            matches=[self.old],repairs=self.rules,reason=suggestion.reason,suggested=suggestion.fields,
            choices={'deterioration':[self.old.fields['deterioration']],'unit':['公尺','平方公尺']},context=self.context))
        self.assertEqual(app.history_fields['remark'].value,self.old.fields['remark'])
        self.assertEqual(app.fields['remark'].value,self.old.fields['remark'])
        self.assertFalse(app.submit_button.disabled)
        app.submit(None)
        submitted=controller.resolve.call_args.args[0]
        self.assertEqual(submitted['d_location'],'A1-L')
        self.assertEqual(submitted['unit'],'平方公尺')
        self.assertEqual(submitted['remark'],self.old.fields['remark'])
        self.assertFalse(hasattr(page,'dialog'))
