import unittest
from types import SimpleNamespace
from bridgefill.core.config import FIELD_IDS
from bridgefill.core.locations import basic_location
from bridgefill.core.remarks import prepare_fields,template_incomplete
from ui_helpers import ui


class NewDefectTemplateTests(unittest.TestCase):
    def setUp(self):
        self.rule=dict(deterioration='其他損傷',pattern='防衝鋼板表面生銹',method='防蝕塗裝',unit='平方公尺',price='200')

    def show(self,app,controller,context=None,**values):
        controller.busy=True
        current={key:'' for key in FIELD_IDS};current.update(values)
        app.show_review(dict(component='橋墩／橋基保護設施',target='構件 P010',current=current,matches=[],repairs=[self.rule],reason='本筆為新增劣化',suggested=None,
            context=context or {'編號':'P010'},choices={'deterioration':['','其他損傷'],'unit':['','平方公尺','公尺']}))

    def test_empty_new_record_seeds_location_and_rule_style(self):
        app,c,p=ui();self.show(app,c)
        self.assertEqual(app.fields['d_location'].value,'P10')
        app.select_repair(self.rule)
        self.assertEqual(app.fields['remark'].value,'P10 防衝鋼板表面生銹。(m*m)')
        self.assertEqual(app.fields['deterioration'].value,'其他損傷')
        self.assertEqual(app.fields['amount'].value,'')
        self.assertEqual(app.fields['price'].value,'200')
        app.submit(None);c.resolve.assert_not_called()
        self.assertIn('補齊',p.snack_bar.content.value)

    def test_manual_location_and_numeric_extent_are_preserved_and_calculated(self):
        app,c,p=ui();self.show(app,c,deterioration='其他損傷')
        app.select_repair(self.rule)
        app.fields['d_location'].value='P10-上下游'
        app.fields_changed(SimpleNamespace(control=app.fields['d_location']))
        self.assertEqual(app.fields['remark'].value,'P10-上下游 防衝鋼板表面生銹。(m*m)')
        app.fields['remark'].value='P10-上下游 防衝鋼板表面生銹。(3m*8m)'
        app.submit(None)
        result=c.resolve.call_args.args[0]
        self.assertEqual(result['d_location'],'P10-上下游')
        self.assertEqual(result['amount'],'24')

    def test_selecting_category_preserves_new_manual_location(self):
        app,c,p=ui();self.show(app,c)
        app.fields['d_location'].value='P10-L,R'
        app.fields['deterioration'].value='其他損傷'
        app.type_changed(None);app.select_repair(self.rule)
        self.assertEqual(app.fields['remark'].value,'P10-L,R 防衝鋼板表面生銹。(m*m)')

    def test_rule_change_keeps_completed_measurements(self):
        app,c,p=ui();self.show(app,c,deterioration='其他損傷')
        app.select_repair(self.rule)
        app.fields['remark'].value='P10-正背 防衝鋼板表面生銹。(3m*8m)'
        app.fields_changed(SimpleNamespace(control=app.fields['remark']))
        app.select_repair(dict(self.rule,pattern='防衝鋼板遺失',price='13000'))
        self.assertEqual(app.fields['remark'].value,'P10-正背 防衝鋼板遺失。(3m*8m)')
        self.assertEqual(app.fields['amount'].value,'24')

    def test_existing_location_and_remark_are_not_replaced_by_basic_seed(self):
        app,c,p=ui();self.show(app,c,deterioration='其他損傷',d_location='P10-上游',remark='P10-上游 生銹。(2m*4m)')
        app.select_repair(self.rule)
        self.assertEqual(app.fields['remark'].value,'P10-上游 防衝鋼板表面生銹。(2m*4m)')

    def test_ambiguous_number_stays_blank_and_template_cannot_be_submitted(self):
        app,c,p=ui();self.show(app,c,context={'編號':'P010｜P011'})
        self.assertFalse(app.fields['d_location'].value)
        app.select_repair(self.rule)
        self.assertEqual(app.fields['remark'].value,'損壞位置 防衝鋼板表面生銹。(m*m)')
        self.assertFalse(app.fields['d_location'].value)
        app.submit(None);c.resolve.assert_not_called()

    def test_basic_seed_formats_defined_component_identifiers(self):
        self.assertEqual(basic_location('橋墩／橋基保護設施',{'編號':'P01｜P010'}),'')
        self.assertEqual(basic_location('橋面板',{'橋孔編號':'S004'}),'S4')
        self.assertEqual(basic_location('主梁',{'橋孔編號':'S004','編號':'S4G1'}),'4G1')
        self.assertEqual(basic_location('橫隔梁',{'橋孔編號':'S004','隔梁編號':'1-3'}),'4D1-3')
        self.assertEqual(basic_location('橋台',{'位置':'橋尾'}),'A2')
        self.assertEqual(basic_location('橋墩／橋基保護設施',{'編號':'P010｜P010N-2'}),'P10N-2')

    def test_templates_follow_units_without_inventing_numbers(self):
        for unit,template in [('平方公尺','(m*m)'),('公尺','(m)'),('立方公尺','(m*m*m)'),('處','(處)')]:
            rule=dict(self.rule,unit=unit)
            fields=prepare_fields({key:'' for key in FIELD_IDS},rule,True,component='橋墩／橋基保護設施',context={'編號':'P010'})
            self.assertEqual(fields['remark'],'P10 防衝鋼板表面生銹。'+template)
            self.assertTrue(template_incomplete(fields['remark']))
            self.assertFalse(fields['amount'])
        self.assertFalse(template_incomplete('P10 防衝鋼板表面生銹。(3m*8m)'))
