from copy import deepcopy
import unittest
import test_documents
from ui_helpers import ui
from bridgefill.core.data import read_repairs
from bridgefill.core.rules_v2 import read_rules
from bridgefill.core.local_files import REFERENCE_DIR,DEFAULT_REPAIR_NAME
from bridgefill.core.policy import automatic_plan
from bridgefill.core.remarks import prepare_fields,crack_count
from bridgefill.core.conditions import condition
from bridgefill.core.documents import export_document


class PlantCrackUpdates(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.rules=read_rules(REFERENCE_DIR/DEFAULT_REPAIR_NAME)  # 1116：由結構化欄位還原備註

    def plan(self,remark,kind='其他損傷',unit='式',amount='1',context=None,component='橋台'):
        return automatic_plan(dict(remark=remark,deterioration=kind,unit=unit,amount=amount),component,[],self.rules,context or {})

    def test_plant_count_uses_remark_instead_of_wrong_unit(self):
        for name,price in [('草本植物生長','500'),('木本植物生長','1000')]:
            result=self.plan(f'A1-L {name}。(2處)',unit='式',amount='99',context={'U':'1'})
            self.assertFalse(result.review,result.reason)
            self.assertEqual((result.fields['amount'],result.fields['unit'],result.fields['price']),('2','處',price))

    def test_plant_area_is_one_lump_sum_instead_of_area_quantity(self):
        for name in ['草本植物生長','木本植物生長']:
            for unit in ['處','式','平方公尺']:
                result=self.plan(f'A1-L {name}。(3m*5m)',unit=unit,amount='15')
                self.assertFalse(result.review,result.reason)
                self.assertEqual((result.fields['amount'],result.fields['unit'],result.fields['price']),('1','式','2000'))
                self.assertIn('(3m*5m)',result.fields['remark'])
                rule=next(r for r in self.rules if r['pattern']==name and r['unit']=='式' and '橋台' in r['components'])
                manual=prepare_fields(dict(result.fields,amount='15'),rule,True,'橋台',{})
                self.assertEqual(manual['amount'],'1')

    def test_generic_vegetation_uses_u_and_remark_quantity_shape(self):
        for urgency,price in [('1','500'),('2','1000')]:
            result=self.plan('A1-L 植生。(1處)',context={'U':urgency})
            self.assertFalse(result.review,result.reason)
            self.assertEqual(result.fields['price'],price)
        self.assertTrue(self.plan('A1-L 植生。(1處)').review)

    def test_crack_type_alias_width_and_multiplier_match_current_excel(self):
        for kind in ['混凝土裂縫','混凝土結構裂縫']:
            for width,price in [('0.2','1500'),('0.3','1800'),('0.5','1800'),('.2','1500'),('.3','1800')]:
                for suffix,amount in [('', '5'),('*3','15')]:
                    result=self.plan(f'A1-L 混凝土裂縫。({width}mm*5m){suffix}',kind,unit='公尺',amount=amount)
                    self.assertFalse(result.review,result.reason)
                    self.assertEqual((result.fields['price'],result.fields['amount'],result.fields['unit']),(price,amount,'公尺'))
                    self.assertTrue(result.fields['remark'].endswith(suffix or '*1'))
                    self.assertEqual(prepare_fields(result.fields,sync_quantity=True)['remark'],result.fields['remark'])

    def test_width_placeholder_is_not_a_measured_value(self):
        for size in ['mm*5m','0.2mm*m','mm*m','0.2mm']:
            self.assertTrue(self.plan(f'A1-L 混凝土裂縫。({size})','混凝土結構裂縫',unit='公尺',amount='5').review)
        self.assertEqual(crack_count('A1-L 混凝土裂縫。(mm*5m)'),'A1-L 混凝土裂縫。(mm*5m)')

    def test_other_thresholds_are_not_changed_by_point_three_boundary(self):
        self.assertEqual(condition({'note':'裂縫寬度大於[0.5mm]使用'},{},'A1 混凝土裂縫。(0.5mm*5m)'),(False,False))

    def test_check_missing_fields_is_all_rows_and_read_only(self):
        f=test_documents.DocumentTests();f.setUp()
        try:
            records=[deepcopy(f.record),deepcopy(f.record)];records[1]['id']='second'
            export_document(f.template,f.output,'job',records,f.history,f.rules)
            app,c,page=ui();c.history=f.history;c.repairs=f.rules;app.start_comparison(f.output)
            app.comparison.entries[1]['fields']['price']=''
            before=deepcopy(app.comparison.entries)
            app.confirm_comparison()
            self.assertEqual(before,app.comparison.entries)
            self.assertEqual(page.dialog.title.value,'本期待補欄位（1 筆）')
            button=page.dialog.content.content.controls[-1]
            self.assertIn('第 2 筆',button.text);self.assertIn('單價',button.text)
            button.on_click(None)
            self.assertEqual(app.editor_controls['price'].value,'')
            app.close_compare_editor(False)
            app.comparison.entries[1]['fields']['price']='300'
            app.confirm_comparison()
            self.assertIn('未發現空白',page.snack_bar.content.value)
            app.comparison.path.unlink(missing_ok=True)
        finally:f.tearDown()
