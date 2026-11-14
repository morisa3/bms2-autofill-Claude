import tempfile
from pathlib import Path
from types import SimpleNamespace
import unittest
from docx import Document
from bridgefill.core.config import FIELD_IDS
from bridgefill.core.data import History,normalize_remark,read_history,repair_components
from bridgefill.core.remarks import prepare_fields,crack_complete
from bridgefill.core.reconcile import scoped_history
from bridgefill.core.policy import automatic_plan
from ui_helpers import ui


class Requirements05Tests(unittest.TestCase):
    def setUp(self):
        self.fields=dict(deterioration='混凝土結構裂縫',d_location='S1-L,R',method='舊工法',amount='2.4',unit='平方公尺',price='100',remark='S1-L,R 鋼筋外露鏽蝕(0.2m*6m)*2')

    def show(self,app,controller,current=None,history=None,rules=None):
        controller.busy=True;controller.history=history or [];controller.repairs=rules or []
        app.show_review(dict(component='橋面板',target='橋面板 S1',current=current or self.fields,matches=history or [],repairs=rules or [],reason='',context={'橋孔編號':'S1'},suggested=None,
            choices={'deterioration':['混凝土結構裂縫','混凝土剝落鋼筋外露'],'unit':['處','平方公尺','公尺']}))

    def test_glued_description_keeps_all_location_suffixes(self):
        for location in ('4G1梁底','P1正','P1-正背','P1正背面','S1-6-4','S1-L,R,中間','1D1-1,1-4,1-6','P1上下游','4G1-All','4G1-頂部-0~5'):
            with self.subTest(location=location):
                result=prepare_fields(dict(self.fields,remark=location+'混凝土剝落鋼筋外露(1處)'))
                expected={'4G1梁底':'4G1-梁底','P1正':'P1-正','P1正背面':'P1-正背面','P1上下游':'P1-上下游'}.get(location,location)
                self.assertEqual(result['d_location'],expected)
                self.assertEqual(result['remark'],expected+' 混凝土剝落鋼筋外露。(1處)')

    def test_girder_alias_preserves_extent_and_position_words(self):
        for suffix in ('-0~35','-L,R','-All','-頂部-0~5','梁底'):
            for prefix in ('S4G1','S4-G1'):
                fields=prepare_fields(dict(self.fields,remark=prefix+suffix+' 混凝土剝落(1處)'),component='主梁')
                self.assertEqual(fields['d_location'],'4G1'+('-梁底' if suffix=='梁底' else suffix))

    def test_crossbeam_prefix_uses_span_without_removing_list_or_suffix(self):
        for old,new in [('S1-1-1','1D1-1'),('S2-2-3','2D2-3'),('1-3,1-4,2-1,3-1','4D1-3,1-4,2-1,3-1')]:
            fields=prepare_fields(dict(self.fields,remark=old+' 混凝土剝落(1處)'),component='橫隔梁',context={'橋孔編號':'S4'})
            self.assertEqual(fields['d_location'],new)

    def test_span_zero_padding_matches_only_same_span(self):
        history=[History('橋面板',dict(self.fields),3,{'位置':'S001'}),History('橋面板',dict(self.fields),4,{'位置':'S002'}),History('橋面板',dict(self.fields,d_location='A1',remark='A1 混凝土剝落(1處)'),5)]
        result=scoped_history(history,'橋面板',{'橋孔編號':'S1'},self.fields['deterioration'])
        self.assertEqual(result,[history[0]])

    def test_cm_crack_width_converts_to_mm_and_calculates_length(self):
        fields=dict(self.fields,unit='公尺',remark='S1 混凝土裂縫（0.02cm＊5m＊4）')
        result=prepare_fields(fields)
        self.assertEqual(result['remark'],'S1 混凝土裂縫。(0.2mm*5m)*4')
        self.assertEqual(result['amount'],'20')
        rule=dict(deterioration=fields['deterioration'],pattern='混凝土裂縫',unit='公尺',method='修復',price='200',note='裂縫寬度小於[0.3mm]使用')
        plan=automatic_plan(fields,'橋面板',[],[rule],{'橋孔編號':'S1'})
        self.assertFalse(plan.review,plan.reason)

    def test_shared_future_cable_sheets_map_to_both_components(self):
        self.assertEqual(repair_components('鋼纜錨碇裝置'),['鋼纜錨碇裝置(橋塔端)','鋼纜錨碇裝置(主梁端)'])
        self.assertEqual(repair_components('鋼纜錨碇裝置（橋塔端）'),['鋼纜錨碇裝置(橋塔端)'])
        self.assertEqual(repair_components('斜張鋼纜'),['鋼纜(吊索)-斜張鋼纜','鋼纜保護套管-斜張鋼纜'])

    def test_previous_good_and_uninspectable_rows_are_not_auto_imported(self):
        with tempfile.TemporaryDirectory() as folder:
            d=Document();t=d.add_table(rows=4,cols=6)
            for row,values in zip(t.rows,[['檢測項目','橋面板','','','',''],['位置','D','E','劣化類型','數量','備註'],['S001','1','','','',''],['S002','','0','','','']]):
                for c,v in zip(row.cells,values):c.text=v
            p=Path(folder)/'prior.docx';d.save(p);history=read_history(p)
            self.assertEqual(len(history),2)
            self.assertEqual(scoped_history(history,'橋面板',{'橋孔編號':'S1'}),[])
            self.assertEqual(scoped_history(history,'橋面板',{'橋孔編號':'S1'},include_status=True)[0].context['_inspection_status'],'前期檢測良好（D＝1）')
            self.assertEqual(history[1].context['_inspection_status'],'前期無法檢測（E＝0）')

    def test_three_item_pages_have_numbers_and_no_continuous_scroll(self):
        app,c,p=ui()
        histories=[History('橋面板',dict(self.fields,price=str(i)),i,{'位置':'S001'}) for i in range(7)]
        rules=[dict(deterioration=self.fields['deterioration'],pattern=f'規則{i}',method='工法',unit='處',price='100') for i in range(7)]
        self.show(app,c,history=histories,rules=rules)
        self.assertTrue(app.history_list.controls[0].tooltip.startswith('1.'))
        app.turn_page('history',1)
        self.assertTrue(app.history_list.controls[0].tooltip.startswith('4.'))
        self.assertEqual(app.page_labels['history'].value,'4–6 / 7')
        app.turn_page('history',1)
        self.assertTrue(app.history_list.controls[0].tooltip.startswith('7.'))
        self.assertEqual(len(app.history_list.controls),3)

    def test_component_popup_cancel_and_confirm_have_separate_drafts(self):
        app,c,p=ui();c.history=[History('主梁',self.fields,3),History('橋面板',self.fields,4)]
        app.refresh_components();app.open_components()
        app.select_components(False);app.close_components(False)
        self.assertEqual(app.selected_components,{'主梁','橋面板'})
        app.open_components();app.select_components(False);app.toggle_component('主梁');app.close_components(True)
        self.assertEqual(app.selected_components,{'主梁'})
        self.assertIn('1／2',app.component_summary.value)

    def test_type_correction_loads_all_rules_and_changes_category_and_style(self):
        app,c,p=ui()
        old=History('橋面板',self.fields,3,{'位置':'S001'})
        rule=dict(deterioration='混凝土剝落鋼筋外露',components=['橋面板'],pattern='混凝土剝落鋼筋外露',unit='平方公尺',method='修復',price='300')
        self.show(app,c,history=[old],rules=[]);c.repairs=[rule]
        app.toggle_type_correction();app.select_history(old);app.select_repair(rule)
        self.assertEqual(app.fields['deterioration'].value,rule['deterioration'])
        self.assertEqual(app.fields['remark'].value,'S1-L,R 混凝土剝落鋼筋外露。(0.2m*6m)*2')
        self.assertEqual(app.fields['amount'].value,'2.4')

    def test_wrong_type_rule_cannot_fall_back_to_original_dropdown_category(self):
        app,c,p=ui();old=History('橋面板',self.fields,3,{'位置':'S001'})
        rule=dict(deterioration='網站不存在的類型',components=['橋面板'],pattern='混凝土剝落鋼筋外露',unit='平方公尺',method='修復',price='300')
        self.show(app,c,history=[old],rules=[]);c.repairs=[rule]
        app.toggle_type_correction();app.select_history(old);app.select_repair(rule)
        self.assertFalse(app.fields['deterioration'].value)
        app.payload['choices']['deterioration'].append('')
        app.submit(None);c.resolve.assert_not_called()

    def test_blank_current_fields_turn_yellow_and_clear_when_filled(self):
        app,c,p=ui();self.show(app,c,current=dict(self.fields,method='',price=''))
        self.assertEqual(app.fields['price'].bgcolor,'#fff0a8')
        app.fields['price'].value='200';app.fields_changed()
        self.assertEqual(app.fields['price'].bgcolor,'white')
        self.assertEqual(app.fields['method'].bgcolor,'#fff0a8')

    def test_crack_template_has_no_invented_dimensions_and_blocks_submit(self):
        app,c,p=ui();self.show(app,c,current=dict(self.fields,remark='S1 混凝土裂縫(1處)',unit='處',amount='1'))
        self.assertTrue(app.crack_button.visible)
        app.apply_crack_template()
        self.assertIn('(mm*m)*數量',app.fields['remark'].value)
        app.submit(None);c.resolve.assert_not_called()
        self.assertFalse(crack_complete({k:str(v.value or '') for k,v in app.fields.items()}))
