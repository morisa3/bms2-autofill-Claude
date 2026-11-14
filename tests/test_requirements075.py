from copy import deepcopy
from pathlib import Path
import unittest
import test_documents
from ui_helpers import ui
from bridgefill.core.alignment import align_records
from bridgefill.core.comparison import ComparisonSession
from bridgefill.core.data import History
from bridgefill.core.documents import export_document,read_generated


class Requirements075Tests(unittest.TestCase):
    def setUp(self):
        self.f=test_documents.DocumentTests();self.f.setUp()

    def tearDown(self):self.f.tearDown()

    def fields(self,style='混凝土剝落',amount='2',location='A1-L',price='300'):
        return dict(self.f.fields,remark=f'{location} {style}。(1m*2m)',amount=amount,d_location=location,price=price)

    def record(self,id,style='混凝土剝落',amount='2',price='300'):
        record=deepcopy(self.f.record);record.update(id=id,current=self.fields(style,amount,price=price))
        return record

    def old(self,style='混凝土剝落',amount='2'):
        return History('橋台',self.fields(style,amount),3,{'位置':'橋頭','編號':'A1'})

    def session(self,records,history,rules=None):
        export_document(self.f.template,self.f.output,'job',records,history,rules or [],review_colors=True)
        return ComparisonSession(self.f.output,history,rules or [],self.f.root/'temp',self.f.root/'out')

    def test_one_prior_three_current_different_order_pairs_only_the_matching_style(self):
        records=[self.record('new1','木本植物生長'),self.record('old'),self.record('new2','混凝土裂縫')]
        rows=align_records(records,[self.old()])
        self.assertEqual([(r['current_index'],r['history_index'],r['kind']) for r in rows],[(0,None,'new'),(1,0,'matched'),(2,None,'new')])

    def test_prior_only_deleted_record_stays_visible_but_not_in_output(self):
        history=[self.old(),self.old('木本植物生長'),self.old('混凝土龜裂')]
        records=[self.record('2','混凝土龜裂'),self.record('1')]
        s=self.session(records,history)
        self.assertEqual(sum(r['kind']=='removed' for r in s.rows),1)
        removed=next(r for r in s.rows if r['kind']=='removed')
        self.assertEqual(removed['history_index'],1)
        for i in range(2):s.select(i);s.confirm()
        output=s.finish('測試橋')
        self.assertEqual([e['id'] for e in read_generated(output)['records']],['2','1'])

    def test_manual_mapping_changes_only_alignment_and_survives_reload(self):
        s=self.session([self.record('new','漏水')],[self.old()])
        before=deepcopy(s.entries[0]['fields'])
        s.map_history(0,0)
        self.assertEqual(s.entries[0]['fields'],before)
        self.assertEqual(len(s.histories(s.entries[0])),1)
        again=ComparisonSession(self.f.output,s.history,[],self.f.root/'temp',self.f.root/'out')
        self.assertEqual(again.pairs,{0:0})
        self.assertEqual(again.entries[0]['fields'],before)
        again.map_history(0,None)
        self.assertEqual(again.entries[0]['color'],'red')
        self.assertEqual(len(again.rows),2)

    def test_ambiguous_duplicates_are_yellow_then_manual_pair_does_not_reuse_prior(self):
        s=self.session([self.record('a'),self.record('b')],[self.old()])
        self.assertEqual([e['color'] for e in s.entries],['yellow','yellow'])
        s.map_history(0,0)
        self.assertEqual([e['color'] for e in s.entries],['normal','red'])
        with self.assertRaisesRegex(ValueError,'已與其他'):s.map_history(1,0)

    def test_same_amount_and_similar_remark_are_normal_without_an_excel_rule(self):
        records=[self.record('a','剝落',amount='3727.5')]
        history=[self.old('剝離',amount='3727.50')]
        s=self.session(records,history)
        self.assertEqual(s.entries[0]['color'],'normal')
        self.assertEqual(s.pairs,{0:0})

    def test_quantity_disambiguates_same_style_without_using_row_order(self):
        rows=align_records([self.record('b',amount='2'),self.record('a',amount='1')],[self.old(amount='1'),self.old(amount='2')])
        self.assertEqual([r['history_index'] for r in rows],[1,0])

    def test_unmatched_current_does_not_copy_prior_fields_or_scope(self):
        new=self.record('new','木本植物生長',amount='9',price='')
        s=self.session([new],[self.old()])
        self.assertIn('木本植物生長',s.entries[0]['fields']['remark'])
        self.assertEqual(s.entries[0]['fields']['amount'],'2')
        self.assertEqual(s.entries[0]['fields']['price'],'')
        self.assertEqual(s.histories(s.entries[0]),[])

    def test_incomplete_price_saves_draft_but_export_names_missing_field(self):
        s=self.session([self.record('a',price='')],[])
        fields=deepcopy(s.entries[0]['fields']);s.edit(fields)
        self.assertFalse(s.entries[0]['confirmed'])
        self.assertEqual(s.entries[0]['fields']['price'],'')
        self.assertTrue(s.finish('測試橋').is_file())
        self.assertIn('單價',s.issues()[0]['detail'])
        with self.assertRaisesRegex(ValueError,'單價'):s.confirm()

    def test_edit_dialog_saves_partial_data_and_uses_current_type_error_button(self):
        s=self.session([self.record('a','木本植物生長',price='')],[])
        app,c,page=ui();c.history=[];c.repairs=[]
        app.start_comparison(self.f.output);app.open_compare_editor(0)
        self.assertEqual(app.editor_correction.text,'本期劣化類型選擇錯誤')
        self.assertEqual(app.editor_controls['price'].bgcolor,'#fff0a8')
        app.editor_controls['method'].value='人工修改工法';app.close_compare_editor(True)
        self.assertFalse(page.dialog.open)
        self.assertEqual(app.comparison.entries[0]['fields']['method'],'人工修改工法')
        self.assertIn('待補',page.snack_bar.content.value)
        app.comparison.path.unlink(missing_ok=True)

    def test_added_red_is_only_on_current_side_and_deleted_is_display_only(self):
        self.session([self.record('a','木本植物生長')],[self.old()])
        app,c,page=ui();c.history=[self.old()];c.repairs=[];app.start_comparison(self.f.output)
        self.assertEqual(app.compare_left.controls[0].bgcolor,'white')
        self.assertEqual(app.compare_right.controls[0].bgcolor,'#ffd8d8')
        self.assertEqual(app.compare_left.controls[0].content.controls[0].value,'[前期未有該劣化樣態]')
        self.assertEqual(app.compare_right.controls[1].content.controls[0].value,'[本期未有該劣化樣態]')
        before=deepcopy(app.comparison.entries)
        app.comparison_select(1);app.compare_apply_rule(self.f.rules[0])
        self.assertEqual(before,app.comparison.entries)
        app.comparison.path.unlink(missing_ok=True)

    def test_new_quantity_range_is_not_overwritten_when_no_excel_rule_exists(self):
        record=self.record('a',amount='9')
        record['current']['remark']='A1-L 混凝土剝落。(1m*9m)'
        s=self.session([record],[self.old(amount='2')])
        self.assertEqual(s.entries[0]['fields']['amount'],'9')
        self.assertIn('9m',s.entries[0]['fields']['remark'])
        self.assertEqual(s.entries[0]['color'],'normal')

    def test_explicit_new_and_removed_do_not_block_confirmed_current_export(self):
        s=self.session([self.record('new','漏水')],[self.old()])
        s.map_history(0,None);s.confirm()
        output=s.finish('測試橋')
        self.assertEqual(len(read_generated(output)['records']),1)

    def test_pairing_uses_edited_current_values_not_immutable_website_snapshot(self):
        history=[History('其他',dict(self.fields('剝離',amount='3727.5',location='整體-All'),deterioration='表面劣化'),3,{'類別':'鋪面','編號':'8'})]
        record=self.record('pavement');record['component']='其他';record['context']={'類別':'鋪面','編號':'8'}
        record['current']=dict(record['current'],deterioration='表面劣化',remark='',amount='1',unit='式',d_location='')
        s=self.session([record],history)
        self.assertEqual(s.entries[0]['color'],'yellow')
        updated=dict(self.fields('剝落',amount='3727.5',location='整體-All'),deterioration='表面劣化')
        s.edit(updated)
        self.assertEqual(s.entries[0]['color'],'normal')
        self.assertEqual(s.pairs,{0:0})
        self.assertEqual(s.entries[0]['current']['amount'],'1')

    def test_manual_mapping_across_components_is_compare_only_and_exports_current_component(self):
        record=self.record('moved');record['component']='翼牆／擋土牆'
        history=[self.old()];before=deepcopy(history[0].fields)
        s=self.session([record],history)
        current=deepcopy(s.entries[0]['fields']);s.map_history(0,0)
        self.assertEqual(current,s.entries[0]['fields'])
        self.assertEqual(history[0].fields,before)
        s.confirm();output=s.finish('測試橋')
        self.assertEqual(read_generated(output)['records'][0]['component'],'翼牆／擋土牆')

    def test_context_is_above_edit_fields_and_below_pager_with_rule_help(self):
        self.session([self.record('a')],[self.old()])
        app,c,page=ui();c.history=[self.old()];app.start_comparison(self.f.output);app.open_compare_editor(0)
        controls=page.dialog.content.content.controls
        self.assertEqual(controls[0].controls[0].content.controls[0].value,'構件缺失規則，點選後可更新備註資訊')
        self.assertIn(app.editor_controls['remark'],controls[-1].controls)
        self.assertTrue(controls[-2].content.value.startswith('構件資訊：'))
        app.comparison.path.unlink(missing_ok=True)

    def test_all_prior_defects_removed_are_display_only_with_no_completed_export(self):
        export_document(self.f.template,self.f.output,'job',[],[self.old()],[],review_colors=True,allow_empty=True)
        s=ComparisonSession(self.f.output,[self.old()],[],self.f.root/'temp',self.f.root/'out')
        self.assertEqual(len(s.entries),0)
        self.assertEqual(len(s.rows),1)
        with self.assertRaisesRegex(ValueError,'不會輸出'):s.finish('測試橋')
        app,c,page=ui();c.history=[self.old()];app.start_comparison(self.f.output)
        self.assertEqual(app.compare_right.controls[0].content.controls[0].value,'[本期未有該劣化樣態]')
        app.comparison.path.unlink(missing_ok=True)

    def test_original_current_import_does_not_reuse_or_overwrite_legacy_auto_draft(self):
        from docx import Document
        from lxml import etree
        import json
        from bridgefill.core.comparison import import_current_document
        from bridgefill.core.documents import NS
        original=Document();original._body._body.insert(0,deepcopy(Document(self.f.template).tables[0]._tbl))
        source=self.f.root/'來源.docx'
        table=original.tables[0];headers=[c.text for c in table.rows[1].cells]
        for key,text in [('劣化類型','混凝土剝落'),('備註','A1-L 混凝土剝落。(1m*2m)'),('數量','2'),('單位','平方公尺'),('D','2'),('E','2')]:table.cell(2,headers.index(key)).text=text
        original.save(source)
        old=import_current_document(source,[self.old()],[],self.f.root/'document')
        doc=Document(old)
        for rel in doc.part.rels.values():
            try:root=etree.fromstring(rel.target_part.blob)
            except (AttributeError,etree.XMLSyntaxError):continue
            if root.tag=='{'+NS+'}manifest':
                manifest=json.loads(root.text);manifest.pop('alignment_version',None)
                root.text=json.dumps(manifest,ensure_ascii=False);rel.target_part._blob=etree.tostring(root)
        doc.save(old);before=old.read_bytes()
        new=import_current_document(source,[self.old()],[],self.f.root/'document')
        self.assertNotEqual(old,new)
        self.assertEqual(old.read_bytes(),before)
        self.assertEqual(read_generated(new,strict=False)['alignment_version'],3)
