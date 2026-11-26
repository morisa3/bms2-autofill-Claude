from copy import deepcopy
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import json
import unittest
import test_documents
from ui_helpers import ui
from bridgefill.core.alignment import align_records
from bridgefill.core.comparison import ComparisonSession
from bridgefill.core.data import History,read_repairs
from bridgefill.core.documents import export_document,read_generated,validate_generated
from bridgefill.core.local_files import REFERENCE_DIR,DEFAULT_REPAIR_NAME
from bridgefill.core.naming import bridge_name,dated_bridge_name


class Requirements076Tests(unittest.TestCase):
    def setUp(self):
        self.f=test_documents.DocumentTests();self.f.setUp()

    def tearDown(self):self.f.tearDown()

    def session(self,records=None,history=None,rules=None):
        history=self.f.history if history is None else history
        rules=self.f.rules if rules is None else rules
        export_document(self.f.template,self.f.output,'job',records or [self.f.record],history,[],review_colors=True)
        return ComparisonSession(self.f.output,history,rules,self.f.root/'temp',self.f.root/'out')

    def test_manual_pair_applies_current_rule_and_preserves_new_range_and_source(self):
        record=deepcopy(self.f.record)
        record['current'].update(remark='A2 混凝土剝落。(3m*4m)',amount='999',d_location='錯位置')
        old=deepcopy(self.f.history[0]);old.context={'位置':'橋尾','編號':'A2'}
        s=self.session([record],[old])
        s.map_history(0,0)
        fields=s.entries[0]['fields']
        self.assertEqual((fields['d_location'],fields['amount'],fields['price']),('A2','12','300'))
        self.assertIn('3m*4m',fields['remark'])
        self.assertEqual(s.entries[0]['current']['amount'],'999')
        self.assertEqual(s.entries[0]['fields']['method'],'新工法')

    def test_importing_assistant_current_document_corrects_rules_without_clicking(self):
        s=self.session()
        self.assertEqual(s.entries[0]['fields']['price'],'300')
        self.assertEqual(s.entries[0]['fields']['method'],'新工法')

    def test_explicit_previous_import_recovers_junk_current_and_reapplies_rule(self):
        record=deepcopy(self.f.record);record['current'].update(remark='亂填',amount='',unit='',price='')
        s=self.session([record]);s.map_history(0,0)
        # 不會在配對時擅自把無法解析的本期改成前期。
        self.assertEqual(s.entries[0]['fields']['remark'],'亂填')
        fields=s.import_previous()
        self.assertEqual((fields['amount'],fields['price'],fields['method']),('2','300','新工法'))
        self.assertEqual(s.entries[0]['current']['remark'],'亂填')

    def test_backup_survives_temp_removal_and_retains_manual_mapping_and_edits(self):
        s=self.session();s.map_history(0,0)
        s.edit(dict(s.entries[0]['fields'],price='777'))
        data=json.loads(s.backup_path.read_text(encoding='utf-8'))
        self.assertEqual(data['manual_pairs'],{'first':0})
        self.assertTrue(data['history']);s.path.unlink()
        restored=ComparisonSession(self.f.output,self.f.history,self.f.rules,self.f.root/'temp',self.f.root/'out')
        self.assertEqual(restored.pairs,{0:0})
        self.assertEqual(restored.entries[0]['fields']['price'],'777')
        restored.finish('測試橋');self.assertTrue(restored.backup_path.exists())
        self.assertFalse(restored.path.exists())

    def test_incomplete_export_succeeds_and_strict_web_import_still_rejects_it(self):
        s=self.session(history=[],rules=[])
        s.edit(dict(s.entries[0]['fields'],price=''))
        output=s.finish('測試橋')
        self.assertTrue(output.exists());self.assertTrue(s.new_defects_path.exists())
        self.assertTrue(s.path.exists());self.assertEqual(s.issues()[0]['detail'],'待補：單價')
        self.assertEqual(read_generated(output,strict=False)['records'][0]['fields']['price'],'')
        with self.assertRaises(ValueError):read_generated(output)

    def test_incomplete_ui_export_lists_actionable_fields_without_starting_web_fill(self):
        s=self.session(history=[],rules=[]);s.edit(dict(s.entries[0]['fields'],price=''))
        app,c,page=ui();c.import_generated=Mock()
        app.start_comparison(self.f.output);app.comparison=s
        app.finish_comparison()
        c.import_generated.assert_not_called()
        self.assertTrue(page.dialog.open)
        self.assertEqual(page.dialog.title.value,'文件已儲存，以下欄位仍待補')
        issue=page.dialog.content.content.controls[-1]
        self.assertIn('第 1 筆｜橋台｜A1-L｜待補：單價',issue.text)
        issue.on_click(None)
        self.assertEqual(app.editor_controls['price'].value,'')

    def test_exported_manual_pairs_survive_import_of_completed_file(self):
        record=deepcopy(self.f.record);record['current']['remark']='A1-L 漏水。(2m)'
        s=self.session([record],rules=[]);s.map_history(0,0)
        output=s.finish('測試橋')
        restored=ComparisonSession(output,self.f.history,[],self.f.root/'temp',self.f.root/'out')
        self.assertEqual(restored.manual_pairs,{'first':0})
        self.assertEqual(restored.pairs,{0:0})

    def test_new_report_contains_only_unpaired_current_grouped_by_component(self):
        records=[]
        for i,component in enumerate(['橋台','橋護欄','橋台','橋台']):
            r=deepcopy(self.f.record);r.update(id=str(i),component=component)
            if i!=3:r['current']['remark']='A1-L 木本植物生長。(1處)'
            r['choices']={};records.append(r)
        s=self.session(records, rules=[])
        s.finish('測試橋')
        from docx import Document
        from bridgefill.core.documents import template_component
        report=Document(s.new_defects_path)
        self.assertEqual([(template_component(t),len(t.rows)-2) for t in report.tables if template_component(t)],[('橋台',2),('橋護欄',1)])
        self.assertEqual(len(read_generated(next((self.f.root/'out').glob('*完成檔案.docx')),strict=False)['records']),4)

    def test_foundations_long_notes_and_concrete_synonyms_match_without_cross_reuse(self):
        context={'位置':'P006','編號':'P06N'}
        kind='混凝土剝落、破碎、鋼筋外露、銹蝕'
        old=[History('橋墩基礎',dict(self.f.fields,deterioration=kind,d_location='P6',amount='25',remark='P6 橋墩基礎混凝土破損鋼筋外露(5m*5m)'),3,context),
             History('橋墩基礎',dict(self.f.fields,deterioration='基礎沖刷、裸露、掏空',d_location='P6基礎裸露',amount='150',remark='P6基礎裸露 (10m*6m*2.5m)，依竣工圖說，樁基礎深度(含樁帽)27.2m'),4,context)]
        records=[dict(id='a',component='橋墩基礎',context=context,current=dict(old[1].fields,amount='1',unit='式',remark='P6基礎裸露 (10m*6m*2.4m)')),
                 dict(id='b',component='橋墩基礎',context=context,current=dict(old[0].fields,remark='P6 混凝土剝落鋼筋外露(5m*5m)'))]
        self.assertEqual([r['history_index'] for r in align_records(records,old)],[1,0])
        # 明確不同劣化仍然不配對。
        records[1]['current']['remark']='P6 木本植物生長(1處)'
        self.assertEqual(align_records(records,old)[1]['kind'],'new')

    def test_date_and_bridge_name_do_not_reuse_source_date_or_drop_brackets(self):
        self.assertEqual(bridge_name('[2025-11-24][龍門大橋]定期檢測表.docx'),'龍門大橋')
        self.assertEqual(dated_bridge_name('[2025-11-24][龍門大橋]定期檢測表.docx',datetime(2026,11,7)),'[2026-11-07][龍門大橋]')

    def test_editor_syncs_location_amount_and_larger_labels_before_edit(self):
        s=self.session();s.entries[0]['fields'].update(d_location='錯位置',amount='999',remark='A2 混凝土剝落。(3m*4m)');s.save()
        app,c,page=ui();c.history=self.f.history;c.repairs=self.f.rules
        app.start_comparison(self.f.output)
        app.comparison.entries[0]['fields'].update(d_location='錯位置',amount='999',remark='A2 混凝土剝落。(3m*4m)')
        app.open_compare_editor(0)
        self.assertEqual(app.editor_controls['d_location'].value,'A2')
        self.assertEqual(app.editor_controls['amount'].value,'12')
        self.assertGreaterEqual(app.editor_controls['amount'].label_style.size,14)
        app.editor_controls['remark'].value='A1-L 混凝土剝落。(2m*5m)'
        app.editor_fields_changed(SimpleNamespace(control=app.editor_controls['remark']))
        self.assertEqual(app.editor_controls['amount'].value,'10')
        app.comparison.path.unlink(missing_ok=True)

    def test_latest_bundled_rules_are_readable(self):
        from bridgefill.core.rules_v2 import read_rules
        self.assertEqual(DEFAULT_REPAIR_NAME,'劣化說明與維修單價1124.xlsx')
        rules=read_rules(REFERENCE_DIR/DEFAULT_REPAIR_NAME)
        self.assertGreater(len(rules),500)

    def test_startup_is_a_normal_window(self):
        app,c,page=ui()
        self.assertFalse(page.window_maximized)
        self.assertFalse(page.window_full_screen)
        self.assertEqual((page.window_width,page.window_height),(1680,1050))
