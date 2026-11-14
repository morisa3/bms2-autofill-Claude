import json
from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
import test_documents
from bridgefill.core.comparison import ComparisonSession, import_current_document
from bridgefill.core.documents import read_generated, export_document, validate_generated, bind_document
from ui_helpers import ui


class ComparisonTests(unittest.TestCase):
    def setUp(self):
        self.fixture=test_documents.DocumentTests();self.fixture.setUp();self.f=self.fixture
        self.draft=self.f.root/'橋梁定期檢測表-自動建立.docx'
        export_document(self.f.template,self.draft,'job',[self.f.record],self.f.history,self.f.rules)

    def tearDown(self):self.fixture.tearDown()

    def session(self,source=None):
        return ComparisonSession(source or self.draft,self.f.history,self.f.rules,self.f.root/'temp',self.f.root/'document')

    def test_saved_edit_resumes_and_final_document_keeps_source_identity(self):
        s=self.session();s.apply_rule(dict(self.f.rules[0],price='555'))
        self.assertFalse(s.entries[0]['confirmed'])
        s=self.session();self.assertEqual(s.entries[0]['fields']['price'],'555')
        self.assertTrue(s.finish('測試橋').is_file())
        s.confirm();path=s.finish('測試橋')
        self.assertTrue(path.name.endswith('[測試橋]定期檢測表-完成檔案.docx'))
        self.assertFalse(s.path.exists())
        plan=validate_generated(read_generated(path),'job',[self.f.record])
        self.assertEqual(plan['first']['price'],'555')
        self.assertTrue(self.draft.exists())

    def test_manual_cancel_keeps_original_draft_and_edit_is_prefilled(self):
        app,c,page=ui();app.controller.history=self.f.history;app.controller.repairs=self.f.rules
        app.start_comparison(self.draft)
        self.assertIsNotNone(app.comparison)
        app.open_compare_editor(0)
        self.assertEqual(app.editor_controls['price'].value,'300')
        app.editor_controls['price'].value='999'
        app.close_compare_editor(False)
        self.assertEqual(app.comparison.entries[0]['fields']['price'],'300')
        app.comparison.path.unlink(missing_ok=True)

    def test_rule_in_manual_editor_changes_preview_only_until_completed(self):
        app,c,page=ui();c.history=self.f.history;c.repairs=self.f.rules
        app.start_comparison(self.draft);app.open_compare_editor(0)
        app.compare_apply_rule(dict(self.f.rules[0],price='777'))
        self.assertEqual(app.editor_controls['price'].value,'777')
        app.close_compare_editor(True)
        self.assertEqual(app.comparison.entries[0]['fields']['price'],'777')
        self.assertTrue(app.comparison.entries[0]['confirmed'])
        app.comparison.path.unlink(missing_ok=True)

    def test_incomplete_size_template_cannot_be_confirmed(self):
        s=self.session();fields=dict(s.entries[0]['fields'],remark='A1-L 混凝土剝落。(m*m)')
        s.edit(fields)
        self.assertFalse(s.entries[0]['confirmed'])
        with self.assertRaises(ValueError):s.confirm()
        self.assertTrue(s.path.exists())

    def test_normal_docx_import_applies_rules_and_binds_uniquely_before_fill(self):
        # 移除助手識別，只留標準 DOCX 表格內容。
        from docx import Document
        original=Document(self.draft);normal=Document()
        normal._body._body.insert(0,deepcopy(original.tables[0]._tbl))
        source=self.f.root/'本期.docx';normal.save(source)
        draft=import_current_document(source,self.f.history,self.f.rules,self.f.root/'normal')
        s=self.session(draft);s.confirm();final=s.finish('本期橋')
        manifest=read_generated(final)
        current=deepcopy(self.f.record)
        current['current']=dict(manifest['records'][0]['current'])
        plan=bind_document(manifest,'real-website-job',[current])
        self.assertEqual(plan['first']['price'],'300')
        with self.assertRaises(ValueError):bind_document(manifest,'job',[current,current])
        changed=deepcopy(current);changed['context']['編號']='A9'
        with self.assertRaises(ValueError):bind_document(manifest,'job',[changed])
        changed=deepcopy(current);changed['current']['price']='999'
        self.assertEqual(bind_document(manifest,'job',[changed])['first']['price'],'300')

    def test_zero_rules_and_new_defect_are_red_and_require_manual_confirmation(self):
        export_document(self.f.template,self.draft,'job',[self.f.record],[],[])
        s=ComparisonSession(self.draft,[],[],self.f.root/'temp',self.f.root/'document')
        self.assertEqual(s.entries[0]['color'],'red')
        self.assertTrue(s.finish('測試橋').is_file())
        s.edit(self.f.fields);self.assertTrue(s.finish('測試橋').is_file())

    def test_matching_history_without_rule_is_normal_not_yellow(self):
        export_document(self.f.template,self.draft,'job',[self.f.record],self.f.history,[])
        s=self.session();self.assertEqual(s.entries[0]['color'],'normal')
        self.assertTrue(s.entries[0]['confirmed'])

    def test_ui_rule_paging_and_no_auto_checkbox_in_sidebar(self):
        app,c,page=ui();c.history=self.f.history;c.repairs=[dict(self.f.rules[0],component='橋台',automatic_ready=True,pattern=f'樣態{i}') for i in range(7)]
        self.assertNotIn(app.auto,app.sidebar.content.controls)
        app.start_comparison(self.draft)
        self.assertEqual(len(app.compare_rules.controls),6)
        app.compare_turn(1);self.assertEqual(app.compare_rule_label.value,'7–7 / 7')
        app.compare_turn(1);self.assertEqual(len(app.compare_rules.controls),1)
        app.comparison.path.unlink(missing_ok=True)

    def test_source_changed_during_review_is_not_silently_accepted(self):
        from docx import Document
        session=self.session();before=session.path.read_bytes()
        doc=Document(self.draft);doc.add_paragraph('外部修改');doc.save(self.draft)
        with self.assertRaisesRegex(ValueError,'來源 DOCX'):session.save()
        self.assertEqual(session.path.read_bytes(),before)

    def test_ordinary_import_reuses_same_initial_document_so_edits_resume(self):
        from docx import Document
        document=Document();source=self.f.root/'current.docx'
        document._body._body.insert(0,deepcopy(Document(self.draft).tables[0]._tbl));document.save(source)
        output=import_current_document(source,self.f.history,self.f.rules,self.f.root/'import')
        session=self.session(output);session.apply_rule(dict(self.f.rules[0],price='888'))
        again=import_current_document(source,self.f.history,self.f.rules,self.f.root/'import')
        self.assertEqual(again,output)
        self.assertEqual(self.session(again).entries[0]['fields']['price'],'888')
