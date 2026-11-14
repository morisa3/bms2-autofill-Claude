from copy import deepcopy
from pathlib import Path
import unittest
from openpyxl import Workbook
import test_documents
from ui_helpers import ui
from bridgefill.core.data import eligible,History,read_repairs
from bridgefill.core.documents import export_document,read_generated
from bridgefill.core.comparison import ComparisonSession
from bridgefill.core.document_scan import build_document_scan
from bridgefill.core.config import FIELD_IDS
from bridgefill.core.progress import digest


class Requirements07Tests(unittest.TestCase):
    def setUp(self):
        self.f=test_documents.DocumentTests();self.f.setUp()

    def tearDown(self):self.f.tearDown()

    def manifest(self,records=None):
        export_document(self.f.template,self.f.output,'job',records or [self.f.record],self.f.history,self.f.rules)
        return read_generated(self.f.output)

    def test_word_der_exclusions_allow_other_column_blank_without_inventing_values(self):
        for d,e in [('1',''),('','0'),('1','bad'),('bad','0')]:self.assertFalse(eligible(d,e))
        for d,e in [('','2'),('2',''),('nan','2'),('bad','2')]:self.assertIsNone(eligible(d,e))
        self.assertTrue(eligible('2','2'))

    def test_excel_rule_ids_are_read_separately_from_website_ids(self):
        book=Workbook();sheet=book.active;sheet.title='橋台'
        sheet.append(['劣化類型','劣化樣態','維修工法','維修單位','單價','備註','劣化ID'])
        sheet.append(['混凝土剝落','混凝土剝落','修復','處',300,'','6-01'])
        path=self.f.root/'rules.xlsx';book.save(path)
        rules=read_repairs(path);self.assertEqual(rules[0]['rule_id'],'6-01')
        self.assertNotIn('website_id',rules[0])

    def test_generated_file_and_completed_document_keep_both_kinds_of_id(self):
        record=dict(self.f.record,website_id='2297410',owner_id='main-row')
        rules=[dict(self.f.rules[0],rule_id='6-01')]
        export_document(self.f.template,self.f.output,'job',[record],self.f.history,rules)
        manifest=read_generated(self.f.output)
        self.assertEqual(manifest['records'][0]['website_id'],'2297410')
        self.assertEqual(manifest['records'][0]['rule_id'],'6-01')
        session=ComparisonSession(self.f.output,self.f.history,rules,self.f.root/'temp',self.f.root/'document')
        session.confirm();output=session.finish('測試橋')
        entry=read_generated(output)['records'][0]
        self.assertEqual((entry['website_id'],entry['rule_id'],entry['owner_id']),('2297410','6-01','main-row'))

    def test_compact_table_has_six_defects_and_continued_component_header(self):
        records=[dict(deepcopy(self.f.record),id=str(i),website_id=str(800000+i)) for i in range(9)]
        self.manifest(records)
        app,c,page=ui();c.history=self.f.history;c.repairs=self.f.rules*7
        app.start_comparison(self.f.output)
        self.assertEqual(len(app.compare_left.controls),6)
        self.assertEqual(len(app.compare_rules.controls),6)
        before=deepcopy(app.comparison.entries)
        app.compare_compact.value=False;app.draw_and_update_comparison()
        self.assertEqual(before,app.comparison.entries)
        self.assertEqual(len(app.compare_right.controls[0].content.controls[-1].controls),7)
        app.compare_section(1)
        self.assertEqual(len(app.compare_left.controls),4)
        self.assertIn('橋台（續）',app.compare_sticky.value)
        self.assertEqual(app.compare_page,1)
        app.comparison.path.unlink(missing_ok=True)

    def test_document_review_uses_document_values_without_reapplying_excel(self):
        app,c,page=ui();c.busy=True;c.history=self.f.history;c.repairs=[dict(self.f.rules[0],price='999')]
        document=History('橋台',dict(self.f.fields,price='300',method='文件工法'),3,self.f.record['context'])
        app.show_review(dict(document_mode=True,website_id='2297410',component='橋台',target='橋台',current=self.f.fields,matches=[document],repairs=c.repairs,
            reason='重複樣態',suggested=document.fields,context=self.f.record['context'],choices=self.f.record['choices']))
        self.assertEqual(app.old_form.content.controls[0].value,'本期修正資料')
        self.assertEqual(app.new_form.content.controls[0].value,'核對結果')
        self.assertEqual(app.fields['price'].value,'300')
        app.submit(None)
        self.assertEqual(c.resolve.call_args.args[0]['price'],'300')

    def test_type_correction_is_next_to_photo_button_on_right(self):
        app,c,page=ui()
        body=app.main_view.controls[1].content
        action=next(row for row in body.controls if hasattr(row,'controls') and app.temporary_button in row.controls)
        self.assertEqual(action.controls.index(app.correction_button)+1,action.controls.index(app.temporary_button))

    def test_guided_scan_single_unique_document_is_automatic(self):
        manifest=self.manifest()
        result=build_document_scan(manifest,'job',[self.f.record])
        self.assertFalse(result['first']['review'])
        self.assertEqual(result['first']['candidates'][0]['fields']['price'],'300')
        manifest['job']='document:unbound'
        self.assertFalse(build_document_scan(manifest,'job',[self.f.record])['first']['review'])
        changed=deepcopy(self.f.record);changed['current']['price']='777'
        self.assertTrue(build_document_scan(manifest,'job',[changed])['first']['review'])
        changed['current']=deepcopy(manifest['records'][0]['fields'])
        self.assertFalse(build_document_scan(manifest,'job',[changed])['first']['review'])

    def test_guided_scan_repeated_style_requires_review_even_when_ids_match(self):
        first=dict(deepcopy(self.f.record),website_id='2297410',owner_id='row')
        second=dict(deepcopy(first),id='second',website_id='2297412')
        manifest=self.manifest([first,second])
        result=build_document_scan(manifest,'job',[first,second])
        self.assertTrue(all(v['review'] for v in result.values()))
        self.assertIn('重複劣化樣態',result['first']['reason'])
        self.assertEqual(len(result['first']['candidates']),1)

    def test_guided_scan_distinct_styles_in_same_component_are_automatic(self):
        first=dict(deepcopy(self.f.record),website_id='2297410',owner_id='row')
        second=dict(deepcopy(first),id='second',website_id='2297412')
        second['current']['remark']='A1-L 鋼筋外露。(1m*2m)'
        export_document(self.f.template,self.f.output,'job',[first,second],self.f.history,self.f.rules)
        manifest=read_generated(self.f.output,strict=False)
        manifest['records'][1]['fields']['remark']='A1-L 鋼筋外露。(1m*2m)'
        result=build_document_scan(manifest,'job',[first,second])
        self.assertFalse(any(v['review'] for v in result.values()))

    def test_guided_scan_missing_duplicate_or_changed_sources_never_guess(self):
        manifest=self.manifest()
        with self.assertRaises(ValueError):build_document_scan(manifest,'other',[self.f.record])
        missing=deepcopy(manifest);missing['records']=[]
        self.assertTrue(build_document_scan(missing,'job',[self.f.record])['first']['review'])
        changed=deepcopy(self.f.record);changed['current']['amount']='9'
        self.assertTrue(build_document_scan(manifest,'job',[changed])['first']['review'])
        ambiguous=deepcopy(manifest);ambiguous['job']='document:unbound'
        ambiguous['records'].append(dict(deepcopy(ambiguous['records'][0]),id='another'))
        self.assertTrue(build_document_scan(ambiguous,'job',[self.f.record])['first']['review'])
