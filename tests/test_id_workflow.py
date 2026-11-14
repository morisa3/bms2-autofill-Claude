from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
from docx import Document
from bridgefill.core.documents import read_generated
from bridgefill.core.scanned_document import export_scan,load_reference
from bridgefill.core.id_document import id_entries
import test_documents


class IdWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.fixture=test_documents.DocumentTests();self.fixture.setUp();self.root=Path(self.fixture.directory.name)
        self.source=self.root/'[2026-11-09][示範橋]本期.docx'
        document=Document(self.fixture.template);table=document.tables[0]
        def fill(row,values):
            for i,value in enumerate(values):row.cells[i].text=str(value)
        fill(table.rows[2],['橋頭','A1','2','2','下載文件錯誤樣態','下載文件工法','99','式','下載文件錯誤備註','2'])
        fill(table.add_row(),['橋尾','A2','1','2','','','','','','1'])
        fill(table.add_row(),['橋尾','A3','2','0','','','','','','1'])
        document.save(self.source)
        self.record=deepcopy(self.fixture.record);self.record['website_id']='2297410'
    def tearDown(self):self.fixture.tearDown()
    def test_scan_is_authoritative_and_download_only_supplements_good_rows_with_visible_ids(self):
        path=export_scan(self.source,'website-job',[self.record],self.root/'document',self.root/'temp')
        self.assertEqual(path.name,self.source.stem+'-掃描資料.docx')
        result=read_generated(path,strict=False);self.assertEqual(len(result['records']),3)
        active=result['records'][0];self.assertEqual(active['website_id'],'2297410')
        self.assertEqual(active['fields']['method'],self.record['current']['method'])
        self.assertNotIn('下載文件',active['fields']['remark'])
        document=Document(path);table=document.tables[active['table']];headers=[c.text for c in table.rows[1].cells]
        self.assertIn('劣化ID',headers);self.assertEqual(table.rows[active['row']].cells[headers.index('劣化ID')].text,'2297410')
        self.assertEqual([e['der'] for e in result['records'][1:]],[['1','2'],['2','0']])
        self.assertTrue(list((self.root/'temp').glob('scan-current-*.json')))
        second=export_scan(self.source,'website-job',[self.record],self.root/'document',self.root/'temp');self.assertNotEqual(second,path)
    def test_raw_reference_does_not_invent_ids_or_apply_rules(self):
        result=load_reference(self.source);self.assertTrue(result['reference_only'])
        self.assertEqual(result['records'][0]['fields']['method'],'下載文件工法')
        self.assertTrue(all(not e.get('website_id') for e in result['records']))
    def test_id_validation_checks_missing_duplicates_and_keeps_good_rows_out(self):
        entry=dict(self.record,fields=self.record['current'])
        good=dict(entry,id='good',website_id='',der=['1','0'],fields={})
        self.assertEqual(list(id_entries({'records':[entry,good]})),['2297410'])
        with self.assertRaisesRegex(ValueError,'劣化 ID'):id_entries({'records':[dict(entry,website_id='')]})
        with self.assertRaisesRegex(ValueError,'重複'):id_entries({'records':[entry,dict(entry,id='other')]})
        with self.assertRaisesRegex(ValueError,'文件待補'):id_entries({'records':[dict(entry,fields={})]})
    def test_scan_scope_is_preserved_and_does_not_treat_unscanned_history_as_deleted(self):
        from bridgefill.core.comparison import ComparisonSession
        from bridgefill.core.data import History
        path=export_scan(self.source,'website-job',[self.record],self.root/'document',self.root/'temp',components=['橋台'])
        prior=self.fixture.history+[History('主梁',dict(self.fixture.fields),1,{})]
        session=ComparisonSession(path,prior,[],temp_dir=self.root/'compare',output_dir=self.root/'out')
        self.assertTrue(all(h.component=='橋台' for h in session.history))
        finished=session.finish('示範橋')
        self.assertEqual(read_generated(finished,strict=False)['scan_scope'],['橋台'])

    def test_scan_rejects_missing_native_id_without_output(self):
        with self.assertRaisesRegex(ValueError,'劣化 ID'):export_scan(self.source,'job',[dict(self.record,website_id='')],self.root/'out',self.root/'temp')
        self.assertFalse((self.root/'out').exists())
