import tempfile
import unittest
from pathlib import Path
from copy import deepcopy
from docx import Document
from bridgefill.core.data import History
from bridgefill.core.documents import export_document,read_generated,validate_generated,shade


class DocumentTests(unittest.TestCase):
    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        self.root=Path(self.directory.name)
        self.template=self.root/'prior.docx';self.output=self.root/'generated.docx'
        document=Document();document.add_paragraph('檢測表')
        table=document.add_table(rows=3,cols=10);table.style='Table Grid'
        table.cell(0,0).text='檢測項目';table.cell(0,1).merge(table.cell(0,9)).text='橋台'
        for c,t in zip(table.rows[1].cells,['位置','編號','D','E','劣化類型','維修工法','數量','單位','備註','U']):c.text=t
        document.save(self.template)
        self.fields=dict(deterioration='混凝土剝落',d_location='A1-L',method='舊工法',amount='2',unit='平方公尺',price='100',remark='A1-L 混凝土剝落。(1m*2m)')
        self.record=dict(id='first',component='橋台',current=dict(self.fields),context={'位置':'橋頭','編號':'A1','DERU':'D:2 E:2 R:1 U:1'},der=['2','2'],choices={'deterioration':['混凝土剝落'],'unit':['平方公尺']})
        self.history=[History('橋台',self.fields,3,{'位置':'橋頭','編號':'A1'})]
        self.rules=[dict(deterioration='混凝土剝落',components=['橋台'],pattern='混凝土剝落',method='新工法',unit='平方公尺',price='300')]

    def tearDown(self):self.directory.cleanup()

    def test_template_price_placement_style_and_roundtrip(self):
        self.assertEqual(export_document(self.template,self.output,'job',[self.record],self.history,self.rules),(1,0))
        doc=Document(self.output);table=doc.tables[0]
        headers=[c.text for c in table.rows[1].cells]
        self.assertEqual(headers[headers.index('單位')+1],'金額')
        self.assertEqual(table.style.name,'Table Grid')
        self.assertEqual(doc.paragraphs[0].text,'檢測表')
        manifest=read_generated(self.output)
        result=validate_generated(manifest,'job',[self.record])
        self.assertEqual(result['first']['price'],'300')
        self.assertEqual(result['first']['method'],'新工法')
        # Word 開啟再存檔仍保留識別清單。
        doc.save(self.output);self.assertEqual(len(read_generated(self.output)['records']),1)

    def test_no_rule_falls_back_to_prior(self):
        record=deepcopy(self.record);record['current'].update(method='',price='',remark='')
        export_document(self.template,self.output,'job',[record],self.history,[])
        self.assertEqual(read_generated(self.output)['records'][0]['fields'],self.fields)

    def test_yellow_new_row_requires_review_then_accepts_manual_fields(self):
        record=deepcopy(self.record);record['current'].update(remark='',method='',price='')
        export_document(self.template,self.output,'job',[record],[],[])
        with self.assertRaisesRegex(ValueError,'黃色|缺少'):read_generated(self.output)
        doc=Document(self.output);table=doc.tables[0];headers=[c.text for c in table.rows[1].cells]
        table.cell(2,headers.index('備註')).text=self.fields['remark']
        table.cell(2,headers.index('維修工法')).text='人工工法'
        table.cell(2,headers.index('金額')).text='450'
        doc.save(self.output)
        result=validate_generated(read_generated(self.output),'job',[record])
        self.assertEqual(result['first']['method'],'人工工法')
        self.assertEqual(result['first']['d_location'],'A1-L')

    def test_reject_wrong_job_order_changed_source_and_table_identity(self):
        second=deepcopy(self.record);second['id']='second';second['context']['編號']='A2'
        export_document(self.template,self.output,'job',[self.record,second],self.history,self.rules)
        manifest=read_generated(self.output)
        for job,records in [('other',[self.record,second]),('job',[second,self.record])]:
            with self.assertRaises(ValueError):validate_generated(manifest,job,records)
        changed=deepcopy(self.record);changed['current']['amount']='3'
        with self.assertRaises(ValueError):validate_generated(manifest,'job',[changed,second])
        doc=Document(self.output);doc.tables[0].cell(2,1).text='A2';doc.save(self.output)
        with self.assertRaisesRegex(ValueError,'位置或編號'):read_generated(self.output)

    def test_moving_two_rows_with_same_component_identity_is_rejected_by_bookmark(self):
        second=deepcopy(self.record);second['id']='second'
        export_document(self.template,self.output,'job',[self.record,second],self.history,self.rules)
        doc=Document(self.output);table=doc.tables[0]
        table._tbl.append(table.rows[2]._tr)
        doc.save(self.output)
        with self.assertRaisesRegex(ValueError,'列結構'):read_generated(self.output)

    def test_replay_can_select_subset_of_generated_components(self):
        second=deepcopy(self.record);second['id']='second';second['component']='橋護欄'
        rules=self.rules+[dict(self.rules[0],components=['橋護欄'])]
        export_document(self.template,self.output,'job',[self.record,second],self.history,rules)
        manifest=read_generated(self.output)
        result=validate_generated(manifest,'job',[self.record],['橋台'])
        self.assertEqual(set(result),{'first'})
        with self.assertRaisesRegex(ValueError,'不在生成文件'):
            validate_generated(manifest,'job',[],['主梁'])

    def test_added_or_removed_rows_are_rejected_before_any_fill(self):
        export_document(self.template,self.output,'job',[self.record],self.history,self.rules)
        doc=Document(self.output);doc.tables[0].add_row();doc.save(self.output)
        with self.assertRaisesRegex(ValueError,'列數'):read_generated(self.output)
