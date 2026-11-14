from copy import deepcopy
import unittest
import test_documents
from ui_helpers import ui
from bridgefill.core.documents import export_document,read_generated,bind_document


class OrderedFillTests(unittest.TestCase):
    def setUp(self):
        self.f=test_documents.DocumentTests();self.f.setUp()

    def tearDown(self):self.f.tearDown()

    def manifest(self,records,drafts=None):
        export_document(self.f.template,self.f.output,'document:source',records,[],self.f.rules,drafts=drafts)
        return read_generated(self.f.output)

    def test_repeated_identical_defects_bind_by_order_and_keep_each_corrected_value(self):
        records=[deepcopy(self.f.record),deepcopy(self.f.record)];records[0]['id']='one';records[1]['id']='two'
        drafts={r['id']:dict(fields=dict(self.f.fields,price=price),reason='人工修訂') for r,price in zip(records,['500','900'])}
        manifest=self.manifest(records,drafts)
        plan=bind_document(manifest,'website-job',records)
        self.assertEqual([plan[r['id']]['price'] for r in records],['500','900'])

    def test_mutable_website_values_do_not_block_authoritative_current_document(self):
        manifest=self.manifest([self.f.record]);current=deepcopy(self.f.record)
        current['current'].update(amount='99',unit='處',method='待修正',remark='備註待修正',price='999')
        plan=bind_document(manifest,'website-job',[current])
        self.assertEqual(plan['first']['price'],'300')
        self.assertEqual(plan['first']['amount'],'2')

    def test_position_der_and_native_id_mismatches_stop_before_any_plan_is_returned(self):
        record=deepcopy(self.f.record);record['website_id']='2297628'
        manifest=self.manifest([record])
        for change in ['position','der','id']:
            current=deepcopy(record)
            if change=='position':current['context']['編號']='A2'
            elif change=='der':current['der']=['3','4']
            else:current['website_id']='2297629'
            with self.assertRaises(ValueError):bind_document(manifest,'website-job',[current])

    def test_automatic_dialog_has_only_current_document_source_and_visible_paced_defaults(self):
        app,c,page=ui();app.open_automatic()
        self.assertEqual([button.text for button in page.dialog.actions],['本期文件導入','返回'])
        self.assertTrue(app.show_fill_page);self.assertEqual(app.fill_delay,.6)
        self.assertEqual(app.automatic_bridge.value,'')
        app.automatic_source('web');c.send.assert_not_called()
        self.assertIn('本期文件導入',page.snack_bar.content.value)
