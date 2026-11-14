import json
from pathlib import Path
import tempfile
import unittest
from bridgefill.core.progress import ProgressStore,digest
from bridgefill.core.data import normalize_remark
from bridgefill.core.remarks import prepare_fields,parse
from ui_helpers import ui
from bridgefill.core.config import FIELD_IDS


class PortableProgressTests(unittest.TestCase):
    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        self.root=Path(self.directory.name)
        self.base=self.root/'temp'/'progress.json'

    def tearDown(self):self.directory.cleanup()

    def store(self,job='job'):
        store=ProgressStore(self.base,temporary=True)
        store.begin(job,'rules')
        return store

    def test_checkpoint_survives_stop_and_reload_then_deletes_only_after_completion(self):
        first=self.store();first.track_rows({'row1':'主梁','row2':'主梁'},['主梁'])
        first.mark('record1','done','hash1');first.mark_row('row1')
        first.mark('record2','pending','hash2')
        path=first.path
        self.assertEqual(path.parent,self.root/'temp')
        self.assertTrue(path.is_file())
        again=self.store()
        self.assertTrue(again.done_row('row1'))
        self.assertEqual(again.record('record2')['status'],'pending')
        self.assertFalse(again.complete_scan())
        again.mark('record2','done','hash2');again.mark_row('row2')
        self.assertTrue(again.complete_scan())
        self.assertFalse(path.exists())

    def test_unselected_incomplete_component_prevents_early_deletion(self):
        store=self.store();store.track_rows({'girder':'主梁','pier':'墩柱與帽梁'},['主梁','墩柱與帽梁'])
        store.mark_row('girder')
        resumed=self.store();resumed.track_rows({'girder':'主梁'},['主梁'])
        self.assertFalse(resumed.complete_scan())
        self.assertEqual(resumed.job['open_rows'],{'pier':'墩柱與帽梁'})
        resumed.track_rows({'pier':'墩柱與帽梁'},['墩柱與帽梁']);resumed.mark_row('pier')
        self.assertTrue(resumed.complete_scan())

    def test_different_bridge_jobs_have_separate_files(self):
        first=self.store('bridge1');first.track_rows({'r1':'橋台'});first.mark('pending','pending','hash')
        second=self.store('bridge2');second.track_rows({'r2':'橋台'});second.mark_row('r2')
        self.assertNotEqual(first.path,second.path)
        self.assertTrue(second.complete_scan())
        self.assertTrue(first.path.exists())

    def test_legacy_progress_is_migrated_with_pending_and_original_backup(self):
        legacy=self.root/'legacy'/'progress.json'
        old=ProgressStore(legacy);old.begin('bridge1','rules');old.mark('r1','pending','hash')
        old.begin('bridge2','rules');old.mark('r2','done','hash2')
        backup=ProgressStore.migrate_legacy(legacy,self.base.parent)
        self.assertTrue(backup.is_file());self.assertFalse(legacy.exists())
        store=self.store('bridge1');self.assertEqual(store.record('r1')['status'],'pending')
        self.assertTrue((self.base.parent/('scan-'+digest('bridge2')+'.json')).is_file())
        self.assertIsNone(ProgressStore.migrate_legacy(legacy,self.base.parent))

    def test_corrupt_temporary_file_is_preserved_and_not_reset(self):
        store=self.store();store.path.write_text('broken',encoding='utf-8')
        with self.assertRaises(ValueError):self.store()
        self.assertEqual(store.path.read_text(),'broken')

    def test_existing_new_checkpoint_wins_over_old_backup_on_migration(self):
        current=self.store();current.mark('r1','pending','new-hash')
        legacy=self.root/'legacy.json';old=ProgressStore(legacy);old.begin('job','rules');old.mark('r1','done','old-hash')
        ProgressStore.migrate_legacy(legacy,self.base.parent)
        self.assertEqual(self.store().record('r1')['fingerprint'],'new-hash')

    def test_skipped_is_finished_but_deferred_is_not(self):
        store=self.store();store.track_rows({'r':'橋台'});store.mark('item','skipped','hash');store.mark_row('r')
        self.assertTrue(store.complete_scan())
        store=self.store();store.track_rows({'r':'橋台'});store.defer('item','hash','r','manual')
        self.assertFalse(store.complete_scan());self.assertTrue(store.path.exists())


class November05RuleTests(unittest.TestCase):
    def test_all_requested_location_tails_are_preserved_when_glued(self):
        for tail in ('正','背','正面','左側','右側','左右側','背面','胸牆','耳牆','頰牆','基礎頂','基礎板','基座','基樁','基礎樁帽','帽梁'):
            with self.subTest(tail=tail):
                self.assertEqual(normalize_remark('P10'+tail+'混凝土剝落(1處)'),'P10-'+tail+' 混凝土剝落。(1處)')

    def test_volume_rule_template_is_common_to_multiple_components(self):
        rule=dict(deterioration='其他損傷',pattern='混凝土剝落',method='混凝土修復',unit='立方公尺',price='3500')
        for component in ('橋台','河道','橋面板','墩柱與帽梁','鋼纜錨碇裝置(橋塔端)'):
            with self.subTest(component=component):
                app,controller,page=ui();controller.busy=True
                current={key:'' for key in FIELD_IDS};current['deterioration']='其他損傷'
                app.show_review(dict(component=component,target='P010',current=current,matches=[],repairs=[rule],reason='',suggested=None,
                    context={'編號':'P010'},choices={'deterioration':['其他損傷'],'unit':['立方公尺']}))
                app.select_repair(rule)
                self.assertEqual(app.fields['remark'].value,'P10 混凝土剝落。(m*m*m)')
                app.fields['remark'].value='P10 混凝土剝落。(2m*3m*4m)';app.submit(None)
                self.assertEqual(controller.resolve.call_args.args[0]['amount'],'24')

    def test_new_hole_unit_can_be_parsed_as_count(self):
        rule=dict(pattern='排水孔堵塞',method='疏通',unit='孔',price='200')
        fields={key:'' for key in FIELD_IDS}
        result=prepare_fields(fields,rule,True,component='橋梁排水設施',context={'編號':'P001'})
        self.assertEqual(result['remark'],'P1 排水孔堵塞。(孔)')
        self.assertEqual(parse('P1 排水孔堵塞(2孔)',rule['pattern'],'孔').amount,'2')
