from copy import deepcopy
from pathlib import Path
from unittest.mock import patch
import json
import unittest
from bridgefill.core.alignment import align_records
from bridgefill.core.conditions import aliases,condition,keyword_groups
from bridgefill.core.comparison import ComparisonSession,import_current_document
from bridgefill.core.data import History,read_repairs,repair_candidates
from bridgefill.core.documents import export_document,read_generated
from bridgefill.core.rules_v2 import read_rules
from bridgefill.core.local_files import REFERENCE_DIR,DEFAULT_REPAIR_NAME
from bridgefill.core.locations import key
from bridgefill.core.policy import automatic_plan
import test_documents
from ui_helpers import ui


class November8PatchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.rules=read_rules(REFERENCE_DIR/DEFAULT_REPAIR_NAME)  # 1116：由結構化欄位還原備註

    def setUp(self):
        self.f=test_documents.DocumentTests();self.f.setUp()

    def tearDown(self):self.f.tearDown()

    def protection(self,remark='',number='P010',kind='防撞鋼板損傷'):
        return dict(id='p10',component='橋墩／橋基保護設施',context={'位置':number},der=['2','4'],choices={},current=dict(self.f.fields,deterioration=kind,d_location=number,amount='1' if not remark else '2',unit='式' if not remark else '平方公尺',price='',remark=remark))

    def previous(self,number='P10',kind='防撞鋼板損傷'):
        return History('橋墩／橋基保護設施',dict(self.f.fields,deterioration=kind,d_location=number,remark=f'{number} 生鏽。(1m*2m)'),3,{'位置':number})

    def test_zero_padding_matches_piers_abutments_and_preserves_suffixes(self):
        for a,b in [('P010','P10'),('A001','A1'),('S001','S1'),('P013N-2','P13N-2')]:self.assertEqual(key(a),key(b))
        self.assertNotEqual(key('P010'),key('P011'))
        self.assertNotEqual(key('P010-L'),key('P10-R'))

    def test_blank_specific_defect_pairs_unique_same_place_without_borrowing_quantity(self):
        record=self.protection();old=self.previous()
        rows=align_records([record],[old])
        self.assertEqual((rows[0]['kind'],rows[0]['history_index']),('matched',0))
        export_document(self.f.template,self.f.output,'job',[record],[old],self.rules,review_colors=True)
        session=ComparisonSession(self.f.output,[old],self.rules,self.f.root/'temp')
        self.assertEqual(session.pairs,{0:0})
        self.assertEqual(session.entries[0]['fields']['amount'],'1')
        self.assertEqual(session.entries[0]['fields']['remark'],'')

    def test_blank_does_not_guess_between_duplicates_or_different_types(self):
        old=self.previous()
        for records,history in [([self.protection()],[old,deepcopy(old)]),([self.protection()],[self.previous(kind='其他損傷')])]:
            self.assertEqual(align_records(records,history)[0]['kind'],'ambiguous')
        record=self.protection(number='P011')
        self.assertEqual(align_records([record],[old])[0]['kind'],'new')

    def test_context_p7_with_remark_p4_is_not_paired_to_p7_or_p4(self):
        record=self.protection('P4 生鏽。(5m*8m)*2','P007')
        rows=align_records([record],[self.previous('P7'),self.previous('P4')])
        self.assertEqual(rows[0]['kind'],'ambiguous')
        self.assertIn('位置不一致',rows[0]['problem'])
        # 人工明確指定後可保留其選擇。
        self.assertEqual(align_records([record],[self.previous('P7')],{'p10':0})[0]['kind'],'matched')

    def test_1107_and_1108_correction_wording_both_apply_price_and_target_type(self):
        # 同時覆蓋 1107 原文與 1108 加入「即可使用」的版本。
        base=next(r for r in self.rules if r['rule_id']=='11-08')
        for note in ['若判讀到關鍵字[生銹]或[生鏽]，但劣化類型填成[其他損傷]，則修正成[防撞鋼板損傷]後使用',base['note']]:
            for word in ['生鏽','生銹','生锈']:
                result=automatic_plan(dict(self.f.fields,deterioration='其他損傷',remark=f'P4 {word}。(5m*8m)*2'),'橋墩／橋基保護設施',[],[dict(base,note=note)],{'位置':'P004'})
                self.assertFalse(result.review,result.reason)
                self.assertEqual((result.rule_id,result.fields['deterioration'],result.fields['price'],result.fields['amount']),('11-08','防撞鋼板損傷','200','80'))

    def test_correction_type_names_are_not_defect_keywords_or_aliases(self):
        rule=next(r for r in self.rules if r['rule_id']=='11-08')
        self.assertNotIn('其他損傷',aliases(rule['note']))
        self.assertNotIn('防撞鋼板損傷',aliases(rule['note']))
        self.assertEqual(repair_candidates([rule],{'remark':'P4 其他損傷。(1m*2m)'}),[])
        self.assertEqual(condition(dict(rule,note=rule['note']+'，若水位異常才使用'),{},'P4 生鏽。(1m*2m)'),(None,True))

    def test_canonical_rust_pattern_keeps_pair_after_rule_replacement(self):
        record=self.protection('P10 防衝鋼板表面生銹。(1m*2m)')
        old=self.previous()
        self.assertEqual(align_records([record],[old])[0]['kind'],'matched')

    def test_manual_foundation_rule_supplies_known_price_without_treating_volume_as_count(self):
        for remark in ['P3 基礎裸露。(10m*6m*1.6m)','P5 基礎樁帽裸露。(10m*6m*2.7m)']:
            fields=dict(self.f.fields,deterioration='基礎沖刷、裸露、掏空',remark=remark,amount='1',unit='式',price='')
            manual_rules=[dict(r,note='手動處理') if r['rule_id']=='5-01' else r for r in self.rules]
            result=automatic_plan(fields,'橋墩基礎',[],manual_rules,{'位置':'P003'})
            self.assertTrue(result.review)
            self.assertEqual((result.rule_id,result.fields['price'],result.fields['amount'],result.fields['unit']),('5-01','2000','1','式'))
            self.assertIn('指定手動',result.reason)
            record=dict(self.f.record,component='橋墩基礎',current=fields,choices={})
            export_document(self.f.template,self.f.output,'job',[record],[],manual_rules)
            entry=read_generated(self.f.output,strict=False)['records'][0]
            self.assertEqual(entry['rule_id'],'5-01')
            self.assertEqual(entry['fields']['price'],'2000')
            self.assertTrue(entry['needs_review'])

    def test_1108v1_foundation_keyword_uses_explicit_count_without_multiplying_dimensions(self):
        for remark,location in [('P3 基礎裸露。(10m*6m*1.6m)','P3'),('P4橋墩基礎裸露基樁掏空(10m*6m*3m)','P4'),('P5 基礎樁帽裸露。(10m*6m*2.7m)','P5'),('P6基礎裸露 (10m*6m*2.4m)','P6')]:
            fields=dict(self.f.fields,deterioration='基礎沖刷、裸露、掏空',remark=remark,amount='1',unit='式',price='')
            result=automatic_plan(fields,'橋墩基礎',[],self.rules,{'位置':location})
            self.assertFalse(result.review,result.reason)
            self.assertEqual((result.rule_id,result.fields['price'],result.fields['amount'],result.fields['unit'],result.fields['d_location']),('5-01','2000','1','式',location))
        for amount in ['', '0','不明']:
            result=automatic_plan(dict(fields,amount=amount),'橋墩基礎',[],self.rules,{'位置':'P6'})
            self.assertTrue(result.review)

    def test_old_auto_cache_is_refreshed_but_manual_values_survive_excel_change(self):
        record=self.protection('P10 生鏽。(1m*2m)',kind='其他損傷');old=self.previous()
        export_document(self.f.template,self.f.output,'job',[record],[old],[])
        folder=self.f.root/'temp'
        s=ComparisonSession(self.f.output,[old],self.rules,folder)
        saved=json.loads(s.path.read_text());saved['comparison_version']=3
        saved['entries'][0]['fields']['price']='';saved['entries'][0]['reason']='待人工：舊判讀'
        s.path.write_text(json.dumps(saved))
        s=ComparisonSession(self.f.output,[old],self.rules,folder)
        self.assertEqual(s.entries[0]['fields']['price'],'200')
        s.edit(dict(s.entries[0]['fields'],price='777'))
        changed=[dict(r,price='201') if r['rule_id']=='11-06' else r for r in self.rules]
        s=ComparisonSession(self.f.output,[old],changed,folder)
        self.assertEqual(s.entries[0]['fields']['price'],'777')

    def test_live_rule_refresh_updates_auto_price_and_keeps_manual_values(self):
        record=self.protection('P10 生鏽。(1m*2m)')
        export_document(self.f.template,self.f.output,'job',[record],[self.previous()],[])
        s=ComparisonSession(self.f.output,[self.previous()],self.rules,self.f.root/'temp')
        changed=[dict(r,price='250') if r['rule_id']=='11-06' else r for r in self.rules]
        s.refresh_repairs(changed)
        self.assertEqual((s.entries[0]['rule_id'],s.entries[0]['fields']['price']),('11-06','250'))
        s.edit(dict(s.entries[0]['fields'],price='777'))
        s.refresh_repairs(self.rules)
        self.assertEqual(s.entries[0]['fields']['price'],'777')

    def test_rules_final_page_is_not_stretched_and_help_is_visible(self):
        app,c,page=ui();c.history=self.f.history;c.repairs=[dict(self.f.rules[0],pattern=f'樣態{i}') for i in range(8)]
        export_document(self.f.template,self.f.output,'job',[self.f.record],self.f.history,[])
        app.start_comparison(self.f.output)
        heights=[control.height for control in app.compare_rules.controls]
        app.compare_turn(1)
        self.assertEqual(len(app.compare_rules.controls),2)
        self.assertEqual([control.height for control in app.compare_rules.controls],heights[:2])
        self.assertTrue(all(not control.expand for control in app.compare_rules.controls))
        self.assertIn('先選本期列',app.compare_rule_panel.content.controls[1].value)
        app.comparison.path.unlink(missing_ok=True)

    def test_plain_document_import_across_dates_reuses_original_draft_and_manual_edits(self):
        from docx import Document
        export_document(self.f.template,self.f.output,'job',[self.f.record],self.f.history,[])
        normal=Document();normal._body._body.insert(0,deepcopy(Document(self.f.output).tables[0]._tbl))
        source=self.f.root/'本期來源.docx';normal.save(source)
        directory=self.f.root/'documents'
        with patch('bridgefill.core.comparison.dated_bridge_name',return_value='[2026-11-07][示範橋]'):
            original=import_current_document(source,self.f.history,self.f.rules,directory)
        s=ComparisonSession(original,self.f.history,self.f.rules,self.f.root/'temp')
        s.edit(dict(s.entries[0]['fields'],price='777'))
        with patch('bridgefill.core.comparison.dated_bridge_name',return_value='[2026-11-08][示範橋]'):
            resumed=import_current_document(source,self.f.history,self.f.rules,directory)
        self.assertEqual(resumed,original)
        restored=ComparisonSession(resumed,self.f.history,self.f.rules,self.f.root/'temp')
        self.assertEqual(restored.entries[0]['fields']['price'],'777')
        before=original.read_bytes()
        normal.add_paragraph('新資料');normal.save(source)
        with patch('bridgefill.core.comparison.dated_bridge_name',return_value='[2026-11-07][示範橋]'):
            different=import_current_document(source,self.f.history,self.f.rules,directory)
        self.assertNotEqual(different,original)
        self.assertEqual(original.read_bytes(),before)
