import unittest
from collections import Counter
from pathlib import Path
from bridgefill.core.data import History,read_repairs
from bridgefill.core.rules_v2 import read_rules
from bridgefill.core.local_files import DEFAULT_REPAIR_NAME,REFERENCE_DIR
from bridgefill.core.alignment import align_records
from bridgefill.core.locations import physical_owner

class Matching1111Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.rules=read_rules(REFERENCE_DIR/DEFAULT_REPAIR_NAME)  # 1116：由結構化欄位還原備註
    def fields(self,owner='P5',remark='積土',amount='8',unit='公尺'):
        return dict(d_location=owner,deterioration='伸縮縫間雜物堆積',remark=owner+' '+remark+' (8m)',amount=amount,unit=unit,method='其他(清除植生)')
    def record(self,owner='P5',**kw):return dict(id=owner,component='伸縮縫',context={'橋墩編號':owner},current=self.fields(owner,**kw))
    def old(self,owner='P5',**kw):return History('伸縮縫',self.fields(owner,**kw),1,{'位置':owner})
    def test_updated_rule_ids_unique(self):
        self.assertGreater(len(self.rules),450)              # 1116：含箱涵、護床工等新增規則
        patterns={}
        for r in self.rules:patterns.setdefault(r['rule_id'],set()).add(r['pattern'])
        self.assertEqual([k for k,v in patterns.items() if len(v)>1],[])   # v3：同一劣化ID只對應一種樣態（多單位、多構件可重複）
        self.assertEqual(next(r['rule_id'] for r in self.rules if r['pattern']=='基礎掏空'),'5-06')   # 1123v1 插入樁帽裸露後重編
    def test_owner_is_exact_and_family_is_inferred(self):
        self.assertEqual(physical_owner({'位置':'橋墩(P)','編號':'005'}),'P5')
        self.assertEqual(physical_owner({'位置':'橋台(A)','編號':'002'}),'A2')
        self.assertEqual(physical_owner({'橋墩編號':'P014'}),'P14')
        rows=align_records([self.record('P14')],[self.old('P001')],repairs=self.rules)
        self.assertIsNone(next(r for r in rows if r['current_index']==0)['history_index'])
    def test_debris_virtual_method_ignores_quantity_unit_difference(self):
        rows=align_records([self.record(remark='積土植生')],[self.old(amount='1',unit='式')],repairs=self.rules)
        self.assertEqual(rows[0]['history_index'],0)
    def test_method_only_multiple_candidates_remain_unpaired(self):
        rows=align_records([self.record(remark='積土植生')],[self.old(amount='1',unit='式'),self.old(amount='2',unit='式')],repairs=self.rules)
        self.assertIsNone(next(r for r in rows if r['current_index']==0)['history_index'])
    def test_removed_same_owner_stays_before_next_owner(self):
        records=[self.record('P5'),self.record('P10')]
        rows=align_records(records,[self.old('P5'),self.old('P5',remark='其他'),self.old('P10')],repairs=self.rules)
        self.assertEqual([(r['current_index'],r['history_index']) for r in rows],[(0,0),(None,1),(1,2)])
    def test_released_history_not_reassigned_automatically(self):
        rows=align_records([self.record()],[self.old()],blocked_history={0},repairs=self.rules)
        self.assertIsNone(next(r for r in rows if r['current_index']==0)['history_index'])
        self.assertEqual(align_records([self.record()],[self.old()],{'P5':0},blocked_history={0})[0]['history_index'],0)
    def test_good_owner_annotation(self):
        rows=align_records([], [self.old()],current_statuses=[dict(component='伸縮縫',context={'位置':'橋墩(P)','編號':'005'},der=['1',''])])
        self.assertEqual(rows[0]['current_message'],'本期改為D=1（良好）')

    def test_component_wide_good_without_pier_number(self):
        old=History('橋梁排水設施',dict(deterioration='排水設施阻塞或功能性降低',d_location='整體-L,R',remark='整體-L,R 阻塞 (57處)'),1,{})
        rows=align_records([], [old],current_statuses=[dict(component='橋梁排水設施',context={},der=['1',''])])
        self.assertEqual(rows[0]['current_message'],'本期改為D=1（良好）')
