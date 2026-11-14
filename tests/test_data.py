import unittest
import tempfile
from pathlib import Path
from docx import Document
from bridgefill.core.data import History, candidates, eligible, normalize_remark, propose, repair_candidates, read_history


class DataTests(unittest.TestCase):
    def test_component_with_merged_report_title(self):
        with tempfile.TemporaryDirectory() as directory:
            doc = Document()
            table = doc.add_table(rows=3, cols=4)
            for row, values in zip(table.rows, [
                    ['檢測項目', '檢測項目', '主梁', '主梁'],
                    ['位置', '位置', '劣化類型', '數量'],
                    ['S001', 'S1G1', '其他損傷', '1']]):
                for cell, value in zip(row.cells, values):
                    cell.text = value
            path = Path(directory) / 'test.docx'
            doc.save(path)
            result = read_history(path)
            self.assertEqual(result[0].component, '主梁')
            self.assertEqual(result[0].fields['d_location'], 'S1G1')

    def test_der_gate(self):
        self.assertFalse(eligible("1", "2"))
        self.assertFalse(eligible("2", "0"))
        self.assertTrue(eligible("2", "2"))
        self.assertIsNone(eligible("", "2"))
        self.assertIsNone(eligible("nan", "2"))

    def test_remark_examples(self):
        cases = {
            "P6-1-上游上部 草本植物生長 (1 處)": "P6-1-上游上部 草本植物生長。(1處)",
            "P1-1,2 防護鋼板生銹。(6m＊2m)＊2": "P1-1,2 防護鋼板生銹。(2m*6m)*2",
            "A2-左側 橋名柱傾斜破損鋼筋外露。 (1處)": "A2-左側 橋名柱傾斜破損鋼筋外露。(1處)",
            "A2-背牆 混凝土剝落鋼筋鏽蝕外露(1m*0.5m)*3": "A2-背牆 混凝土剝落鋼筋鏽蝕外露。(0.5m*1m)*3",
            "3G2-0~25 混凝土裂縫。(1m*1mm)*10": "3G2-0~25 混凝土裂縫。(1mm*1m)*10",
            "P2 基礎裸露。(6m*6m*3.0m)": "P2 基礎裸露。(6m*6m*3.0m)",
            "S3-L,R,中間 混凝土剝落鋼筋外露。(35m＊10m)": "S3-L,R,中間 混凝土剝落鋼筋外露。(10m*35m)",
            "1G1-0~35 火害。(35m＊2m)": "1G1-0~35 火害。(2m*35m)",
            "A1 破損 (1m*20cm)*2": "A1 破損。(20cm*1m)*2",
            "A1 破損 (5*2)": "A1 破損。(2*5)",
            "A1 破損 (5m*2)": "A1 破損。(5m*2)",
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(normalize_remark(source), expected)
                self.assertEqual(normalize_remark(expected), expected)

    def test_ambiguous_history_is_not_arbitrarily_selected(self):
        history = [History("橋護欄", {"deterioration": "混凝土剝落", "amount": "2", "d_location": p}, 3)
                   for p in ("A1", "A2")]
        self.assertEqual(len(candidates(history, "橋護欄", {"deterioration": "混凝土剝落", "amount": "2.0"})), 2)
        self.assertEqual(candidates(history, "橋護欄", {"d_location": "A2"}), [history[1]])
        self.assertFalse(candidates(history, "橋護欄", {"amount": "3"}))

    def test_preserve_new_description_and_zero_price(self):
        old = History("橋護欄", {"remark": "A1 舊損傷。(1處)", "price": "500", "method": "修復"}, 3)
        result = propose({"remark": "A1 新損傷 (2 處)", "price": "0"}, old)
        self.assertEqual(result["remark"], "A1 新損傷。(2處)")
        self.assertEqual(result["price"], "0")
        self.assertEqual(result["method"], "修復")

    def test_repairs_keep_unit_ambiguity_and_longest_pattern(self):
        rules = [{"pattern": "混凝土剝落", "unit": "處"},
                 {"pattern": "混凝土剝落鋼筋外露", "unit": "處"},
                 {"pattern": "混凝土剝落鋼筋外露", "unit": "平方公尺"}]
        self.assertEqual(len(repair_candidates(rules, {"remark": "A1 混凝土剝落鋼筋外露。(1處)"})), 2)
        self.assertEqual(repair_candidates(rules, {"remark": "混凝土剝落鋼筋外露", "unit": "處"}), [rules[1]])

if __name__ == "__main__":
    unittest.main()
