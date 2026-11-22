"""下橋路徑與輔助設備（core.access，3.0）：組合文字、讀回、存檔、比對資料輸出與同步總表。"""
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from bridgefill.core import access as ac
from bridgefill.core import bridge_list as bl


class AccessTextTests(unittest.TestCase):
    def test_user_examples(self):
        c = ac.Access('C', ['A1下游側'], '架設二截梯', ['長筒雨鞋'])
        self.assertEqual(ac.path_text(c), 'C\nA1下游側架設二截梯下橋\n水深需要長筒雨鞋')
        self.assertEqual(ac.aux_text(c), '')
        b = ac.Access('B', equipment=['橋檢車'])
        self.assertEqual(ac.path_text(b), 'B\n需用橋檢車檢測')
        self.assertEqual(ac.aux_text(b), 'B\n橋檢車')

    def test_multiple_positions_and_note(self):
        d = ac.Access('D', ['A2上游側', 'A1上游側'], '既有檢修梯', note='注意蜂窩')
        self.assertEqual(ac.path_text(d), 'D\nA1上游側、A2上游側既有檢修梯下橋\n注意蜂窩')
        self.assertEqual(ac.path_text(ac.Access('D', ['A1上游側'])), 'D\nA1上游側徒步下橋')
        self.assertEqual(ac.path_text(ac.Access()), '')

    def test_parse_round_trip(self):
        for value in (ac.Access('C', ['A1下游側'], '架設二截梯', ['長筒雨鞋']),
                      ac.Access('B', equipment=['橋檢車', '工作船']),
                      ac.Access('A', equipment=['空拍機'], note='河道深'),
                      ac.Access('D', ['A1上游側', 'A2下游側'], '徒步', ['青蛙裝'], '注意')):
            back = ac.parse(ac.path_text(value), ac.aux_text(value))
            self.assertEqual(ac.path_text(back), ac.path_text(value))
            self.assertEqual(ac.aux_text(back), ac.aux_text(value))

    def test_parse_free_text(self):
        self.assertEqual(ac.parse('D\n上下游又爬梯'), ac.Access('D', note='上下游又爬梯'))
        self.assertEqual(ac.parse('', 'B\n橋檢車').level, 'B')
        self.assertEqual(ac.parse('D 徒步').level, 'D')
        self.assertTrue(ac.parse('', '').empty)

    def test_save_load_with_manual_text(self):
        with tempfile.TemporaryDirectory() as folder:
            value = ac.Access('C', ['A1下游側'], '架設一截梯')
            self.assertIsNone(ac.load(folder, '測試橋', '2026-11-05'))
            ac.save(folder, '測試橋', '2026-11-05', value, 'C\n手改的文字', '')
            self.assertEqual(ac.load(folder, '測試橋', '2026-11-05'), value)
            self.assertEqual(ac.saved_texts(folder, '測試橋', '2026-11-05'), ('C\n手改的文字', ''))
            self.assertEqual(ac.saved_texts(folder, '別的橋', '2026-11-05'), ('', ''))


class AccessListTests(unittest.TestCase):
    def test_written_to_list_cells(self):
        from tests.test_bridge_list_sync import HEADS, sheet_file
        with tempfile.TemporaryDirectory() as folder:
            path = sheet_file(folder)
            sources = bl.Sources(22870, {'橋梁名稱': '莿桐N29', '橋梁編號': '13-005CS'}, {}, [], [], [], [], [], [], 113,
                                 [], '', '', None, True, True)
            sources.access_path, sources.access_aux = 'B\n需用橋檢車檢測', 'B\n橋檢車'
            bl.write(path, sources)
            row = [c.value for c in load_workbook(path).active[2]]
            self.assertEqual(row[HEADS.index('下橋路徑(ABCD)')], 'B\n需用橋檢車檢測')
            self.assertEqual(row[HEADS.index('輔助設備(A/B)')], 'B\n橋檢車')
            sources.access_path = sources.access_aux = ''                   # 沒設定：不清掉總表原有內容
            bl.write(path, sources)
            self.assertEqual(load_workbook(path).active[2][HEADS.index('下橋路徑(ABCD)')].value, 'B\n需用橋檢車檢測')


if __name__ == '__main__':
    unittest.main()
