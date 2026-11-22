"""縣市橋梁清單轉換（tools/bridge_list.py）：舊格式依範本改欄位、資料不遺失、欄杆至河床高度公式（含多跨）。"""
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from bridge_list import convert, railing_formula  # noqa: E402

OLD = ['序號', '二代系統ID', '橋梁編號\n 須修正', '橋梁名稱', '道路等級', '路線', '橋長', '跨越物體', '下橋路徑(ABCD)', '輔助設備(A/B)',
       '歷年持續觀測劣化\n (U≦2)', '歷年持續觀測劣化\n (U≧3)', '115年檢測新增劣化\n (U≦2)', '115年檢測新增劣化\n (U≧3)', '歷年維修',
       '已維修構件狀況', '梁下淨高(m)', '欄杆至河床高度(m)', '欄杆高度(m)', '橋面板厚(m)', '大梁高度(m)', '基本資料變動\n (ch5.1)',
       '應維修構件\n (ch5.2)', '應維修處理建議', '河工設施損壞\n (ch5.4)', '護欄高度不足(V)\n (ch5.5)', '橋下堆積物(V)\n (ch5.6)',
       '疑似鹽害橋梁(V)', '六張基本照片拍攝', '漏電檢測照片拍攝', '橋名牌損壞、遺失、內容錯誤', '是否有限重牌', '其他檢測特殊情形']
NEW = (OLD[:6] + ['跨數', '橋長', '最大淨寬', '橋孔面積\n(橋長*淨寬)', '跨越物體', '下橋路徑(ABCD)', '輔助設備(A/B)',
                  '前期觀測劣化\n(U=1)', '前期觀測劣化\n(U=2)', '前期觀測劣化\n (U≧3)', '115年新增劣化\n (U=1)',
                  '115年新增劣化\n (U=2)', '115年新增劣化\n (U≧3)', '過去維修紀錄\n(列出過去維修紀錄)', '已維修構件狀況',
                  '前期梁下/橋下\n淨高(m)(113年)', '本期梁下/橋下\n淨高(m)(115年)'] + OLD[17:29] + OLD[30:])


def sample(folder):
    workbook = Workbook()
    template = workbook.active
    template.title = '西螺'
    template.append(NEW)
    template.append([1, 1, 'A', '甲橋', None, None, 2, 6, 5.5] + [None] * 13 + [1.1, None, 0.8, 1, 0] + [None] * 11)
    old = workbook.create_sheet('莿桐')
    old.append(OLD)
    old.append([1, 22454, '13-001CT', '莿桐N6', '村里道路', '甘厝村', 30, '新虎尾溪', 'D', None, '1.伸縮縫滲水', '無',
                '1.S3橋面板剝落', '無', '110年度', '無', 'P1=3.5\n P2=3.4', 'P1=4.8\n P2=4.7', 0.7, 0.15, 0.45,
                None, None, None, None, 'v', 'v', None, None, None, None, None, None])
    old.append([2, 22455, '13-002CS', '二崙N1', '村里道路', '', 8, '', None, None, '無', '無', '無', '無', '無', '無',
                'P1:11.1m\n P2:5.5m', None, 0.8, 0.15, 1.75] + [None] * 12)
    path = Path(folder) / 'list.xlsx'
    workbook.save(path)
    return path


class BridgeListTests(unittest.TestCase):
    def test_convert_moves_columns_without_losing_data(self):
        with tempfile.TemporaryDirectory() as folder:
            report = convert(sample(folder), Path(folder) / 'out.xlsx')
            self.assertIn('莿桐：轉換 2 列（沒有對應欄位：漏電檢測照片拍攝）', report)
            self.assertIn('西螺：淨高 1 列從本期欄移到前期欄（舊值為前期量測值）', report)
            book = load_workbook(Path(folder) / 'out.xlsx')
            sheet = book['莿桐']
            heads = [c.value for c in sheet[1]]
            self.assertEqual(heads, NEW)
            row = dict(zip(heads, (c.value for c in sheet[2])))
            self.assertIsNone(row['前期觀測劣化\n(U=1)'])
            self.assertEqual(row['前期觀測劣化\n(U=2)'], '1.伸縮縫滲水')                 # 舊格式 U≦2 → U=2
            self.assertEqual(row['115年新增劣化\n (U=2)'], '1.S3橋面板剝落')
            self.assertEqual(row['過去維修紀錄\n(列出過去維修紀錄)'], '110年度')
            self.assertEqual(row['前期梁下/橋下\n淨高(m)(113年)'], 'P01=3.50\nP02=3.40')     # P 補兩位數、兩位小數
            self.assertEqual((row['護欄高度不足(V)\n (ch5.5)'], row['下橋路徑(ABCD)']), ('v', 'D'))
            self.assertNotIn('漏電檢測照片拍攝', heads)
            self.assertTrue(str(row['欄杆至河床高度(m)']).startswith('=IF('))
            self.assertTrue(str(row['橋孔面積\n(橋長*淨寬)']).startswith('=IF('))
            self.assertIn('A級（間接檢測）', sheet['L1'].comment.text)               # 下橋路徑判定標準
            self.assertIn('B級', sheet['M1'].comment.text)

    def test_formula_length_fits_excel(self):
        formula = railing_formula(2, {'淨高': 'W', '前期淨高': 'V', '欄杆高度': 'Y', '橋面板厚': 'Z', '大梁高度': 'AA'})
        self.assertLess(len(formula), 8192)                    # Excel 公式長度上限
        self.assertNotIn('LET(', formula)                      # 舊版 Excel 也能用
        self.assertNotIn('REGEX', formula)

    @unittest.skipUnless(shutil.which('soffice'), '需要 LibreOffice 計算公式')
    def test_formula_values(self):
        with tempfile.TemporaryDirectory() as folder:
            convert(sample(folder), Path(folder) / 'out.xlsx')
            subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', str(Path(folder) / 'calc'),
                            str(Path(folder) / 'out.xlsx')], capture_output=True, timeout=180)
            book = load_workbook(Path(folder) / 'calc' / 'out.xlsx', data_only=True)
            self.assertEqual(book['西螺']['X2'].value, 2.9)                              # 1.1+0.8+1+0
            self.assertEqual(book['西螺']['J2'].value, 33)                               # 橋孔面積 6×5.5
            self.assertEqual(book['莿桐']['X2'].value, 'P01=4.80\nP02=4.70')              # 多跨逐墩、兩位小數
            self.assertEqual(book['莿桐']['X3'].value, 'P01=13.80\nP02=8.20')             # P1:11.1m 寫法


if __name__ == '__main__':
    unittest.main()


class TemplateImportTests(unittest.TestCase):
    """⑥ 步驟 1：建立空白總表範本、導入前期總表（新增／解除管養、清空本期欄）（2026-11-24）。"""

    def test_new_template(self):
        from bridgefill.core.list_template import new_template
        with tempfile.TemporaryDirectory() as folder:
            path = new_template(Path(folder) / 't.xlsx', 115, ['西螺', '莿桐'])
            book = load_workbook(path)
            self.assertEqual(book.sheetnames, ['西螺', '莿桐'])
            heads = [c.value for c in book['西螺'][1]]
            self.assertIn('115年新增劣化\n(U=1)', heads)
            self.assertEqual(heads[-1], '系統填寫完成')
            self.assertIn('A級', book['西螺']['L1'].comment.text)

    def test_import_prior_with_current_list(self):
        from bridgefill.core.list_template import import_prior, new_template
        with tempfile.TemporaryDirectory() as folder:
            prior = sample(folder)                                     # 西螺：甲橋；莿桐：莿桐N6、二崙N1
            book = load_workbook(prior)
            book['莿桐']['T2'] = '有值'                                  # 本期欄（舊格式沒有，測新格式清空用不到）
            book.save(prior)
            current = new_template(Path(folder) / 'c.xlsx', 116, ['莿桐'])
            listing = load_workbook(current)
            listing['莿桐'].append([1, 22454, None, '莿桐N6'])
            listing['莿桐'].append([2, None, None, '全新橋'])
            listing.save(current)
            report = import_prior(prior, Path(folder) / 'new.xlsx', 116, current)
            self.assertIn('莿桐：新增 全新橋（第 4 列）', report)
            self.assertTrue(any('二崙N1 本期清單沒有' in line for line in report))
            sheet = load_workbook(Path(folder) / 'new.xlsx')['莿桐']
            self.assertEqual([sheet.cell(r, 1).value for r in (2, 3, 4)], [1, '2（解除管養）', '3（新增）'])
            self.assertEqual(sheet.cell(4, 4).value, '全新橋')
            self.assertIn('116年新增劣化\n(U=1)', [c.value for c in sheet[1]])
