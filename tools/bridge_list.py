"""命令列：縣市橋梁總表依範本工作表轉換格式（實作在 bridgefill/core/list_template.py）。

用法：python tools/bridge_list.py 來源.xlsx 輸出.xlsx [範本工作表名稱，預設 西螺]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from bridgefill.core.list_template import *  # noqa: E402,F401,F403
from bridgefill.core.list_template import convert, railing_formula  # noqa: E402,F401

if __name__ == '__main__':
    for line in convert(*sys.argv[1:4]):
        print(line)
