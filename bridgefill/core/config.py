from pathlib import Path
import os
import sys
from .local_files import APP_DIR

BASE_URL = "https://bms2.iot.gov.tw/bms3/"
APP_VERSION = "1.00"
# 備註與數量規則已更新，重新檢查舊版完成項目；pending 仍保留。
RULES_VERSION = "0.9.0"
LOGIN_URL = BASE_URL + "public/public/login/"
DATA_DIR = APP_DIR/"data" if getattr(sys,"frozen",False) else Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "BridgeFill"
FIELD_IDS = ("deterioration", "d_location", "method", "amount", "unit", "price", "remark")
LABELS = dict(zip(FIELD_IDS, ("劣化類型", "損壞位置", "維修工法", "數量", "單位", "單價", "備註")))
# 僅固定構件表格 ID；列、DERU 與動作欄位置皆由實際表格辨識。
COMPONENTS = {
    "引道路堤": "approach_bank",
    "引道路堤護欄": "approach_guard",
    "引道道路堤保護措施": "approach_protect",
    "河道": "waterway", "橋台基礎": "abutment_fund",
    "橋台": "abutment", "翼牆／擋土牆": "wingwall",
    "護床工": "found_protect", "伸縮縫": "expansion",
    "橋梁排水設施": "drainages", "橋護欄": "guardrail",
    "支承／支承墊／阻尼裝置": "bearing", "防落設施": "brake",
    "橋墩／橋基保護設施": "pier_protect", "橋墩基礎": "pier_fund",
    "墩柱與帽梁": "column", "主梁": "girder",
    "橫隔梁": "beam", "橋面板": "slab", "其他": "others",
}
ALIASES = {"引道路堤保護措施": "引道道路堤保護措施", "橋墩墩體／帽梁": "墩柱與帽梁"}

ALIASES.update({"引道路堤保護設施": "引道道路堤保護措施",
                "引道路堤護欄保護措施": "引道道路堤保護措施",
                "橋墩／帽梁": "墩柱與帽梁",
                "橋面版": "橋面板"})

COMPONENTS.update({
    "鋼纜錨碇裝置(橋塔端)":"pylon_anchor",
    "鋼纜錨碇裝置(主梁端)":"stay_cable",
    "鋼纜(吊索)-斜張鋼纜":"stay_cable",
    "鋼纜保護套管-斜張鋼纜":"stay_cable",
})
# 箱涵構件（2026-11-15，鹿-13 網頁確認）
COMPONENTS.update({"端牆": "headwall", "中隔板": "midwall", "頂板": "top_slab", "底板": "bottom_slab"})
COMPONENT_SECTIONS={
    "鋼纜錨碇裝置(橋塔端)":['x_pylon_anchor'],
    "鋼纜錨碇裝置(主梁端)":['x_girder_anchor','x_stay_cable_anchor'],
    "鋼纜(吊索)-斜張鋼纜":['x_stay_cable'],
    "鋼纜保護套管-斜張鋼纜":['x_stay_cable_cover'],
}
SHARED_REPAIR_SHEETS={
    "鋼纜錨碇裝置":["鋼纜錨碇裝置(橋塔端)","鋼纜錨碇裝置(主梁端)"],
    "斜張鋼纜":["鋼纜(吊索)-斜張鋼纜","鋼纜保護套管-斜張鋼纜"],
}
