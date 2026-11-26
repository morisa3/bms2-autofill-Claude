# 公路橋梁填報助手（v4.0）

協助檢測員在「車行橋梁管理資訊系統」（bms2.iot.gov.tw）比對前後期劣化、補齊本期資料、生成檢測員意見（詳細版／精簡版）、上傳基本照片、依劣化 ID 回填網站，並支援本機 Excel 與 Google 試算表多人即時同步。

- 討論確認的需求與規則：[docs/規格.md](docs/規格.md)
- 完整操作說明：[docs/操作說明.md](docs/操作說明.md)
- 版本紀錄（目前 `v4.0.000.0`）：[docs/版本紀錄.md](docs/版本紀錄.md)
- Google 試算表 Apps Script 同步設定：[docs/Google試算表同步設定.md](docs/Google試算表同步設定.md)（腳本：[tools/google_sheets/Code.gs](tools/google_sheets/Code.gs)）
- 劣化維修規則 Excel（`data/劣化說明與維修單價v3.xlsx`，直接維護）
- **主程式**：雙擊 `填報助手.vbs`（不開 CMD 視窗）或 `填報助手.bat`

## 架構

```
bridgefill/
  core/   純邏輯：v2/v3 Excel 規則、8 張設定工作表、備註格式、配對、意見生成（詳細／精簡）、
          橋梁總表 38 欄同步、Google Apps Script 同步、排程監看與重複橋名分流
          （嚴格限制：不得匯入 PySide6 或任何瀏覽器模組）
  site/   網站頁面判讀（純邏輯）：登入狀態、網頁解析（pages.py）、唯讀掃描流程與白名單（scan.py）
  ui/     PySide6 介面：主視窗（①～⑥ 分頁＋狀態徽章）、內嵌瀏覽器、隱藏頁面讀取器
```

## 開啟

- 單一執行檔：雙擊 `公路橋梁填報助手-v4.0.000.0.exe`。
- 原始碼版：雙擊 **`填報助手.vbs`**（不開 CMD 視窗；第一次會先以 `填報助手.bat` 建立環境）。
- 版本號：`4.x.yyy.0`（x 新功能、yyy 錯誤修正），見 [docs/版本紀錄.md](docs/版本紀錄.md)。

## 單一執行檔（Windows）

需要正式封裝時，手動執行 GitHub Actions「建置 Windows 單一執行檔」（不會在每次推送時自動打包），在 Windows 上打包（PyInstaller onefile），
並以 `--self-test` 確認內嵌瀏覽器能啟動；完成後在該次執行的 Artifacts 下載「公路橋梁填報助手-v版本號」（含 exe、操作說明、版本紀錄、Google 試算表設定與 `Code.gs`、`data\` 規則表）。

## 執行測試

```bash
pip install -r requirements-dev.txt
python -m pytest -q --ignore=tests/test_main_window.py --ignore=tests/test_browser_behavior.py --ignore=tests/test_filler.py --ignore=tests/test_login_window.py --ignore=tests/test_scan_window.py
```
