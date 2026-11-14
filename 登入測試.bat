@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 公路橋梁填報助手 - 內嵌登入測試

if not exist ".venv\Scripts\python.exe" (
    echo 第一次執行：建立 Python 環境，約需 1 到 3 分鐘...
    py -3.12 -m venv .venv 2>/dev/null || python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
    echo 找不到 Python。請先安裝 Python 3.12（勾選 Add to PATH）後再執行。
    pause
    exit /b 1
)

echo 檢查套件...
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
    echo 套件安裝失敗，請確認網路連線後再試一次。
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m bridgefill.ui.login_test
if errorlevel 1 pause
