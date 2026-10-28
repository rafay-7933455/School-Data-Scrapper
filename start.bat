@echo off
setlocal
cd /d "%~dp0"
title PESRP Scraper

if not exist "venv\Scripts\python.exe" (
    echo The app is not installed yet. Please run install.bat first.
    echo.
    pause
    exit /b 1
)

rem Skip Streamlit's first-run email prompt, which would block startup
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    mkdir "%USERPROFILE%\.streamlit" 2>nul
    > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

echo Starting PESRP School Data Extractor...
echo Your browser will open automatically.
echo Keep this window open while using the app. Close it to stop the app.
echo.

"venv\Scripts\python.exe" -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false

echo.
echo The app has stopped.
pause
