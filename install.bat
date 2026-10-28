@echo off
setlocal
cd /d "%~dp0"
title PESRP Scraper - Install

echo ============================================
echo   PESRP School Data Extractor - Installer
echo ============================================
echo.

call :find_python
if defined PY goto :have_python

echo Python was not found. Installing it automatically ...
echo.
call "%~dp0install_python.bat" nopause
call :find_python
if defined PY goto :have_python

echo.
echo [ERROR] Python could not be installed automatically.
echo Install Python 3.10 or newer from https://www.python.org/downloads/
echo IMPORTANT: tick "Add python.exe to PATH" during installation,
echo then run this file again.
echo.
pause
exit /b 1

:have_python
if exist "venv\Scripts\python.exe" goto :install_packages
echo Creating private environment ...
%PY% -m venv venv
if errorlevel 1 (
    echo [ERROR] Could not create the environment.
    pause
    exit /b 1
)

:install_packages
echo Installing required packages. This may take a few minutes ...
"venv\Scripts\python.exe" -m pip install --upgrade pip
"venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Installation failed. Check the internet connection and try again.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Installation complete.
echo   Double-click start.bat to launch the app.
echo ============================================
echo.
pause
exit /b 0

:find_python
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if defined PY exit /b 0
python --version >nul 2>&1 && set "PY=python"
if defined PY exit /b 0
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set PY="%LocalAppData%\Programs\Python\Python312\python.exe"
exit /b 0
