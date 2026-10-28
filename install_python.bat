@echo off
setlocal
cd /d "%~dp0"
title Python Setup

set "NOPAUSE=%~1"
set "PYVER=3.12.10"
set "PYURL=https://www.python.org/ftp/python/%PYVER%/python-%PYVER%-amd64.exe"
set "PYEXE=%LocalAppData%\Programs\Python\Python312\python.exe"
set "INSTALLER=%TEMP%\python-%PYVER%-amd64.exe"

echo ============================================
echo   Python Check and Installer
echo ============================================
echo.

py -3 --version >nul 2>&1 && goto :already
python --version >nul 2>&1 && goto :already
if exist "%PYEXE%" goto :already

if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 goto :unsupported

echo Python was not found. Downloading Python %PYVER% from python.org ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYURL%' -OutFile '%INSTALLER%' -UseBasicParsing"
if errorlevel 1 goto :download_failed
if not exist "%INSTALLER%" goto :download_failed

echo Installing Python. Please wait, this takes 1 to 2 minutes ...
start /wait "" "%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_pip=1 Include_test=0 Shortcuts=0
del "%INSTALLER%" >nul 2>&1
if not exist "%PYEXE%" goto :install_failed

echo.
echo Python %PYVER% was installed successfully.
goto :done_ok

:already
echo Python is already installed. Nothing to do.
goto :done_ok

:unsupported
echo [ERROR] 32-bit Windows is not supported. Please use a 64-bit computer.
goto :done_fail

:download_failed
echo.
echo [ERROR] Could not download Python. Check the internet connection and try again.
echo Or download it manually from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH" during installation.
goto :done_fail

:install_failed
echo.
echo [ERROR] Python installation did not complete.
echo Please install it manually from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH" during installation.
goto :done_fail

:done_ok
echo.
if /i not "%NOPAUSE%"=="nopause" pause
exit /b 0

:done_fail
echo.
if /i not "%NOPAUSE%"=="nopause" pause
exit /b 1
