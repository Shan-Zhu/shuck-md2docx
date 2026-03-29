@echo off
chcp 65001 >nul 2>&1
title shuck-md2docx Installer

:: ============================================
:: Auto-elevate to admin if not already
:: ============================================
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"
echo.
echo ============================================
echo   shuck-md2docx Installer
echo ============================================
echo.

:: ============================================
:: 1. Check Python
:: ============================================
echo [1/4] Checking Python...
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Please install Python 3.8+ from https://www.python.org/downloads/
    echo         Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo         Found: %%i
echo.

:: ============================================
:: 2. Check Pandoc
:: ============================================
echo [2/4] Checking Pandoc...
pandoc --version >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Pandoc is not installed or not in PATH.
    echo         Please install from https://pandoc.org/installing.html
    echo.
    pause
    exit /b 1
)
for /f "tokens=1,2" %%a in ('pandoc --version 2^>^&1') do (
    if "%%a"=="pandoc" echo         Found: pandoc %%b
    goto :pandoc_done
)
:pandoc_done
echo.

:: ============================================
:: 3. Install python-docx
:: ============================================
echo [3/4] Installing python-docx...
pip install python-docx >nul 2>&1
if %errorLevel% neq 0 (
    echo [WARNING] pip install failed, trying with --user flag...
    pip install --user python-docx >nul 2>&1
)
echo         Done.
echo.

:: ============================================
:: 4. Generate and import registry
:: ============================================
echo [4/4] Registering context menu...
python "%~dp0setup.py" >nul 2>&1
if not exist "%~dp0install.reg" (
    echo [ERROR] Failed to generate registry file.
    pause
    exit /b 1
)
reg import "%~dp0install.reg" >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Failed to import registry. Try running as administrator.
    pause
    exit /b 1
)
echo         Done.
echo.

echo ============================================
echo   Installation complete!
echo   Right-click any .md file to convert.
echo ============================================
echo.
pause
