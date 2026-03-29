@echo off
chcp 65001 >nul 2>&1
title shuck-md2docx Uninstaller

:: Auto-elevate to admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"
echo.
echo ============================================
echo   shuck-md2docx Uninstaller
echo ============================================
echo.

:: Generate uninstall.reg if missing
if not exist "%~dp0uninstall.reg" (
    python "%~dp0setup.py" >nul 2>&1
)

if exist "%~dp0uninstall.reg" (
    reg import "%~dp0uninstall.reg" >nul 2>&1
    echo Context menu entry removed.
) else (
    reg delete "HKCR\SystemFileAssociations\.md\shell\md2docx" /f >nul 2>&1
    echo Context menu entry removed (direct).
)

echo.
echo ============================================
echo   Uninstall complete.
echo ============================================
echo.
pause
