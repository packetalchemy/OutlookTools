@echo off
title OutlookTools Installer v1.2.0
color 0A
echo ============================================
echo    OutlookTools v1.2.0 Installer
echo    Open-Source Outlook Add-in
echo ============================================
echo.

:: No admin rights required — per-user installation

echo [1/4] Finding OutlookTools.dll...
set DLL_PATH=%~dp0OutlookTools.dll
if not exist "%DLL_PATH%" (
    echo [ERROR] OutlookTools.dll not found!
    echo Make sure this installer is in the same folder as OutlookTools.dll
    pause
    exit /b 1
)
echo       Found: %DLL_PATH%

echo.
echo [2/4] Registering DLL with Windows...
regasm "%DLL_PATH%" /codebase /tlb >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Registration failed!
    pause
    exit /b 1
)
echo       Registered successfully!

echo.
echo [3/4] Closing Outlook (if running)...
taskkill /f /im outlook.exe >nul 2>&1
echo       Done.

echo.
echo [4/4] Creating Start Menu shortcut...
set SHORTCUT_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\OutlookTools
mkdir "%SHORTCUT_DIR%" >nul 2>&1
echo OutlookTools installed successfully! > "%SHORTCUT_DIR%\Info.txt"
echo.
echo ============================================
echo    Installation Complete!
echo ============================================
echo.
echo Please open Outlook now.
echo The "OutlookTools" tab will appear.
echo.
echo To uninstall: run Uninstall.bat
echo.
pause
