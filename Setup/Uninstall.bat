@echo off
title OutlookTools Uninstaller
color 0C
echo ============================================
echo    OutlookTools Uninstaller
echo ============================================
echo.

:: Check admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] This uninstaller needs Administrator rights!
    echo Right-click and select "Run as Administrator"
    pause
    exit /b 1
)

echo [1/3] Finding OutlookTools.dll...
set DLL_PATH=%~dp0OutlookTools.dll
if not exist "%DLL_PATH%" (
    echo [WARNING] DLL not found, trying to unregister by name...
    regasm OutlookTools.dll /unregister >nul 2>&1
) else (
    echo       Found: %DLL_PATH%
    echo.
    echo [2/3] Unregistering DLL...
    regasm "%DLL_PATH%" /unregister >nul 2>&1
)
echo       Unregistered!

echo.
echo [3/3] Closing Outlook (if running)...
taskkill /f /im outlook.exe >nul 2>&1
echo       Done.

echo.
echo ============================================
echo    Uninstall Complete!
echo ============================================
echo.
echo Please open Outlook to verify.
echo The OutlookTools tab should be gone.
echo.
pause
