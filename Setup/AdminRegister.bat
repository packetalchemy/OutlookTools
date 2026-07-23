@echo off
title OutlookTools - Admin Registration Test
color 0E
echo ========================================================
echo    Admin Registration Test
echo    Run this as ADMINISTRATOR (one time only!)
echo ========================================================
echo.
echo This registers the DLL system-wide using regasm.
echo After this, the add-in should work.
echo.

set DLL_PATH=%LOCALAPPDATA%\OutlookTools\OutlookTools.dll
set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe

if not exist "%REGASM%" (
    echo regasm not found!
    pause
    exit /b 1
)

echo [1] Closing Outlook...
taskkill /f /im outlook.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo [2] Registering DLL (this needs admin)...
"%REGASM%" "%DLL_PATH%" /codebase /tlb
echo.
echo Exit code: %errorLevel%
echo.

echo [3] Starting Outlook...
timeout /t 3 /nobreak >nul
start outlook

echo.
echo If OutlookTools tab appears, the DLL is fine!
echo The issue was HKCU registration only.
echo.
pause
