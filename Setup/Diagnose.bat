@echo off
title OutlookTools - Diagnostic Tool
color 0B
echo ========================================================
echo    OutlookTools Diagnostic
echo ========================================================
echo.

echo [1/5] Checking DLL location...
set DLL_PATH=%LOCALAPPDATA%\OutlookTools\OutlookTools.dll
if exist "%DLL_PATH%" (
    echo   [OK] DLL found: %DLL_PATH%
    dir "%DLL_PATH%" | find "OutlookTools.dll"
) else (
    echo   [FAIL] DLL NOT FOUND at %DLL_PATH%
    echo   Copy OutlookTools.dll to this folder first!
)
echo.

echo [2/5] Checking CLSID registry...
reg query "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /s 2>nul
if %errorLevel% neq 0 (
    echo   [FAIL] CLSID not found in HKCU\Software\Classes
) else (
    echo   [OK] CLSID found
)
echo.

echo [3/5] Checking ProgId registry...
reg query "HKCU\Software\Classes\OutlookTools.AddIn" /s 2>nul
if %errorLevel% neq 0 (
    echo   [FAIL] ProgId not found
) else (
    echo   [OK] ProgId found
)
echo.

echo [4/5] Checking Outlook Add-in entry...
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /s 2>nul
if %errorLevel% neq 0 (
    echo   [FAIL] Outlook add-in entry not found
) else (
    echo   [OK] Outlook entry found
)
echo.

echo [5/5] Checking HKLM (system-wide) for conflicts...
reg query "HKLM\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /s 2>nul
if %errorLevel% equ 0 (
    echo   [WARN] HKLM has a CLSID entry - may conflict!
    echo   Run as Admin to clean: regsvr32 /u "%DLL_PATH%"
) else (
    echo   [OK] No HKLM conflict
)
echo.

echo ========================================================
echo    Checking if regasm can load the DLL...
echo ========================================================
echo.
set REGASM64=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
set REGASM32=C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe

if exist "%REGASM64%" (
    echo   Testing 64-bit regasm...
    "%REGASM64%" "%DLL_PATH%" /? 2>nul | find "RegAsm" >nul
    echo   64-bit regasm available
)

if exist "%REGASM32%" (
    echo   Testing 32-bit regasm...
    echo   32-bit regasm available
)
echo.

echo ========================================================
echo    Summary: If any [FAIL] above, that's the problem!
echo ========================================================
echo.
echo To fix manually, run this in PowerShell (as Admin):
echo   regasm "%DLL_PATH%" /codebase /tlb
echo.
echo Or double-click RegisterOutlookTools.reg
echo.
pause
