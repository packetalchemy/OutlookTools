@echo off
title OutlookTools - Full Diagnostic
color 0B
echo ========================================================
echo    OutlookTools - Full Diagnostic Report
echo    Run this and send me the output!
echo ========================================================
echo.

echo === [1] DLL File Check ===
set DLL_PATH=%LOCALAPPDATA%\OutlookTools\OutlookTools.dll
if exist "%DLL_PATH%" (
    echo   [OK] DLL exists: %DLL_PATH%
    dir "%DLL_PATH%" | find "OutlookTools"
) else (
    echo   [FAIL] DLL NOT FOUND at %DLL_PATH%
)
echo.

echo === [2] CLSID Registry (HKCU) ===
reg query "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /s 2>nul
if %errorLevel% neq 0 echo   [FAIL] CLSID not found
echo.

echo === [3] InprocServer32 ===
reg query "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /s 2>nul
if %errorLevel% neq 0 echo   [FAIL] InprocServer32 not found
echo.

echo === [4] TypeLib ===
reg query "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\TypeLib" /s 2>nul
if %errorLevel% neq 0 echo   [FAIL] TypeLib not found
echo.

echo === [5] ProgId ===
reg query "HKCU\Software\Classes\OutlookTools.AddIn" /s 2>nul
if %errorLevel% neq 0 echo   [FAIL] ProgId not found
echo.

echo === [6] Outlook Add-in (16.0) ===
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /s 2>nul
if %errorLevel% neq 0 echo   [FAIL] Outlook 16.0 Add-in entry not found
echo.

echo === [7] Outlook Add-in (no version) ===
reg query "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn" /s 2>nul
if %errorLevel% neq 0 echo   [INFO] No-version path (OK, we use 16.0)
echo.

echo === [8] HKLM CLSID (check for conflicts) ===
reg query "HKLM\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /s 2>nul
if %errorLevel% equ 0 (
    echo   [WARN] HKLM has a CLSID entry - may conflict!
) else (
    echo   [OK] No HKLM conflict
)
echo.

echo === [9] HKLM Outlook Addins ===
reg query "HKLM\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /s 2>nul
if %errorLevel% equ 0 (
    echo   [WARN] HKLM has an add-in entry - may conflict!
) else (
    echo   [OK] No HKLM add-in entry
)
echo.

echo === [10] Resiliency (disabled items) ===
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" /s 2>nul
if %errorLevel% equ 0 (
    echo   [WARN] DisabledItems found - Outlook may have disabled add-ins!
) else (
    echo   [OK] No disabled items
)
echo.

echo === [11] Outlook Version Check ===
reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport 2>nul
if %errorLevel% neq 0 (
    reg query "HKLM\Software\Microsoft\Office\16.0\Common\ProductVersion" /v LastVersion 2>nul
)
echo.

echo === [12] All COM Add-ins listed by Outlook ===
echo   (Check Outlook > File > Options > Add-ins > COM Add-ins)
echo   Compare with what's in the registry:
echo.
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins" 2>nul
echo.
reg query "HKLM\Software\Microsoft\Office\16.0\Outlook\Addins" 2>nul
echo.

echo ========================================================
echo    Copy ALL of the above text and send it to me!
echo ========================================================
echo.
pause
