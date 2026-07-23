@echo off
title OutlookTools - Nuclear Fix
color 0A
echo ========================================================
echo    OutlookTools - Nuclear Fix
echo    This should DEFINITELY work!
echo ========================================================
echo.

echo [1/6] Closing Outlook...
taskkill /f /im outlook.exe >nul 2>&1
timeout /t 3 /nobreak >nul
echo   Done.
echo.

echo [2/6] Cleaning ALL old entries...
reg delete "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /f >nul 2>&1
reg delete "HKCU\Software\Classes\OutlookTools" /f >nul 2>&1
reg delete "HKCU\Software\Classes\OutlookTools.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" /f >nul 2>&1
echo   All old entries removed.
echo.

echo [3/6] Copying ALL DLLs (not just OutlookTools.dll)...
set INSTALL_DIR=%LOCALAPPDATA%\OutlookTools
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
set RELEASE_DIR=%~dp0..\OutlookTools\bin\Release

if not exist "%RELEASE_DIR%\OutlookTools.dll" (
    set RELEASE_DIR=%~dp0..\..\OutlookTools\bin\Release
)

if exist "%RELEASE_DIR%\OutlookTools.dll" (
    copy "%RELEASE_DIR%\*.dll" "%INSTALL_DIR%\" /Y >nul 2>&1
    copy "%RELEASE_DIR%\*.pdb" "%INSTALL_DIR%\" /Y >nul 2>&1
    echo   Copied all files from: %RELEASE_DIR%
    echo   Files in install dir:
    dir "%INSTALL_DIR%\*.dll" /B 2>nul
) else (
    echo   [ERROR] Release folder not found!
    echo   Expected: %RELEASE_DIR%
    echo   Build the project first!
    pause
    exit 1
)
echo.

echo [4/6] Registering COM (regasm)...
set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
if not exist "%REGASM%" set REGASM=C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe

if exist "%REGASM%" (
    "%REGASM%" "%INSTALL_DIR%\OutlookTools.dll" /codebase /tlb >nul 2>&1
    echo   regasm done (exit: %errorLevel%)
) else (
    echo   regasm not found
)
echo.

echo [5/6] Creating registry entries...
set CLSID={8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}
set TYPELIB={96A197CC-38FD-4899-8EFF-5B263FE8BB8D}

reg add "HKCU\Software\Classes\CLSID\%CLSID%" /ve /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\ProgId" /ve /t REG_SZ /d "OutlookTools.AddIn" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\TypeLib" /ve /t REG_SZ /d "%TYPELIB%" /f >nul

reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Both" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Class" /t REG_SZ /d "OutlookTools.ThisAddIn" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "Assembly" /t REG_SZ /d "OutlookTools, Version=1.2.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "RuntimeVersion" /t REG_SZ /d "v4.0.30319" /f >nul
reg add "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v "CodeBase" /t REG_SZ /d "file:///%INSTALL_DIR:\=/%/OutlookTools.dll" /f >nul

reg add "HKCU\Software\Classes\OutlookTools.AddIn" /ve /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Classes\OutlookTools.AddIn\CLSID" /ve /t REG_SZ /d "%CLSID%" /f >nul

reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "Description" /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "FriendlyName" /t REG_SZ /d "OutlookTools v1.2.1" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul

echo   Registry entries created.
echo.

echo [6/6] Verification...
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v LoadBehavior >nul 2>&1
if %errorLevel% equ 0 (
    echo   [OK] Outlook add-in entry exists
) else (
    echo   [FAIL] Outlook add-in entry missing
)

if exist "%INSTALL_DIR%\OutlookTools.dll" (
    echo   [OK] DLL exists
) else (
    echo   [FAIL] DLL missing
)

reg query "HKCU\Software\Classes\CLSID\%CLSID%\InprocServer32" /v CodeBase >nul 2>&1
if %errorLevel% equ 0 (
    echo   [OK] CodeBase registered
) else (
    echo   [FAIL] CodeBase missing
)

echo.
echo ========================================================
echo    DONE! Starting Outlook in 5 seconds...
echo ========================================================
echo.
timeout /t 5 /nobreak >nul
start outlook
