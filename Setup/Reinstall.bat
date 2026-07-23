@echo off
title OutlookTools - Clean Reinstall
color 0A
echo ========================================================
echo    OutlookTools - Clean Reinstall
echo ========================================================
echo.

echo [1/4] Closing Outlook...
taskkill /f /im outlook.exe >nul 2>&1
timeout /t 2 /nobreak >nul
echo   Done.
echo.

echo [2/4] Cleaning old registration...
reg delete "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /f >nul 2>&1
reg delete "HKCU\Software\Classes\OutlookTools.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /f >nul 2>&1
echo   Old entries removed.
echo.

echo [3/4] Registering with regasm (best effort)...
set DLL_PATH=%LOCALAPPDATA%\OutlookTools\OutlookTools.dll
if not exist "%DLL_PATH%" (
    set DLL_PATH=%~dp0..\OutlookTools\bin\Release\OutlookTools.dll
)

set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
if not exist "%REGASM%" set REGASM=C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe

if exist "%REGASM%" (
    "%REGASM%" "%DLL_PATH%" /codebase /tlb >nul 2>&1
    echo   regasm completed (exit: %errorLevel%)
) else (
    echo   regasm not found, skipping
)
echo.

echo [4/4] Creating registry entries (HKCU)...
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /ve /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\ProgId" /ve /t REG_SZ /d "OutlookTools.AddIn" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Both" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /v "Class" /t REG_SZ /d "OutlookTools.ThisAddIn" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /v "Assembly" /t REG_SZ /d "OutlookTools, Version=1.2.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\InprocServer32" /v "RuntimeVersion" /t REG_SZ /d "v4.0.30319" /f >nul

reg add "HKCU\Software\Classes\CLSID\{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}\TypeLib" /ve /t REG_SZ /d "{96A197CC-38FD-4899-8EFF-5B263FE8BB8D}" /f >nul

reg add "HKCU\Software\Classes\TypeLib\{96A197CC-38FD-4899-8EFF-5B263FE8BB8D}\1.0" /ve /t REG_SZ /d "OutlookTools Type Library" /f >nul
reg add "HKCU\Software\Classes\TypeLib\{96A197CC-38FD-4899-8EFF-5B263FE8BB8D}\1.0\0" /ve /t REG_SZ /d "" /f >nul
reg add "HKCU\Software\Classes\TypeLib\{96A197CC-38FD-4899-8EFF-5B263FE8BB8D}\1.0\HELPDIR" /ve /t REG_SZ /d "%LOCALAPPDATA%\OutlookTools" /f >nul

reg add "HKCU\Software\Classes\OutlookTools.AddIn" /ve /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Classes\OutlookTools.AddIn\CLSID" /ve /t REG_SZ /d "{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}" /f >nul

reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "Description" /t REG_SZ /d "OutlookTools" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "FriendlyName" /t REG_SZ /d "OutlookTools v1.2.1" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookTools.AddIn" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul

echo   Registry entries created.
echo.

echo ========================================================
echo    Reinstall complete!
echo ========================================================
echo.
echo Starting Outlook in 3 seconds...
timeout /t 3 /nobreak >nul
start outlook
