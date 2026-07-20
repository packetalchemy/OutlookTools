@echo off
echo ============================================
echo    OutlookTools Installer
echo    Run as Administrator!
echo ============================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
pause
