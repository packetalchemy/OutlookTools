@echo off
echo ============================================
echo    OutlookTools Installer
echo    No Admin Rights Required!
echo ============================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
pause
