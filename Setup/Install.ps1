<#
.SYNOPSIS
OutlookTools — Per-User Installer (NO Admin Rights Required)
.DESCRIPTION
Registers OutlookTools as a COM Add-in for current user only.
No administrator privileges needed.
#>

# Close Outlook
Write-Host "Closing Outlook..." -ForegroundColor Cyan
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Copy DLL to user's local folder
$InstallDir = "$env:LOCALAPPDATA\OutlookTools"
if (!(Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }

$SourcePath = "$PSScriptRoot\..\OutlookTools\bin\Release\OutlookTools.dll"

if (Test-Path $SourcePath) {
    Copy-Item $SourcePath "$InstallDir\OutlookTools.dll" -Force
    Write-Host "DLL copied to $InstallDir" -ForegroundColor Green
} else {
    Write-Host "ERROR: OutlookTools.dll not found at $SourcePath" -ForegroundColor Red
    Write-Host "Please build the project first (Build > Release)" -ForegroundColor Yellow
    pause
    exit 1
}

# Register DLL (user-level, no admin needed)
Write-Host "Registering DLL..." -ForegroundColor Cyan
$regPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$dllPath = "$InstallDir\OutlookTools.dll"

if (Test-Path $regPath) {
    & $regPath $dllPath /codebase /tlb 2>&1 | Out-Null
} else {
    & C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe $dllPath /codebase /tlb 2>&1 | Out-Null
}

# Add Registry Entry (HKCU = per-user, NO admin needed)
Write-Host "Creating registry entry..." -ForegroundColor Cyan
$AddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "Description" -Value "OutlookTools" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "OutlookTools" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  OutlookTools installed successfully!" -ForegroundColor Green
Write-Host "  No admin rights needed!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Opening Outlook..."
Start-Process outlook
