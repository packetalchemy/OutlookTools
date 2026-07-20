<#
.SYNOPSIS
Installs OutlookTools as a COM Add-in for Outlook.
.DESCRIPTION
Registers the OutlookTools.dll as a COM Add-in.
Must run as Administrator.
#>

# Self-elevate to Administrator
param([switch]$Elevated)

if (-not $Elevated) {
    $Script = $MyInvocation.MyCommand.Path
    $Process = Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$Script`" -Elevated" -Verb RunAs -PassThru

    # Wait for elevation to complete
    $Process.WaitForExit()
    exit $Process.ExitCode
}

# Check if already installed
$AddinKey = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn"
if (Test-Path $AddinKey) {
    Write-Host "Removing old installation..." -ForegroundColor Yellow
    regasm OutlookTools.dll /unregister 2>&1 | Out-Null
    Remove-Item -Path $AddinKey -Recurse -Force -ErrorAction SilentlyContinue
}

# Close Outlook
Write-Host "Closing Outlook..." -ForegroundColor Cyan
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Register DLL
Write-Host "Registering OutlookTools.dll..." -ForegroundColor Cyan
$regPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

if (Test-Path $regPath) {
    & $regPath OutlookTools.dll /codebase /tlb | Out-Null
} else {
    & C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe OutlookTools.dll /codebase /tlb | Out-Null
}

# Add Registry Entry
Write-Host "Creating registry entry..." -ForegroundColor Cyan
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "Description" -Value "OutlookTools - Open Source Outlook Plugin" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "OutlookTools" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  OutlookTools installed successfully!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Opening Outlook..."
Start-Process outlook
