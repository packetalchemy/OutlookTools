<#
.SYNOPSIS
OutlookTools — Per-User Installer (NO Admin Rights Required)
.DESCRIPTION
Registers OutlookTools as a COM Add-in for current user only.
Uses regasm + manual HKCU registration for reliability.
No administrator privileges needed.
#>

$ErrorActionPreference = "Stop"

Write-Host "============================================" -ForegroundColor White
Write-Host "    OutlookTools Installer" -ForegroundColor White
Write-Host "    No Admin Rights Required!" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host ""

# === Step 1: Close Outlook ===
Write-Host "[1/5] Closing Outlook..." -ForegroundColor Cyan
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2
Write-Host "      Done." -ForegroundColor Green

# === Step 2: Copy DLL ===
Write-Host "[2/5] Copying DLL..." -ForegroundColor Cyan
$InstallDir = "$env:LOCALAPPDATA\OutlookTools"
if (!(Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }

$SourcePath = "$PSScriptRoot\..\OutlookTools\bin\Release\OutlookTools.dll"

if (Test-Path $SourcePath) {
    Copy-Item $SourcePath "$InstallDir\OutlookTools.dll" -Force
    Write-Host "      DLL -> $InstallDir\OutlookTools.dll" -ForegroundColor Green
} else {
    Write-Host "      ERROR: OutlookTools.dll not found!" -ForegroundColor Red
    Write-Host "      Expected: $SourcePath" -ForegroundColor Yellow
    Write-Host "      Build the project first (VS 2022 -> Build -> Release)" -ForegroundColor Yellow
    Read-Host -Prompt "      Press Enter to exit"
    exit 1
}

# === Step 3: Register COM with regasm ===
Write-Host "[3/5] Registering COM (regasm)..." -ForegroundColor Cyan
$dllPath = "$InstallDir\OutlookTools.dll"

# Try 64-bit first, then 32-bit
$regasmPaths = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
)

$regasmOk = $false
foreach ($regPath in $regasmPaths) {
    if (Test-Path $regPath) {
        Write-Host "      Using: $regPath" -ForegroundColor DarkGray
        $output = & $regPath $dllPath /codebase /tlb 2>&1
        if ($LASTEXITCODE -eq 0) {
            $regasmOk = $true
            Write-Host "      regasm OK!" -ForegroundColor Green
            break
        } else {
            Write-Host "      regasm failed (exit $LASTEXITCODE): $output" -ForegroundColor Yellow
        }
    }
}

if (-not $regasmOk) {
    Write-Host "      WARNING: regasm failed. Trying manual registration..." -ForegroundColor Yellow
}

# === Step 4: Manual HKCU Registration (always run, ensures reliability) ===
Write-Host "[4/5] Registering in HKCU (per-user)..." -ForegroundColor Cyan

# The CLSID from ThisAddIn.cs [Guid] attribute
$CLSID = "{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}"
$ProgId = "OutlookTools.AddIn"

# CLSID entries
$clsidBase = "HKCU:\Software\Classes\CLSID\$CLSID"
New-Item -Path "$clsidBase" -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "(Default)" -Value "OutlookTools Add-In" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "ProgId" -Value $ProgId -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "TypeLib" -Value "{00020813-0000-0000-C000-000000000046}" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "Version" -Value "1.2" -PropertyType String -Force | Out-Null

# InprocServer32 — tells COM how to load the .NET DLL
$inprocBase = "$clsidBase\InprocServer32"
New-Item -Path "$inprocBase" -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "(Default)" -Value "mscoree.dll" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "ThreadingModel" -Value "Both" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "Class" -Value "OutlookTools.ThisAddIn" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "Assembly" -Value "OutlookTools, Version=1.2.0.0, Culture=neutral, PublicKeyToken=null" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "RuntimeVersion" -Value "v4.0.30319" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "CodeBase" -Value "file:///$($dllPath -replace '\\','/')" -PropertyType String -Force | Out-Null

# ProgId -> CLSID mapping
$progIdKey = "HKCU:\Software\Classes\$ProgId"
New-Item -Path "$progIdKey\CLSID" -Force | Out-Null
New-ItemProperty -Path "$progIdKey" -Name "(Default)" -Value $CLSID -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$progIdKey\CLSID" -Name "(Default)" -Value $CLSID -PropertyType String -Force | Out-Null

Write-Host "      HKCU registration complete." -ForegroundColor Green

# === Step 5: Verify & Add Outlook Add-in Entry ===
Write-Host "[5/5] Registering with Outlook..." -ForegroundColor Cyan
$AddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "Description" -Value "OutlookTools - Open Source Outlook Add-in" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "OutlookTools v1.2.1" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

# === Verify registration ===
Write-Host ""
Write-Host "Verifying registration..." -ForegroundColor Cyan

$verifyClsid = Get-ItemProperty -Path "$clsidBase" -ErrorAction SilentlyContinue
$verifyAddin = Get-ItemProperty -Path $AddinKey -ErrorAction SilentlyContinue
$verifyDll = Test-Path $dllPath

if ($verifyClsid -and $verifyAddin -and $verifyDll) {
    Write-Host "  [OK] CLSID registered: $CLSID" -ForegroundColor Green
    Write-Host "  [OK] ProgId registered: $ProgId" -ForegroundColor Green
    Write-Host "  [OK] DLL exists: $dllPath" -ForegroundColor Green
    Write-Host "  [OK] Outlook add-in entry: LoadBehavior=3" -ForegroundColor Green
} else {
    Write-Host "  [WARN] Some registrations may be missing" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  OutlookTools installed successfully!" -ForegroundColor Green
Write-Host "  No admin rights needed!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Opening Outlook..." -ForegroundColor Cyan
Start-Process outlook
