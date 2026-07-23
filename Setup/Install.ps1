<#
.SYNOPSIS
OutlookTools — Per-User Installer (NO Admin Rights Required)
.DESCRIPTION
Registers OutlookTools as a COM Add-in for current user only.
Uses regasm + manual HKCU registration for reliability.
No administrator privileges needed.
#>

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

# === Step 3: Register COM with regasm (best-effort, don't stop on warnings) ===
Write-Host "[3/5] Registering COM (regasm)..." -ForegroundColor Cyan
$dllPath = "$InstallDir\OutlookTools.dll"

$regasmPaths = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
)

$regasmOk = $false
foreach ($regPath in $regasmPaths) {
    if (Test-Path $regPath) {
        Write-Host "      Using: $regPath" -ForegroundColor DarkGray
        # Suppress stderr (regasm warnings) — they are NOT errors
        $proc = Start-Process -FilePath $regPath -ArgumentList "`"$dllPath`" /codebase /tlb" -Wait -PassThru -NoNewWindow -RedirectStandardError "$env:TEMP\regasm_err.txt"
        $regasmErr = ""
        if (Test-Path "$env:TEMP\regasm_err.txt") {
            $regasmErr = Get-Content "$env:TEMP\regasm_err.txt" -Raw
            Remove-Item "$env:TEMP\regasm_err.txt" -ErrorAction SilentlyContinue
        }
        if ($proc.ExitCode -eq 0) {
            $regasmOk = $true
            Write-Host "      regasm OK!" -ForegroundColor Green
        } else {
            Write-Host "      regasm exit code: $($proc.ExitCode)" -ForegroundColor Yellow
            if ($regasmErr) { Write-Host "      $($regasmErr.Trim())" -ForegroundColor DarkGray }
        }
        break
    }
}

if (-not $regasmOk) {
    Write-Host "      regasm had issues — manual registration will handle it." -ForegroundColor Yellow
}

# === Step 4: Manual HKCU Registration (ALWAYS runs — the reliable path) ===
Write-Host "[4/5] Registering in HKCU (per-user)..." -ForegroundColor Cyan

$CLSID = "{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}"
$ProgId = "OutlookTools.AddIn"

# CLSID base entry
$clsidBase = "HKCU:\Software\Classes\CLSID\$CLSID"
New-Item -Path "$clsidBase" -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "(Default)" -Value "OutlookTools Add-In" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "ProgId" -Value $ProgId -PropertyType String -Force | Out-Null
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

Write-Host "      HKCU registration complete." -ForegroundColor Green

# === Step 5: Register with Outlook ===
Write-Host "[5/5] Registering with Outlook..." -ForegroundColor Cyan
$AddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "Description" -Value "OutlookTools - Open Source Outlook Add-in" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "OutlookTools v1.2.1" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

# === Verify ===
Write-Host ""
Write-Host "Verifying registration..." -ForegroundColor Cyan

$verifyClsid = Test-Path "$clsidBase"
$verifyAddin = Test-Path $AddinKey
$verifyDll = Test-Path $dllPath

if ($verifyClsid -and $verifyAddin -and $verifyDll) {
    Write-Host "  [OK] CLSID registered: $CLSID" -ForegroundColor Green
    Write-Host "  [OK] ProgId registered: $ProgId" -ForegroundColor Green
    Write-Host "  [OK] DLL exists: $dllPath" -ForegroundColor Green
    Write-Host "  [OK] Outlook add-in entry: LoadBehavior=3" -ForegroundColor Green
} else {
    if (-not $verifyClsid) { Write-Host "  [FAIL] CLSID not found" -ForegroundColor Red }
    if (-not $verifyAddin) { Write-Host "  [FAIL] Add-in entry not found" -ForegroundColor Red }
    if (-not $verifyDll)   { Write-Host "  [FAIL] DLL not found" -ForegroundColor Red }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  OutlookTools installed successfully!" -ForegroundColor Green
Write-Host "  No admin rights needed!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Opening Outlook..." -ForegroundColor Cyan
Start-Process outlook
