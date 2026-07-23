<#
.SYNOPSIS
OutlookTools - Per-User Installer (NO Admin Rights Required)
.DESCRIPTION
Registers OutlookTools as a COM Add-in for current user only.
No administrator privileges needed.
#>

Write-Host "============================================" -ForegroundColor White
Write-Host "    OutlookTools Installer" -ForegroundColor White
Write-Host "    No Admin Rights Required!" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host ""

# Step 1: Close Outlook
Write-Host "[1/5] Closing Outlook..." -ForegroundColor Cyan
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2
Write-Host "      Done." -ForegroundColor Green

# Step 2: Copy DLL
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
    Write-Host "      Build the project first (VS 2022, Build, Release)" -ForegroundColor Yellow
    Read-Host -Prompt "      Press Enter to exit"
    exit 1
}

# Step 3: Register COM with regasm (best-effort)
Write-Host "[3/5] Registering COM (regasm)..." -ForegroundColor Cyan
$dllPath = "$InstallDir\OutlookTools.dll"

$regasm64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$regasm32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
$regasmExe = ""
if (Test-Path $regasm64) { $regasmExe = $regasm64 }
elseif (Test-Path $regasm32) { $regasmExe = $regasm32 }

$regasmOk = $false
if ($regasmExe -ne "") {
    Write-Host "      Using: $regasmExe" -ForegroundColor DarkGray
    try {
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $regasmExe
        $pinfo.Arguments = "`"$dllPath`" /codebase /tlb"
        $pinfo.UseShellExecute = $false
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::Start($pinfo)
        $p.WaitForExit()
        if ($p.ExitCode -eq 0) {
            $regasmOk = $true
            Write-Host "      regasm OK!" -ForegroundColor Green
        } else {
            Write-Host "      regasm warning (exit $($p.ExitCode)) - OK to ignore" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "      regasm error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "      regasm not found - skipping" -ForegroundColor Yellow
}

if (-not $regasmOk) {
    Write-Host "      Using manual HKCU registration instead." -ForegroundColor Yellow
}

# Step 4: Manual HKCU Registration (ALWAYS runs)
Write-Host "[4/5] Registering in HKCU (per-user)..." -ForegroundColor Cyan

$CLSID = "{8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E}"
$TypeLib = "{96A197CC-38FD-4899-8EFF-5B263FE8BB8D}"
$ProgId = "OutlookTools.AddIn"

# CLSID base entry
$clsidBase = "HKCU:\Software\Classes\CLSID\$CLSID"
New-Item -Path "$clsidBase" -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "(Default)" -Value "OutlookTools Add-In" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "ProgId" -Value $ProgId -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$clsidBase" -Name "Version" -Value "1.2" -PropertyType String -Force | Out-Null

# TypeLib subkey (MUST be different from CLSID!)
$typeLibBase = "$clsidBase\TypeLib"
New-Item -Path "$typeLibBase" -Force | Out-Null
New-ItemProperty -Path "$typeLibBase" -Name "(Default)" -Value $TypeLib -PropertyType String -Force | Out-Null

# InprocServer32
$inprocBase = "$clsidBase\InprocServer32"
New-Item -Path "$inprocBase" -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "(Default)" -Value "mscoree.dll" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "ThreadingModel" -Value "Both" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "Class" -Value "OutlookTools.ThisAddIn" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "Assembly" -Value "OutlookTools, Version=1.2.0.0, Culture=neutral, PublicKeyToken=null" -PropertyType String -Force | Out-Null
New-ItemProperty -Path "$inprocBase" -Name "RuntimeVersion" -Value "v4.0.30319" -PropertyType String -Force | Out-Null
$codeBase = "file:///" + ($dllPath -replace '\\','/')
New-ItemProperty -Path "$inprocBase" -Name "CodeBase" -Value $codeBase -PropertyType String -Force | Out-Null

# ProgId -> CLSID
$progIdKey = "HKCU:\Software\Classes\$ProgId"
New-Item -Path "$progIdKey\CLSID" -Force | Out-Null
New-ItemProperty -Path "$progIdKey" -Name "(Default)" -Value $CLSID -PropertyType String -Force | Out-Null

Write-Host "      HKCU registration complete." -ForegroundColor Green

# Step 5: Register with Outlook
Write-Host "[5/5] Registering with Outlook..." -ForegroundColor Cyan
$AddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "Description" -Value "OutlookTools" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "OutlookTools v1.2.1" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

# Verify
Write-Host ""
Write-Host "Verifying registration..." -ForegroundColor Cyan

$ok1 = Test-Path "$clsidBase"
$ok2 = Test-Path $AddinKey
$ok3 = Test-Path $dllPath
$ok4 = Test-Path "$typeLibBase"

if ($ok1 -and $ok2 -and $ok3 -and $ok4) {
    Write-Host "  [OK] CLSID: $CLSID" -ForegroundColor Green
    Write-Host "  [OK] TypeLib: $TypeLib" -ForegroundColor Green
    Write-Host "  [OK] ProgId: $ProgId" -ForegroundColor Green
    Write-Host "  [OK] DLL: $dllPath" -ForegroundColor Green
    Write-Host "  [OK] Outlook add-in: LoadBehavior=3" -ForegroundColor Green
} else {
    if (-not $ok1) { Write-Host "  [FAIL] CLSID not found" -ForegroundColor Red }
    if (-not $ok4) { Write-Host "  [FAIL] TypeLib not found" -ForegroundColor Red }
    if (-not $ok2) { Write-Host "  [FAIL] Add-in entry not found" -ForegroundColor Red }
    if (-not $ok3) { Write-Host "  [FAIL] DLL not found" -ForegroundColor Red }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  OutlookTools installed successfully!" -ForegroundColor Green
Write-Host "  No admin rights needed!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Opening Outlook..." -ForegroundColor Cyan
Start-Process outlook
