<#
.SYNOPSIS
    Automated 1-Click Installer for Excel Highlighter Add-in.
    Copies excel-highlighter.xlam and images/ to %APPDATA%\Microsoft\AddIns\ExcelHighlighter
    and activates it automatically in the Windows Registry for Excel.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = Get-Location }

$xlamSource = Join-Path $scriptDir "excel-highlighter.xlam"
$imagesSource = Join-Path $scriptDir "images"
if (-not (Test-Path $imagesSource)) {
    $imagesSource = Join-Path (Join-Path $scriptDir "customUI") "images"
}

# Always build fresh excel-highlighter.xlam from src/ and customUI/
$buildScript = Join-Path (Join-Path $scriptDir "scripts") "build-xlam.ps1"
if (Test-Path $buildScript) {
    Write-Host "Building latest excel-highlighter.xlam..." -ForegroundColor Yellow
    & $buildScript
}

if (-not (Test-Path $xlamSource)) {
    Write-Error "Could not find excel-highlighter.xlam. Please build or download it first."
    exit 1
}

$targetDir = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$targetXlam = Join-Path $targetDir "excel-highlighter.xlam"
$targetImages = Join-Path $targetDir "images"

Write-Host "Installing Excel-Highlighter Add-in..." -ForegroundColor Cyan

# Terminate running Excel instances to release file locks
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "Terminating running Excel process(es) to release file lock..." -ForegroundColor Yellow
    Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# Clear Excel's ribbon cache so the Highlighter tab reliably appears on the
# next launch. A stale Excel.officeUI is the most common cause of a missing
# custom ribbon tab (it can be written to either %APPDATA% or %LOCALAPPDATA%).
$officeUiPaths = @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)
foreach ($uiPath in $officeUiPaths) {
    if (Test-Path $uiPath) {
        Remove-Item $uiPath -Force -ErrorAction SilentlyContinue
        Write-Host "  Cleared ribbon cache: $uiPath" -ForegroundColor Green
    }
}

# Copy files to target folder and root AddIns folder
$rootAddinDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$rootXlam = Join-Path $rootAddinDir "excel-highlighter.xlam"
$rootImages = Join-Path $rootAddinDir "images"

Copy-Item $xlamSource $targetXlam -Force
Copy-Item $xlamSource $rootXlam -Force
Unblock-File -Path $targetXlam -ErrorAction SilentlyContinue
Unblock-File -Path $rootXlam -ErrorAction SilentlyContinue
Write-Host "  Copied and unblocked add-in at: $targetXlam" -ForegroundColor Green
Write-Host "  Copied and unblocked add-in at: $rootXlam" -ForegroundColor Green

if (Test-Path $imagesSource) {
    if (-not (Test-Path $targetImages)) { New-Item -ItemType Directory -Path $targetImages -Force | Out-Null }
    if (-not (Test-Path $rootImages)) { New-Item -ItemType Directory -Path $rootImages -Force | Out-Null }
    Get-ChildItem -Path $imagesSource -File | Copy-Item -Destination $targetImages -Force
    Get-ChildItem -Path $imagesSource -File | Copy-Item -Destination $rootImages -Force
    Write-Host "  Copied UI icons to: $targetImages" -ForegroundColor Green
}

# Registry activation for Excel 2010+ (16.0 = 2016+, 15.0 = 2013, 14.0 = 2010)
$officeVersions = @("16.0", "15.0", "14.0")
$activatedCount = 0

foreach ($ver in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $regPath) {
        $props = Get-ItemProperty -Path $regPath
        $openKeys = $props.psobject.Properties | Where-Object { $_.Name -like "OPEN*" }
        
        $targetKeyName = $null
        foreach ($prop in $openKeys) {
            if ($prop.Value -like "*excel-highlighter*" -or $prop.Value -like "*excel-crosshair-highlighter*") {
                $targetKeyName = $prop.Name
                break
            }
        }
        
        if ($null -eq $targetKeyName) {
            $nextIndex = 0
            while ($true) {
                $keyName = if ($nextIndex -eq 0) { "OPEN" } else { "OPEN$nextIndex" }
                if (-not ($props.psobject.Properties.Name -contains $keyName)) {
                    $targetKeyName = $keyName
                    break
                }
                $nextIndex++
            }
        }

        # Excel strictly requires the /R switch in the OPEN registry value
        # to load the add-in on startup. Without /R, Excel ignores the entry.
        $regValue = '/R "' + $targetXlam + '"'
        Set-ItemProperty -Path $regPath -Name $targetKeyName -Value $regValue -Force | Out-Null
        Write-Host "  Activated in Excel $ver Registry ($targetKeyName -> $regValue)" -ForegroundColor Green
        $activatedCount++
    }
}

Write-Host ""
Write-Host "`nSUCCESS! Excel-Highlighter Add-in has been installed." -ForegroundColor Green
Write-Host "Open Microsoft Excel and the 'Highlighter' tab will appear automatically." -ForegroundColor Cyan
