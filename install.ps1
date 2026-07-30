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

# Copy files
Copy-Item $xlamSource $targetXlam -Force
Unblock-File -Path $targetXlam -ErrorAction SilentlyContinue
Write-Host "  Copied and unblocked add-in at: $targetXlam" -ForegroundColor Green

if (Test-Path $imagesSource) {
    if (-not (Test-Path $targetImages)) {
        New-Item -ItemType Directory -Path $targetImages -Force | Out-Null
    }
    Copy-Item (Join-Path $imagesSource "*") $targetImages -Force
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

        $regValue = '/R "' + $targetXlam + '"'
        Set-ItemProperty -Path $regPath -Name $targetKeyName -Value $regValue -Force | Out-Null
        Write-Host "  Activated in Excel $ver Registry ($targetKeyName -> $targetXlam)" -ForegroundColor Green
        $activatedCount++
    }
}

Write-Host ""
Write-Host "`nSUCCESS! Excel-Highlighter Add-in has been installed." -ForegroundColor Green
Write-Host "Open Microsoft Excel and the 'Highlighter' tab will appear automatically." -ForegroundColor Cyan
