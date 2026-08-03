<#
.SYNOPSIS
    Automated Uninstaller for Excel Highlighter Add-in.
    Removes registry activation keys and deletes installed files.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$targetDir = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter"
$targetXlam = Join-Path $targetDir "excel-highlighter.xlam"

Write-Host "Uninstalling Excel Highlighter Add-in..." -ForegroundColor Cyan

# Remove Registry Entries
$officeVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $regPath) {
        $props = Get-ItemProperty -Path $regPath
        $openKeys = $props.psobject.Properties | Where-Object { $_.Name -like "OPEN*" }
        foreach ($prop in $openKeys) {
            if ($prop.Value -like "*excel-highlighter*" -or $prop.Value -like "*excel-crosshair-highlighter*") {
                Remove-ItemProperty -Path $regPath -Name $prop.Name -Force | Out-Null
                Write-Host "  Removed registry activation key $($prop.Name) from Excel $ver" -ForegroundColor Green
            }
        }
    }
}

# Remove installation directory
if (Test-Path $targetDir) {
    Remove-Item $targetDir -Recurse -Force | Out-Null
    Write-Host "  Removed installation files from $targetDir" -ForegroundColor Green
}

# --- Remove the COM ribbon add-in registration (ExcelHighlighter.Ribbon) -----
# The COM add-in delivers the Highlighter tab on Excel 2024 / recent 365
# builds where the VBA ribbon path is never queried. Remove its ProgID,
# CLSID and Excel COM add-in keys so Excel no longer loads it.
$comProgId = "ExcelHighlighter.Ribbon"
$comClsid = "6E2F4A11-83C5-4B9D-9A07-2D51C8E4F0B6"

# Excel COM add-in keys (both the version-independent path and 16.0)
foreach ($base in @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Addins",
    "HKCU:\Software\Microsoft\Office\14.0\Excel\Addins"
)) {
    $k = Join-Path $base $comProgId
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed COM add-in key $k" -ForegroundColor Green
    }
}

# HKCU\Software\Classes registration (ProgID + CLSID) written by install.ps1
foreach ($p in @(
    "HKCU:\Software\Classes\$comProgId",
    "HKCU:\Software\Classes\CLSID\{$comClsid}"
)) {
    if (Test-Path $p) {
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed COM registration $p" -ForegroundColor Green
    }
}

# Clear any "Disabled Items" entry for the COM add-in too
foreach ($ver in $officeVersions) {
    $disabledPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $disabledPath) {
        Get-ItemProperty -Path $disabledPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSObject.Properties } |
            Where-Object { $_.Name -notlike "PS*" } |
            ForEach-Object {
                $bytes = $_.Value
                if ($bytes -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($bytes)
                    if ($text -match $comProgId -or $text -match $comClsid) {
                        Remove-ItemProperty -Path $disabledPath -Name $_.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "  Cleared 'Disabled Items' entry for COM add-in in Excel $ver" -ForegroundColor Green
                    }
                }
            }
    }
}

# Remove the old duplicate copy some earlier installer versions left directly
# in the AddIns root (fixed in 2.0.0 - only one copy is written now).
$legacyRootXlam = Join-Path $env:APPDATA "Microsoft\AddIns\excel-highlighter.xlam"
$legacyRootImages = Join-Path $env:APPDATA "Microsoft\AddIns\images"
if (Test-Path $legacyRootXlam) {
    Remove-Item $legacyRootXlam -Force -ErrorAction SilentlyContinue
    Write-Host "  Removed legacy duplicate copy from AddIns root" -ForegroundColor Green
}
if (Test-Path $legacyRootImages) {
    Remove-Item $legacyRootImages -Recurse -Force -ErrorAction SilentlyContinue
}

# Clear any "Disabled Items" entry so a future reinstall isn't silently
# blacklisted by Excel from a past crash.
foreach ($ver in $officeVersions) {
    $disabledPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $disabledPath) {
        Get-ItemProperty -Path $disabledPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSObject.Properties } |
            Where-Object { $_.Name -notlike "PS*" } |
            ForEach-Object {
                $bytes = $_.Value
                if ($bytes -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($bytes)
                    if ($text -like "*excel-highlighter*" -or $text -like "*ExcelHighlighter*" -or $text -like "*excel-crosshair-highlighter*") {
                        Remove-ItemProperty -Path $disabledPath -Name $_.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "  Cleared 'Disabled Items' entry in Excel $ver" -ForegroundColor Green
                    }
                }
            }
    }
}

Write-Host ""
Write-Host "Uninstallation complete. Excel Highlighter Add-in has been removed." -ForegroundColor Green
