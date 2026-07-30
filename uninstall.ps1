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
            if ($prop.Value -like "*excel-highlighter.xlam*") {
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

Write-Host ""
Write-Host "Uninstallation complete. Excel Highlighter Add-in has been removed." -ForegroundColor Green
