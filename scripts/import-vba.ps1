<#
.SYNOPSIS
    Re-imports every module from src/ into an already-existing .xlam,
    replacing whatever is currently there. Faster than build-xlam.ps1 for
    day-to-day iteration since it doesn't rebuild the RibbonX package.

.PARAMETER XlamPath
    Path to the .xlam to update. Defaults to
    excel-highlighter.xlam in the repo root.

.NOTES
    Requires "Trust access to the VBA project object model" enabled.
    The target .xlam must not be open in another Excel instance.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$XlamPath
)

$ErrorActionPreference = "Stop"

if (-not $RepoRoot) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    if (-not $scriptDir) { $scriptDir = Get-Location }
    $RepoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
}
if (-not $XlamPath) {
    $XlamPath = Join-Path $RepoRoot "excel-highlighter.xlam"
}

if (-not (Test-Path $XlamPath)) {
    throw "Could not find $XlamPath. Run build-xlam.ps1 first, or pass -XlamPath explicitly."
}

$srcDir = Join-Path $RepoRoot "src"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($XlamPath)
    try {
        $vbProject = $workbook.VBProject
    } catch {
        $vbProject = $null
    }
    if ($null -eq $vbProject -or $null -eq $vbProject.VBComponents) {
        throw "Access to the Excel VBA Project object model is blocked or disabled. Please enable 'Trust access to the VBA project object model' in Excel (File > Options > Trust Center > Trust Center Settings > Macro Settings)."
    }

    # Remove every component except ThisWorkbook (which can't be removed,
    # only overwritten) before reimporting, so stale renamed/deleted modules
    # don't linger.
    $toRemove = @()
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -ne "ThisWorkbook") {
            $toRemove += $component
        }
    }
    foreach ($component in $toRemove) {
        Write-Host "  Removing existing $($component.Name)"
        $vbProject.VBComponents.Remove($component)
    }

    # Overwrite ThisWorkbook's code.
    $thisWorkbookSrc = Get-Content (Join-Path $srcDir "ThisWorkbook.cls") -Raw
    $thisWorkbookBody = ($thisWorkbookSrc -split "(?ms)^Attribute VB_Exposed.*?\r?\n", 2)[1]
    $codeModule = $vbProject.VBComponents.Item("ThisWorkbook").CodeModule
    $codeModule.DeleteLines(1, [Math]::Max($codeModule.CountOfLines, 1))
    $codeModule.AddFromString($thisWorkbookBody)

    # Reimport everything else.
    $eventAppTemp = Join-Path $env:TEMP "EventApp.cls"
    Copy-Item (Join-Path $srcDir "EventApp.cls") $eventAppTemp -Force
    $vbProject.VBComponents.Import($eventAppTemp) | Out-Null
    Remove-Item $eventAppTemp -Force

    $standardModules = @(
        "Constants.bas", "Logging.bas", "Utilities.bas", "Settings.bas",
        "HighlightEngine.bas", "AddinHost.bas", "RibbonCallbacks.bas",
        "ColourPicker.bas", "SelectionHistory.bas", "Profiles.bas"
    )
    foreach ($moduleFile in $standardModules) {
        Write-Host "  Importing $moduleFile"
        $vbProject.VBComponents.Import((Join-Path $srcDir $moduleFile)) | Out-Null
    }

    $workbook.Save()
    $workbook.Close($false)
}
finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Import complete. $XlamPath updated in place." -ForegroundColor Green
