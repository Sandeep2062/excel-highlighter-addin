<#
.SYNOPSIS
    Exports every VBA module from an open .xlam back out to src/ as text,
    so changes made in the VBA editor can be committed to git.

.PARAMETER XlamPath
    Path to the .xlam to export from. Defaults to
    excel-highlighter.xlam in the repo root.

.NOTES
    Requires "Trust access to the VBA project object model" enabled.
    Overwrites files under src/ - review the diff before committing.
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
    throw "Could not find $XlamPath. Pass -XlamPath explicitly if it lives elsewhere."
}

$srcDir = Join-Path $RepoRoot "src"
if (-not (Test-Path $srcDir)) { New-Item -ItemType Directory -Path $srcDir | Out-Null }

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

    foreach ($component in $vbProject.VBComponents) {

        $extension = switch ($component.Type) {
            1 { ".bas" }  # vbext_ct_StdModule
            2 { ".cls" }  # vbext_ct_ClassModule
            3 { ".frm" }  # vbext_ct_MSForm (not used in this project, kept for completeness)
            100 { ".cls" } # vbext_ct_Document (ThisWorkbook)
            default { $null }
        }

        if ($null -eq $extension) {
            Write-Warning "Skipping $($component.Name) - unrecognised component type $($component.Type)"
            continue
        }

        $targetPath = Join-Path $srcDir "$($component.Name)$extension"
        Write-Host "  Exporting $($component.Name)$extension"
        $component.Export($targetPath)
    }

    $workbook.Close($false)
}
finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Export complete. Review changes under src/ before committing." -ForegroundColor Green
