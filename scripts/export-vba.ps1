<#
.SYNOPSIS
    Exports every VBA module from an open .xlam back out to src/ as text,
    so changes made in the VBA editor can be committed to git.

.PARAMETER XlamPath
    Path to the .xlam to export from. Defaults to
    excel-crosshair-highlighter.xlam in the repo root.

.NOTES
    Requires "Trust access to the VBA project object model" enabled.
    Overwrites files under src/ - review the diff before committing.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")),
    [string]$XlamPath = (Join-Path $RepoRoot "excel-crosshair-highlighter.xlam")
)

$ErrorActionPreference = "Stop"

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
    $vbProject = $workbook.VBProject

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
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Export complete. Review changes under src/ before committing." -ForegroundColor Green
