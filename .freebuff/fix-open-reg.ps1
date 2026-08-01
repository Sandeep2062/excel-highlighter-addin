# Removes the /R (read-only) flag from the Excel add-in OPEN registry value.
# Excel opens OPEN values prefixed with /R in read-only mode, which blocks
# saving the add-in after signing it in the VBE.
$ErrorActionPreference = "Stop"

$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$xlam = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"

if (-not (Test-Path $regPath)) {
    Write-Host "Excel Options registry key not found: $regPath"
    exit 0
}

$props = Get-ItemProperty -Path $regPath

# Find the OPEN value that points at our add-in (OPEN, OPEN1, OPEN2, ...)
$targetName = $null
foreach ($prop in $props.psobject.Properties) {
    if ($prop.Name -like "OPEN*" -and $prop.Value -like "*excel-highlighter*") {
        $targetName = $prop.Name
        break
    }
}

if ($null -eq $targetName) {
    Write-Host "No OPEN value points at excel-highlighter.xlam - nothing to fix."
    exit 0
}

$oldValue = $props.$targetName
Write-Host "Current $targetName : $oldValue"

# Strip a leading /R or /F flag, keep just the quoted path.
$newValue = $oldValue
if ($newValue -match '^/[RF]\s+"(.*)"$') {
    $newValue = '"' + $matches[1] + '"'
}

if ($newValue -eq $oldValue) {
    Write-Host "No /R flag present - already writable."
    exit 0
}

Set-ItemProperty -Path $regPath -Name $targetName -Value $newValue -Force
Write-Host "Fixed  $targetName : $newValue"
Write-Host "DONE - the add-in will now load read-write (no /R flag)."
