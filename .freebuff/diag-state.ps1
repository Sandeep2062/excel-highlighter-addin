$ErrorActionPreference = "Continue"
Write-Host "=== All Excel processes ==="
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) {
    $p | Select-Object Id, MainWindowTitle, StartTime, Responding | Format-Table -AutoSize | Out-String | Write-Host
} else {
    Write-Host "No Excel running"
}

Write-Host "=== Deployed file state ==="
$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$i = Get-Item $f
Write-Host "Exists: $(Test-Path $f)"
Write-Host "IsReadOnly: $($i.IsReadOnly)"
Write-Host "LastWrite: $($i.LastWriteTime)"
Write-Host "Size: $($i.Length)"
try {
    $fs = [System.IO.File]::Open($f, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    Write-Host "Lock probe: WRITABLE"
    $fs.Close()
} catch {
    Write-Host "Lock probe: LOCKED - $($_.Exception.Message)"
}
