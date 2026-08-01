$ErrorActionPreference = 'Continue'

$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (-not (Test-Path $log)) { Write-Host "NO LOG FILE"; exit 0 }

Write-Host "=== Every line mentioning ribbon/Ribbon/onLoad/onload ==="
Select-String -Path $log -Pattern 'ribbon|onLoad|onload' -CaseSensitive:$false | ForEach-Object {
    Write-Host "  $($_.Line)"
}

Write-Host "`n=== Every ERROR line ever ==="
Select-String -Path $log -Pattern 'ERROR' | ForEach-Object {
    Write-Host "  $($_.Line)"
}

Write-Host "`n=== Total line count and first/last timestamps ==="
$all = Get-Content $log
Write-Host "  Lines: $($all.Count)"
Write-Host "  First: $($all[0])"
Write-Host "  Last:  $($all[$all.Count-1])"

Write-Host "`n=== The current session's startup block (last StartUp forward) ==="
$idx = $all.Count - 1
$lastStart = -1
for ($i = $all.Count - 1; $i -ge 0; $i--) {
    if ($all[$i] -like '*StartUp*') { $lastStart = $i; break }
}
if ($lastStart -ge 0) {
    for ($j = $lastStart; $j -lt $all.Count; $j++) { Write-Host "  $($all[$j])" }
} else {
    Write-Host "  (no StartUp found?)"
}
