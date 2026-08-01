$ErrorActionPreference = 'Continue'

Write-Host "=== 1. Log tail (last 20) - did onLoad fire, did XLSTART add-in load? ==="
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (Test-Path $log) {
    Get-Content $log -Tail 20 | Write-Host
    Write-Host "`n  'Ribbon UI attached' total ever: $((Select-String -Path $log -Pattern 'Ribbon UI attached').Count)"
} else { Write-Host "  no log" }

Write-Host "`n=== 2. AddInLoadTimes - did the XLSTART minimal test add-in load? ==="
$lt = Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Excel\AddInLoadTimes" -ErrorAction SilentlyContinue
if ($lt) {
    $lt.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
        $isZero = ($_.Value | Where-Object { $_ -ne 0 } | Measure-Object).Count -eq 0
        Write-Host "  $($_.Name) = $($_.Value) $(if ($isZero) { '<-- ALL ZEROS (did not load)' })"
    }
}

Write-Host "`n=== 3. Excel process state ==="
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) {
    $p | Select-Object Id, MainWindowTitle, Path | Format-List | Out-String | Write-Host
} else { Write-Host "  Excel not running" }

Write-Host "`n=== 4. CustomUIValidationCache after full clear ==="
$ck = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $ck) {
    Get-ItemProperty $ck -ErrorAction SilentlyContinue | Select-Object * -ExcludeProperty PS* | Format-List | Out-String | Write-Host
} else { Write-Host "  (key still absent - nothing re-validated)" }

Write-Host "`n=== 5. Excel.officeUI state ==="
foreach ($path in @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)) {
    if (Test-Path $path) {
        $i = Get-Item $path
        Write-Host "  EXISTS: $path ($($i.Length) bytes, $($i.LastWriteTime))"
    } else { Write-Host "  absent: $path" }
}

Write-Host "`n=== 6. Does the deployed add-in's customUI14.xml exist & is it the ORIGINAL Highlighter XML? ==="
Add-Type -AssemblyName System.IO.Compression.FileSystem
$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$copy = Join-Path $env:TEMP "post-check.xlam"
try {
    Copy-Item $deployed $copy -Force -ErrorAction Stop
    $z = [System.IO.Compression.ZipFile]::OpenRead($copy)
    $e = $z.GetEntry('customUI/customUI14.xml')
    $sr = New-Object System.IO.StreamReader($e.Open())
    $xml = $sr.ReadToEnd(); $sr.Close()
    $z.Dispose()
    Write-Host "  customUI14.xml present: True, contains tabHighlighter: $($xml.Contains('tabHighlighter')), onLoad: $($xml.Contains('RibbonCallbacks.onLoad'))"
    Remove-Item $copy -Force
} catch {
    Write-Host "  ERROR: $($_.Exception.Message)"
}

Write-Host "`n=== done ==="
