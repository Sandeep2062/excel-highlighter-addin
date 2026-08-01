$ErrorActionPreference = 'Continue'

$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
$before = if (Test-Path $log) { (Get-Content $log).Count } else { 0 }

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    # AutomationSecurity values: 1=Low(allow all), 2=ByUI, 3=ForceDisable
    try { $excel.AutomationSecurity = 1 } catch { Write-Host "  (could not set AutomationSecurity)" }

    Write-Host "=== AddIns2 already present in this instance ==="
    foreach ($ai in $excel.AddIns2) {
        Write-Host "  Name='$($ai.Name)' Installed=$($ai.Installed) Path='$($ai.Path)'"
    }
    Write-Host "  Workbooks at start: $($excel.Workbooks.Count)"

    # Use the minimal-test copy (fresh file, not yet registered)
    $testCopy = Join-Path $env:TEMP "ribbon-test\test.xlam"
    if (-not (Test-Path $testCopy)) {
        Write-Host "test.xlam missing - building from deployed copy"
        $src = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
        Copy-Item $src $testCopy -Force
    }
    Write-Host "`n=== AddIns2.Add minimal test add-in: $testCopy ==="
    $ai = $excel.AddIns2.Add($testCopy, $true)
    Write-Host "  Added: Name='$($ai.Name)' Installed=$($ai.Installed)"
    Start-Sleep -Seconds 8

    Write-Host "`n=== Now re-add the REAL deployed add-in via AddIns2 ==="
    $real = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
    $ai2 = $excel.AddIns2.Add($real, $true)
    Write-Host "  Added: Name='$($ai2.Name)' Installed=$($ai2.Installed)"
    Start-Sleep -Seconds 6

    Write-Host "`n=== Workbooks open now ==="
    for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
        $wb = $excel.Workbooks.Item($i)
        Write-Host "  $($wb.Name) IsAddin=$($wb.IsAddin)"
    }
} catch {
    Write-Host "  ERROR: $($_.Exception.Message)"
} finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $false } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}

Start-Sleep -Seconds 2
Write-Host "`n=== New log entries ==="
if (Test-Path $log) {
    $all = Get-Content $log
    if ($all.Count -gt $before) {
        $all[($before)..($all.Count-1)] | ForEach-Object { Write-Host "  $_" }
    } else {
        Write-Host "  (no new entries)"
    }
    Write-Host "`n=== 'Ribbon UI attached' total ever: $((Select-String -Path $log -Pattern 'Ribbon UI attached').Count) ==="
} else {
    Write-Host "  no log file"
}

Write-Host "`n=== done ==="
