$ErrorActionPreference = 'Continue'

$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
$before = if (Test-Path $log) { (Get-Content $log).Count } else { 0 }

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    try { $excel.AutomationSecurity = 1 } catch {}

    $xlam = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"

    # Method 1: AddIns2.Add + Installed=$true (the Add-ins dialog path)
    Write-Host "=== Method 1: AddIns2.Add (standard add-in dialog path) ==="
    $ai = $null
    try {
        $ai = $excel.AddIns2.Add($xlam, $true)
        Write-Host "  AddIns2.Add returned: Name='$($ai.Name)' Installed=$($ai.Installed)"
    } catch {
        Write-Host "  AddIns2.Add failed: $($_.Exception.Message)"
        try {
            $ai = $excel.AddIns.Add($xlam, $true)
            Write-Host "  AddIns.Add (fallback): Name='$($ai.Name)' Installed=$($ai.Installed)"
        } catch {
            Write-Host "  AddIns.Add also failed: $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 6
    Write-Host "  Workbooks now: $($excel.Workbooks.Count)"
    if ($excel.Workbooks.Count -gt 0) {
        Write-Host "  First workbook IsAddin: $($excel.Workbooks.Item(1).IsAddin)"
    }

    # Method 2 (fallback): register test xlam copy via AddIns
    Write-Host "`n=== Method 2: AddIns2.Add on the minimal-test copy ==="
    $testCopy = Join-Path $env:TEMP "ribbon-test\test.xlam"
    if (Test-Path $testCopy) {
        try {
            $ai2 = $excel.AddIns2.Add($testCopy, $true)
            Write-Host "  Added test copy. Name='$($ai2.Name)' Installed=$($ai2.Installed)"
            Start-Sleep -Seconds 5
        } catch {
            Write-Host "  Failed: $($_.Exception.Message)"
        }
    } else {
        Write-Host "  test.xlam not found (skipping)"
    }

    Start-Sleep -Seconds 2
} catch {
    Write-Host "  OUTER ERROR: $($_.Exception.Message)"
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
