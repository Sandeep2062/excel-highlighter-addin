$ErrorActionPreference = 'Continue'

$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
$before = if (Test-Path $log) { (Get-Content $log).Count } else { 0 }

$fresh = "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"
Write-Host "Fresh build to test: $fresh ($((Get-Item $fresh).Length) bytes)"

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    try { $excel.AutomationSecurity = 1 } catch {}  # 1=Low, allow macros

    Write-Host "`n=== AddIns2.Add FRESH build (no signatures, clean build) ==="
    try {
        $ai = $excel.AddIns2.Add($fresh, $true)
        Write-Host "  Added: Name='$($ai.Name)' Installed=$($ai.Installed)"
    } catch {
        Write-Host "  AddIns2.Add failed: $($_.Exception.Message)"
        # Try opening directly
        try {
            $wb = $excel.Workbooks.Open($fresh)
            Write-Host "  Workbooks.Open worked instead. IsAddin=$($wb.IsAddin)"
        } catch {
            Write-Host "  Workbooks.Open also failed: $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 8

    Write-Host "`n=== Workbooks open now ==="
    for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
        $wb = $excel.Workbooks.Item($i)
        Write-Host "  $($wb.Name) IsAddin=$($wb.IsAddin)"
    }
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
