$ErrorActionPreference = "Stop"
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    Start-Sleep -Seconds 2
    Write-Host "Workbooks open in a fresh COM instance: $($excel.Workbooks.Count)"
    foreach ($wb in $excel.Workbooks) {
        Write-Host "  - '$($wb.Name)'  IsAddin=$($wb.IsAddin)  FullName=$($wb.FullName)"
    }
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    Start-Sleep -Milliseconds 800
}
