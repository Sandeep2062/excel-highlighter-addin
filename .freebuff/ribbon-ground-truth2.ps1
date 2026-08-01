$ErrorActionPreference = 'Continue'

function Get-RibbonTabs($excel) {
    $tabs = @()
    try {
        foreach ($cb in $excel.CommandBars) {
            try {
                $n = [string]$cb.Name
                if ($n -like 'Tab*' -or $n -like 'tab*') { $tabs += $n }
            } catch {}
        }
    } catch {
        return @("ENUM-ERROR: $($_.Exception.Message)")
    }
    return $tabs
}

Write-Host "=== PART 1: RUNNING Excel (user session) ==="
$running = $null
try {
    $running = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    Write-Host "  Attached OK. Visible=$($running.Visible)"
    $tabs = Get-RibbonTabs $running
    Write-Host "  All ribbon tabs found ($($tabs.Count)):"
    $tabs | ForEach-Object { Write-Host "    - $_" }
    $hl = $tabs | Where-Object { $_ -like '*ighlighter*' -or $_ -like '*ighlight*' }
    if ($hl) { Write-Host "  >>> HIGHLIGHTER TAB EXISTS: $hl" } else { Write-Host "  >>> NO highlighter tab in running Excel" }
} catch {
    Write-Host "  Attach failed: $($_.Exception.Message)"
}

Write-Host "`n=== PART 2: FRESH hidden instance ==="
$fresh = $null
try {
    $fresh = New-Object -ComObject Excel.Application
    $fresh.Visible = $false
    $fresh.DisplayAlerts = $false
    $fresh.EnableEvents = $false
    try { $fresh.AutomationSecurity = 1 } catch {}

    $xlam = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
    $wb = $fresh.Workbooks.Open($xlam)
    Write-Host "  Add-in opened. IsAddin=$($wb.IsAddin)"
    $tabs = Get-RibbonTabs $fresh
    Write-Host "  All ribbon tabs in fresh instance ($($tabs.Count)):"
    $tabs | ForEach-Object { Write-Host "    - $_" }
    $hl = $tabs | Where-Object { $_ -like '*ighlighter*' -or $_ -like '*ighlight*' }
    if ($hl) { Write-Host "  >>> HIGHLIGHTER TAB EXISTS in fresh instance: $hl" } else { Write-Host "  >>> NO highlighter tab in fresh instance" }
    $wb.Close($false)
} catch {
    Write-Host "  Fresh instance ERROR: $($_.Exception.Message)"
} finally {
    if ($fresh) { try { $fresh.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fresh) | Out-Null } catch {} }
}

Write-Host "`n=== PART 3: CustomUIValidationCache entries for our addin ==="
$ck = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $ck) {
    $props = Get-ItemProperty $ck -ErrorAction SilentlyContinue
    $props.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
        Write-Host "  $($_.Name) = $($_.Value)"
    }
} else {
    Write-Host "  (key not present)"
}

Write-Host "`n=== PART 4: images folder next to deployed add-in ==="
$imgDir = Join-Path (Split-Path $xlam) "images"
if (Test-Path $imgDir) {
    Write-Host "  images folder EXISTS: $imgDir"
    Get-ChildItem $imgDir -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "    - $($_.Name) ($($_.Length) bytes)" }
} else {
    Write-Host "  images folder MISSING: $imgDir"
}
