$ErrorActionPreference = 'Continue'

function Test-RibbonTab($excel, $label) {
    Write-Host "=== $label ==="
    try {
        $found = $false
        foreach ($cbName in @('tabHighlighter', 'Highlighter')) {
            try {
                $cb = $excel.CommandBars.Item($cbName)
                if ($cb) {
                    Write-Host "  FOUND CommandBar: '$cbName' - NameLocal='$($cb.NameLocal)' Visible=$($cb.Visible)"
                    $found = $true
                }
            } catch { }
        }
        if (-not $found) {
            Write-Host "  tabHighlighter NOT in CommandBars"
        }
        # List a few tabs that always exist, to prove COM enumeration works
        $sample = @()
        try {
            for ($i = 1; $i -le 40; $i++) {
                $cb = $excel.CommandBars.Item($i)
                if ($cb.Type -eq 1 -and $cb.Name -like 'tab*') { $sample += $cb.Name }
            }
        } catch {}
        Write-Host "  sample ribbon tabs: $($sample -join ', ')"
    } catch {
        Write-Host "  ERROR enumerating: $($_.Exception.Message)"
    }
}

Write-Host "=== PART 1: Attach to RUNNING Excel (user's session) ==="
$running = $null
try {
    $running = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    Write-Host "  Attached. Visible=$($running.Visible), ActiveWorkbook=$($running.ActiveWorkbook.Name)"
    Write-Host "  AddIns count: $($running.AddIns.Count)"
    foreach ($ai in $running.AddIns) {
        Write-Host "    AddIn: Name='$($ai.Name)' Installed=$($ai.Installed) Path='$($ai.Path)'"
    }
    Test-RibbonTab $running "RUNNING Excel"
} catch {
    Write-Host "  Could not attach to running Excel: $($_.Exception.Message)"
}

Write-Host "`n=== PART 2: FRESH hidden Excel instance (clean load test) ==="
$fresh = $null
try {
    $fresh = New-Object -ComObject Excel.Application
    $fresh.Visible = $false
    $fresh.DisplayAlerts = $false
    $fresh.EnableEvents = $false
    try { $fresh.AutomationSecurity = 1 } catch {}

    $xlam = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
    Write-Host "  Opening add-in: $xlam"
    $wb = $fresh.Workbooks.Open($xlam)
    Write-Host "  Opened. IsAddin=$($wb.IsAddin)"

    $found = $false
    try {
        $cb = $fresh.CommandBars.Item('tabHighlighter')
        if ($cb) {
            Write-Host "  FRESH INSTANCE: tabHighlighter FOUND (ribbon WOULD build)"
            $found = $true
        }
    } catch {}
    if (-not $found) {
        Write-Host "  FRESH INSTANCE: tabHighlighter NOT found - ribbon build FAILS even on clean load"
        # Try building the ribbon explicitly
        try {
            $ui = $fresh.CustomUI
            Write-Host "  CustomUI object accessible"
        } catch {
            Write-Host "  CustomUI error: $($_.Exception.Message)"
        }
    }
    $wb.Close($false)
} catch {
    Write-Host "  FRESH INSTANCE ERROR: $($_.Exception.Message)"
} finally {
    if ($fresh) {
        try { $fresh.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($fresh) | Out-Null } catch {}
    }
}
Write-Host "`n=== done ==="
