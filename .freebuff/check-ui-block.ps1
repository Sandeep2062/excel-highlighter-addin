$ErrorActionPreference = 'Continue'

Write-Host "=== 1. Did the reset apply? ==="
$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
if (Test-Path $deployed) {
    $i = Get-Item $deployed
    Write-Host "  Deployed: $($i.Length) bytes, LastWrite $($i.LastWriteTime) (clean build = 130834 bytes)"
} else {
    Write-Host "  deployed MISSING"
}
$xlstart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$min = Join-Path $xlstart "minimal-ribbon-test.xlam"
Write-Host "  Minimal test add-in: $(if (Test-Path $min) { 'PRESENT (' + (Get-Item $min).Length + ' bytes)' } else { 'MISSING - user may have removed it' })"
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
Write-Host "  Excel running now: $(if ($p) { 'YES - PID ' + ($p.Id -join ',') } else { 'no' })"

Write-Host "`n=== 2. Log tail (did onLoad EVER fire after reset?) ==="
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (Test-Path $log) {
    Get-Content $log -Tail 12 | Write-Host
    Write-Host "`n  'Ribbon UI attached' total: $((Select-String -Path $log -Pattern 'Ribbon UI attached').Count)"
}

Write-Host "`n=== 3. DEEP registry sweep: every value whose name mentions Disable/UI/Ribbon/Extensib/Custom/Addin under Office 16.0 ==="
$roots = @(
    'HKCU:\Software\Microsoft\Office\16.0',
    'HKLM:\Software\Microsoft\Office\16.0',
    'HKCU:\Software\Policies\Microsoft\Office\16.0',
    'HKLM:\Software\Policies\Microsoft\Office\16.0',
    'HKCU:\Software\Microsoft\Office\Excel',
    'HKLM:\Software\Microsoft\Office\Excel'
)
foreach ($root in $roots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem $root -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $item = $_
        try {
            $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
            if ($props) {
                $props.psobject.Properties | Where-Object {
                    $_.Name -notlike 'PS*' -and
                    ($_.Name -match 'Disable|UI|Ribbon|Extensib|CustomUI|Addin' -or
                     "$($_.Value)" -match 'Disable.*UI|UI.*Extensib')
                } | ForEach-Object {
                    Write-Host "  [$($item.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','')] $($_.Name) = $($_.Value)"
                }
            }
        } catch {}
    }
}

Write-Host "`n=== 4. Trust Center files on disk ==="
foreach ($d in @(
    "$env:APPDATA\Microsoft\Office\TrustCenter",
    "$env:APPDATA\Microsoft\Office\16.0\TrustCenter",
    "$env:LOCALAPPDATA\Microsoft\Office\TrustCenter"
)) {
    Write-Host "--- $d ---"
    if (Test-Path $d) { Get-ChildItem $d -Recurse -ErrorAction SilentlyContinue | Select-Object FullName, Length, LastWriteTime | Format-List | Out-String | Write-Host } else { Write-Host "  (not present)" }
}

Write-Host "`n=== 5. All Excel/Office installs and click-to-run state ==="
try {
    $c2r = Get-ItemProperty "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($c2r) {
        $c2r.psobject.Properties | Where-Object { $_.Name -match 'ProductReleaseIds|VersionToReport|Platform|InstallationPath' } | ForEach-Object {
            Write-Host "  C2R: $($_.Name) = $($_.Value)"
        }
    }
} catch {}
Get-Item "HKLM:\Software\Classes\Excel.Application\CurVer" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  Excel.Application CurVer: $((Get-ItemProperty $_.PSPath).'(default)')"
}

Write-Host "`n=== done ==="
