$ErrorActionPreference = 'Continue'

Write-Host "=== Full Resiliency tree ==="
$res = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency"
if (Test-Path $res) {
    Get-ChildItem $res -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $item = $_
        Write-Host "--- $($item.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','') ---"
        $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
        if ($props) {
            $props.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                Write-Host "  $($_.Name) = $($_.Value)"
            }
        } else {
            Write-Host "  (empty)"
        }
    }
} else {
    Write-Host "  Resiliency key not present"
}

Write-Host "`n=== XLSTART folders ==="
@(
    "$env:APPDATA\Microsoft\Excel\XLSTART",
    "$env:ProgramFiles\Microsoft Office\root\Office16\XLSTART",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\XLSTART"
) | ForEach-Object {
    Write-Host "--- $_ ---"
    if (Test-Path $_) {
        Get-ChildItem $_ -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $($_.Name)" }
    } else { Write-Host "  (not present)" }
}

Write-Host "`n=== Any registry value mentioning excel-highlighter under HKCU Office ==="
$base = "HKCU:\Software\Microsoft\Office\16.0"
if (Test-Path $base) {
    Get-ChildItem $base -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $item = $_
        try {
            $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
            if ($props) {
                $props.psobject.Properties | Where-Object {
                    $_.Name -notlike 'PS*' -and
                    ($_.Name -like '*ighlighter*' -or "$($_.Value)" -like '*ighlighter*')
                } | ForEach-Object {
                    Write-Host "  $($item.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::','') :: $($_.Name) = $($_.Value)"
                }
            }
        } catch {}
    }
}
Write-Host "`n=== done ==="
