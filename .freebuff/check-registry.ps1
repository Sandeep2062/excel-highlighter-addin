$ErrorActionPreference = "Continue"
Write-Host "=== Resiliency / DisabledItems ==="
$keys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\StartupItems"
)
foreach ($k in $keys) {
    Write-Host "`n$k :"
    if (Test-Path $k) {
        $v = Get-ItemProperty $k -ErrorAction SilentlyContinue
        $v.psobject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
            Write-Host "  $($_.Name) = $($_.Value)"
        }
        $subs = Get-ChildItem $k -ErrorAction SilentlyContinue
        foreach ($sub in $subs) {
            Write-Host "  [subkey] $($sub.PSChildName)"
            $sv = Get-ItemProperty $sub.PSPath -ErrorAction SilentlyContinue
            $sv.psobject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
                Write-Host "    $($_.Name) = $($_.Value)"
            }
        }
    } else {
        Write-Host "  (not present)"
    }
}
Write-Host "`n=== Add-in Manager keys ==="
$am = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager"
if (Test-Path $am) {
    $v = Get-ItemProperty $am -ErrorAction SilentlyContinue
    $v.psobject.Properties | Where-Object { $_.Name -notlike "PS*" } | ForEach-Object {
        Write-Host "  $($_.Name) = $($_.Value)"
    }
} else {
    Write-Host "  (not present)"
}
Write-Host "`n=== Log tail ==="
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (Test-Path $log) { Get-Content $log -Tail 6 | Out-String } else { Write-Host "no log" }
