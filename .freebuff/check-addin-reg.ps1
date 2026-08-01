$ErrorActionPreference = 'Continue'

Write-Host "=== HKCU Excel Options OPEN values ==="
$opts = Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -ErrorAction SilentlyContinue
$opts.psobject.Properties | Where-Object { $_.Name -like 'OPEN*' } | ForEach-Object {
    Write-Host "  $($_.Name) = $($_.Value)"
}

Write-Host "`n=== Add-in Manager registry (HKLM) ==="
$am = "HKLM:\Software\Microsoft\Office\Excel\Addins"
if (Test-Path $am) {
    Get-ChildItem $am -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        Write-Host "  $($_.PSChildName): LoadBehavior=$($v.LoadBehavior) FriendlyName=$($v.FriendlyName) Manifest=$($v.Manifest)"
    }
} else { Write-Host "  (not present)" }

Write-Host "`n=== XLAM copies on disk ==="
$candidates = @(
    (Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"),
    "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"
)
foreach ($c in $candidates) {
    if (Test-Path $c) {
        $i = Get-Item $c
        Write-Host "  EXISTS: $c"
        Write-Host "    Size: $($i.Length)  LastWrite: $($i.LastWriteTime)"
    } else {
        Write-Host "  missing: $c"
    }
}

Write-Host "`n=== Disabled Items (Resiliency) ==="
$dk = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
if (Test-Path $dk) {
    Get-ItemProperty $dk -ErrorAction SilentlyContinue | Format-List | Out-String | Write-Host
} else { Write-Host "  (none)" }

Write-Host "`n=== AddInLoadTimes detail ==="
$lt = Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Excel\AddInLoadTimes" -ErrorAction SilentlyContinue
$lt.psobject.Properties | Where-Object { $_.Name -like '*excel*' -or $_.Name -like '*xlam*' } | ForEach-Object {
    Write-Host "  $($_.Name) = $($_.Value)"
}
