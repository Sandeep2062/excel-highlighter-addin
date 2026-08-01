$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
Write-Host "=== Deployed xlam: $f ==="
Write-Host "Exists: $(Test-Path $f)  Size: $((Get-Item $f).Length) bytes  LastWrite: $((Get-Item $f).LastWriteTime)"

$z = [System.IO.Compression.ZipFile]::OpenRead($f)
Write-Host "`n--- ALL entries ---"
$z.Entries | ForEach-Object { Write-Host "  $($_.FullName) ($($_.Length) bytes)" }

Write-Host "`n--- _rels/.rels ---"
$rels = $z.GetEntry("_rels/.rels")
if ($rels) { $sr = New-Object System.IO.StreamReader($rels.Open()); Write-Host $sr.ReadToEnd(); $sr.Close() }

Write-Host "`n--- customUI/customUI14.xml (first 1200 chars) ---"
$cu = $z.GetEntry("customUI/customUI14.xml")
if ($cu) {
    $sr2 = New-Object System.IO.StreamReader($cu.Open())
    $c = $sr2.ReadToEnd(); $sr2.Close()
    Write-Host $c.Substring(0, [Math]::Min(1200, $c.Length))
    Write-Host "..."
    Write-Host "(total length: $($c.Length) chars)"
} else {
    Write-Host "NOT PRESENT - this is why the ribbon doesn't load!"
}

$z.Dispose()

Write-Host "`n=== Add-in log tail (20) ==="
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (Test-Path $log) { Get-Content $log -Tail 20 | Out-String } else { Write-Host "no log" }

Write-Host "`n=== Excel processes ==="
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) { $p | Select-Object Id, MainWindowTitle, StartTime | Format-Table -AutoSize | Out-String } else { Write-Host "none" }

Write-Host "=== Excel.officeUI / ribbon cache ==="
Get-ChildItem (Join-Path $env:APPDATA "Microsoft\Office") -Filter "*officeUI*" -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize | Out-String
Get-ChildItem (Join-Path $env:LOCALAPPDATA "Microsoft\Office") -Filter "*officeUI*" -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize | Out-String

Write-Host "=== OPEN registry value ==="
Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Excel\Options" -ErrorAction SilentlyContinue | Select-Object OPEN, OPEN1, OPEN2 | Format-List | Out-String -Width 250

Write-Host "=== Add-in Manager ==="
$am = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager"
if (Test-Path $am) { Get-ItemProperty $am -ErrorAction SilentlyContinue | Format-List | Out-String } else { Write-Host "key not present" }

Write-Host "=== Add-in Load (load behavior) ==="
$al = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Load"
if (Test-Path $al) { Get-ChildItem $al | ForEach-Object { Write-Host $_.PSChildName; Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue | Format-List | Out-String } } else { Write-Host "key not present" }
