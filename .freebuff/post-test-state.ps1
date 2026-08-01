$ErrorActionPreference = 'Continue'

Write-Host "=== 1. Excel.officeUI cache state after the user's test session ==="
foreach ($p in @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)) {
    if (Test-Path $p) {
        $i = Get-Item $p
        Write-Host "  EXISTS: $p ($($i.Length) bytes, LastWrite $($i.LastWriteTime))"
        $content = Get-Content $p -Raw
        Write-Host "  Mentions Highlighter: $($content -match 'Highlighter|tabHighlighter')"
        Write-Host "  Mentions Minimal: $($content -match 'Minimal|tabMinimalTest')"
        Write-Host "  First 400 chars: $($content.Substring(0, [Math]::Min(400, $content.Length)))"
    } else {
        Write-Host "  absent: $p"
    }
}

Write-Host "`n=== 2. Event log Office/Excel events around the test session (18:20-18:35 today) ==="
try {
    $from = (Get-Date).Date.AddHours(18).AddMinutes(15)
    $to = (Get-Date).Date.AddHours(18).AddMinutes(40)
    $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$from; EndTime=$to} -ErrorAction SilentlyContinue |
        Where-Object { $_.ProviderName -like '*Office*' -or $_.ProviderName -like '*Excel*' -or $_.Message -like '*Excel*' -or $_.Message -like '*add-in*' -or $_.Message -like '*addin*' }
    if ($events) {
        $events | ForEach-Object {
            Write-Host "  [$($_.TimeCreated.ToString('HH:mm:ss'))] $($_.ProviderName): $($_.LevelDisplayName)"
            Write-Host "    $($_.Message.Substring(0, [Math]::Min(400, $_.Message.Length)))"
        }
    } else {
        Write-Host "  (no matching events)"
    }
} catch {
    Write-Host "  error: $($_.Exception.Message)"
}

Write-Host "`n=== 3. Resiliency StartupItems / full tree again ==="
$res = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency"
if (Test-Path $res) {
    Get-ChildItem $res -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        $vals = if ($props) { ($props.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ' } else { '(empty)' }
        Write-Host "  $($_.PSChildName): $vals"
    }
} else { Write-Host "  (not present)" }

Write-Host "`n=== 4. Other Office UI caches (.pip, officeUI variants) ==="
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Office" -Filter "*.pip" -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String | Write-Host
Get-ChildItem "$env:APPDATA\Microsoft\Office" -Filter "*.pip" -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String | Write-Host
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Office" -Filter "*officeUI*" -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String | Write-Host

Write-Host "`n=== 5. The minimal test add-in's actual content ==="
$min = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\minimal-ribbon-test.xlam"
if (Test-Path $min) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($min)
        $z.Entries | ForEach-Object { Write-Host "  $($_.FullName) ($($_.Length) bytes)" }
        $z.Dispose()
    } catch { Write-Host "  cannot open as zip: $($_.Exception.Message)" }
}

Write-Host "`n=== done ==="
