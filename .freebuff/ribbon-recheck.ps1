$ErrorActionPreference = 'Continue'

Write-Host "=== 1. Excel processes ==="
$p = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($p) {
    $p | Select-Object Id, MainWindowTitle, StartTime, Responding | Format-Table -AutoSize | Out-String | Write-Host
} else {
    Write-Host "  No Excel running - good, cache can be safely cleared and won't be rewritten."
}

Write-Host "`n=== 2. Ribbon cache state ==="
$paths = @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)
foreach ($pp in $paths) {
    if (Test-Path $pp) {
        $item = Get-Item $pp
        Write-Host "  EXISTS: $pp (LastWrite: $($item.LastWriteTime))"
    } else {
        Write-Host "  absent: $pp"
    }
}

Write-Host "`n=== 3. Add-in log tail (last 30 lines) ==="
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
if (Test-Path $log) {
    Get-Content $log -Tail 30 | Write-Host
} else {
    Write-Host "  NO LOG FILE"
}

Write-Host "`n=== 4. Deployed file state ==="
$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
if (Test-Path $f) {
    $i = Get-Item $f
    Write-Host "  Path: $f"
    Write-Host "  Size: $($i.Length)  LastWrite: $($i.LastWriteTime)  IsReadOnly: $($i.IsReadOnly)"
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($f)
        $names = @($z.Entries | ForEach-Object { $_.FullName })
        $z.Dispose()
        Write-Host "  customUI14.xml present: $($names -contains 'customUI14.xml')"
        Write-Host "  customUI.xml present:   $($names -contains 'customUI/customUI.xml')"
        $sigParts = @('xl/vbaProjectSignature.bin','xl/vbaProjectSignatureAgile.bin','xl/vbaProjectSignatureV3.bin') | Where-Object { $names -contains $_ }
        if ($sigParts) { Write-Host "  SIGNATURE PARTS PRESENT: $($sigParts -join ', ')" } else { Write-Host "  signature parts: NONE" }
    } catch {
        Write-Host "  Could not open zip: $($_.Exception.Message)"
    }
} else {
    Write-Host "  MISSING: $f"
}

Write-Host "`n=== 5. Add-in registry load state ==="
$loadKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Load"
if (Test-Path $loadKey) {
    Get-ChildItem $loadKey | ForEach-Object {
        $v = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue)
        Write-Host "  $($_.PSChildName) = $($v.'(default)')"
    }
} else {
    Write-Host "  (no Add-in Load key)"
}
