$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
Write-Host "=== Deployed xlam: $f ==="
Write-Host "Exists: $(Test-Path $f)  Size: $((Get-Item $f).Length) bytes"

# Open with read sharing so we can peek even while Excel has it open.
try {
    $fs = New-Object System.IO.FileStream($f, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $z = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
    Write-Host "`n--- ALL entries ---"
    $z.Entries | ForEach-Object { Write-Host "  $($_.FullName) ($($_.Length) bytes)" }

    Write-Host "`n--- _rels/.rels ---"
    $rels = $z.GetEntry("_rels/.rels")
    if ($rels) { $sr = New-Object System.IO.StreamReader($rels.Open()); Write-Host $sr.ReadToEnd(); $sr.Close() }

    Write-Host "`n--- customUI/customUI14.xml exists? ---"
    $cu = $z.GetEntry("customUI/customUI14.xml")
    Write-Host ($(if ($cu) { "YES ($($cu.Length) bytes)" } else { "NOT PRESENT - RIBBON CANNOT LOAD!" }))

    Write-Host "`n--- [Content_Types].xml ---"
    $ct = $z.GetEntry("[Content_Types].xml")
    if ($ct) { $sr2 = New-Object System.IO.StreamReader($ct.Open()); Write-Host $sr2.ReadToEnd(); $sr2.Close() }

    $z.Dispose(); $fs.Dispose()
} catch {
    Write-Host "ERROR reading file: $($_.Exception.Message)"
}

Write-Host "`n=== Excel.officeUI caches (both locations) ==="
Get-ChildItem (Join-Path $env:APPDATA "Microsoft\Office") -Filter "*officeUI*" -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize | Out-String
Get-ChildItem (Join-Path $env:LOCALAPPDATA "Microsoft\Office") -Filter "*officeUI*" -ErrorAction SilentlyContinue | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize | Out-String

Write-Host "=== install.ps1 officeUI handling ==="
$install = "D:\Github\excel-highlighter-addin\install.ps1"
if (Test-Path $install) {
    Select-String -Path $install -Pattern "officeUI|\.officeUI" -Context 2,2 | ForEach-Object { $_.ToString() }
} else { Write-Host "install.ps1 not found" }
