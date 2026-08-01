param(
    [string]$Path
)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Host "=== Deep inspection of: $Path ===" -ForegroundColor Cyan
$z = [System.IO.Compression.ZipFile]::OpenRead($Path)

# Check workbook.xml
Write-Host "`n=== xl/workbook.xml ===" -ForegroundColor Yellow
$wbEntry = $z.GetEntry("xl/workbook.xml")
if ($wbEntry) {
    $sr = New-Object System.IO.StreamReader($wbEntry.Open())
    $content = $sr.ReadToEnd(); $sr.Close()
    Write-Host $content
} else {
    Write-Host "MISSING!" -ForegroundColor Red
}

# Check workbook.xml.rels
Write-Host "`n=== xl/_rels/workbook.xml.rels ===" -ForegroundColor Yellow
$wbRelsEntry = $z.GetEntry("xl/_rels/workbook.xml.rels")
if ($wbRelsEntry) {
    $sr = New-Object System.IO.StreamReader($wbRelsEntry.Open())
    $content = $sr.ReadToEnd(); $sr.Close()
    Write-Host $content
} else {
    Write-Host "MISSING!" -ForegroundColor Red
}

# Check docProps
Write-Host "`n=== docProps/app.xml ===" -ForegroundColor Yellow
$appEntry = $z.GetEntry("docProps/app.xml")
if ($appEntry) {
    $sr = New-Object System.IO.StreamReader($appEntry.Open())
    $content = $sr.ReadToEnd(); $sr.Close()
    Write-Host $content
} else {
    Write-Host "MISSING!" -ForegroundColor Red
}

Write-Host "`n=== docProps/core.xml ===" -ForegroundColor Yellow
$coreEntry = $z.GetEntry("docProps/core.xml")
if ($coreEntry) {
    $sr = New-Object System.IO.StreamReader($coreEntry.Open())
    $content = $sr.ReadToEnd(); $sr.Close()
    Write-Host $content
} else {
    Write-Host "MISSING!" -ForegroundColor Red
}

# Check customUI.xml (the old Office 2007 one - should it exist?)
Write-Host "`n=== customUI/customUI.xml (Office 2007 ribbon) ===" -ForegroundColor Yellow
$ui2007Entry = $z.GetEntry("customUI/customUI.xml")
if ($ui2007Entry) {
    Write-Host "EXISTS: $($ui2007Entry.Length) bytes"
    $sr = New-Object System.IO.StreamReader($ui2007Entry.Open())
    $content = $sr.ReadToEnd(); $sr.Close()
    # Show first 500 chars
    if ($content.Length -gt 500) {
        Write-Host $content.Substring(0, 500)
        Write-Host "... (truncated)"
    } else {
        Write-Host $content
    }
} else {
    Write-Host "Not present (normal)"
}

# Check if vbaProject.bin contains any XML signature residue
Write-Host "`n=== xl/vbaProject.bin signature check ===" -ForegroundColor Yellow
$vbaEntry = $z.GetEntry("xl/vbaProject.bin")
if ($vbaEntry) {
    Write-Host "Size: $($vbaEntry.Length) bytes"
    $stream = $vbaEntry.Open()
    $bytes = New-Object byte[] $vbaEntry.Length
    $stream.Read($bytes, 0, $vbaEntry.Length) | Out-Null
    $stream.Close()
    
    # Check for ASCII signature strings
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
    if ($ascii -match 'PKI#X509') { Write-Host "  FOUND: PKI#X509 (Authenticode signature header)" -ForegroundColor Red }
    if ($ascii -match 'MSPST') { Write-Host "  FOUND: MSPST (Office signature block)" -ForegroundColor Red }
    if ($ascii -match 'XmlSignature') { Write-Host "  FOUND: XmlSignature" -ForegroundColor Red }
    
    # Look for OLE compound document signature
    if ($bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) {
        Write-Host "  Format: OLE Compound Document (normal for vbaProject.bin)" -ForegroundColor Green
    }
    
    # Check for VBA project signature marker
    $hexFirst50 = ($bytes[0..49] | ForEach-Object { $_.ToString("X2") }) -join ' '
    Write-Host "  First 50 bytes: $hexFirst50"
    
    # Look for empty/zeroed signature regions (leftover from failed signing)
    $zeroRegions = 0
    for ($i = 0; $i -lt [Math]::Min(1000, $bytes.Length); $i++) {
        if ($bytes[$i] -eq 0) { $zeroRegions++ }
    }
    Write-Host "  Zero bytes in first 1000: $zeroRegions"
}

$z.Dispose()
Write-Host "`nDone." -ForegroundColor Cyan
