param(
    [string]$Path
)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path $Path)) {
    Write-Host "ERROR: File not found: $Path" -ForegroundColor Red
    exit 1
}

Write-Host "=== Inspecting: $Path ===" -ForegroundColor Cyan
Write-Host "Size: $((Get-Item $Path).Length) bytes"

$z = [System.IO.Compression.ZipFile]::OpenRead($Path)

Write-Host "`n=== All zip entries ===" -ForegroundColor Yellow
$z.Entries | ForEach-Object {
    Write-Host "  $($_.FullName)  ($($_.Length) bytes)"
}

Write-Host "`n=== Signature-related entries ===" -ForegroundColor Yellow
$sigEntries = $z.Entries | Where-Object { $_.FullName -match 'signature|vbaProject|vbaData' }
if ($sigEntries) {
    $sigEntries | ForEach-Object {
        Write-Host "  FOUND: $($_.FullName)  ($($_.Length) bytes)" -ForegroundColor Red
    }
} else {
    Write-Host "  None found (clean)" -ForegroundColor Green
}

Write-Host "`n=== _rels/.rels ===" -ForegroundColor Yellow
$relsEntry = $z.GetEntry("_rels/.rels")
if ($relsEntry) {
    $sr = New-Object System.IO.StreamReader($relsEntry.Open())
    $relsContent = $sr.ReadToEnd()
    $sr.Close()
    
    # Check for signature relationship
    if ($relsContent -match 'digital-signature') {
        Write-Host "  WARNING: Contains digital-signature relationship!" -ForegroundColor Red
    } else {
        Write-Host "  No signature relationship (clean)" -ForegroundColor Green
    }
    
    # Check for customUI relationship
    if ($relsContent -match 'extensibility') {
        Write-Host "  Has extensibility relationship (good)" -ForegroundColor Green
    } else {
        Write-Host "  WARNING: Missing extensibility relationship!" -ForegroundColor Red
    }
    
    Write-Host "`n  Full content:"
    Write-Host $relsContent
} else {
    Write-Host "  MISSING _rels/.rels!" -ForegroundColor Red
}

Write-Host "`n=== [Content_Types].xml ===" -ForegroundColor Yellow
$ctEntry = $z.GetEntry("[Content_Types].xml")
if ($ctEntry) {
    $sr = New-Object System.IO.StreamReader($ctEntry.Open())
    $ctContent = $sr.ReadToEnd()
    $sr.Close()
    
    # Check for customUI content type
    if ($ctContent -match 'customui') {
        Write-Host "  Has customUI content type (good)" -ForegroundColor Green
    } else {
        Write-Host "  WARNING: Missing customUI content type!" -ForegroundColor Red
    }
    
    # Check for signature content type
    if ($ctContent -match 'signature') {
        Write-Host "  WARNING: Contains signature content type!" -ForegroundColor Red
    } else {
        Write-Host "  No signature content type (clean)" -ForegroundColor Green
    }
    
    Write-Host "`n  Full content:"
    Write-Host $ctContent
} else {
    Write-Host "  MISSING [Content_Types].xml!" -ForegroundColor Red
}

Write-Host "`n=== customUI/customUI14.xml ===" -ForegroundColor Yellow
$uiEntry = $z.GetEntry("customUI/customUI14.xml")
if ($uiEntry) {
    $sr = New-Object System.IO.StreamReader($uiEntry.Open())
    $uiContent = $sr.ReadToEnd()
    $sr.Close()
    Write-Host "  Content:"
    Write-Host $uiContent
} else {
    Write-Host "  MISSING customUI/customUI14.xml!" -ForegroundColor Red
}

$z.Dispose()
Write-Host "`nDone." -ForegroundColor Cyan
