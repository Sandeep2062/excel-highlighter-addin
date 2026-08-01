$ErrorActionPreference = "Stop"
$src = "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"
$dst = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"

Write-Host "Source: $src"
Write-Host "Dest  : $dst"

# Make sure no Excel holds the file, and clear any read-only attribute.
$item = Get-Item $dst -ErrorAction SilentlyContinue
if ($item -and $item.IsReadOnly) {
    Set-ItemProperty -Path $dst -Name IsReadOnly -Value $false
    Write-Host "Cleared read-only attribute on destination."
}

Copy-Item $src $dst -Force
Write-Host "Copied. Size: $((Get-Item $dst).Length) bytes"

Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead($dst)
$sig = $z.GetEntry("xl/vbaProjectSignature.bin")
$sigA = $z.GetEntry("xl/vbaProjectSignatureAgile.bin")
$sigV = $z.GetEntry("xl/vbaProjectSignatureV3.bin")
$customUI = $z.GetEntry("customUI/customUI14.xml")
Write-Host "vbaProjectSignature.bin  : $($null -ne $sig) (want False)"
Write-Host "vbaProjectSignatureAgile  : $($null -ne $sigA) (want False)"
Write-Host "vbaProjectSignatureV3     : $($null -ne $sigV) (want False)"
Write-Host "customUI14.xml            : $($null -ne $customUI) (want True)"
$z.Dispose()

try {
    $fs = [System.IO.File]::Open($dst, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    Write-Host "Lock probe: WRITABLE"
    $fs.Close()
} catch {
    Write-Host "Lock probe: LOCKED - $($_.Exception.Message)"
}
