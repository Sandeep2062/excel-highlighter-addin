# Verifies whether the VBA signature inside the deployed xlam actually covers
# the vbaProject.bin bytes (VALID) or was invalidated by a later save (STALE).
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.IO.Compression.FileSystem

$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$z = [System.IO.Compression.ZipFile]::OpenRead($f)

$e1 = $z.GetEntry("xl/vbaProjectSignature.bin")
$e2 = $z.GetEntry("xl/vbaProject.bin")
if (-not $e1 -or -not $e2) { Write-Host "Signature or project part missing"; $z.Dispose(); exit 1 }

$sig = New-Object byte[] $e1.Length
$s = $e1.Open(); [void]$s.Read($sig, 0, $sig.Length); $s.Close()
$proj = New-Object byte[] $e2.Length
$s2 = $e2.Open(); [void]$s2.Read($proj, 0, $proj.Length); $s2.Close()
$z.Dispose()

Write-Host "vbaProjectSignature.bin : $($sig.Length) bytes"
Write-Host "vbaProject.bin           : $($proj.Length) bytes"

# --- locate the PKCS#7 SignedData: "30 82 <len> 06 09 2a 86 48 86 f7 0d 01 07 02"
$oid = @(0x06, 0x09, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x07, 0x02)
$oidPos = -1
for ($i = 0; $i -lt $sig.Length - $oid.Length; $i++) {
    $ok = $true
    for ($j = 0; $j -lt $oid.Length; $j++) {
        if ($sig[$i + $j] -ne $oid[$j]) { $ok = $false; break }
    }
    if ($ok) { $oidPos = $i; break }
}
if ($oidPos -lt 0) { Write-Host "SignedData OID not found"; exit 1 }

# SEQUENCE header should be 4 bytes before the OID: 30 82 LL LL
$start = $oidPos - 4
if ($start -lt 0 -or $sig[$start] -ne 0x30) {
    Write-Host "Unexpected layout: no SEQUENCE before OID (start=$start) - trying from OID-2"
    $start = $oidPos - 2
}
$lenHigh = $sig[$start + 2]
$lenLow  = $sig[$start + 3]
$len = ([int]$lenHigh -shl 8) -bor [int]$lenLow
$total = $start + 4 + $len
Write-Host "PKCS#7 at offset $start, length 0x$($len.ToString('X4')) ($len bytes) -> slice [$start..$total) = $($total - $start) bytes"

if ($total -gt $sig.Length -or $len -le 10) {
    Write-Host "Suspicious length ($len) - falling back to OID-2 with next-bytes length"
    $start = $oidPos - 2
    $lenHigh = $sig[$start + 2]
    $lenLow  = $sig[$start + 3]
    $len = ([int]$lenHigh -shl 8) -bor [int]$lenLow
    $total = $start + 4 + $len
    Write-Host "  retry: PKCS#7 at offset $start, length $len -> slice $($total - $start) bytes"
}

$p7bytes = New-Object byte[] ($total - $start)
[Array]::Copy($sig, $start, $p7bytes, 0, $total - $start)

$cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
try {
    $cms.Decode($p7bytes)
} catch {
    Write-Host "Decode failed: $($_.Exception.Message)"
    # last resort: slice everything from $start to end of file
    $p7bytes = New-Object byte[] ($sig.Length - $start)
    [Array]::Copy($sig, $start, $p7bytes, 0, $sig.Length - $start)
    $cms.Decode($p7bytes)
    Write-Host "  (decoded using remainder of file instead)"
}

Write-Host "CMS content type : $($cms.ContentInfo.ContentType)"
$signer = $cms.SignerInfos[0]
Write-Host "Signer subject   : $($signer.Certificate.Subject)"
Write-Host "Digest algorithm : $($signer.DigestAlgorithm.FriendlyName)"

try {
    $cms.CheckSignature($true)
    Write-Host "CMS signature    : VERIFIES (cert chain checked)"
} catch {
    Write-Host "Chain check      : $($_.Exception.Message)"
    try {
        $cms.CheckSignature($false)
        Write-Host "CMS signature    : VERIFIES (crypto only)"
    } catch {
        Write-Host "CMS signature    : INVALID - $($_.Exception.Message)"
    }
}

# --- compare digest embedded in the SPC indirect data with MD5(vbaProject.bin)
$content = $cms.ContentInfo.Content
$md5 = [System.Security.Cryptography.MD5]::Create()
$projHash = $md5.ComputeHash($proj)
$hexProj = [BitConverter]::ToString($projHash).Replace("-", "")
Write-Host ""
Write-Host "MD5(vbaProject.bin) = $hexProj"

$found = $false
for ($i = 0; $i -lt $content.Length - 18; $i++) {
    if ($content[$i] -eq 0x04 -and $content[$i + 1] -eq 0x10) {
        $cand = New-Object byte[] 16
        [Array]::Copy($content, $i + 2, $cand, 0, 16)
        $hexCand = [BitConverter]::ToString($cand).Replace("-", "")
        if ($hexCand -eq $hexProj) {
            Write-Host "PROJECT DIGEST MATCH at content offset $i -> signature COVERS the current project (VALID)."
            $found = $true
            break
        }
    }
}
if (-not $found) {
    Write-Host "NO embedded digest matches MD5(vbaProject.bin) -> signature is STALE (project was modified/saved after signing)."
}
