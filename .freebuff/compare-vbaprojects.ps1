$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$repo = "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"

function Extract-VbaProject($xlam, $out) {
    $copy = Join-Path $env:TEMP "vp-src.xlam"
    Copy-Item $xlam $copy -Force
    $zip = [System.IO.Compression.ZipFile]::OpenRead($copy)
    $e = $zip.GetEntry('xl/vbaProject.bin')
    if (-not $e) { Write-Host "  NO vbaProject.bin in $xlam"; return $null }
    $fs = [System.IO.File]::Create($out)
    $s = $e.Open()
    $s.CopyTo($fs)
    $fs.Close(); $s.Close()
    $zip.Dispose()
    Remove-Item $copy -Force
    return $out
}

$depOut = Join-Path $env:TEMP "vp-deployed.bin"
$repoOut = Join-Path $env:TEMP "vp-repo.bin"

Write-Host "=== Extract vbaProject.bin ==="
Extract-VbaProject $deployed $depOut | Out-Null
Extract-VbaProject $repo $repoOut | Out-Null
Write-Host "  Deployed vbaProject.bin: $((Get-Item $depOut).Length) bytes"
Write-Host "  Repo/fresh vbaProject.bin: $((Get-Item $repoOut).Length) bytes"

Write-Host "`n=== Search both for key ribbon strings ==="
foreach ($pair in @(@('DEPLOYED', $depOut), @('REPO', $repoOut))) {
    $label = $pair[0]
    $path = $pair[1]
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
    Write-Host "--- $label ---"
    foreach ($needle in @('RibbonCallbacks', 'onLoad', 'Ribbon UI attached', 'GetProfilesContent', 'tabHighlighter', 'ThisWorkbook', 'Event sink attached')) {
        $found = $ascii.Contains($needle)
        Write-Host "  '$needle': $found"
    }
}

Write-Host "`n=== Diff the two projects (first differing offset) ==="
$b1 = [System.IO.File]::ReadAllBytes($depOut)
$b2 = [System.IO.File]::ReadAllBytes($repoOut)
$diffCount = 0
$firstDiff = -1
for ($i = 0; $i -lt [Math]::Max($b1.Length, $b2.Length); $i++) {
    $v1 = if ($i -lt $b1.Length) { $b1[$i] } else { -1 }
    $v2 = if ($i -lt $b2.Length) { $b2[$i] } else { -1 }
    if ($v1 -ne $v2) {
        if ($firstDiff -lt 0) { $firstDiff = $i }
        $diffCount++
    }
}
Write-Host "  Total differing bytes: $diffCount (of $([Math]::Max($b1.Length, $b2.Length)))"
Write-Host "  First difference at offset: $firstDiff (0x$($firstDiff.ToString('X')))"
if ($firstDiff -ge 0) {
    $ctx1 = $b1[($firstDiff-16)..[Math]::Min($firstDiff+16, $b1.Length-1)] -join ' '
    $ctx2 = $b2[($firstDiff-16)..[Math]::Min($firstDiff+16, $b2.Length-1)] -join ' '
    Write-Host "  DEPLOYED context: $ctx1"
    Write-Host "  REPO context:     $ctx2"
}

Remove-Item $depOut, $repoOut -Force -ErrorAction SilentlyContinue
Write-Host "`n=== done ==="
