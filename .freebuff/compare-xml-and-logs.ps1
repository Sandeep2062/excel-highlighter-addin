$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$source = "D:\Github\excel-highlighter-addin\customUI\customUI14.xml"
$repoXlam = "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"

Write-Host "=== 1. Deployed customUI14.xml vs source (byte diff) ==="
$copy = Join-Path $env:TEMP "deployed-diff.xlam"
Copy-Item $deployed $copy -Force
$zip = [System.IO.Compression.ZipFile]::OpenRead($copy)
$e = $zip.GetEntry('customUI/customUI14.xml')
$sr = New-Object System.IO.StreamReader($e.Open())
$deployedXml = $sr.ReadToEnd(); $sr.Close()
$zip.Dispose()
Remove-Item $copy -Force

$srcXml = Get-Content $source -Raw -Encoding UTF8
Write-Host "  Deployed XML bytes: $([System.Text.Encoding]::UTF8.GetByteCount($deployedXml))"
Write-Host "  Source XML bytes:   $([System.Text.Encoding]::UTF8.GetByteCount($srcXml))"
if ($deployedXml.Trim() -eq $srcXml.Trim()) {
    Write-Host "  >>> IDENTICAL content (trimmed) - deployed XML is the real source XML"
} else {
    Write-Host "  >>> DIFFERENT content! Here's the diff of first differing line:"
    $dLines = $deployedXml -split "`n"
    $sLines = $srcXml -split "`n"
    for ($i = 0; $i -lt [Math]::Max($dLines.Count, $sLines.Count); $i++) {
        if ($dLines[$i].Trim() -ne $sLines[$i].Trim()) {
            Write-Host "    Line $($i+1):"
            Write-Host "      DEPLOYED: $($dLines[$i])"
            Write-Host "      SOURCE:   $($sLines[$i])"
            break
        }
    }
}

Write-Host "`n=== 2. Repo-root xlam customUI check ==="
if (Test-Path $repoXlam) {
    $copy2 = Join-Path $env:TEMP "repo-diff.xlam"
    Copy-Item $repoXlam $copy2 -Force
    $z2 = [System.IO.Compression.ZipFile]::OpenRead($copy2)
    $names = @($z2.Entries | ForEach-Object { $_.FullName })
    Write-Host "  customUI14.xml present: $($names -contains 'customUI/customUI14.xml')"
    Write-Host "  vbaProject.bin present: $($names -contains 'xl/vbaProject.bin')"
    Write-Host "  signature parts: $(@('xl/vbaProjectSignature.bin','xl/vbaProjectSignatureAgile.bin','xl/vbaProjectSignatureV3.bin') | Where-Object { $names -contains $_ } | ForEach-Object { $_ } | Out-String)"
    $e2 = $z2.GetEntry('xl/vbaProject.bin')
    Write-Host "  vbaProject.bin size: $($e2.Length)"
    $z2.Dispose()
    Remove-Item $copy2 -Force
} else { Write-Host "  repo xlam missing" }

Write-Host "`n=== 3. Ribbon-related registry keys ==="
@(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Ribbon",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
) | ForEach-Object {
    Write-Host "--- $_ ---"
    if (Test-Path $_) {
        $props = Get-ItemProperty $_ -ErrorAction SilentlyContinue
        if ($props) {
            $props.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                Write-Host "  $($_.Name) = $($_.Value)"
            }
        }
    } else { Write-Host "  (not present)" }
}

Write-Host "`n=== 4. Windows Application event log - Excel/Office errors (last 3 days) ==="
try {
    $since = (Get-Date).AddDays(-3)
    $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$since} -ErrorAction SilentlyContinue |
        Where-Object { $_.ProviderName -like '*Office*' -or $_.ProviderName -like '*Excel*' -or $_.Message -like '*Excel*' } |
        Select-Object -First 15
    if ($events) {
        $events | ForEach-Object {
            Write-Host "  [$($_.TimeCreated.ToString('MM-dd HH:mm'))] $($_.ProviderName): $($_.LevelDisplayName)"
            Write-Host "    $($_.Message.Substring(0, [Math]::Min(300, $_.Message.Length)))"
        }
    } else {
        Write-Host "  (no matching events)"
    }
} catch {
    Write-Host "  Event log check failed: $($_.Exception.Message)"
}

Write-Host "`n=== done ==="
