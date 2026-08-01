param(
    [string]$DeployedPath,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Host "=== Building Full-XML Test (no onLoad) ===" -ForegroundColor Cyan

Copy-Item $DeployedPath $OutputPath -Force
Write-Host "Copied deployed add-in"

$z = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)

# Read the source customUI14.xml
$sourceXml = Get-Content "D:\Github\excel-highlighter-addin\customUI\customUI14.xml" -Raw

# Remove the onLoad attribute from <customUI> tag
$modifiedXml = $sourceXml -replace 'onLoad="RibbonCallbacks\.onLoad"', ''

# Delete existing entries
$e14 = $z.GetEntry("customUI/customUI14.xml"); if ($e14) { $e14.Delete() }
$e07 = $z.GetEntry("customUI/customUI.xml"); if ($e07) { $e07.Delete() }

# Write modified customUI14.xml
$new14 = $z.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w14 = New-Object System.IO.StreamWriter($new14.Open())
$w14.Write($modifiedXml)
$w14.Close()

# Write same for Office 2007 version
$sourceXml07 = Get-Content "D:\Github\excel-highlighter-addin\customUI\customUI.xml" -Raw
$modifiedXml07 = $sourceXml07 -replace 'onLoad="RibbonCallbacks\.onLoad"', ''
$new07 = $z.CreateEntry("customUI/customUI.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w07 = New-Object System.IO.StreamWriter($new07.Open())
$w07.Write($modifiedXml07)
$w07.Close()

$z.Dispose()

# Verify
Write-Host "`n=== Verifying ===" -ForegroundColor Yellow
$z2 = [System.IO.Compression.ZipFile]::OpenRead($OutputPath)
$uiEntry = $z2.GetEntry("customUI/customUI14.xml")
$sr = New-Object System.IO.StreamReader($uiEntry.Open())
$content = $sr.ReadToEnd(); $sr.Close()
$z2.Dispose()

if ($content -match 'onLoad') {
    Write-Host "ERROR: onLoad still present!" -ForegroundColor Red
} else {
    Write-Host "OK: onLoad attribute removed" -ForegroundColor Green
}

Write-Host "`nOutput: $OutputPath"
Write-Host "File size: $((Get-Item $OutputPath).Length) bytes"
Write-Host ""
Write-Host "This has the FULL ribbon XML (all controls, all callbacks)" -ForegroundColor Yellow
Write-Host "but WITHOUT the onLoad attribute on <customUI>." -ForegroundColor Yellow
Write-Host "If the Highlighter tab appears now, the issue is onLoad resolution." -ForegroundColor Yellow
Write-Host "If it still doesn't appear, the issue is elsewhere in the XML." -ForegroundColor Yellow
