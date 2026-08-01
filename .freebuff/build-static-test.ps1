param(
    [string]$DeployedPath,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Host "=== Building Static Ribbon Test Add-in ===" -ForegroundColor Cyan

# Copy the deployed clean add-in
if (-not (Test-Path $DeployedPath)) {
    Write-Host "ERROR: Deployed add-in not found: $DeployedPath" -ForegroundColor Red
    exit 1
}

Copy-Item $DeployedPath $OutputPath -Force
Write-Host "Copied deployed add-in to: $OutputPath"

# Open the zip and replace customUI14.xml with a completely static version
$z = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)

# Delete existing customUI14.xml
$existingEntry = $z.GetEntry("customUI/customUI14.xml")
if ($existingEntry) { $existingEntry.Delete() }

# Create new static customUI14.xml - NO callbacks, NO onLoad, just static buttons
$newEntry = $z.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$writer = New-Object System.IO.StreamWriter($newEntry.Open())
$writer.Write(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabStaticTest" label="Static Test" insertAfterMso="TabHome">
        <group id="grpStaticTest" label="Static Test">
          <button id="btnStaticHello"
                  label="Hello"
                  size="large"
                  imageMso="HappyFace"
                  onAction="StaticOnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@)
$writer.Close()

# Also update customUI.xml (Office 2007 version) with same static content
$existingUi07 = $z.GetEntry("customUI/customUI.xml")
if ($existingUi07) { $existingUi07.Delete() }

$newUi07 = $z.CreateEntry("customUI/customUI.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$writer2 = New-Object System.IO.StreamWriter($newUi07.Open())
$writer2.Write(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="tabStaticTest" label="Static Test" insertAfterMso="TabHome">
        <group id="grpStaticTest" label="Static Test">
          <button id="btnStaticHello"
                  label="Hello"
                  size="large"
                  imageMso="HappyFace"
                  onAction="StaticOnAction"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@)
$writer2.Close()

$z.Dispose()

# Verify
Write-Host "`n=== Verifying package ===" -ForegroundColor Yellow
$z2 = [System.IO.Compression.ZipFile]::OpenRead($OutputPath)

foreach ($entry in $z2.Entries) {
    Write-Host "  $($entry.FullName) ($($entry.Length) bytes)"
}

# Verify the customUI14.xml content
$uiEntry = $z2.GetEntry("customUI/customUI14.xml")
$sr = New-Object System.IO.StreamReader($uiEntry.Open())
$content = $sr.ReadToEnd(); $sr.Close()
Write-Host "`ncustomUI14.xml content:"
Write-Host $content

# Verify _rels
$relsEntry = $z2.GetEntry("_rels/.rels")
$sr = New-Object System.IO.StreamReader($relsEntry.Open())
$relsContent = $sr.ReadToEnd(); $sr.Close()
if ($relsContent -match 'extensibility') {
    Write-Host "`n_rels/.rels: has extensibility relationship (good)" -ForegroundColor Green
} else {
    Write-Host "`n_rels/.rels: MISSING extensibility relationship!" -ForegroundColor Red
}

# Verify content types
$ctEntry = $z2.GetEntry("[Content_Types].xml")
$sr = New-Object System.IO.StreamReader($ctEntry.Open())
$ctContent = $sr.ReadToEnd(); $sr.Close()
if ($ctContent -match 'customUI') {
    Write-Host "[Content_Types].xml: has customUI content type (good)" -ForegroundColor Green
} else {
    Write-Host "[Content_Types].xml: MISSING customUI content type!" -ForegroundColor Red
}

$z2.Dispose()

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Output: $OutputPath"
Write-Host "File size: $((Get-Item $OutputPath).Length) bytes"
Write-Host ""
Write-Host "This is a STATIC ribbon test - no callbacks, no VBA dependency." -ForegroundColor Yellow
Write-Host "If this tab appears, the problem is in the callback wiring." -ForegroundColor Yellow
Write-Host "If this tab also does NOT appear, the problem is in the xlam package structure itself." -ForegroundColor Yellow
