$ErrorActionPreference = 'Continue'

Write-Host "=========================================================="
Write-Host " Excel Highlighter - Valid Ribbon Test (round 2)"
Write-Host "=========================================================="

$excelProc = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Host "ERROR: Excel is running (PID(s): $($excelProc.Id -join ', ')). Close ALL Excel windows first." -ForegroundColor Red
    exit 1
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# The deployed file is NOW the clean fresh build (130834 bytes, no signatures).
$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
if (-not (Test-Path $deployed)) { Write-Host "ERROR: deployed file missing" -ForegroundColor Red; exit 1 }
$depLen = (Get-Item $deployed).Length
Write-Host "Deployed file: $depLen bytes"
if ($depLen -ne 130834) {
    Write-Host "WARNING: deployed file is NOT the expected clean 130834-byte build." -ForegroundColor Yellow
    Write-Host "If it changed (e.g. re-signed), the test below would inherit stale signature parts." -ForegroundColor Yellow
    Write-Host "Continuing anyway - but re-run the ribbon-reset.ps1 first if in doubt." -ForegroundColor Yellow
}

# 1. Build a VALID minimal test: copy the clean build, replace ONLY the
#    customUI14.xml content with a minimal ribbon (different tab id/label).
#    The package stays otherwise identical (docProps, vbaProject, rels, etc.)
$xlstart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlstart)) { New-Item -ItemType Directory -Path $xlstart -Force | Out-Null }
$minOut = Join-Path $xlstart "minimal-ribbon-test.xlam"
Copy-Item $deployed $minOut -Force

$minimalXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabMinimalTest2" label="Minimal Test">
        <group id="grpMinimalTest2" label="Static">
          <button id="btnMinimalTest2" label="Static Button" imageMso="HappyFace"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@

$zip = [System.IO.Compression.ZipFile]::Open($minOut, [System.IO.Compression.ZipArchiveMode]::Update)
$entry = $zip.GetEntry('customUI/customUI14.xml')
if ($entry) { $entry.Delete() } else { Write-Host "WARNING: customUI/customUI14.xml not found in deployed copy!" }
$ne = $zip.CreateEntry('customUI/customUI14.xml')
$sw = New-Object System.IO.StreamWriter($ne.Open())
$sw.Write($minimalXml); $sw.Close()
$zip.Dispose()
Write-Host "Built valid minimal test add-in: $minOut ($((Get-Item $minOut).Length) bytes)"

# Verify package integrity of the test add-in
$z2 = [System.IO.Compression.ZipFile]::OpenRead($minOut)
$names = @($z2.Entries | ForEach-Object { $_.FullName })
$uiEntry = $z2.GetEntry('customUI/customUI14.xml')
$sr = New-Object System.IO.StreamReader($uiEntry.Open())
$uiContent = $sr.ReadToEnd(); $sr.Close()
$z2.Dispose()
Write-Host "  entries: docProps present: $($names -contains 'docProps/core.xml') | vbaProject: $($names -contains 'xl/vbaProject.bin') | customUI14: $($names -contains 'customUI/customUI14.xml')"
Write-Host "  customUI content has tabMinimalTest2: $($uiContent.Contains('tabMinimalTest2'))"

# 2. Clear the ENTIRE CustomUIValidationCache (documented fix: stale entries
#    there can make Office permanently skip an add-in's custom UI; it is only
#    a cache and other add-ins will simply re-validate)
Write-Host "`n=== Clearing CustomUIValidationCache (entire key - it is only a cache) ==="
$ck = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $ck) {
    Remove-Item $ck -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  Removed: $ck"
} else {
    Write-Host "  not present"
}

# 3. Clear ribbon cache files again
Write-Host "`n=== Clearing Excel.officeUI ==="
foreach ($p in @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)) {
    if (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue; Write-Host "  Cleared: $p" } else { Write-Host "  absent: $p" }
}

# 4. Also clear the Office 'RibbonX' state if present
Write-Host "`n=== Clearing RibbonX state keys if present ==="
foreach ($k in @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\StartupItems",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
)) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue; Write-Host "  Removed: $k" } else { Write-Host "  absent: $k" }
}

Write-Host ""
Write-Host "=========================================================="
Write-Host " DONE. Restart Excel and check the ribbon tabs."
Write-Host "=========================================================="
Write-Host ""
Write-Host "This test add-in is a VALID copy of the real add-in with only the"
Write-Host "ribbon XML swapped (it still carries the full VBA project, so it will also"
Write-Host "log StartUp/hotkey events - those are normal and expected)."
Write-Host ""
Write-Host "IMPORTANT: the result is VISUAL ONLY. The minimal ribbon has no onLoad"
Write-Host "callback, so the 'Ribbon UI attached' log line will NOT fire for it."
Write-Host "Judge the test purely by whether a 'Minimal Test' tab appears:"
Write-Host ""
Write-Host "  - 'Minimal Test' tab appears:      Excel CAN build custom ribbons ->"
Write-Host "                                     the Highlighter XML itself is the issue."
Write-Host "  - 'Minimal Test' does NOT appear:  Excel/Office on this machine is broken"
Write-Host "                                     at the UI level -> Office repair."
Write-Host ""
Write-Host "Also check: does the REAL 'Highlighter' tab appear in the same session?"
Write-Host "  - Both appear: the stale cache/signature was the problem - fully fixed."
Write-Host "  - Only 'Highlighter': odd but means the minimal XML specifically fails."
Write-Host ""
Write-Host "To remove the test add-in later, delete:"
Write-Host "  $minOut"
