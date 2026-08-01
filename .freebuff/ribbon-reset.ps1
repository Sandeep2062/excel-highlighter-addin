$ErrorActionPreference = 'Continue'

Write-Host "=========================================================="
Write-Host " Excel Highlighter - Ribbon Reset & Diagnostic"
Write-Host "=========================================================="

# 0. Refuse if Excel is running (file would be locked)
$excelProc = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Host "ERROR: Excel is running (PID(s): $($excelProc.Id -join ', '))." -ForegroundColor Red
    Write-Host "Close ALL Excel windows first (File > Exit, and check the system tray), then re-run this script."
    exit 1
}

$deployed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$fresh = "D:\Github\excel-highlighter-addin\excel-highlighter.xlam"

# 1. Redeploy the clean, unsigned fresh build over the stale-signed deployed file
Write-Host "`n=== 1. Redeploying clean build ==="
if (-not (Test-Path $fresh)) { Write-Host "ERROR: fresh build missing: $fresh" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $deployed)) { Write-Host "ERROR: deployed file missing: $deployed" -ForegroundColor Red; exit 1 }
# sign-xlam.ps1 may have locked the deployed file read-only to protect a
# previous signature - clear that before overwriting.
try { Set-ItemProperty -Path $deployed -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
try { Set-ItemProperty -Path $fresh -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
Copy-Item $fresh $deployed -Force
Write-Host "  Replaced deployed file with clean build ($((Get-Item $deployed).Length) bytes)"

# Verify the redeployed file's ribbon wiring
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead($deployed)
$names = @($z.Entries | ForEach-Object { $_.FullName })
$hasUI = $names -contains 'customUI/customUI14.xml'
$hasVBA = $names -contains 'xl/vbaProject.bin'
$hasSig = @('xl/vbaProjectSignature.bin','xl/vbaProjectSignatureAgile.bin','xl/vbaProjectSignatureV3.bin') | Where-Object { $names -contains $_ }
$rels = $z.GetEntry('_rels/.rels')
$sr = New-Object System.IO.StreamReader($rels.Open())
$relsText = $sr.ReadToEnd(); $sr.Close()
$z.Dispose()
Write-Host "  customUI14.xml: $hasUI | vbaProject.bin: $hasVBA | signature parts: $(if ($hasSig) { $hasSig -join ',' } else { 'NONE (good)' })"
Write-Host "  rels references customUI14: $($relsText.Contains('customUI/customUI14.xml'))"

# 2. Clear ribbon UI cache (both locations)
Write-Host "`n=== 2. Clearing Excel.officeUI ribbon cache ==="
foreach ($p in @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)) {
    if (Test-Path $p) {
        Remove-Item $p -Force -ErrorAction SilentlyContinue
        Write-Host "  Cleared: $p"
    } else {
        Write-Host "  absent: $p"
    }
}

# 3. Clear any CustomUIValidationCache entry for THIS add-in only
# (a stale entry there can block custom UI loading per Microsoft docs; we only
# touch values naming our add-in, leaving other add-ins' validation state alone)
Write-Host "`n=== 3. Clearing CustomUIValidationCache entry for this add-in ==="
$ck = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $ck) {
    $removed = $false
    $props = Get-ItemProperty $ck -ErrorAction SilentlyContinue
    if ($props) {
        $props.psobject.Properties | Where-Object {
            $_.Name -notlike 'PS*' -and $_.Name -like '*ighlighter*'
        } | ForEach-Object {
            Remove-ItemProperty -Path $ck -Name $_.Name -ErrorAction SilentlyContinue
            Write-Host "  Removed value: $($_.Name)"
            $removed = $true
        }
    }
    if (-not $removed) { Write-Host "  no entry for this add-in (other add-ins' entries left untouched)" }
} else {
    Write-Host "  not present"
}

# 4. Build a MINIMAL ribbon-only test add-in and drop it into XLSTART
Write-Host "`n=== 4. Installing minimal ribbon test add-in into XLSTART ==="
$xlstart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
if (-not (Test-Path $xlstart)) { New-Item -ItemType Directory -Path $xlstart -Force | Out-Null }

$minimal = @'
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

$ct = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.addin.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.customUI+xml"/>
</Types>
'@

$rels2 = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/07/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>
'@

$wbxml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'@

$wbRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@

$sheet = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>
'@

$minOut = Join-Path $xlstart "minimal-ribbon-test.xlam"
if (Test-Path $minOut) { Remove-Item $minOut -Force }
$zip = [System.IO.Compression.ZipFile]::Open($minOut, [System.IO.Compression.ZipArchiveMode]::Create)
function Add-TextEntry($z, $path, $content) {
    $e = $z.CreateEntry($path)
    $sw = New-Object System.IO.StreamWriter($e.Open())
    $sw.Write($content); $sw.Close()
}
Add-TextEntry $zip "[Content_Types].xml" $ct
Add-TextEntry $zip "_rels/.rels" $rels2
Add-TextEntry $zip "xl/workbook.xml" $wbxml
Add-TextEntry $zip "xl/_rels/workbook.xml.rels" $wbRels
Add-TextEntry $zip "xl/worksheets/sheet1.xml" $sheet
Add-TextEntry $zip "customUI/customUI14.xml" $minimal
$zip.Dispose()
Write-Host "  Built: $minOut ($((Get-Item $minOut).Length) bytes)"

# Sanity check: re-open the zip and confirm the customUI entry is present with expected content
$check = [System.IO.Compression.ZipFile]::OpenRead($minOut)
$ce = $check.GetEntry('customUI/customUI14.xml')
if ($ce) {
    $cr = New-Object System.IO.StreamReader($ce.Open())
    $cContent = $cr.ReadToEnd(); $cr.Close()
    Write-Host "  Sanity: customUI14.xml present, contains tabMinimalTest2: $($cContent.Contains('tabMinimalTest2'))"
} else {
    Write-Host "  Sanity FAIL: customUI/customUI14.xml missing from built test add-in!" -ForegroundColor Red
}
$check.Dispose()

# 5. Report next steps
Write-Host ""
Write-Host "=========================================================="
Write-Host " DONE. Now restart Excel and check the ribbon tabs."
Write-Host "=========================================================="
Write-Host ""
Write-Host "You should now see TWO custom tabs:"
Write-Host "  1. 'Minimal Test'  - a static ribbon with one button"
Write-Host "                     (proves Excel can build custom ribbons at all)"
Write-Host "  2. 'Highlighter'   - the real add-in ribbon"
Write-Host ""
Write-Host "Tell me which of these appear after restart. That single answer"
Write-Host "tells us whether the problem is Excel-side or add-in-side."
Write-Host ""
Write-Host "To remove the test add-in later, just delete:"
Write-Host "  $minOut"
