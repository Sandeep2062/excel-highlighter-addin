$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$outDir = Join-Path $env:TEMP "ribbon-file-tests"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$minimalXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabFileTest" label="File UI Test">
        <group id="grpFileTest" label="Static">
          <button id="btnFileTest" label="Static Button" imageMso="HappyFace"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@

function Add-CustomUI($filePath, $relEntryPath, $relType, $relId) {
    # Add customUI/customUI14.xml and wire it into _rels/.rels + [Content_Types].xml
    $zip = [System.IO.Compression.ZipFile]::Open($filePath, [System.IO.Compression.ZipArchiveMode]::Update)

    # 1. Add the customUI14.xml part
    $ne = $zip.CreateEntry('customUI/customUI14.xml')
    $sw = New-Object System.IO.StreamWriter($ne.Open())
    $sw.Write($minimalXml); $sw.Close()

    # 2. Update _rels/.rels
    $rels = $zip.GetEntry('_rels/.rels')
    $sr = New-Object System.IO.StreamReader($rels.Open())
    $relsText = $sr.ReadToEnd(); $sr.Close()
    if (-not $relsText.Contains('ui/extensibility')) {
        $newRel = '<Relationship Id="' + $relId + '" Type="' + $relType + '" Target="customUI/customUI14.xml"/>'
        $relsText = $relsText.Replace('</Relationships>', $newRel + '</Relationships>')
        $rels.Delete()
        $ne2 = $zip.CreateEntry('_rels/.rels')
        $sw2 = New-Object System.IO.StreamWriter($ne2.Open())
        $sw2.Write($relsText); $sw2.Close()
    }

    # 3. Update [Content_Types].xml
    $ct = $zip.GetEntry('[Content_Types].xml')
    $sr2 = New-Object System.IO.StreamReader($ct.Open())
    $ctText = $sr2.ReadToEnd(); $sr2.Close()
    if (-not $ctText.Contains('/customUI/customUI14.xml')) {
        $newOverride = '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.customUI+xml"/>'
        $ctText = $ctText.Replace('</Types>', $newOverride + '</Types>')
        $ct.Delete()
        $ne3 = $zip.CreateEntry('[Content_Types].xml')
        $sw3 = New-Object System.IO.StreamWriter($ne3.Open())
        $sw3.Write($ctText); $sw3.Close()
    }
    $zip.Dispose()
}

# --- Build the .xlsx via Excel COM (guarantees a valid base package) ---------
Write-Host "=== Building ribbon-test.xlsx ==="
$excel = $null
$xlsxPath = Join-Path $outDir "ribbon-test.xlsx"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
    $wb.SaveAs($xlsxPath, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    Write-Host "  Base xlsx created: $xlsxPath ($((Get-Item $xlsxPath).Length) bytes)"
    Add-CustomUI $xlsxPath '_rels/.rels' 'http://schemas.microsoft.com/office/2007/07/relationships/ui/extensibility' 'rIdCustomUI14'
    Write-Host "  customUI added. New size: $((Get-Item $xlsxPath).Length) bytes"
} catch {
    Write-Host "  ERROR: $($_.Exception.Message)"
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {} }
}

# --- Build the .docx via Word COM (guarantees a valid base package) ----------
Write-Host "`n=== Building ribbon-test.docx ==="
$word = $null
$docxPath = Join-Path $outDir "ribbon-test.docx"
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    if (Test-Path $docxPath) { Remove-Item $docxPath -Force }
    $doc.SaveAs2($docxPath, 12)  # 12 = wdFormatXMLDocument (.docx)
    $doc.Close($false)
    Write-Host "  Base docx created: $docxPath ($((Get-Item $docxPath).Length) bytes)"
    Add-CustomUI $docxPath '_rels/.rels' 'http://schemas.microsoft.com/office/2007/07/relationships/ui/extensibility' 'rIdCustomUI14'
    Write-Host "  customUI added. New size: $((Get-Item $docxPath).Length) bytes"
} catch {
    Write-Host "  ERROR: $($_.Exception.Message)"
} finally {
    if ($word) { try { $word.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {} }
}

# --- Verify both files can be read as OPC packages ----------------------------
Write-Host "`n=== Verification ==="
$allOk = $true
foreach ($f in @($xlsxPath, $docxPath)) {
    if (Test-Path $f) {
        try {
            $z = [System.IO.Compression.ZipFile]::OpenRead($f)
            $names = @($z.Entries | ForEach-Object { $_.FullName })
            $ui = $z.GetEntry('customUI/customUI14.xml')
            $ok = $ui -ne $null
            $z.Dispose()
            Write-Host "  $(Split-Path $f -Leaf): $($names.Count) entries, customUI14.xml=$ok"
            if (-not $ok) { $allOk = $false }
        } catch {
            Write-Host "  $(Split-Path $f -Leaf): READ ERROR $($_.Exception.Message)"
            $allOk = $false
        }
    } else {
        Write-Host "  $(Split-Path $f -Leaf): FILE NOT CREATED"
        $allOk = $false
    }
}
if (-not $allOk) {
    Write-Host "`nERROR: one or more test files failed to build with customUI injected." -ForegroundColor Red
    Write-Host "Do NOT use them for the test - re-run after closing Excel/Word if they were open." -ForegroundColor Red
    exit 1
}

Write-Host "`nDone. Files ready:"
Write-Host "  $xlsxPath"
Write-Host "  $docxPath"
