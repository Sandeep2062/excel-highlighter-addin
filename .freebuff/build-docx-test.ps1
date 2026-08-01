$ErrorActionPreference = 'Continue'

# Guard: an orphaned/running Word instance can re-lock the file or hang COM.
$wordProc = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($wordProc) {
    Write-Host "WARNING: Word is currently running (PID(s): $($wordProc.Id -join ', '))." -ForegroundColor Yellow
    Write-Host "If a previous run of this script left it orphaned, it is safe to close it." -ForegroundColor Yellow
    Write-Host "Closing it now to avoid COM hangs..." -ForegroundColor Yellow
    Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$outDir = Join-Path $env:TEMP "ribbon-file-tests"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$docxPath = Join-Path $outDir "ribbon-test.docx"
if (Test-Path $docxPath) { Remove-Item $docxPath -Force }

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

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0          # wdAlertsNone
    $word.ScreenUpdating = $false
    $word.AutomationSecurity = 1     # msoAutomationSecurityLow
    $doc = $word.Documents.Add()
    $doc.SaveAs2($docxPath, 12)      # wdFormatXMLDocument
    $doc.Close(0)                    # wdDoNotSaveChanges
    $doc = $null
    Write-Host "Base docx created: $docxPath ($((Get-Item $docxPath).Length) bytes)"
} catch {
    Write-Host "ERROR creating base: $($_.Exception.Message)"
    exit 1
} finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($word) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
    }
}
Start-Sleep -Seconds 2

# Inject customUI14.xml
$zip = [System.IO.Compression.ZipFile]::Open($docxPath, [System.IO.Compression.ZipArchiveMode]::Update)
$ne = $zip.CreateEntry('customUI/customUI14.xml')
$sw = New-Object System.IO.StreamWriter($ne.Open())
$sw.Write($minimalXml); $sw.Close()

$rels = $zip.GetEntry('_rels/.rels')
$sr = New-Object System.IO.StreamReader($rels.Open())
$relsText = $sr.ReadToEnd(); $sr.Close()
if (-not $relsText.Contains('ui/extensibility')) {
    $newRel = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/07/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    $relsText = $relsText.Replace('</Relationships>', $newRel + '</Relationships>')
    $rels.Delete()
    $ne2 = $zip.CreateEntry('_rels/.rels')
    $sw2 = New-Object System.IO.StreamWriter($ne2.Open())
    $sw2.Write($relsText); $sw2.Close()
}

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
Write-Host "customUI injected. New size: $((Get-Item $docxPath).Length) bytes"

# Verify
$z = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
$ui = $z.GetEntry('customUI/customUI14.xml')
$ok = $ui -ne $null
$z.Dispose()
Write-Host "Verify customUI14.xml present: $ok"
if (-not $ok) { exit 1 }
Write-Host "DONE: $docxPath"
Write-Host "(If this script ever hangs, kill WINWORD and re-run - it is idempotent.)"
