$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$src = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$tmp = Join-Path $env:TEMP "ribbon-test"
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
New-Item -ItemType Directory -Path $tmp | Out-Null

# 1. Copy the deployed add-in
$out = Join-Path $tmp "test.xlam"
Copy-Item $src $out -Force
Write-Host "Test copy: $((Get-Item $out).Length) bytes"

# 2. Replace customUI/customUI14.xml IN PLACE inside the zip
$minimal = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonCallbacks.onLoad">
  <ribbon>
    <tabs>
      <tab id="tabMinimalTest" label="Minimal Test">
        <group id="grpMinimalTest" label="Test">
          <button id="btnMinimalTest" label="Hello" imageMso="HappyFace" onAction="RibbonCallbacks.OnAbout_Action"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@

$zip = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Update)
$entry = $zip.GetEntry('customUI/customUI14.xml')
if ($entry) {
    $entry.Delete()
} else {
    Write-Host "WARNING: customUI/customUI14.xml not found in zip!"
}
$newEntry = $zip.CreateEntry('customUI/customUI14.xml')
$sw = New-Object System.IO.StreamWriter($newEntry.Open())
$sw.Write($minimal)
$sw.Close()
$zip.Dispose()
Write-Host "customUI14.xml replaced in place. New size: $((Get-Item $out).Length) bytes"

# Sanity: verify
$check = [System.IO.Compression.ZipFile]::OpenRead($out)
$e2 = $check.GetEntry('customUI/customUI14.xml')
if ($e2) {
    $sr = New-Object System.IO.StreamReader($e2.Open())
    $c = $sr.ReadToEnd(); $sr.Close()
    Write-Host "Sanity: entry present, contains tabMinimalTest: $($c.Contains('tabMinimalTest'))"
} else {
    Write-Host "Sanity FAIL: entry missing"
}
$check.Dispose()

# 3. Note current log length
$log = Join-Path $env:APPDATA "ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log"
$before = if (Test-Path $log) { (Get-Content $log).Count } else { 0 }

# 4. Open in a FRESH VISIBLE Excel instance
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    Write-Host "Opening test add-in in visible Excel..."
    $wb = $excel.Workbooks.Open($out)
    Write-Host "  Opened. IsAddin=$($wb.IsAddin)"
    Start-Sleep -Seconds 8
    $wb.Close($false)
} catch {
    Write-Host "  ERROR: $($_.Exception.Message)"
} finally {
    if ($excel) { try { $excel.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {} }
}

# 5. Check the log
Start-Sleep -Seconds 1
Write-Host "`n=== New log entries after test ==="
if (Test-Path $log) {
    $all = Get-Content $log
    if ($all.Count -gt $before) {
        $all[($before)..($all.Count-1)] | ForEach-Object { Write-Host "  $_" }
    } else {
        Write-Host "  (no new entries)"
    }
    $hits = Select-String -Path $log -Pattern 'Ribbon UI attached'
    Write-Host "`n=== 'Ribbon UI attached' total occurrences ever: $($hits.Count) ==="
} else {
    Write-Host "  no log file"
}

# 6. CustomUIValidationCache after test
Write-Host "`n=== CustomUIValidationCache after test ==="
$ck = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $ck) {
    Get-ItemProperty $ck -ErrorAction SilentlyContinue | Select-Object * -ExcludeProperty PS* | Format-List | Out-String | Write-Host
} else { Write-Host "  (not present)" }

Write-Host "`n=== done ==="
