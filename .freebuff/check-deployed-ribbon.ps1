$ErrorActionPreference = 'Continue'

$f = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
$copy = Join-Path $env:TEMP "deployed-check2.xlam"
Copy-Item $f $copy -Force -ErrorAction Stop

Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead($copy)

function Read-Entry($zip, $name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open())
    $txt = $sr.ReadToEnd(); $sr.Close()
    return $txt
}

$ui14 = Read-Entry $z "customUI/customUI14.xml"
if ($ui14 -eq $null) {
    Write-Host "customUI/customUI14.xml is MISSING!"
} else {
    Write-Host "customUI/customUI14.xml size: $($ui14.Length) chars"
    try {
        $xml = [xml]$ui14
        Write-Host "  XML parses OK"
        Write-Host "  onLoad = $($xml.customUI.onLoad)"
        Write-Host "  tabs = $($xml.customUI.ribbon.tabs.tab.id -join ', ')"
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('a','http://schemas.microsoft.com/office/2009/07/customui')
        $tabs = $xml.SelectNodes('//a:tab',$ns)
        Write-Host "  tab count via namespace: $($tabs.Count)"
        $controls = $xml.SelectNodes('//*[@onAction]',$ns)
        Write-Host "  onAction callbacks referenced: $($controls.Count)"
        $getFns = $xml.SelectNodes('//*[@getLabel or @getPressed or @getEnabled or @getImage or @getContent]',$ns)
        Write-Host "  get* callbacks referenced: $($getFns.Count)"
    } catch {
        Write-Host "  XML INVALID: $($_.Exception.Message)"
        Write-Host "  --- first 500 chars ---"
        Write-Host $ui14.Substring(0, [Math]::Min(500, $ui14.Length))
    }
}

Write-Host "`n=== workbook.xml.rels (add-in relation from workbook) ==="
$wbRels = Read-Entry $z "xl/_rels/workbook.xml.rels"
if ($wbRels) { Write-Host $wbRels } else { Write-Host "  (none)" }

Write-Host "`n=== vbaProject.bin.rels ==="
$vpRels = Read-Entry $z "xl/_rels/vbaProject.bin.rels"
if ($vpRels) { Write-Host $vpRels } else { Write-Host "  (none)" }

$z.Dispose()
Remove-Item $copy -Force -ErrorAction SilentlyContinue
