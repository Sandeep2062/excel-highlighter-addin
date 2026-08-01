param(
    [string]$DeployedPath,
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "=== Building Progressive Ribbon Tests ===" -ForegroundColor Cyan

# Read source XMLs
$source14 = Get-Content "D:\Github\excel-highlighter-addin\customUI\customUI14.xml" -Raw
$source07 = Get-Content "D:\Github\excel-highlighter-addin\customUI\customUI.xml" -Raw

# Strip XML comments and normalize whitespace for easier manipulation
function Strip-Comments($xml) {
    return $xml -replace '<!--[\s\S]*?-->', ''
}

$clean14 = Strip-Comments $source14
$clean07 = Strip-Comments $source07

# ---- Test 1: Full XML, no onLoad ----
# Already built as no-onload-test.xlam, skip

# ---- Test 2: Full XML, no onLoad, no dynamicMenu ----
Write-Host "`n--- Test 2: No onLoad, no dynamicMenu ---" -ForegroundColor Yellow
$test2xml = $clean14 -replace 'onLoad="RibbonCallbacks\.onLoad"', ''

# Remove the dynamicMenu block
$test2xml = $test2xml -replace '(?s)<dynamicMenu\s+id="ddlProfiles"[^/]*/>', ''
$test2xml = $test2xml -replace '\n\s*\n', "`n"

$test2path = Join-Path $OutputDir "test2-no-dynamic.xlam"
Copy-Item $DeployedPath $test2path -Force
$z = [System.IO.Compression.ZipFile]::Open($test2path, [System.IO.Compression.ZipArchiveMode]::Update)
$e = $z.GetEntry("customUI/customUI14.xml"); if ($e) { $e.Delete() }
$e = $z.GetEntry("customUI/customUI.xml"); if ($e) { $e.Delete() }
$n = $z.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w = New-Object System.IO.StreamWriter($n.Open()); $w.Write($test2xml); $w.Close()
$n2 = $z.CreateEntry("customUI/customUI.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w2 = New-Object System.IO.StreamWriter($n2.Open()); $w2.Write($test2xml); $w2.Close()
$z.Dispose()
Write-Host "  Built: $test2path ($( (Get-Item $test2path).Length) bytes)"

# ---- Test 3: Only the Controls group (4 toggle buttons + 1 button with callbacks) ----
Write-Host "`n--- Test 3: Controls group only ---" -ForegroundColor Yellow

$test3xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabHighlighter" label="Highlighter" insertAfterMso="TabHome">
        <group id="grpControls" label="Controls">
          <toggleButton id="btnToggleHighlight"
                        size="large"
                        imageMso="ConditionalFormattingMenu"
                        getLabel="RibbonCallbacks.GetToggle_Label"
                        screentip="Enable or disable the crosshair highlighter"
                        supertip="Turns highlighting on for every open workbook, or off entirely."
                        getPressed="RibbonCallbacks.GetToggle_Pressed"
                        onAction="RibbonCallbacks.OnToggle_Action"/>
          <toggleButton id="btnModeRow"
                        size="large"
                        label="Row"
                        imageMso="TableSelectRow"
                        screentip="Highlight active row"
                        supertip="Highlights the entire row containing the active cell."
                        getPressed="RibbonCallbacks.GetMode_Pressed"
                        getEnabled="RibbonCallbacks.GetMode_Enabled"
                        onAction="RibbonCallbacks.OnMode_Action"/>
          <toggleButton id="btnModeColumn"
                        size="large"
                        label="Column"
                        imageMso="TableSelectColumn"
                        screentip="Highlight active column"
                        supertip="Highlights the entire column containing the active cell."
                        getPressed="RibbonCallbacks.GetMode_Pressed"
                        getEnabled="RibbonCallbacks.GetMode_Enabled"
                        onAction="RibbonCallbacks.OnMode_Action"/>
          <toggleButton id="btnModeCrosshair"
                        size="large"
                        label="Crosshair"
                        imageMso="ViewGridlines"
                        screentip="Highlight row and column"
                        supertip="Highlights both the row and the column of the active cell."
                        getPressed="RibbonCallbacks.GetMode_Pressed"
                        getEnabled="RibbonCallbacks.GetMode_Enabled"
                        onAction="RibbonCallbacks.OnMode_Action"/>
          <toggleButton id="btnModeCell"
                        size="large"
                        label="Cell"
                        imageMso="FillColor"
                        screentip="Highlight active cell only"
                        supertip="Highlights only the active cell."
                        getPressed="RibbonCallbacks.GetMode_Pressed"
                        getEnabled="RibbonCallbacks.GetMode_Enabled"
                        onAction="RibbonCallbacks.OnMode_Action"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@

$test3path = Join-Path $OutputDir "test3-controls-only.xlam"
Copy-Item $DeployedPath $test3path -Force
$z = [System.IO.Compression.ZipFile]::Open($test3path, [System.IO.Compression.ZipArchiveMode]::Update)
$e = $z.GetEntry("customUI/customUI14.xml"); if ($e) { $e.Delete() }
$e = $z.GetEntry("customUI/customUI.xml"); if ($e) { $e.Delete() }
$n = $z.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w = New-Object System.IO.StreamWriter($n.Open()); $w.Write($test3xml); $w.Close()
$n2 = $z.CreateEntry("customUI/customUI.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$w2 = New-Object System.IO.StreamWriter($n2.Open()); $w2.Write($test3xml); $w2.Close()
$z.Dispose()
Write-Host "  Built: $test3path ($( (Get-Item $test3path).Length) bytes)"

# ---- Verify all 3 tests ----
Write-Host "`n=== Package verification ===" -ForegroundColor Cyan
foreach ($tp in @(
    (Join-Path $OutputDir "test2-no-dynamic.xlam"),
    (Join-Path $OutputDir "test3-controls-only.xlam")
)) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($tp)
    $ui = $z.GetEntry("customUI/customUI14.xml")
    $sr = New-Object System.IO.StreamReader($ui.Open())
    $c = $sr.ReadToEnd(); $sr.Close()
    $hasRel = $false
    $rels = $z.GetEntry("_rels/.rels")
    if ($rels) {
        $sr2 = New-Object System.IO.StreamReader($rels.Open())
        $rc = $sr2.ReadToEnd(); $sr2.Close()
        $hasRel = $rc -match 'extensibility'
    }
    $hasCT = $false
    $ct = $z.GetEntry("[Content_Types].xml")
    if ($ct) {
        $sr3 = New-Object System.IO.StreamReader($ct.Open())
        $cc = $sr3.ReadToEnd(); $sr3.Close()
        $hasCT = $cc -match 'customUI'
    }
    $z.Dispose()
    
    $tabCount = ([regex]::Matches($c, '<tab ')).Count
    $groupCount = ([regex]::Matches($c, '<group ')).Count
    $btnCount = ([regex]::Matches($c, '(?:toggleButton|button|gallery|dynamicMenu) ')).Count
    $cbCount = ([regex]::Matches($c, 'RibbonCallbacks\.\w+')).Count
    
    Write-Host "`n  $(Split-Path $tp -Leaf):" -ForegroundColor Yellow
    Write-Host "    Tabs: $tabCount, Groups: $groupCount, Controls: $btnCount, Callbacks: $cbCount"
    Write-Host "    Rels: $hasRel, ContentTypes: $hasCT"
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "  static-ribbon-test.xlam  -> 0 callbacks, no tab appeared? (re-test needed)"
Write-Host "  no-onload-test.xlam      -> 57 callbacks, onLoad removed"  
Write-Host "  test2-no-dynamic.xlam    -> ~50 callbacks, no onLoad, no dynamicMenu"
Write-Host "  test3-controls-only.xlam -> 15 callbacks, Controls group only"
Write-Host ""
Write-Host "Test in order: start with test3, then test2, then no-onload-test." -ForegroundColor Yellow
Write-Host "The FIRST one that shows the 'Highlighter' tab = that's the complexity threshold." -ForegroundColor Yellow
