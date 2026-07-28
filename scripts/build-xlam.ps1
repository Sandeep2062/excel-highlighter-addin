<#
.SYNOPSIS
    Builds excel-crosshair-highlighter.xlam from the text-exported VBA
    modules under src/ and the RibbonX definition under customUI/.

.DESCRIPTION
    Drives Excel via COM automation. Requires:
      - Excel installed on the machine running this script
      - "Trust access to the VBA project object model" enabled
        (File > Options > Trust Center > Trust Center Settings > Macro Settings)

    Because the VBA Extensibility object model can only import RibbonX
    through a saved file's OPC package (not directly via COM), this script
    saves the workbook as .xlam first, then post-processes the resulting
    zip package to inject customUI14.xml and its image relationships. This
    mirrors what a tool like Custom UI Editor for Microsoft Office does.

.NOTES
    Idempotent: re-running overwrites the previous build output.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")),
    [string]$OutputName = "excel-crosshair-highlighter.xlam"
)

$ErrorActionPreference = "Stop"

$srcDir      = Join-Path $RepoRoot "src"
$customUiDir = Join-Path $RepoRoot "customUI"
$outputPath  = Join-Path $RepoRoot $OutputName

Write-Host "Building $OutputName from $srcDir ..." -ForegroundColor Cyan

# --- Step 1: build the raw .xlam with all VBA modules imported ----------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()

    # Remove default extra sheets, leave exactly one - keeps the add-in lean.
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $vbProject = $workbook.VBProject

    # ThisWorkbook.cls content is merged into the existing ThisWorkbook module
    # rather than imported as a new one - Excel doesn't allow importing a
    # second ThisWorkbook.
    $thisWorkbookSrc = Get-Content (Join-Path $srcDir "ThisWorkbook.cls") -Raw
    $thisWorkbookBody = ($thisWorkbookSrc -split "(?ms)^Attribute VB_Exposed.*?\r?\n", 2)[1]
    $vbProject.VBComponents.Item("ThisWorkbook").CodeModule.AddFromString($thisWorkbookBody)

    # EventApp.cls is a genuinely new class module.
    $eventAppTemp = Join-Path $env:TEMP "EventApp.cls"
    Copy-Item (Join-Path $srcDir "EventApp.cls") $eventAppTemp -Force
    $vbProject.VBComponents.Import($eventAppTemp) | Out-Null
    Remove-Item $eventAppTemp -Force

    # Standard modules import directly.
    $standardModules = @(
        "Constants.bas",
        "Logging.bas",
        "Utilities.bas",
        "Settings.bas",
        "HighlightEngine.bas",
        "AddinHost.bas",
        "RibbonCallbacks.bas"
    )

    foreach ($moduleFile in $standardModules) {
        $path = Join-Path $srcDir $moduleFile
        Write-Host "  Importing $moduleFile"
        $vbProject.VBComponents.Import($path) | Out-Null
    }

    $workbook.IsAddin = $true
    $workbook.SaveAs($outputPath, 55)   # 55 = xlOpenXMLAddIn (.xlam)
    $workbook.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Base .xlam written to $outputPath" -ForegroundColor Green

# --- Step 2: inject the RibbonX customization into the OPC package -----------
Write-Host "Injecting customUI14.xml ..." -ForegroundColor Cyan

Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($outputPath, "Update")
try {
    # customXml/ part
    $customUiEntry = $zip.CreateEntry("customUI/customUI14.xml")
    $writer = New-Object System.IO.StreamWriter($customUiEntry.Open())
    $writer.Write((Get-Content (Join-Path $customUiDir "customUI14.xml") -Raw))
    $writer.Close()

    # _rels/.rels needs a relationship pointing at the customUI part.
    $relsEntry = $zip.GetEntry("_rels/.rels")
    $relsXml = [xml](New-Object System.IO.StreamReader($relsEntry.Open())).ReadToEnd()
    $ns = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

    $newRel = $relsXml.CreateElement("Relationship", $relsXml.DocumentElement.NamespaceURI)
    $newRel.SetAttribute("Id", "rIdCustomUI14")
    $newRel.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
    $newRel.SetAttribute("Target", "customUI/customUI14.xml")
    $relsXml.DocumentElement.AppendChild($newRel) | Out-Null

    $relsEntry.Delete()
    $newRelsEntry = $zip.CreateEntry("_rels/.rels")
    $writer = New-Object System.IO.StreamWriter($newRelsEntry.Open())
    $writer.Write($relsXml.OuterXml)
    $writer.Close()
}
finally {
    $zip.Dispose()
}

# --- Step 3: copy the images folder alongside the .xlam -----------------------
$imagesTarget = Join-Path $RepoRoot "images"
Copy-Item (Join-Path $customUiDir "images") $imagesTarget -Recurse -Force

Write-Host ""
Write-Host "Done. $OutputName and its images/ folder are ready in $RepoRoot" -ForegroundColor Green
Write-Host "Copy both to your add-ins location and enable the add-in in Excel Options." -ForegroundColor Green
