<#
.SYNOPSIS
    Builds excel-highlighter.xlam from the text-exported VBA
    modules under src/ and the RibbonX definition under customUI/.

.DESCRIPTION
    Drives Excel via COM automation. Requires:
      - Excel installed on the machine running this script
      - "Trust access to the VBA project object model" enabled
        (File > Options > Trust Center > Trust Center Settings > Macro Settings)

    The script saves the workbook as .xlam first, then rebuilds the
    OPC package from scratch by extracting to a temp directory, adding
    the customUI parts, and re-zipping. This avoids the stale-entry
    problem that can occur when modifying a zip in-place.

    IMPORTANT: ZipFile.CreateFromDirectory uses backslashes in entry
    paths on Windows, but OPC packages require forward slashes. This
    script manually creates the zip with forward-slash paths to ensure
    Excel can find the customUI parts.

.NOTES
    Idempotent: re-running overwrites the previous build output.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$OutputName = "excel-highlighter.xlam"
)

$ErrorActionPreference = "Stop"

if (-not $RepoRoot) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    if (-not $scriptDir) { $scriptDir = Get-Location }
    $RepoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
}

$srcDir      = Join-Path $RepoRoot "src"
$customUiDir = Join-Path $RepoRoot "customUI"
$outputPath  = Join-Path $RepoRoot $OutputName

Write-Host "Building $OutputName from $srcDir ..." -ForegroundColor Cyan

# --- Step 1: build the raw .xlam with all VBA modules imported ----------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 1 # 1 = msoAutomationSecurityLow

$tempXlam = $null

try {
    $workbook = $excel.Workbooks.Add()

    # Remove default extra sheets, leave exactly one - keeps the add-in lean.
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    try {
        $vbProject = $workbook.VBProject
    } catch {
        $vbProject = $null
    }
    if ($null -eq $vbProject -or $null -eq $vbProject.VBComponents) {
        throw "Access to the Excel VBA Project object model is blocked or disabled. Please enable 'Trust access to the VBA project object model' in Excel (File > Options > Trust Center > Trust Center Settings > Macro Settings)."
    }

    # ThisWorkbook.cls content is merged into the existing ThisWorkbook module
    # rather than imported as a new one - Excel doesn't allow importing a
    # second ThisWorkbook.
    $thisWorkbookSrc = Get-Content (Join-Path $srcDir "ThisWorkbook.cls") -Raw
    $thisWorkbookBody = ($thisWorkbookSrc -split "(?ms)^Attribute VB_Exposed.*?\r?\n", 2)[1]
    $vbProject.VBComponents.Item("ThisWorkbook").CodeModule.AddFromString($thisWorkbookBody)

    # NOTE: src/ is checked out of git with LF-only line endings. That's fine for
    # AddFromString above (which normalises internally) and mostly fine for plain
    # .bas code, but VBComponents.Import() reading a class module's raw
    # VERSION/BEGIN/MultiUse/END header straight off disk needs real CRLF line
    # breaks in that header block, or the project loader can mis-parse it and
    # throw "Compile error: Expected: end of statement" the first time the
    # project tries to compile. Normalise to CRLF in a throwaway temp copy
    # before importing anything from disk, so this can't bite us again
    # regardless of how the working copy's line endings were checked out.
    function Import-ModuleWithCrlf {
        param([string]$SourcePath, [string]$FileName)
        $raw = Get-Content -Path $SourcePath -Raw
        $crlf = [System.Text.RegularExpressions.Regex]::Replace($raw, "\r\n|\r|\n", "`r`n")
        $tempPath = Join-Path $env:TEMP $FileName
        [System.IO.File]::WriteAllText($tempPath, $crlf, (New-Object System.Text.UTF8Encoding($false)))
        $vbProject.VBComponents.Import($tempPath) | Out-Null
        Remove-Item $tempPath -Force
    }

    # EventApp.cls is a genuinely new class module.
    Import-ModuleWithCrlf -SourcePath (Join-Path $srcDir "EventApp.cls") -FileName "EventApp.cls"

    # Standard modules import directly.
    $standardModules = @(
        "Constants.bas",
        "Logging.bas",
        "Utilities.bas",
        "Settings.bas",
        "HighlightEngine.bas",
        "AddinHost.bas",
        "RibbonCallbacks.bas",
        "ColourPicker.bas",
        "SelectionHistory.bas",
        "Profiles.bas"
    )

    foreach ($moduleFile in $standardModules) {
        Write-Host "  Importing $moduleFile"
        Import-ModuleWithCrlf -SourcePath (Join-Path $srcDir $moduleFile) -FileName $moduleFile
    }

    $tempXlam = Join-Path $env:TEMP ("excel_build_" + [Guid]::NewGuid().ToString("N") + ".xlam")
    if (Test-Path $tempXlam) { Remove-Item $tempXlam -Force }

    $workbook.IsAddin = $true
    $workbook.RemovePersonalInformation = $false

    try {
        # Title is what shows as the add-in's display name in File > Options > Add-ins.
        # Author/Company are what Excel reads for the Publisher column there - set both
        # so it shows up regardless of which one Excel's dialog happens to prefer.
        $workbook.BuiltinDocumentProperties.Item("Author").Value = "Sandeep Khadka"
        $workbook.BuiltinDocumentProperties.Item("Title").Value = "Excel Highlighter"
        $workbook.BuiltinDocumentProperties.Item("Comments").Value = "Non-destructive row, column and crosshair cell highlighter for Microsoft Excel."
        $workbook.BuiltinDocumentProperties.Item("Company").Value = "Sandeep Khadka"
    } catch {}

    $workbook.SaveAs($tempXlam, 55)   # 55 = xlOpenXMLAddIn (.xlam)
    $workbook.Close($false)
}
finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    # Give Excel a moment to fully release file locks.
    Start-Sleep -Milliseconds 500
}

Write-Host "Base .xlam written to $tempXlam" -ForegroundColor Green

# --- Step 2: rebuild the OPC package with customUI injected -------------------
Write-Host "Injecting customUI XML files into OPC package..." -ForegroundColor Cyan

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Instead of modifying the zip in-place (which can leave orphaned central
# directory entries when deleting and re-creating files), we extract
# everything to a temp directory, add/modify the customUI parts, and
# create a fresh zip. This is the most reliable approach.
$extractDir = Join-Path $env:TEMP ("excel_build_extract_" + [Guid]::NewGuid().ToString("N"))
if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

try {
    # Extract the original xlam to the temp directory.
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tempXlam, $extractDir)

    # 2a. Copy customUI XML files into the package.
    $customUi14Path = Join-Path $customUiDir "customUI14.xml"
    $customUi07Path = Join-Path $customUiDir "customUI.xml"

    $pkgCustomUiDir = Join-Path $extractDir "customUI"
    if (-not (Test-Path $pkgCustomUiDir)) {
        New-Item -ItemType Directory -Path $pkgCustomUiDir -Force | Out-Null
    }

    if (Test-Path $customUi14Path) {
        Copy-Item $customUi14Path (Join-Path $pkgCustomUiDir "customUI14.xml") -Force
        Write-Host "  Added customUI/customUI14.xml"
    }
    if (Test-Path $customUi07Path) {
        Copy-Item $customUi07Path (Join-Path $pkgCustomUiDir "customUI.xml") -Force
        Write-Host "  Added customUI/customUI.xml"
    }

    # 2b. Update _rels/.rels to reference customUI14.xml.
    $relsPath = Join-Path $extractDir "_rels\.rels"
    $relsContent = [System.IO.File]::ReadAllText($relsPath)

    $relsContent = $relsContent -replace '(?i)<Relationship[^>]+ui/extensibility[^>]+/>', ''

    $relType14 = "http://schemas.microsoft.com/office/2007/07/relationships/ui/extensibility"
    $addedRel = '<Relationship Id="rIdCustomUI14" Type="' + $relType14 + '" Target="customUI/customUI14.xml"/>'
    $relsContent = $relsContent.Replace('</Relationships>', "$addedRel</Relationships>")

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($relsPath, $relsContent, $utf8NoBom)
    Write-Host "  Updated _rels/.rels"

    # 2c. Update [Content_Types].xml to declare customUI14.xml.
    $contentTypesPath = Join-Path $extractDir "[Content_Types].xml"
    $contentTypes = [System.IO.File]::ReadAllText($contentTypesPath)

    $contentTypes = $contentTypes -replace '(?i)<Override[^>]+customUI[^>]+/>', ''

    $addedOverride = '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.customUI+xml"/>'
    $contentTypes = $contentTypes.Replace('</Types>', "$addedOverride</Types>")

    [System.IO.File]::WriteAllText($contentTypesPath, $contentTypes, $utf8NoBom)
    Write-Host "  Updated [Content_Types].xml"

    # 2d. Update docProps/core.xml for Publisher and Description in Excel Add-ins dialog
    $corePropsPath = Join-Path $extractDir "docProps\core.xml"
    if (Test-Path $corePropsPath) {
        $coreXml = [System.IO.File]::ReadAllText($corePropsPath)
        if (-not $coreXml.Contains("<dc:creator>")) {
            $coreXml = $coreXml.Replace('</cp:coreProperties>', '<dc:creator>Sandeep Khadka</dc:creator><dc:title>Excel Highlighter</dc:title><dc:description>Non-destructive row, column and crosshair cell highlighter for Microsoft Excel.</dc:description></cp:coreProperties>')
        } else {
            $coreXml = $coreXml -replace '<dc:creator>[^<]*</dc:creator>', '<dc:creator>Sandeep Khadka</dc:creator>'
        }
        if (-not $coreXml.Contains("<dc:title>")) {
            $coreXml = $coreXml.Replace('</cp:coreProperties>', '<dc:title>Excel Highlighter</dc:title><dc:description>Non-destructive row, column and crosshair cell highlighter for Microsoft Excel.</dc:description></cp:coreProperties>')
        } else {
            $coreXml = $coreXml -replace '<dc:title>[^<]*</dc:title>', '<dc:title>Excel Highlighter</dc:title>'
            $coreXml = $coreXml -replace '<dc:description>[^<]*</dc:description>', '<dc:description>Non-destructive row, column and crosshair cell highlighter for Microsoft Excel.</dc:description>'
        }
        [System.IO.File]::WriteAllText($corePropsPath, $coreXml, $utf8NoBom)
        Write-Host "  Updated docProps/core.xml"
    }

    # 2e. Create a fresh zip with forward-slash paths (required by OPC).
    $freshXlam = Join-Path $env:TEMP ("excel_build_fresh_" + [Guid]::NewGuid().ToString("N") + ".xlam")
    if (Test-Path $freshXlam) { Remove-Item $freshXlam -Force }

    # We must manually create the zip to ensure forward-slash entry paths.
    # ZipFile.CreateFromDirectory uses backslashes on Windows, which breaks
    # OPC package parsing in Excel.
    $zipStream = [System.IO.File]::Open($freshXlam, [System.IO.FileMode]::CreateNew)
    $zipArchive = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)

    try {
        # Get all files recursively from the extract directory.
        $allFiles = [System.IO.Directory]::GetFiles($extractDir, '*', [System.IO.SearchOption]::AllDirectories)
        $basePath = $extractDir.TrimEnd('\') + '\'

        foreach ($filePath in $allFiles) {
            # Get the relative path and convert backslashes to forward slashes.
            $relativePath = $filePath.Substring($basePath.Length).Replace('\', '/')
            $entry = $zipArchive.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object System.IO.BinaryWriter($entry.Open())
            $writer.Write([System.IO.File]::ReadAllBytes($filePath))
            $writer.Close()
        }
    } finally {
        $zipArchive.Dispose()
        $zipStream.Close()
    }

    Write-Host "  Fresh OPC package created with forward-slash paths" -ForegroundColor Green
}
finally {
    # Clean up temp files.
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue }
    if ($tempXlam -and (Test-Path $tempXlam)) { Remove-Item $tempXlam -Force -ErrorAction SilentlyContinue }
}

# Copy the fresh xlam to the output path.
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}
Copy-Item $freshXlam $outputPath -Force
Remove-Item $freshXlam -Force

Write-Host "Build successful: $outputPath" -ForegroundColor Green

# --- Step 3: copy the images folder alongside the .xlam -----------------------
$imagesTarget = Join-Path $RepoRoot "images"
if (-not (Test-Path $imagesTarget)) {
    New-Item -ItemType Directory -Path $imagesTarget -Force | Out-Null
}
Copy-Item (Join-Path $customUiDir "images\*") $imagesTarget -Force

Write-Host ""
Write-Host "Done. $OutputName and its images/ folder are ready in $RepoRoot" -ForegroundColor Green
Write-Host "Copy both to your add-ins location and enable the add-in in Excel Options." -ForegroundColor Green
