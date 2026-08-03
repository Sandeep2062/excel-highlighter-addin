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

# Whether to inject the declarative customUI14.xml package part into the .xlam.
#
# DEFAULT: $true (part ON) - the customUI part is the standard, documented
# mechanism for Excel add-in ribbons, and the build emits the classic package
# WITH the part. ThisWorkbook also implements IRibbonExtensibility as
# belt-and-braces: when the part is present Excel uses it and never calls
# GetCustomUI (the part masks the interface), so the VBA interface is simply
# never invoked; on builds/versions that ignore the part, GetCustomUI is there
# as a fallback.
#
# WHY THE DEFAULT CHANGED (session findings, Excel 2024 build 16.0.20228,
# see CHANGELOG 2.1.3): earlier rounds blamed the digital signature (WRONG -
# an UNSIGNED copy of ASAP Utilities renders its tab from the same folder,
# proven by stripping its signature parts) and the customUI part (also not the
# gate - even a minimal static part is ignored on this build). The confirmed
# difference between this add-in and the one that renders (ASAP Utilities) is
# that ASAP's VBA project is PASSWORD-LOCKED (precompiled). Excel's ribbon
# engine discovers IRibbonExtensibility from the compiled type info; an
# unlocked source-only project is queried before lazy compilation exposes the
# interface, so Excel silently treats the add-in as having no ribbon. The part
# remains the default because it is the classic documented mechanism; the
# interface stays as a fallback. For the ribbon on Excel 2024 / recent 365,
# run scripts/lock-vba-project.ps1 once after installing.
#
# Set NO_CUSTOMUI_PART=1 to build the VBA-only package (no part) for
# diagnostics. Read from an env var rather than a [bool] CLI param: invoking
# through bash mangles $-prefixed values ($false expands to empty) and
# PowerShell -File mode raises ParameterArgumentTransformationError for some
# bool spellings - an env var is immune to all of that.
$WithCustomUIPart = -not ($env:NO_CUSTOMUI_PART -eq "1" -or $env:NO_CUSTOMUI_PART -ieq "true")

$ErrorActionPreference = "Stop"

if (-not $RepoRoot) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    if (-not $scriptDir) { $scriptDir = Get-Location }
    $RepoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
}

# Shared customUI part wiring helper (Test-XlamRibbonWiring / Repair-XlamRibbonWiring).
$utilsScript = Join-Path $RepoRoot "scripts\xlam-ribbon-utils.ps1"
if (Test-Path $utilsScript) { . $utilsScript }
else { throw "Shared helper not found: $utilsScript" }

$srcDir      = Join-Path $RepoRoot "src"
$customUiDir = Join-Path $RepoRoot "customUI"
$outputPath  = Join-Path $RepoRoot $OutputName

Write-Host "Building $OutputName from $srcDir ..." -ForegroundColor Cyan

# --- Step 0: static VBA source guard -------------------------------------------
# The VBE compiles lazily, so the Application.Run sweep in Step 3 only catches
# errors in code it actually executes. Class modules (ThisWorkbook, Sheet1,
# EventApp) are not Application.Run-callable, and a compile error in a sibling
# proc of a probed module can still ship. The exported sources are plain text,
# so scan them for the two classes of construct VBA rejects at compile time but
# that a save never surfaces: (1) module-level Const/Declare/Type/Enum placed
# AFTER the first procedure (only legal in the declarations section - the exact
# 2.3.0 mid-module-Const bug), and (2) function calls in Const initializers
# (VBA requires a constant expression - Chr$(30) raises "Constant expression
# required"). Belt-and-braces that fails the build before Excel even opens.
function Test-VbaStaticSource {
    param([string]$Path)
    $problems = @()
    $lines = Get-Content -Path $Path
    $seenProc = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $trimmed = $lines[$i].TrimStart()
        $indent  = $lines[$i].Length - $trimmed.Length
        # Module-level statements are unindented (column 0).
        if ($indent -eq 0 -and $trimmed) {
            if ($trimmed -match '^(Public|Private|Friend)?\s*(Sub|Function|Property)\s') { $seenProc = $true; continue }
            if ($seenProc -and $trimmed -match '^(Public|Private)?\s*(Const|Declare|Type|Enum)\b') {
                $problems += "line $($i + 1): module-level '$trimmed' appears after the first procedure - Const/Declare/Type/Enum are only legal in the declarations section"
            }
        }
        # Function calls are illegal in Const initializers (constant expression
        # required). Strip trailing comments first so a paren inside a comment
        # can't false-positive.
        $codeOnly = $trimmed -replace ".*$", ''
        if ($codeOnly -match '^(Public|Private)?\s*Const\s+[A-Za-z_][A-Za-z0-9_]*\s*=.*\b[A-Za-z_][A-Za-z0-9_]*\$?\s*\(') {
            $problems += "line $($i + 1): Const initializer calls a function (illegal in VBA - constant expression required): $trimmed"
        }
    }
    return $problems
}

$vbaSources = @(Get-ChildItem -Path $srcDir -File | Where-Object { $_.Extension -in '.bas', '.cls' })
$staticIssues = @()
foreach ($vbaFile in $vbaSources) {
    foreach ($p in (Test-VbaStaticSource $vbaFile.FullName)) { $staticIssues += "$($vbaFile.Name): $p" }
}
if ($staticIssues.Count -gt 0) {
    foreach ($issue in $staticIssues) { Write-Host "  STATIC: $issue" -ForegroundColor Yellow }
    throw "BUILD FAILED: static VBA source guard found $($staticIssues.Count) issue(s) - the project would fail to compile at load time. Fix the source and rebuild."
}
Write-Host "  Static source guard passed ($($vbaSources.Count) modules scanned)" -ForegroundColor Green

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

    # Ensure the Microsoft Office object library is referenced. ThisWorkbook
    # implements IRibbonExtensibility (the interface Excel's ribbon engine uses
    # to get the ribbon XML for .xlam VBA add-ins), and that interface is defined
    # in the Office object library - without the reference the project fails to
    # compile and the add-in silently breaks at load (the exact failure class this
    # whole investigation has been fighting). Fresh workbooks normally include it,
    # but be explicit so the build can't quietly produce a broken add-in on any
    # Office build/version.
    # NOTE: Reference.Guid is returned WITH curly braces, so the constant must
    # be braced too or the has-check below never matches and AddFromGuid throws
    # "reference already exists" on every build (harmless but noisy, and it hid
    # whether the reference actually existed). The braced form also works as the
    # argument to AddFromGuid.
    $msoGuid = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
    $hasMso = $false
    foreach ($ref in $vbProject.References) {
        if ($ref.Guid -ieq $msoGuid) { $hasMso = $true; break }
    }
    if (-not $hasMso) {
        try {
            $vbProject.References.AddFromGuid($msoGuid, 1, 0) | Out-Null
            Write-Host "  Added Microsoft Office object library reference (IRibbonExtensibility)"
        } catch {
            Write-Host "  WARNING: could not add Office object library reference: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        # If the reference still isn't there the add-in will fail to compile at
        # load (Implements IRibbonExtensibility on ThisWorkbook needs it). Fail
        # loudly rather than silently shipping a broken package - that is the
        # exact failure class this whole investigation has been fighting.
        foreach ($ref in $vbProject.References) {
            if ($ref.Guid -ieq $msoGuid) { $hasMso = $true; break }
        }
        if (-not $hasMso) {
            throw "Microsoft Office object library reference could not be added - ThisWorkbook implements IRibbonExtensibility and the build would not compile. Add it manually via Tools > References and re-run."
        }
    }

    # ThisWorkbook.cls content is merged into the existing ThisWorkbook module
    # rather than imported as a new one - Excel doesn't allow importing a
    # second ThisWorkbook. ThisWorkbook implements IRibbonExtensibility (see
    # its header for why the interface must live on the root IDispatch object
    # rather than a standalone class or the package part).
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

    # Give the VBA project a friendly name so the VBE shows
    # "Highlighter (excel-highlighter.xlam)" instead of the default
    # "VBAProject (excel-highlighter.xlam)". Safe to change: every runtime
    # entry point (ribbon callbacks, Application.Run, OnKey, OnTime, the
    # context-menu OnAction) references procedures by module-qualified name
    # only - no code anywhere prefixes a macro with the project name.
    try {
        $vbProject.Name = "Highlighter"
        Write-Host "  Renamed VBA project to 'Highlighter'"
    } catch {
        Write-Host "  WARNING: could not rename VBA project: $($_.Exception.Message)" -ForegroundColor Yellow
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

    # Hoisted here so step 2d (docProps) works on every path, including when
    # the customUI part is skipped entirely (-WithCustomUIPart:$false).
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    # 2a. Copy customUI14.xml into the package.
    #
    # The customUI part is the standard documented ribbon mechanism for Excel
    # add-ins and is the DEFAULT build. (Session finding: a part present with
    # correct relationships is processed by Excel - the earlier "silently
    # ignored" conclusion was drawn from invalid start-screen harness runs and
    # from an unsigned build on Excel 2024 build 16.0.20228, where ALL customUI
    # delivery is suppressed for unsigned add-ins. The gate is the digital
    # signature / Trusted Publisher, not the part.) The part masks
    # IRibbonExtensibility.GetCustomUI on ThisWorkbook (Excel uses the part and
    # never calls GetCustomUI), which is fine - the interface stays in the
    # project purely as a fallback for builds that do not honour the part.
    # Set NO_CUSTOMUI_PART=1 to skip the part (diagnostic VBA-only build).
    # customUI/customUI14.xml is the single source of truth either way.
    if ($WithCustomUIPart) {
        $customUi14Path = Join-Path $customUiDir "customUI14.xml"
        $pkgCustomUiDir = Join-Path $extractDir "customUI"
        if (-not (Test-Path $pkgCustomUiDir)) {
            New-Item -ItemType Directory -Path $pkgCustomUiDir -Force | Out-Null
        }

        if (Test-Path $customUi14Path) {
            Copy-Item $customUi14Path (Join-Path $pkgCustomUiDir "customUI14.xml") -Force
            Write-Host "  Added customUI/customUI14.xml"
        }

# 2b. Reference customUI14.xml. We declare the ui/extensibility
# relationship in BOTH locations - on the workbook part
# (xl/_rels/workbook.xml.rels, where the ribbon engine discovers the
# part for Excel documents including add-ins) and at the package root
# (_rels/.rels, the convention some tooling/documentation uses) - as
# belt-and-braces. Earlier session observations that the part was
# "ignored" were made on UNSIGNED builds of Excel 2024 build
# 16.0.20228, where all customUI delivery is suppressed regardless of
# relationship location, so no clean conclusion about the location was
# ever drawn; declaring both is the safe choice.
        $relType14 = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

        # NOTE: the relationship Target is relative to the SOURCE part. From
        # xl/_rels/workbook.xml.rels the target is ../customUI/customUI14.xml
        # (relative to xl/); from the package root _rels/.rels it must be
        # customUI/customUI14.xml (no leading ../ - that would resolve above
        # the package root and be invalid per OPC).
        $wbRel14 = '<Relationship Id="rIdCustomUI14" Type="' + $relType14 + '" Target="../customUI/customUI14.xml"/>'
        $rootRel14 = '<Relationship Id="rIdCustomUI14" Type="' + $relType14 + '" Target="customUI/customUI14.xml"/>'

        # Workbook-part relationship (the location the ribbon engine follows for
        # Excel documents - verified this session to be the only sensible place
        # after the package-root-only variant was ignored).
        $workbookRelsPath = Join-Path $extractDir "xl\_rels\workbook.xml.rels"
        if (Test-Path $workbookRelsPath) {
            $wbRels = [System.IO.File]::ReadAllText($workbookRelsPath)
            $wbRels = $wbRels -replace '(?i)<Relationship[^>]+ui/extensibility[^>]+/>', ''
            $wbRels = $wbRels.Replace('</Relationships>', "$wbRel14</Relationships>")
            [System.IO.File]::WriteAllText($workbookRelsPath, $wbRels, $utf8NoBom)
            Write-Host "  Updated xl/_rels/workbook.xml.rels (ui/extensibility -> customUI14.xml)"
        } else {
            Write-Warning "  xl/_rels/workbook.xml.rels not found - customUI part will not be discoverable!"
        }

        # Package-root relationship too, with the CORRECT root-relative target
        # (harmless belt-and-braces; some Excel versions/docs look here).
        $relsPath = Join-Path $extractDir "_rels\.rels"
        if (Test-Path $relsPath) {
            $relsContent = [System.IO.File]::ReadAllText($relsPath)
            $relsContent = $relsContent -replace '(?i)<Relationship[^>]+ui/extensibility[^>]+/>', ''
            $relsContent = $relsContent.Replace('</Relationships>', "$rootRel14</Relationships>")
            [System.IO.File]::WriteAllText($relsPath, $relsContent, $utf8NoBom)
            Write-Host "  Updated _rels/.rels (also added package-root reference)"
        }

        # 2c. Update [Content_Types].xml to declare customUI14.xml.
        $contentTypesPath = Join-Path $extractDir "[Content_Types].xml"
        $contentTypes = [System.IO.File]::ReadAllText($contentTypesPath)

        $contentTypes = $contentTypes -replace '(?i)<Override[^>]+customUI[^>]+/>', ''

        $addedOverride = '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.customUI+xml"/>'

        $contentTypes = $contentTypes.Replace('</Types>', "$addedOverride</Types>")

        [System.IO.File]::WriteAllText($contentTypesPath, $contentTypes, $utf8NoBom)
        Write-Host "  Updated [Content_Types].xml"
    } else {
        Write-Host "  Skipping customUI package part (NO_CUSTOMUI_PART=1) - ribbon supplied via IRibbonExtensibility.GetCustomUI only"
    }

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

# --- Step 2f: assert the ribbon wiring is complete ----------------------------
# Excel only processes the customUI part when the ui/extensibility relationship
# exists in xl/_rels/workbook.xml.rels (the build writes it, but assert so a
# future change can never silently ship a present-but-unwired part - which
# masks IRibbonExtensibility.GetCustomUI while never being processed, i.e. an
# add-in whose ribbon silently never loads).
if ($WithCustomUIPart) {
    $wiring = Test-XlamRibbonWiring $outputPath
    if (-not $wiring.Ok) {
        throw "BUILD FAILED: ribbon wiring incomplete - Part=$($wiring.Part) ContentType=$($wiring.ContentType) RootRel=$($wiring.RootRel) WorkbookRel=$($wiring.WorkbookRel). The add-in would ship with a present-but-unwired customUI part (masks GetCustomUI, ribbon never loads)."
    }
    Write-Host "  Ribbon wiring verified: part + content-type + root/workbook relationships present" -ForegroundColor Green
}

# --- Step 3: compile validation -------------------------------------------------
# The VBA project is never compiled by the build (VBE compiles lazily), so a
# project with a compile error - e.g. a wrong Implements signature, a broken
# reference, a typo - would save fine and only fail at load time, showing as
# "installed but dead" with no ribbon tab and no error. Exactly the failure
# class this whole investigation has been fighting. Open the built .xlam in a
# fresh Excel instance and execute a known macro: if the project does not
# compile, Application.Run throws "Cannot run the macro... macro may not be
# available" and we fail the build here instead of shipping a dead add-in.
Write-Host "Validating built add-in compiles (direct open + Application.Run)..." -ForegroundColor Cyan

$checkExcel = New-Object -ComObject Excel.Application
$checkExcel.Visible = $false
$checkExcel.DisplayAlerts = $false
$checkExcel.EnableEvents = $false
$checkExcel.AutomationSecurity = 1 # 1 = msoAutomationSecurityLow

try {
    # If a previously-installed copy of this add-in is loaded via the OPEN
    # registry key in this session, close it first so Application.Run below
    # unambiguously resolves to the workbook we just built (module-qualified
    # macro names are project-scoped; two AddinHost modules would be ambiguous).
    # Defensive: an add-in caught mid-load can throw "object invalid or no
    # longer set" - that must not fail the build for an unrelated reason (the
    # fresh workbook we open next becomes active anyway, which is what
    # Application.Run resolves first).
    try {
        foreach ($wb in $checkExcel.Workbooks) {
            if ($wb.IsAddin -and $wb.Name -like "excel-highlighter*") {
                $wb.Close($false)
            }
        }
    } catch {
        Write-Host "  Note: could not close previously-installed copy during compile check: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $checkWb = $checkExcel.Workbooks.Open($outputPath)
    try {
        # Sweep a representative entry point from EVERY module. The old check
        # only ran AddinHost.IsRunning, which never exercised the rest of the
        # module - a compile error anywhere else (e.g. the mid-module Const in
        # AddinHost that broke every proc after it in 2.3.0) shipped silently
        # and only surfaced as "installed but dead" at load time. VBA compiles
        # a procedure plus its callees lazily, so running these probes also
        # forces compilation of the modules they depend on (RibbonCallbacks,
        # HighlightEngine, Utilities, ...). Any failure fails the build.
        $compileProbes = @(
            @{ Name = "AddinHost.IsRunning" },
            @{ Name = "AddinHost.AddinFolder" },
            @{ Name = "AddinHost.AddContextMenu" },      # exercises CommandBar + Office constants
            @{ Name = "AddinHost.RemoveContextMenu" },
            @{ Name = "HighlightEngine.ActiveWorkbookHighlightState" },
            @{ Name = "Settings.EffectiveRGB" },
            @{ Name = "Logging.LogPath" },
            @{ Name = "SelectionHistory.CanGoBack" },
            @{ Name = "Profiles.ProfileCount" },
            @{ Name = "ColourPicker.ClearSwatchCache" }
        )
        foreach ($probe in $compileProbes) {
            try {
                $probeResult = $checkExcel.Run($probe.Name)
                Write-Host "  Compile probe OK: $($probe.Name) -> $probeResult"
            } catch {
                throw "BUILD FAILED: compile probe '$($probe.Name)' threw '$($_.Exception.Message)'. The project would show as installed but dead with no ribbon tab. Fix the VBA source and rebuild."
            }
        }
        Write-Host "  Compile check passed ($($compileProbes.Count) probes across all modules)" -ForegroundColor Green
    } catch {
        throw "BUILD FAILED: the built add-in does not compile - Application.Run threw '$($_.Exception.Message)'. The project would show as installed but dead with no ribbon tab. Fix the VBA source and rebuild."
    } finally {
        $checkWb.Close($false)
    }
}
finally {
    try { $checkExcel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($checkExcel) | Out-Null
}

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
