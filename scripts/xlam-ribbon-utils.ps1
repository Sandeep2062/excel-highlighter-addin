<#
.SYNOPSIS
    Shared helpers for verifying and repairing the customUI ribbon part wiring
    inside an .xlam OPC package.

.DESCRIPTION
    An Excel add-in's ribbon is delivered either by the declarative
    customUI14.xml package part or by ThisWorkbook's IRibbonExtensibility
    interface. The declarative part only works when it is properly wired:

      - customUI/customUI14.xml exists in the package
      - [Content_Types].xml declares it:
        <Override PartName="/customUI/customUI14.xml"
                  ContentType="application/vnd.ms-office.customUI+xml"/>
      - _rels/.rels has the ui/extensibility relationship
        (Target="customUI/customUI14.xml")
      - xl/_rels/workbook.xml.rels has the ui/extensibility relationship
        (Target="../customUI/customUI14.xml")

    WHY THIS MATTERS (session findings, Excel 2024 build 16.0.20228):
    Excel RE-SAVES an add-in file when you sign it in the VBA editor
    (Tools > Digital Signature, then save). That re-save regenerates
    xl/_rels/workbook.xml.rels and drops the customUI relationship, because
    Excel does not understand the customUI part. The result is a file that
    still CONTAINS the part but no longer points at it from the workbook.

    A present-but-unwired part is worse than no part at all: Excel sees a
    customUI part in the package and therefore never falls back to
    IRibbonExtensibility.GetCustomUI (the part masks the interface), while
    the unwired part itself is never processed. The observable symptom is
    exactly the one this whole investigation chased: no ribbon tab, onLoad
    never fired, GetCustomUI never called, no error.

    Repair-XlamRibbonWiring re-injects whatever is missing, PRESERVING every
    other entry byte-for-byte (in particular xl/vbaProject.bin and the
    xl/vbaProjectSignature*.bin parts), so a valid VBA digital signature -
    whose digest is over the vbaProject.bin CONTENT, not the zip bytes -
    survives the repair.

.NOTES
    Dot-source this file from build-xlam.ps1, sign-xlam.ps1 and install.ps1:
        . "$PSScriptRoot\xlam-ribbon-utils.ps1"
#>

Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

#-------------------------------------------------------------------------------
# Test-XlamRibbonWiring
# Description : Inspects an .xlam package and reports which pieces of the
#               customUI ribbon wiring are present.
# Returns     : PSCustomObject with Part, ContentType, RootRel, WorkbookRel,
#               Ok (all four present).
#-------------------------------------------------------------------------------
function Test-XlamRibbonWiring {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )
    $result = [PSCustomObject]@{
        Path        = $Path
        Part        = $false
        ContentType = $false
        RootRel     = $false
        WorkbookRel = $false
        Ok          = $false
    }
    if (-not (Test-Path $Path)) { return $result }
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $result.Part = $null -ne $zip.GetEntry("customUI/customUI14.xml")

            $ct = $zip.GetEntry("[Content_Types].xml")
            if ($ct) {
                $sr = New-Object System.IO.StreamReader($ct.Open())
                $ctXml = $sr.ReadToEnd(); $sr.Close()
                # Loose name-only check, matching Repair's detection: the
                # Override for /customUI/customUI14.xml must exist, regardless
                # of attribute ordering.
                $result.ContentType = $ctXml -match 'customUI/customUI14\.xml'
            }

            foreach ($relEntry in @("_rels/.rels", "xl/_rels/workbook.xml.rels")) {
                $e = $zip.GetEntry($relEntry)
                if (-not $e) { continue }
                $sr = New-Object System.IO.StreamReader($e.Open())
                $relXml = $sr.ReadToEnd(); $sr.Close()
                $found = $relXml -match 'office/2007/relationships/ui/extensibility" Target="[^"]*customUI/customUI14\.xml"'
                if ($relEntry -eq "_rels/.rels") { $result.RootRel = $found }
                else { $result.WorkbookRel = $found }
            }
            $result.Ok = $result.Part -and $result.ContentType -and $result.RootRel -and $result.WorkbookRel
        } finally { $zip.Dispose() }
    } catch {
        Write-Warning "Test-XlamRibbonWiring could not inspect '$Path': $($_.Exception.Message)"
    }
    return $result
}

#-------------------------------------------------------------------------------
# Remove-XlamSignatureParts
# Description : Reads $Path and writes $OutPath with every VBA signature part
#               (xl/vbaProjectSignature*.bin, including the legacy / Agile / V3
#               variants) and their [Content_Types].xml Overrides removed. Used
#               to give the signing flow a clean slate: repeated manual VBE
#               signing attempts with an incomplete 'Remove' step can leave
#               MULTIPLE conflicting signature parts in one file, which Excel's
#               add-in loader can refuse outright even though a direct open
#               tolerates it - exactly the state observed on the 13:58 build
#               (three parts: vbaProjectSignature.bin + Agile + V3). Every other
#               entry (incl. xl/vbaProject.bin and the customUI wiring) is
#               copied byte-for-byte.
# Returns     : $OutPath on success, $null on failure.
#-------------------------------------------------------------------------------
function Remove-XlamSignatureParts {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$OutPath
    )
    if (-not (Test-Path $Path)) { throw "Source add-in not found: $Path" }

    $srcBytes = [System.IO.File]::ReadAllBytes($Path)
    $inZip = New-Object System.IO.Compression.ZipArchive(
        [System.IO.MemoryStream]::new($srcBytes),
        [System.IO.Compression.ZipArchiveMode]::Read)

    $outStream = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::Create)
    $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)

    $removedCount = 0

    try {
        foreach ($entry in $inZip.Entries) {
            $name = $entry.FullName
            if ($name -match '(?i)^xl/vbaProjectSignature[^/]*\.bin$') {
                $removedCount++
                continue   # drop this signature part
            }
            $newEntry = $outZip.CreateEntry($name)
            $src = $entry.Open()
            $dst = $newEntry.Open()
            if ($name -eq "[Content_Types].xml") {
                $sr = New-Object System.IO.StreamReader($src)
                $xml = $sr.ReadToEnd(); $sr.Close()
                # Drop any Override that referenced a signature part.
                $xml = $xml -replace '(?i)<Override[^>]+PartName="/xl/vbaProjectSignature[^"]*"[^>]*/>', ''
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($xml)
                $dst.Write($bytes, 0, $bytes.Length)
            } else {
                $src.CopyTo($dst)
            }
            $src.Close(); $dst.Close()
        }
    } finally {
        $outZip.Dispose(); $outStream.Close(); $inZip.Dispose()
    }

    if ($removedCount -eq 0) {
        Write-Host "  No signature parts found - file is already unsigned (clean slate)." -ForegroundColor DarkGray
    } else {
        Write-Host "  Removed $removedCount signature part(s) - clean slate for signing." -ForegroundColor Green
    }
    return $OutPath
}

#-------------------------------------------------------------------------------
# Repair-XlamRibbonWiring
# Description : Reads $Path, re-injects any missing customUI ribbon wiring
#               (part, content-type override, root + workbook relationships),
#               and writes the result to $OutPath. Every other entry is copied
#               byte-for-byte, so vbaProject.bin and its signature parts are
#               preserved exactly.
# Returns     : $OutPath on success, $null on failure.
#-------------------------------------------------------------------------------
function Repair-XlamRibbonWiring {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$CustomUiXmlSource,
        [Parameter(Mandatory = $true)][string]$OutPath
    )
    if (-not (Test-Path $Path)) { throw "Source add-in not found: $Path" }
    if (-not (Test-Path $CustomUiXmlSource)) { throw "customUI14.xml source not found: $CustomUiXmlSource" }

    $srcBytes = [System.IO.File]::ReadAllBytes($Path)
    $inZip = New-Object System.IO.Compression.ZipArchive(
        [System.IO.MemoryStream]::new($srcBytes),
        [System.IO.Compression.ZipArchiveMode]::Read)

    $outStream = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::Create)
    $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)

    $xmlSource = [System.IO.File]::ReadAllText($CustomUiXmlSource)

    try {
        foreach ($entry in $inZip.Entries) {
            $name = $entry.FullName
            $newEntry = $outZip.CreateEntry($name)
            $src = $entry.Open()
            $dst = $newEntry.Open()

            if ($name -eq "customUI/customUI14.xml") {
                # Replace the (possibly stale) part with the current source.
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($xmlSource)
                $dst.Write($bytes, 0, $bytes.Length)
            } elseif ($name -eq "[Content_Types].xml") {
                $sr = New-Object System.IO.StreamReader($src)
                $xml = $sr.ReadToEnd(); $sr.Close()
                if ($xml -notmatch 'customUI/customUI14\.xml') {
                    $override = '<Override PartName="/customUI/customUI14.xml" ContentType="application/vnd.ms-office.customUI+xml"/>'
                    $xml = $xml.Replace("</Types>", "$override</Types>")
                }
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($xml)
                $dst.Write($bytes, 0, $bytes.Length)
            } elseif ($name -eq "_rels/.rels") {
                $sr = New-Object System.IO.StreamReader($src)
                $xml = $sr.ReadToEnd(); $sr.Close()
                if ($xml -notmatch 'ui/extensibility') {
                    $rel = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
                    $xml = $xml.Replace("</Relationships>", "$rel</Relationships>")
                }
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($xml)
                $dst.Write($bytes, 0, $bytes.Length)
            } elseif ($name -eq "xl/_rels/workbook.xml.rels") {
                $sr = New-Object System.IO.StreamReader($src)
                $xml = $sr.ReadToEnd(); $sr.Close()
                if ($xml -notmatch 'ui/extensibility') {
                    $rel = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="../customUI/customUI14.xml"/>'
                    $xml = $xml.Replace("</Relationships>", "$rel</Relationships>")
                }
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($xml)
                $dst.Write($bytes, 0, $bytes.Length)
            } else {
                $src.CopyTo($dst)
            }
            $src.Close(); $dst.Close()
        }

        # If the package had no customUI part at all, add it now.
        $hasPart = $null -ne $inZip.GetEntry("customUI/customUI14.xml")
        if (-not $hasPart) {
            $pkgDir = $outZip.CreateEntry("customUI/customUI14.xml")
            $w = New-Object System.IO.StreamWriter($pkgDir.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $w.Write($xmlSource); $w.Close()
        }
    } finally {
        $outZip.Dispose(); $outStream.Close(); $inZip.Dispose()
    }

    $verify = Test-XlamRibbonWiring $OutPath
    if (-not $verify.Ok) {
        Write-Warning "Repair-XlamRibbonWiring wrote $OutPath but wiring is still incomplete: Part=$($verify.Part) ContentType=$($verify.ContentType) RootRel=$($verify.RootRel) WorkbookRel=$($verify.WorkbookRel)"
        return $null
    }
    Write-Host "  Ribbon wiring repaired: part + content-type + root/workbook relationships all present in $OutPath" -ForegroundColor Green
    return $OutPath
}
