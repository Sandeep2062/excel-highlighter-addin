<#
.SYNOPSIS
    Automated 1-Click Installer for Excel Highlighter Add-in.
    Copies excel-highlighter.xlam and images/ to %APPDATA%\Microsoft\AddIns\ExcelHighlighter
    and activates it automatically in the Windows Registry for Excel.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = Get-Location }

$xlamSource = Join-Path $scriptDir "excel-highlighter.xlam"
$imagesSource = Join-Path $scriptDir "images"
if (-not (Test-Path $imagesSource)) {
    $imagesSource = Join-Path (Join-Path $scriptDir "customUI") "images"
}

# Always build fresh excel-highlighter.xlam from src/ and customUI/
$buildScript = Join-Path (Join-Path $scriptDir "scripts") "build-xlam.ps1"
if (Test-Path $buildScript) {
    Write-Host "Building latest excel-highlighter.xlam..." -ForegroundColor Yellow
    & $buildScript
}

if (-not (Test-Path $xlamSource)) {
    Write-Error "Could not find excel-highlighter.xlam. Please build or download it first."
    exit 1
}

$targetDir = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$targetXlam = Join-Path $targetDir "excel-highlighter.xlam"
$targetImages = Join-Path $targetDir "images"

Write-Host "Installing Excel-Highlighter Add-in..." -ForegroundColor Cyan

# Terminate running Excel instances to release file locks
$excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host "Terminating running Excel process(es) to release file lock..." -ForegroundColor Yellow
    Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# Clear Excel's ribbon cache so the Highlighter tab reliably appears on the
# next launch. A stale Excel.officeUI is the most common cause of a missing
# custom ribbon tab (it can be written to either %APPDATA% or %LOCALAPPDATA%).
$officeUiPaths = @(
    (Join-Path $env:APPDATA "Microsoft\Office\Excel.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Excel.officeUI")
)
foreach ($uiPath in $officeUiPaths) {
    if (Test-Path $uiPath) {
        Remove-Item $uiPath -Force -ErrorAction SilentlyContinue
        Write-Host "  Cleared ribbon cache: $uiPath" -ForegroundColor Green
    }
}

# --- Clear "Disabled Items" -----------------------------------------------
# If the add-in ever threw an error while Excel was loading it (a bad build,
# a crash mid-save, an interrupted install), Excel silently blacklists it in
# the Resiliency\DisabledItems registry key and will keep refusing to load
# it - including the ribbon tab - on every future launch, with no visible
# error. This is the single most common reason a previously-working add-in
# "just disappears". We proactively clear any entry referencing this add-in.
$officeVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $officeVersions) {
    $disabledPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $disabledPath) {
        Get-ItemProperty -Path $disabledPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSObject.Properties } |
            Where-Object { $_.Name -notlike "PS*" } |
            ForEach-Object {
                # DisabledItems values are binary blobs containing the file path
                # as part of the structure; a raw byte-string match is enough
                # here since we only care about entries mentioning our add-in.
                $bytes = $_.Value
                if ($bytes -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($bytes)
                    if ($text -like "*excel-highlighter*" -or $text -like "*ExcelHighlighter*" -or $text -like "*excel-crosshair-highlighter*") {
                        Remove-ItemProperty -Path $disabledPath -Name $_.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "  Cleared stale 'Disabled Items' entry in Excel $ver" -ForegroundColor Green
                    }
                }
            }
    }
}

# Copy files to the add-in folder. (Older builds also copied a second
# duplicate straight into the AddIns root, which just left two files to keep
# in sync and made uninstall incomplete - we only ever write one copy.)

# If the currently-installed copy carries a valid VBA digital signature,
# warn that this rebuild clobbers it. The signature is required for the
# custom ribbon on Excel 2024 / recent Microsoft 365 builds (see the
# signature section below); install.bat rebuilds unsigned, so the user must
# re-run sign-xlam.ps1 afterwards. (Session finding: a stale/invalid
# signature is silently treated as unsigned and the ribbon is suppressed.)
$wasSigned = $false
if (Test-Path $targetXlam) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $z = [System.IO.Compression.ZipFile]::OpenRead($targetXlam)
        try {
            foreach ($p in @("xl/vbaProjectSignature.bin", "xl/vbaProjectSignatureAgile.bin", "xl/vbaProjectSignatureV3.bin")) {
                if ($z.GetEntry($p)) { $wasSigned = $true; break }
            }
        } finally { $z.Dispose() }
    } catch {}
}

Copy-Item $xlamSource $targetXlam -Force
Unblock-File -Path $targetXlam -ErrorAction SilentlyContinue
Write-Host "  Copied and unblocked add-in at: $targetXlam" -ForegroundColor Green

if ($wasSigned) {
    Write-Host "  NOTE: the previous install was signed, but this rebuild is unsigned - the" -ForegroundColor Yellow
    Write-Host "        signature does not survive a rebuild. If you need the ribbon on Excel" -ForegroundColor Yellow
    Write-Host "        2024 / recent 365, run  .\scripts\sign-xlam.ps1 -Deploy -Trust  after this." -ForegroundColor Yellow
}

if (Test-Path $imagesSource) {
    if (-not (Test-Path $targetImages)) { New-Item -ItemType Directory -Path $targetImages -Force | Out-Null }
    Get-ChildItem -Path $imagesSource -File | Copy-Item -Destination $targetImages -Force
    Write-Host "  Copied UI icons to: $targetImages" -ForegroundColor Green
}

# The ribbon XML is embedded in the .xlam package as the customUI14.xml part
# (build-xlam.ps1 default). Excel reads the part directly; ThisWorkbook's
# IRibbonExtensibility.GetCustomUI stays in the project as a fallback for
# builds that do not honour the part. We ALSO copy customUI14.xml next to the
# .xlam - GetCustomUI reads it from there if it is ever invoked, so this copy
# must stay in place.
$ribbonXmlSource = Join-Path $scriptDir "customUI\customUI14.xml"
if (Test-Path $ribbonXmlSource) {
    Copy-Item $ribbonXmlSource (Join-Path $targetDir "customUI14.xml") -Force
    Write-Host "  Copied ribbon definition (customUI14.xml) alongside add-in (read by GetCustomUI at load time)" -ForegroundColor Green
} else {
    Write-Host "  WARNING: customUI/customUI14.xml not found - the add-in will fall back to a minimal ribbon." -ForegroundColor Yellow
}

# --- Validate / repair the deployed add-in's customUI ribbon wiring -----------
# CRITICAL (session finding, Excel 2024 build 16.0.20228): Excel re-saves the
# add-in when it is signed in the VBE and saved, and that re-save drops the
# ui/extensibility relationship from xl/_rels/workbook.xml.rels. A file that
# still CONTAINS the customUI part but no longer points at it from the
# workbook is the worst case: Excel sees the part and never falls back to
# IRibbonExtensibility.GetCustomUI (the part masks the interface), while the
# unwired part itself is never processed - no ribbon tab, no onLoad, no
# GetCustomUI call, no error. We therefore verify the deployed file's wiring
# here and repair it in place if it is missing.
$utilsScript = Join-Path $scriptDir "scripts\xlam-ribbon-utils.ps1"
# NOTE: both Test-Path calls MUST be parenthesized - a bare command followed
# by -and would bind -and as a *parameter* of Test-Path (PowerShell gotcha),
# which crashed the installer after the file copy with
# "A parameter cannot be found that matches parameter name 'and'."
if ((Test-Path $utilsScript) -and (Test-Path $ribbonXmlSource)) {
    . $utilsScript
    $wiring = Test-XlamRibbonWiring $targetXlam
    if (-not $wiring.Ok) {
        Write-Host "  Ribbon wiring incomplete (Part=$($wiring.Part) ContentType=$($wiring.ContentType) RootRel=$($wiring.RootRel) WorkbookRel=$($wiring.WorkbookRel)) - repairing..." -ForegroundColor Yellow
        $tmpRepair = Join-Path $env:TEMP ("install-repair-" + [Guid]::NewGuid().ToString("N") + ".xlam")
        $repaired = Repair-XlamRibbonWiring -Path $targetXlam -CustomUiXmlSource $ribbonXmlSource -OutPath $tmpRepair
        if ($repaired -and (Test-Path $tmpRepair)) {
            Copy-Item $tmpRepair $targetXlam -Force
            Remove-Item $tmpRepair -Force -ErrorAction SilentlyContinue
            Write-Host "  Repaired: customUI part + content-type + root/workbook relationships are now present." -ForegroundColor Green
        } else {
            Remove-Item $tmpRepair -Force -ErrorAction SilentlyContinue
            Write-Host "  WARNING: could not repair ribbon wiring - the Highlighter tab may not appear." -ForegroundColor Red
        }
    } else {
        Write-Host "  Ribbon wiring verified: customUI part + relationships present." -ForegroundColor Green
    }
} else {
    Write-Host "  WARNING: xlam-ribbon-utils.ps1 or customUI14.xml missing - skipping ribbon wiring check." -ForegroundColor Yellow
}

# --- Trust the publisher certificate (ribbon gate on Excel 2024+) -------------
# Empirically established on Excel 2024 build 16.0.20228 in valid interactive
# sessions: ALL customUI delivery for an UNSIGNED add-in is suppressed - the
# declarative customUI part, VBA IRibbonExtensibility on the root object, and
# even a bare no-VBA add-in carrying only a minimal part all fail to render
# their tab, with no error, while macros still run (VBAWarnings=1). The one
# add-in whose ribbon DOES render on the test machine (ASAP Utilities) is
# digitally signed by a certificate present in the user's Trusted Publishers
# store. So: Excel 2024 gates the ribbon on a Trusted Publisher signature.
#
# The VBA object model exposes no signing API, so the actual signing is a
# one-time manual step (scripts/sign-xlam.ps1, Tools > Digital Signature in the
# VBE). But the certificate must be TRUSTED for that signature to count, and
# importing it into the current user's Trusted Publishers store is fully
# automatable and needs no elevation. We do that here if the publisher's
# code-signing certificate exists (created by sign-xlam.ps1 on first run); if
# no certificate exists yet, the add-in still loads and highlights on all
# builds (the confirmed /R fix is version-independent) and the user can
# create+sign later if their Excel 2024/365 build suppresses the ribbon.
$publisherName = "Sandeep Khadka"
try {
    $pubCert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -like "CN=$publisherName*" -and $_.NotAfter -gt (Get-Date) } |
        Select-Object -First 1
    if ($pubCert) {
        $alreadyTrusted = Get-ChildItem Cert:\CurrentUser\TrustedPublisher -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $pubCert.Thumbprint }
        if (-not $alreadyTrusted) {
            $cerPath = Join-Path $env:TEMP ("pubcert-" + $pubCert.Thumbprint + ".cer")
            try {
                Export-Certificate -Cert $pubCert -FilePath $cerPath -Force | Out-Null
                Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher -ErrorAction Stop | Out-Null
                Write-Host "  Trusted publisher certificate imported (ribbon gate on Excel 2024+): $publisherName" -ForegroundColor Green
            } catch {
                Write-Warning "  Could not import publisher certificate into Trusted Publishers: $($_.Exception.Message)"
            } finally {
                Remove-Item $cerPath -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "  Publisher certificate already trusted: $publisherName" -ForegroundColor Green
        }
    }
} catch {
    Write-Warning "  Certificate check skipped: $($_.Exception.Message)"
}

# Registry activation for Excel 2010+ (16.0 = 2016+, 15.0 = 2013, 14.0 = 2010)
$activatedCount = 0

foreach ($ver in $officeVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $regPath) {
        $props = Get-ItemProperty -Path $regPath

        # Sweep ALL OPEN* values referencing this add-in and remove them
        # (old versions may have left several - e.g. a stale /R-prefixed value
        # in one slot and a plain one in another - and Excel processes every
        # OPEN* value, so a leftover /R value in any slot would still block
        # loading). Then write exactly one clean plain-path value.
        $openKeys = @($props.psobject.Properties | Where-Object { $_.Name -like "OPEN*" })
        foreach ($prop in $openKeys) {
            if ($prop.Value -like "*excel-highlighter*" -or $prop.Value -like "*excel-crosshair-highlighter*") {
                Remove-ItemProperty -Path $regPath -Name $prop.Name -Force -ErrorAction SilentlyContinue
                Write-Host "  Removed stale registry entry $($prop.Name) (was: $($prop.Value))" -ForegroundColor Yellow
            }
        }

        # If a removal silently failed (registry locked / access denied), a
        # stale /R value would survive and keep blocking the add-in while the
        # free-slot search below simply writes a NEW value. Fail loudly - this
        # is exactly the failure class this whole fix targets.
        $props = Get-ItemProperty -Path $regPath
        $stillThere = @($props.psobject.Properties | Where-Object {
            $_.Name -like "OPEN*" -and ($_.Value -like "*excel-highlighter*" -or $_.Value -like "*excel-crosshair-highlighter*")
        })
        if ($stillThere.Count -gt 0) {
            Write-Warning "  A stale registry entry for this add-in could not be removed ($($stillThere[0].Name)); Excel may still refuse to load it. Close Excel and re-run the installer."
        }

        $targetKeyName = $null
        $nextIndex = 0
        while ($true) {
            $keyName = if ($nextIndex -eq 0) { "OPEN" } else { "OPEN$nextIndex" }
            if (-not ($props.psobject.Properties.Name -contains $keyName)) {
                $targetKeyName = $keyName
                break
            }
            $nextIndex++
        }

        # NOTE: the OPEN value must NOT carry the /R switch. Empirically
        # verified this session: with /R the add-in showed as installed and
        # checked in the Add-ins dialog but never loaded at all (no
        # Workbook_Open, no ribbon tab, no error) - exactly the reported
        # symptom. Without /R it loads. ASAP Utilities is registered in this
        # same key WITHOUT /R and loads fine. Plain quoted path is correct.
        $regValue = '"' + $targetXlam + '"'
        Set-ItemProperty -Path $regPath -Name $targetKeyName -Value $regValue -Force | Out-Null
        Write-Host "  Activated in Excel $ver Registry ($targetKeyName -> $regValue)" -ForegroundColor Green
        $activatedCount++
    }
}

# --- COM ribbon add-in (Excel 2024 / recent 365) -------------------------------
# Empirically established on Excel 2024 (build 16.0.20228): the VBA ribbon
# path (customUI package part AND ThisWorkbook IRibbonExtensibility.GetCustomUI)
# is silently never queried for an .xlam - the add-in loads, macros run, but
# Excel never asks for ribbon XML, so the Highlighter tab never appears and no
# error is ever shown. The COM add-in below delivers the SAME customUI14.xml
# via IRibbonExtensibility, which Excel DOES query on every build, and every
# ribbon callback delegates back into the VBA add-in through Application.Run.
# (The .xlam's own customUI part/interface is left in place as the classic
# mechanism for older builds - harmless where it is ignored, and the COM
# add-in renders the tab on builds that need it.)
Write-Host ""
Write-Host "Building COM ribbon add-in (needed for the Highlighter tab on Excel 2024 / recent 365)..." -ForegroundColor Yellow
$comRibbonScript = Join-Path $scriptDir "scripts\build-comribbon.ps1"
if (Test-Path $comRibbonScript) {
    & $comRibbonScript -RepoRoot $scriptDir
} else {
    Write-Host "  WARNING: scripts\build-comribbon.ps1 not found - skipping COM ribbon add-in." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "`nSUCCESS! Excel-Highlighter Add-in has been installed." -ForegroundColor Green
Write-Host "Open Microsoft Excel and the 'Highlighter' tab will appear automatically." -ForegroundColor Cyan
Write-Host ""
Write-Host "If the tab still does not appear, in order:" -ForegroundColor Yellow
Write-Host "  1. Check File > Options > Add-ins > Manage: Disabled Items (Excel silently" -ForegroundColor Yellow
Write-Host "     blacklists add-ins that ever errored at load - the installer clears our" -ForegroundColor Yellow
Write-Host "     entries, but if Excel has re-added one, remove it here)." -ForegroundColor Yellow
Write-Host "  2. Check File > Options > Add-ins > Manage: COM Add-ins > Go... - ensure" -ForegroundColor Yellow
Write-Host "     'Excel Highlighter Ribbon' is checked (the COM add-in delivers the tab" -ForegroundColor Yellow
Write-Host "     on Excel 2024 / recent 365)." -ForegroundColor Yellow
Write-Host "  3. Check File > Options > Trust Center > Trust Center Settings > Macro" -ForegroundColor Yellow
Write-Host "     Settings - macros must be allowed for the ribbon buttons to work." -ForegroundColor Yellow
