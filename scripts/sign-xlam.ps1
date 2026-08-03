<#
.SYNOPSIS
    Signs the Excel Highlighter VBA project with a self-signed
    certificate so the Publisher column in Excel's Add-ins dialog
    shows your name instead of being blank.

.DESCRIPTION
    Excel shows the "Publisher" of a VBA add-in (.xlam) from the
    digital signature of its VBA project - NOT from document
    properties (Author/Company) and NOT from an Authenticode file
    signature. This script:

      1. Makes sure a self-signed code-signing certificate with your
         name exists in the current user's Personal certificate store,
         creating one automatically if it is missing.
      2. Opens the add-in in Excel and brings up the Visual Basic
         Editor, where you sign the project once:
             Tools > Digital Signature... > Choose... > "Sandeep Khadka" > OK
      3. Saves the add-in, then VERIFIES the signature with Excel's own
         verdict: the file is opened in a fresh hidden instance and
         `VBProject.Signed` is read on the fully-readable project - the
         only trustworthy check (the signature digest verifier is
         unreliable for VBA Agile/V3 signatures and is kept only as a
         fallback). Presence of the signature parts alone is NOT proof -
         Excel silently invalidates the signature whenever it re-saves
         the project after you sign it (e.g. via the "Remove personal
         information from file properties on save" privacy option), which
         leaves the parts behind but breaks the signature. That is the
         usual reason the Publisher column stays blank despite "signing".
      4. Locks the signed add-in file read-only so Excel cannot re-save
         it and invalidate the signature again. Use -NoLock to skip.

    The short manual click-through inside the VBE is unavoidable - the
    VBA object model does not expose signing - but it is a one-time
    four-click step.

.PARAMETER XlamPath
    Path to the .xlam to sign. Defaults to the installed copy at
    %APPDATA%\Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam,
    falling back to excel-highlighter.xlam in the repo root.

.PARAMETER PublisherName
    The name to show as the Publisher. Defaults to "Sandeep Khadka".

.PARAMETER Deploy
    After signing, copy the signed .xlam over the installed copy under
    %APPDATA%\Microsoft\AddIns\ExcelHighlighter. Used when signing a
    file that isn't already the installed copy (e.g. the repo build).

.PARAMETER Trust
    Also install the certificate into the current user's Trusted Root
    and Trusted Publishers stores so Excel does not warn that the
    publisher is unverified. Trusted Root requires administrator
    rights; if elevation is unavailable the script warns and continues
    (the Publisher column still displays correctly either way).

.PARAMETER Rebuild
    Run scripts/build-xlam.ps1 first, then sign the fresh build. Note
    that a rebuild produces an unsigned add-in, so this is the correct
    order: build, sign, deploy.

.PARAMETER NoLock
    Do NOT set the read-only lock on the signed add-in. By default the
    script locks the file so Excel cannot re-save it and silently
    invalidate the signature.

.EXAMPLE
    .\scripts\sign-xlam.ps1

    Signs the installed add-in at
    %APPDATA%\Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam.

.EXAMPLE
    .\scripts\sign-xlam.ps1 -XlamPath .\excel-highlighter.xlam -Deploy -Trust

    Signs the repo build, deploys it over the installed copy, and
    trusts the certificate (run elevated for the Trusted Root step).

.NOTES
    - The signature lives in the file itself, so it survives copies -
      but running scripts/build-xlam.ps1 (or install.ps1, which builds)
      regenerates the file unsigned. Re-run this script after any
      rebuild or reinstall.
    - Self-signed certificates expire (one year by default). When the
      signature stops being trusted, delete the old certificate, create
      a fresh one, and re-sign.
    - Close Excel before signing: a running Excel instance locks the
      add-in file, and the script refuses to start while the file is
      locked. If Excel dies mid-save, the script still verifies the
      file so you always get a definitive answer.
#>

[CmdletBinding()]
param(
    [string]$XlamPath,
    [string]$PublisherName = "Sandeep Khadka",
    [switch]$Deploy,
    [switch]$Trust,
    [switch]$Rebuild,
    [switch]$NoLock
)

$ErrorActionPreference = "Stop"

# --- resolve paths -----------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = Get-Location }
$repoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path

if ($Rebuild) {
    $buildScript = Join-Path $repoRoot "scripts\build-xlam.ps1"
    if (-not (Test-Path $buildScript)) { throw "build-xlam.ps1 not found at $buildScript" }
    Write-Host "Rebuilding the add-in first..." -ForegroundColor Cyan
    & $buildScript -RepoRoot $repoRoot
}

if (-not $XlamPath) {
    $installed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
    $repoXlam  = Join-Path $repoRoot "excel-highlighter.xlam"
    # With -Rebuild the freshly built repo copy is the file to sign (the build
    # just regenerated it); without it, the installed copy is the natural
    # default. Prefer the fresh build so '-Rebuild -Deploy' signs the new file
    # and pushes it over the installed copy - otherwise the script would sign
    # the OLD installed copy and the rebuild would never ship (the -Deploy
    # copy step is skipped when the source IS the install target).
    if ($Rebuild -and (Test-Path $repoXlam)) { $XlamPath = $repoXlam }
    elseif (Test-Path $installed) { $XlamPath = $installed }
    elseif (Test-Path $repoXlam) { $XlamPath = $repoXlam }
}
if (-not $XlamPath -or -not (Test-Path $XlamPath)) {
    throw "Could not find the add-in to sign. Build it first (scripts\build-xlam.ps1 or install.bat), or pass -XlamPath explicitly."
}
$XlamPath = (Resolve-Path $XlamPath).Path

# Shared zip-surgery helpers (Remove-XlamSignatureParts, Test-XlamRibbonWiring,
# Repair-XlamRibbonWiring). Dot-sourced ONCE here - used by the clean-slate
# strip, the ribbon re-injection and the deploy wiring check below.
$utilsScript = Join-Path $scriptDir "xlam-ribbon-utils.ps1"
if (Test-Path $utilsScript) { . $utilsScript }
else { throw "Shared helper not found: $utilsScript" }

Write-Host ""
Write-Host "Add-in to sign : $XlamPath" -ForegroundColor Cyan
Write-Host "Publisher name  : $PublisherName" -ForegroundColor Cyan

# --- pre-flight: refuse while a REAL Excel is running, clear lock, probe ---
# A running Excel instance auto-loads the add-in (via the OPEN registry value)
# and holds the file open. Saving from a second instance then crashes with
# RPC_E_DISCONNECTED, so refuse to start while Excel is in use.
#
# BUT "Excel is running" is not the same as "an Excel window is open". COM
# automation (build-xlam.ps1, earlier aborted runs of this script) frequently
# leaves windowless zombie EXCEL.EXE processes behind: no window, no tray
# icon, invisible in Task Manager's Apps list (they hide under
# Details / Background processes), yet still holding the add-in file open.
# Telling the user to "close Excel" for those is a dead end - there is
# nothing to close. So split the check:
#   - visible instance(s) with a window -> refuse, tell the user to close them
#   - only windowless zombie(s)         -> report PID + start time, then offer
#     to kill them (COM-automation leftovers, no user data to lose - killing
#     them is what "close Excel" would do anyway)
$runningExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($runningExcel.Count -gt 0) {
    $visible = @($runningExcel | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($visible.Count -gt 0) {
        throw "Excel is currently running with a visible window (PID(s): $($visible.Id -join ', ')). A running Excel instance locks the add-in file (it auto-loads it via the registry). Close ALL Excel windows (File > Exit, including the system tray) and re-run this script."
    }
    # Only windowless processes. They may be our own build's instances that
    # are still shutting down (just Quit()'d) - give them a moment to exit on
    # their own before treating them as zombies. Also re-check for visible
    # instances: one could appear during the wait (user opens Excel, or a
    # zombie surfaces a recovery dialog) and must be refused with the clear
    # message, not left to the generic file-lock probe.
    $zombies = @($runningExcel | Where-Object { $_.MainWindowHandle -eq 0 })
    $deadline = (Get-Date).AddSeconds(5)
    while ($zombies.Count -gt 0 -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $runningExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        $visible = @($runningExcel | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($visible.Count -gt 0) {
            throw "Excel is currently running with a visible window (PID(s): $($visible.Id -join ', ')). A running Excel instance locks the add-in file (it auto-loads it via the registry). Close ALL Excel windows (File > Exit, including the system tray) and re-run this script."
        }
        $zombies = @($runningExcel | Where-Object { $_.MainWindowHandle -eq 0 })
    }
    if ($zombies.Count -gt 0) {
        Write-Host "Found $($zombies.Count) windowless EXCEL.EXE process(es) - leftovers from COM automation (no window, no tray icon, but they lock the add-in file):" -ForegroundColor Yellow
        foreach ($z in $zombies) {
            $started = try { $z.StartTime.ToString("HH:mm:ss") } catch { "unknown" }
            Write-Host "  PID $($z.Id) - started $started" -ForegroundColor Yellow
        }
        # Windowless instances are usually COM-automation leftovers, but they
        # can also be automation servers hosting a workbook with unsaved
        # changes - killing them discards that. Default is N for that reason.
        $answer = Read-Host "Kill these windowless Excel processes and continue? (y/N - killing may discard unsaved changes in an automated session)"
        if ($answer -notmatch '^(y|yes)$') {
            throw "Excel processes are still running. Close ALL Excel windows AND end any EXCEL.EXE in Task Manager > Details, then re-run this script."
        }
        foreach ($z in $zombies) {
            try { Stop-Process -Id $z.Id -Force -ErrorAction Stop; Write-Host "  Killed EXCEL PID $($z.Id)." -ForegroundColor Green }
            catch { Write-Warning "  Could not kill EXCEL PID $($z.Id): $($_.Exception.Message)" }
        }
        Start-Sleep -Milliseconds 1500
    }
}

# A previous run may have locked the file read-only to protect the signature.
# Clear that attribute first so we can write the freshly signed project.
try {
    if ((Get-Item $XlamPath).IsReadOnly) {
        Set-ItemProperty -Path $XlamPath -Name IsReadOnly -Value $false
        Write-Host "Cleared the read-only lock from a previous signing run." -ForegroundColor Yellow
    }
} catch {}

# A running Excel instance holds the add-in file open. Opening it anyway and
# saving later is what causes the Excel COM instance to die mid-save
# (RPC_E_DISCONNECTED). Refuse to start instead.
$probe = $null
try {
    $probe = [System.IO.File]::Open($XlamPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
} catch {
    throw "The add-in file is locked, which usually means Excel is still running with it loaded. Close ALL Excel windows (and check the system tray) and re-run this script."
} finally {
    if ($probe) { $probe.Close(); $probe.Dispose() }
}

# --- step 1: make sure the certificate exists --------------------------------
function Get-PublisherCert {
    param([string]$Name)
    Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -like "CN=$Name*" -and $_.NotAfter -gt (Get-Date) } |
        Select-Object -First 1
}

$cert = Get-PublisherCert $PublisherName
if (-not $cert) {
    Write-Host "No code-signing certificate named '$PublisherName' found - creating one..." -ForegroundColor Yellow
    try {
        $cert = New-SelfSignedCertificate `
            -Subject "CN=$PublisherName" `
            -FriendlyName $PublisherName `
            -Type CodeSigningCert `
            -KeyUsage DigitalSignature `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
    } catch {
        throw "Could not create a certificate automatically ($($_.Exception.Message)). " +
              "Create one named '$PublisherName' with the Office SelfCert tool " +
              "(SELFCERT.EXE, installed with Office under " +
              "C:\Program Files\Microsoft Office\root\Office16\) and re-run this script."
    }
    Write-Host "Certificate created: $($cert.Subject)  (expires $($cert.NotAfter.ToShortDateString()))" -ForegroundColor Green
} else {
    Write-Host "Using existing certificate: $($cert.Subject)  (expires $($cert.NotAfter.ToShortDateString()))" -ForegroundColor Green
}

if ($Trust) {
    Write-Host ""
    Write-Host "Trusting the certificate..." -ForegroundColor Cyan
    $cerPath = Join-Path $env:TEMP ("publisher-cert-" + $cert.Thumbprint + ".cer")
    try { Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null } catch {}

    try {
        Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher -ErrorAction Stop | Out-Null
        Write-Host "  Added to Trusted Publishers." -ForegroundColor Green
    } catch {
        Write-Warning "  Could not add to Trusted Publishers: $($_.Exception.Message)"
    }
    try {
        Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root -ErrorAction Stop | Out-Null
        Write-Host "  Added to Trusted Root Certification Authorities." -ForegroundColor Green
    } catch {
        Write-Warning "  Could not add to Trusted Root (needs administrator rights). The Publisher column still shows the name - this step only avoids the 'publisher unverified' prompt. Re-run as Administrator if you want it."
    }
    Remove-Item $cerPath -Force -ErrorAction SilentlyContinue
}

# --- step 1b: ensure the VBA signing defaults exist ---------------------------
# Session finding (Excel 2024 build 16.0.20228): a project signed without
# V1HashEnhanced consistently comes back Stale - signature parts exist but
# Excel's own verdict (fresh open, VBProject.Signed on the readable project)
# is False, and the decoded signature's DigestAlgorithm reads empty. Microsoft's
# documentation ("Digitally sign your VBA macro project") prescribes exactly this
# key: HKCU\SOFTWARE\Microsoft\VBA\Security\V1HashEnhanced, DWORD, 1=SHA1,
# 2=SHA256, 3=SHA384, 4=SHA512. We default to SHA256. A TimeStampURL (RFC 3161)
# makes the signature survive certificate expiry. Both are current-user keys -
# reversible, no admin needed. Existing values are left untouched.
$vbaSecurityKey = "HKCU:\Software\Microsoft\VBA\Security"
try {
    if (-not (Test-Path $vbaSecurityKey)) { New-Item -Path $vbaSecurityKey -Force | Out-Null }
    if (-not (Get-ItemProperty $vbaSecurityKey -Name V1HashEnhanced -ErrorAction SilentlyContinue)) {
        New-ItemProperty -Path $vbaSecurityKey -Name V1HashEnhanced -PropertyType DWord -Value 2 -Force | Out-Null
        Write-Host "Set V1HashEnhanced=2 (SHA256) - the documented fix for VBA signatures Excel rejects." -ForegroundColor Green
    }
    if (-not (Get-ItemProperty $vbaSecurityKey -Name TimeStampURL -ErrorAction SilentlyContinue)) {
        New-ItemProperty -Path $vbaSecurityKey -Name TimeStampURL -PropertyType String -Value "http://timestamp.digicert.com" -Force | Out-Null
        New-ItemProperty -Path $vbaSecurityKey -Name TimeStampRetryCount -PropertyType DWord -Value 3 -Force | Out-Null
        New-ItemProperty -Path $vbaSecurityKey -Name TimeStampRetryDelay -PropertyType DWord -Value 1000 -Force | Out-Null
        Write-Host "Set TimeStampURL (RFC 3161) so the signature survives cert expiry." -ForegroundColor Green
    }
} catch {
    Write-Warning "Could not ensure VBA signing registry keys: $($_.Exception.Message)"
}

# --- report the current signature state so the user knows what to expect -----
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
$hasSigParts = $false
try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($XlamPath)
    try {
        foreach ($part in @("xl/vbaProjectSignature.bin", "xl/vbaProjectSignatureAgile.bin", "xl/vbaProjectSignatureV3.bin")) {
            if ($zip.GetEntry($part)) { $hasSigParts = $true; break }
        }
    } finally { $zip.Dispose() }
} catch {}

# Strip ANY leftover signature parts before opening for signing, so the VBE
# signing starts from a provably clean slate. Session finding (Excel 2024
# build 16.0.20228): repeated manual VBE signing with an incomplete 'Remove'
# step leaves MULTIPLE conflicting signature parts (legacy + Agile + V3) in
# one file, and Excel's ADD-IN LOADER refuses such a file outright (it never
# even loads the add-in - no StartUp in the log) while a direct workbook open
# tolerates it. Remove-XlamSignatureParts preserves vbaProject.bin and the
# customUI wiring byte-for-byte, so this is safe on an already-valid signature
# too - but for THIS flow we want a clean unsigned file so the one signature
# we are about to apply is the only one.
if ($hasSigParts -and (Get-Command Remove-XlamSignatureParts -ErrorAction SilentlyContinue)) {
    $tmpClean = Join-Path $env:TEMP ("excel-clean-" + [Guid]::NewGuid().ToString("N") + ".xlam")
    try {
        $cleaned = Remove-XlamSignatureParts -Path $XlamPath -OutPath $tmpClean
        if ($cleaned -and (Test-Path $tmpClean)) {
            Copy-Item $tmpClean $XlamPath -Force
            Write-Host "Stripped leftover signature parts - signing from a clean slate." -ForegroundColor Green
            $hasSigParts = $false
        }
    } catch {
        Write-Warning "Could not strip leftover signature parts: $($_.Exception.Message)"
    } finally {
        Remove-Item $tmpClean -Force -ErrorAction SilentlyContinue
    }
}

if ($hasSigParts) {
    Write-Host ""
    Write-Host "NOTE: the add-in file already contains a signature. If the Digital Signature dialog" -ForegroundColor Yellow
    Write-Host "      shows an existing signature, click 'Remove' first, then 'Choose...' and re-select" -ForegroundColor Yellow
    Write-Host "      '$PublisherName' - a leftover signature does NOT count." -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "Current state: unsigned (no existing signature) - clean signing." -ForegroundColor Green
}

# --- step 2: open the add-in and guide the one manual signing step -----------
Write-Host ""
Write-Host "Opening the add-in in Excel and showing the Visual Basic Editor..." -ForegroundColor Cyan
Write-Host ""
Write-Host "In the Visual Basic Editor, do the following:" -ForegroundColor Yellow
Write-Host "  0. Make sure 'Highlighter (excel-highlighter.xlam)' is selected in the"
Write-Host "     Project Explorer pane (left side). Press Ctrl+R if you can't see it."
Write-Host "  1. Menu: Tools > Digital Signature..."
Write-Host "     - If the dialog ALREADY shows a signature, click 'Remove' first."
Write-Host "       (A leftover signature from before does NOT count - it must be"
Write-Host "       removed and re-applied so it matches the current project.)"
Write-Host "  2. Click the 'Choose...' button"
Write-Host "  3. Select '$PublisherName' and click OK"
Write-Host "  4. Click OK to close the Digital Signature dialog"
Write-Host "  5. NOW SAVE the add-in yourself: press Ctrl+S inside the VBA editor"
Write-Host "     (or File > Save). Saving from inside the VBE is reliable - the"
Write-Host "     script's own save has crashed before (RPC_E_DISCONNECTED), so we"
Write-Host "     prefer you save it natively."
Write-Host ""
Write-Host "Then come back here and press Enter to verify." -ForegroundColor Yellow
Write-Host "(Do NOT close the Excel window that opened - the script needs it. Just switch back to this terminal.)" -ForegroundColor DarkGray

$excel = $null
$wb = $null
$saveError = $null
$origWriteTime = (Get-Item $XlamPath).LastWriteTime
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true   # visible so Alt+F11 is always available, even if the VBE auto-open fails
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 1 } catch {}   # 1 = low, loads macros without prompts
    # Session finding: the add-in auto-loads via the OPEN registry value, so a
    # fresh Excel instance ALREADY has a copy open before we open $XlamPath.
    # The VBE then shows TWO projects both named 'Highlighter
    # (excel-highlighter.xlam)' and the user can sign/save the WRONG one -
    # which is how a signature ends up on a different project than the file we
    # verify (verdict: parts exist but Signed=False). Close the auto-loaded
    # copy first so the VBE shows exactly one project with an unambiguous name.
    try {
        foreach ($_wb in $excel.Workbooks) {
            if ($_wb.IsAddin -and $_wb.Name -like "excel-highlighter*") {
                $_wb.Close($false)
                Write-Host "Closed the auto-loaded copy of the add-in so the VBE shows only the file being signed." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Warning "Could not close the auto-loaded add-in copy: $($_.Exception.Message)"
    }
    try {
        $wb = $excel.Workbooks.Open($XlamPath)
    } catch {
        throw "Could not open '$XlamPath'. Excel is probably still running and holding the add-in file open - close Excel completely and re-run this script. (Details: $($_.Exception.Message))"
    }

    try {
        $excel.VBE.MainWindow.Visible = $true
    } catch {
        Write-Warning "Could not show the VBA editor automatically - press Alt+F11 inside the Excel window that opened."
    }

    [void](Read-Host "Press Enter after you have signed AND saved the project in the VBA editor (Ctrl+S)")

    # Advisory check only - and now with the correct interpretation. The VBE's
    # VBProject.Signed is authoritative on a fully READABLE project: False means
    # the signature genuinely did not stick. (The known false negative applies
    # only to LOCKED/password-protected projects like ASAP Utilities - where the
    # property reads False because Excel cannot evaluate it. Our project is
    # always readable, so on this file Signed=False is a real failure.)
    $isSigned = $false
    $componentsReadable = $false
    try { $isSigned = [bool]$wb.VBProject.Signed } catch { $isSigned = $false }
    try { $componentsReadable = $wb.VBProject.VBComponents.Count -gt 0 } catch { $componentsReadable = $false }
    if ($isSigned) {
        Write-Host "Signature detected in the VBA project." -ForegroundColor Green
    } elseif ($componentsReadable) {
        Write-Warning "VBProject.Signed reports no signature, and the project is fully readable - this is a REAL failure, not a false negative. The signature did not stick."
        Write-Host "  Redo the dialog carefully: Tools > Digital Signature > (Remove if shown) >" -ForegroundColor Yellow
        Write-Host "  Choose... > '$PublisherName' > OK > OK > then Ctrl+S to save." -ForegroundColor Yellow
    } else {
        Write-Warning "VBProject.Signed reports no signature. The project is not readable here (locked?) - the saved file is the real test."
    }

    # If the user saved from inside the VBE (Ctrl+S), the file has changed and
    # we don't need to save again via COM (which has crashed before). Only fall
    # back to a COM save if the file is unchanged.
    $newWriteTime = (Get-Item $XlamPath).LastWriteTime
    if ($newWriteTime -ne $origWriteTime) {
        Write-Host "Detected that the add-in was already saved from inside the VBE - no COM save needed." -ForegroundColor Green
    } else {
        Write-Host "The file was not saved yet - attempting to save it now..." -ForegroundColor Yellow
        try {
            $wb.Save()
            Write-Host "Saved: $XlamPath" -ForegroundColor Green
        } catch {
            $saveError = $_
            Write-Warning "Save reported an error: $($_.Exception.Message). If you signed and saved in the VBE, the signature may already be in the file - verifying the file now."
        }
    }
} finally {
    if ($wb) {
        try { $wb.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    Start-Sleep -Milliseconds 500
}

# --- step 2b: re-inject the customUI ribbon wiring into the SIGNED file ------
# CRITICAL (session finding, Excel 2024 build 16.0.20228): Excel re-saves the
# add-in when you sign it in the VBE and save. That re-save regenerates
# xl/_rels/workbook.xml.rels and DROPS the customUI relationship (Excel does
# not understand the customUI part). The file then contains the part but no
# longer points at it from the workbook - and a present-but-unwired part masks
# IRibbonExtensibility.GetCustomUI while never being processed, so the ribbon
# silently never loads (no tab, onLoad never fired, GetCustomUI never called,
# no error - exactly the symptom this whole investigation chased).
#
# Re-inject the wiring into the saved file NOW (after the Excel save, before
# any read-only lock). The repair copies every other package entry byte-for-
# byte, so xl/vbaProject.bin and the signature parts are untouched and the
# signature digest we verify in step 3 stays valid.
if (Get-Command Repair-XlamRibbonWiring -ErrorAction SilentlyContinue) {
    $customUiSource = Join-Path $repoRoot "customUI\customUI14.xml"
    if (Test-Path $customUiSource) {
        $tmpRepaired = Join-Path $env:TEMP ("excel-repaired-" + [Guid]::NewGuid().ToString("N") + ".xlam")
        $repaired = $null
        try {
            $repaired = Repair-XlamRibbonWiring -Path $XlamPath -CustomUiXmlSource $customUiSource -OutPath $tmpRepaired
        } catch {
            Write-Warning "Ribbon re-injection failed: $($_.Exception.Message)"
        }
        if ($repaired -and (Test-Path $tmpRepaired)) {
            try {
                Copy-Item $tmpRepaired $XlamPath -Force
                Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
                Write-Host "Ribbon wiring re-injected after signing save (Excel had stripped it)." -ForegroundColor Green
            } catch {
                Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
                Write-Warning "Repaired copy written but could not replace $XlamPath ($($_.Exception.Message)) - close Excel fully and re-run the signing script."
            }
        } else {
            Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
            Write-Warning "Could not re-inject ribbon wiring - the signed add-in may still load without its ribbon tab. Check customUI/customUI14.xml exists and re-run."
        }
    } else {
        Write-Warning "customUI14.xml not found at $customUiSource - cannot re-inject ribbon wiring."
    }
} else {
    Write-Warning "Repair-XlamRibbonWiring not available - skipping ribbon re-injection."
}

# --- step 3: verify the signature from the saved file itself -----------------
# The authoritative check. Merely finding xl/vbaProjectSignature*.bin parts is
# NOT enough: Excel leaves the parts behind even after it re-saves the project
# and silently invalidates the signature. We decode the CMS blob and compare
# the digest embedded in its SPC indirect-data structure against the current
# vbaProject.bin bytes. Returns:
#   "Valid" - parts present AND digest matches the current project
#   "Stale" - parts present but the project was modified/saved after signing
#   "None"  - no signature parts
function Test-AddinSignatureState {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return "None" }
    Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $projEntry = $zip.GetEntry("xl/vbaProject.bin")
            if (-not $projEntry) { return "None" }

            $proj = New-Object byte[] $projEntry.Length
            $s2 = $projEntry.Open(); [void]$s2.Read($proj, 0, $proj.Length); $s2.Close()

            # Check whichever signature parts exist (legacy / Agile / V3).
            $sigEntries = @(
                $zip.GetEntry("xl/vbaProjectSignature.bin"),
                $zip.GetEntry("xl/vbaProjectSignatureAgile.bin"),
                $zip.GetEntry("xl/vbaProjectSignatureV3.bin")
            ) | Where-Object { $_ -ne $null }
            if ($sigEntries.Count -eq 0) { return "None" }

            # Pre-compute candidate digests of the project. VBA historically
            # uses MD5; be tolerant of SHA-1/SHA-256 from newer Office builds.
            $md5    = [System.Security.Cryptography.MD5]::Create()
            $sha1   = [System.Security.Cryptography.SHA1]::Create()
            $sha256 = [System.Security.Cryptography.SHA256]::Create()
            $hashMap = @{
                ([BitConverter]::ToString($md5.ComputeHash($proj)).Replace("-", ""))    = "MD5"
                ([BitConverter]::ToString($sha1.ComputeHash($proj)).Replace("-", ""))   = "SHA1"
                ([BitConverter]::ToString($sha256.ComputeHash($proj)).Replace("-", "")) = "SHA256"
            }

            foreach ($sigEntry in $sigEntries) {
                $sig = New-Object byte[] $sigEntry.Length
                $s = $sigEntry.Open(); [void]$s.Read($sig, 0, $sig.Length); $s.Close()

                # locate the PKCS#7 SignedData blob: 30 82 <len> 06 09 2a 86 48 86 f7 0d 01 07 02
                $oid = @(0x06, 0x09, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x07, 0x02)
                $oidPos = -1
                for ($i = 0; $i -lt $sig.Length - $oid.Length; $i++) {
                    $ok = $true
                    for ($j = 0; $j -lt $oid.Length; $j++) {
                        if ($sig[$i + $j] -ne $oid[$j]) { $ok = $false; break }
                    }
                    if ($ok) { $oidPos = $i; break }
                }
                if ($oidPos -lt 0) { continue }

                $start = $oidPos - 4
                if ($start -lt 0 -or $sig[$start] -ne 0x30) { $start = $oidPos - 2 }
                $len = ([int]$sig[$start + 2] -shl 8) -bor [int]$sig[$start + 3]
                $total = $start + 4 + $len
                if ($total -gt $sig.Length -or $len -le 10) {
                    $start = $oidPos - 2
                    $len = ([int]$sig[$start + 2] -shl 8) -bor [int]$sig[$start + 3]
                    $total = $start + 4 + $len
                }
                $p7bytes = New-Object byte[] ($total - $start)
                [Array]::Copy($sig, $start, $p7bytes, 0, $total - $start)

                $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
                $cms.Decode($p7bytes)
                $content = $cms.ContentInfo.Content

                # Scan the SPC indirect data for OCTET STRINGs (16/20/32 bytes)
                # whose value matches one of the project digests.
                for ($i = 0; $i -lt $content.Length - 33; $i++) {
                    if ($content[$i] -ne 0x04) { continue }
                    $digestLen = $content[$i + 1]
                    if ($digestLen -ne 16 -and $digestLen -ne 20 -and $digestLen -ne 32) { continue }
                    if ($i + 2 + $digestLen -gt $content.Length) { continue }
                    $cand = New-Object byte[] $digestLen
                    [Array]::Copy($content, $i + 2, $cand, 0, $digestLen)
                    $hexCand = [BitConverter]::ToString($cand).Replace("-", "")
                    if ($hashMap.ContainsKey($hexCand)) { return "Valid" }
                }
            }
            return "Stale"
        } finally { $zip.Dispose() }
    } catch {
        return "None"
    }
}

Write-Host ""
Write-Host "Verifying the saved file..." -ForegroundColor Cyan

# --- step 3a: Excel's own verdict (authoritative) ----------------------------
# The digest verifier below is known UNRELIABLE (it reports 'Stale' even for
# ASAP Utilities, whose ribbon demonstrably renders) - VBA Agile/V3 signatures
# do not embed a plain MD5/SHA1/SHA256 of vbaProject.bin, which is what it
# scans for. Excel itself is the only trustworthy judge: open the file in a
# fresh hidden instance and read VBProject.Signed. On our fully-readable
# project that verdict is authoritative - False means the signature genuinely
# did not land (the false-negative case only applies to locked projects).
function Test-FreshOpenSignature {
    param([string]$Path)
    # Presence of signature parts is decided from the package itself (zip) -
    # it distinguishes "no signature was ever applied" from "parts exist but
    # Excel doesn't consider them valid", which have different messages.
    $hasParts = $false
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            foreach ($p in @("xl/vbaProjectSignature.bin", "xl/vbaProjectSignatureAgile.bin", "xl/vbaProjectSignatureV3.bin")) {
                if ($zip.GetEntry($p)) { $hasParts = $true; break }
            }
        } finally { $zip.Dispose() }
    } catch {}
    $checkExcel = $null
    $checkWb = $null
    try {
        $checkExcel = New-Object -ComObject Excel.Application
        $checkExcel.Visible = $false
        $checkExcel.DisplayAlerts = $false
        $checkExcel.EnableEvents = $false
        try { $checkExcel.AutomationSecurity = 1 } catch {}
        # Close any copy of this add-in already loaded via the OPEN key in this
        # instance (the compile-check pattern), then open the file directly.
        try {
            foreach ($wb in $checkExcel.Workbooks) {
                if ($wb.IsAddin -and $wb.Name -like "excel-highlighter*") { $wb.Close($false) }
            }
        } catch {}
        $checkWb = $checkExcel.Workbooks.Open($Path)
        $signed = $false
        $readable = $false
        try { $signed = [bool]$checkWb.VBProject.Signed } catch { $signed = $false }
        try { $readable = $checkWb.VBProject.VBComponents.Count -gt 0 } catch { $readable = $false }
        return [PSCustomObject]@{ Signed = $signed; Readable = $readable; HasParts = $hasParts }
    } catch {
        return [PSCustomObject]@{ Signed = $false; Readable = $false; HasParts = $hasParts; Error = $_.Exception.Message }
    } finally {
        try { if ($checkWb) { $checkWb.Close($false) } } catch {}
        try { if ($checkExcel) { $checkExcel.Quit() } } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($checkExcel) | Out-Null } catch {}
    }
}

$fresh = Test-FreshOpenSignature $XlamPath
$sigState = "None"
if ($fresh.Readable) {
    if ($fresh.Signed) {
        $sigState = "Valid"
        Write-Host "Excel's own verdict (fresh open): VALID signature on the readable project." -ForegroundColor Green
    } elseif ($fresh.HasParts) {
        $sigState = "Stale"
        Write-Host "Excel's own verdict (fresh open): signature parts exist but the readable project has NO valid signature." -ForegroundColor Yellow
    } else {
        $sigState = "None"
        Write-Host "Excel's own verdict (fresh open): no signature parts at all - the signature did not land." -ForegroundColor Yellow
    }
} else {
    Write-Warning "Could not read the project in a fresh Excel instance ($($fresh.Error)) - falling back to the cryptographic digest check (known unreliable)."
    $sigState = Test-AddinSignatureState $XlamPath
}

switch ($sigState) {
    "Valid" {
        Write-Host "VERIFIED: the add-in file carries a VALID VBA digital signature issued to '$PublisherName'." -ForegroundColor Green
        if (-not $NoLock) {
            try {
                Set-ItemProperty -Path $XlamPath -Name IsReadOnly -Value $true
                Write-Host "Locked the add-in read-only so Excel cannot re-save it and invalidate the signature." -ForegroundColor Green
                Write-Host "(Re-run this script to clear the lock and re-sign. Use -NoLock to skip the lock.)" -ForegroundColor DarkGray
            } catch {
                Write-Warning "Could not set the read-only lock: $($_.Exception.Message)"
            }
        }
        Write-Host "Restart Excel and check File > Options > Add-ins - the Publisher column now shows '$PublisherName'." -ForegroundColor Green
    }
    "Stale" {
        Write-Warning "PARTIAL: the file contains signature parts, but Excel does not consider them valid - the signature did NOT stick."
        Write-Host "The most common causes, in order:" -ForegroundColor Yellow
        Write-Host "  1. The Digital Signature dialog was not completed - you must click 'Choose...'," -ForegroundColor Yellow
        Write-Host "     select '$PublisherName', and click OK on BOTH dialogs before saving." -ForegroundColor Yellow
        Write-Host "  2. The project was saved AFTER signing (Excel re-saving the add-in, e.g. via the" -ForegroundColor Yellow
        Write-Host "     'Remove personal information from file properties on save' privacy option," -ForegroundColor Yellow
        Write-Host "     breaks the digest) - the read-only lock this script applies prevents that." -ForegroundColor Yellow
        Write-Host "  3. The wrong project was selected in the Project Explorer." -ForegroundColor Yellow
        Write-Host "Re-run this script and sign again - leftover parts are now stripped automatically," -ForegroundColor Yellow
        Write-Host "so each attempt starts from a clean slate." -ForegroundColor Yellow
        if ($saveError) {
            Write-Host ""
            Write-Host "The save itself also reported an error ($($saveError.Exception.Message)) - the file may" -ForegroundColor Yellow
            Write-Host "have been locked by another Excel instance. Close ALL Excel windows and re-run this script." -ForegroundColor Yellow
        }
    }
    default {
        Write-Warning "VERIFIED: the add-in file still has NO VBA digital signature."
        Write-Host "Likely causes and fixes:" -ForegroundColor Yellow
        Write-Host "  - The Digital Signature dialog was not completed - make sure you clicked"
        Write-Host "    'Choose...', selected '$PublisherName', and clicked OK on BOTH dialogs."
        Write-Host "  - The wrong project was selected. In the Project Explorer (Ctrl+R),"
        Write-Host "    single-click 'Highlighter (excel-highlighter.xlam)' BEFORE opening"
        Write-Host "    Tools > Digital Signature."
        Write-Host "  - Excel's 'Remove personal information from file properties on save'"
        Write-Host "    privacy option can strip signatures on save (File > Options > Trust"
        Write-Host "    Center > Trust Center Settings > Privacy Options). Untick it if set."
        if ($saveError) {
            Write-Host ""
            Write-Host "The save itself also reported an error ($($saveError.Exception.Message)) - the file may" -ForegroundColor Yellow
            Write-Host "have been locked by another Excel instance. Close ALL Excel windows and re-run this script." -ForegroundColor Yellow
        }
    }
}

# --- step 4: deploy over the installed copy ----------------------------------
$installedTarget = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
if ($Deploy -and $XlamPath -ne $installedTarget) {
    $targetDir = Split-Path $installedTarget
    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    try {
        # Verify the wiring in the file being deployed BEFORE copying - a signed
        # file that lost its customUI rels (Excel re-save during signing) would
        # otherwise ship a present-but-unwired part whose ribbon never loads.
        $deployCheck = Test-XlamRibbonWiring $XlamPath
        if (-not $deployCheck.Ok -and (Get-Command Repair-XlamRibbonWiring -ErrorAction SilentlyContinue)) {
            $customUiSource = Join-Path $repoRoot "customUI\customUI14.xml"
            $tmpRepaired = Join-Path $env:TEMP ("excel-deploy-repair-" + [Guid]::NewGuid().ToString("N") + ".xlam")
            $repaired = Repair-XlamRibbonWiring -Path $XlamPath -CustomUiXmlSource $customUiSource -OutPath $tmpRepaired
            if ($repaired) {
                Copy-Item $repaired $XlamPath -Force
                Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
                Write-Host "Ribbon wiring repaired before deploy." -ForegroundColor Green
            } else {
                Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
                Write-Warning "Deployed file is missing customUI wiring and repair failed - the ribbon tab may not appear."
            }
        }
        Copy-Item $XlamPath $installedTarget -Force
        Write-Host "Deployed signed add-in to: $installedTarget" -ForegroundColor Green
        if ($sigState -eq "Valid" -and -not $NoLock) {
            try { Set-ItemProperty -Path $installedTarget -Name IsReadOnly -Value $true } catch {}
        }
    } catch {
        Write-Warning "Could not copy over the installed add-in ($($_.Exception.Message)). Excel may still be running - close it and copy the signed file manually."
    }
}

Write-Host ""
if ($sigState -eq "Valid") {
    Write-Host "(Tip: if you ever re-run install.bat, it rebuilds the add-in and removes the signature - run this script again afterwards. The read-only lock is removed automatically on the next sign run.)" -ForegroundColor DarkGray
} else {
    Write-Host "(Nothing was lost - an unsigned add-in behaves exactly as before. Re-run this script to try again.)" -ForegroundColor DarkGray
}
