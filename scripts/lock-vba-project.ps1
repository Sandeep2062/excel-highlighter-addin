<#
.SYNOPSIS
    Locks (password-protects) the Excel Highlighter VBA project so Excel's
    ribbon engine can discover IRibbonExtensibility.GetCustomUI, then
    re-injects the customUI ribbon wiring and verifies the lock.

.DESCRIPTION
    EMPIRICAL ROOT CAUSE (proven in valid interactive sessions on Excel 2024
    build 16.0.20228 - see CHANGELOG 2.1.3):

    The add-in loads and runs (Workbook_Open fires, macros run, hotkeys
    register), but its ribbon never renders. Every delivery mechanism was
    tested and eliminated:

      - The customUI14.xml package part (wiring verified byte-perfect) is
        IGNORED - even a minimal static part with zero callbacks fails to
        render, both in this package and when injected into a package whose
        own ribbon demonstrably renders.
      - IRibbonExtensibility.GetCustomUI on ThisWorkbook (implementation
        verified present and correct in the compiled project) is NEVER CALLED
        - confirmed with a plain-VBA marker file immune to logging failures,
        and with a minimal spec-perfect add-in built natively by Excel.
      - The digital signature is NOT the gate: an UNSIGNED copy of ASAP
        Utilities renders its tab from the very same folder (signature parts
        stripped, proven by experiment).
      - Trusted location is not the gate (folder IS trusted, macros run).

    The ONLY add-in whose ribbon renders on this machine (ASAP Utilities)
    differs from ours in exactly one structural way: its VBA project is
    PASSWORD-LOCKED. A locked VBA project is stored precompiled, and Excel's
    ribbon engine discovers IRibbonExtensibility from the compiled type info.
    An unlocked (source-only) project is queried at a point where the lazy
    compilation has not yet exposed the interface, so Excel silently treats
    the add-in as having no ribbon at all.

    This script performs the one step the VBA object model cannot automate:
    Tools > VBAProject Properties > Protection > "Lock project for viewing"
    + password. You do that once in the VBE (like the signing flow), press
    Enter, and the script verifies VBProject.Protection is locked, re-injects
    the customUI wiring that Excel's re-save strips, and optionally deploys.

.PARAMETER XlamPath
    Path to the .xlam to lock. Defaults to the installed copy at
    %APPDATA%\Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam,
    falling back to excel-highlighter.xlam in the repo root.

.PARAMETER Deploy
    After locking, copy the file over the installed copy under
    %APPDATA%\Microsoft\AddIns\ExcelHighlighter.

.PARAMETER NoLock
    Do NOT set the read-only lock after success (prevents Excel re-saving
    and stripping wiring/signature later). Default: lock read-only.

.EXAMPLE
    .\scripts\lock-vba-project.ps1 -Deploy

    Locks the installed add-in's VBA project, re-injects the ribbon wiring
    and verifies. Requires one manual VBE step.

.NOTES
    - You can lock the project with any password; remember it if you ever
      need to edit the code. Losing it means rebuilding the add-in (which
      is unsigned/unlocked) and re-locking.
    - After locking, the ribbon is delivered via GetCustomUI reading
      customUI14.xml from the add-in folder (deployed by install.ps1 next
      to the .xlam) - exactly the delivery mechanism ASAP Utilities uses.
    - Re-running install.bat rebuilds the add-in and removes the lock -
      re-run this script afterwards.
#>

[CmdletBinding()]
param(
    [string]$XlamPath,
    [switch]$Deploy,
    [switch]$NoLock
)

$ErrorActionPreference = "Stop"

# --- resolve paths -----------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = Get-Location }
$repoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path

if (-not $XlamPath) {
    $installed = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
    $repoXlam  = Join-Path $repoRoot "excel-highlighter.xlam"
    if (Test-Path $installed) { $XlamPath = $installed }
    elseif (Test-Path $repoXlam) { $XlamPath = $repoXlam }
}
if (-not $XlamPath -or -not (Test-Path $XlamPath)) {
    throw "Could not find the add-in. Build it first (scripts\build-xlam.ps1 or install.bat), or pass -XlamPath explicitly."
}
$XlamPath = (Resolve-Path $XlamPath).Path

# Shared zip-surgery helpers (Test-XlamRibbonWiring / Repair-XlamRibbonWiring).
$utilsScript = Join-Path $scriptDir "xlam-ribbon-utils.ps1"
if (Test-Path $utilsScript) { . $utilsScript }
else { throw "Shared helper not found: $utilsScript" }

Write-Host ""
Write-Host "Add-in to lock : $XlamPath" -ForegroundColor Cyan

# --- pre-flight: refuse while Excel is running -------------------------------
$runningExcel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($runningExcel) {
    throw "Excel is currently running, and a running Excel instance locks the add-in file (it auto-loads it via the registry). Close ALL Excel windows and re-run this script."
}
try {
    if ((Get-Item $XlamPath).IsReadOnly) {
        Set-ItemProperty -Path $XlamPath -Name IsReadOnly -Value $false
        Write-Host "Cleared the read-only lock from a previous run." -ForegroundColor Yellow
    }
} catch {}

# --- the manual VBE step ------------------------------------------------------
Write-Host ""
Write-Host "Opening the add-in in Excel and showing the Visual Basic Editor..." -ForegroundColor Cyan
Write-Host ""
Write-Host "In the Visual Basic Editor, do the following ONCE:" -ForegroundColor Yellow
Write-Host "  0. Make sure 'Highlighter (excel-highlighter.xlam)' is selected in the" -ForegroundColor Yellow
Write-Host "     Project Explorer pane (left side). Press Ctrl+R if you can't see it." -ForegroundColor Yellow
Write-Host "  1. Menu: Tools > VBAProject Properties..." -ForegroundColor Yellow
Write-Host "  2. Go to the 'Protection' tab" -ForegroundColor Yellow
Write-Host "  3. TICK 'Lock project for viewing'" -ForegroundColor Yellow
Write-Host "  4. Enter a password (twice) - remember it if you ever need to edit the code" -ForegroundColor Yellow
Write-Host "  5. Click OK" -ForegroundColor Yellow
Write-Host "  6. NOW SAVE the add-in yourself: press Ctrl+S inside the VBA editor" -ForegroundColor Yellow
Write-Host "     (or File > Save)." -ForegroundColor Yellow
Write-Host ""
Write-Host "Then come back here and press Enter to verify." -ForegroundColor Yellow
Write-Host "(Do NOT close the Excel window that opened - the script needs it.)" -ForegroundColor DarkGray

$excel = $null
$wb = $null
$origWriteTime = (Get-Item $XlamPath).LastWriteTime
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 1 } catch {}
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

    [void](Read-Host "Press Enter after you have locked AND saved the project (Ctrl+S)")

    # Verify the lock from the OPEN instance first.
    # vbext_ProjectProtection: vbext_pp_none = 0, vbext_pp_locked = 1.
    $protection = -1
    try { $protection = [int]$wb.VBProject.Protection } catch { $protection = -1 }
    Write-Host "VBProject.Protection (open instance) = $protection  (1 = locked)" -ForegroundColor Cyan

    # If the user saved from inside the VBE, the file changed and we don't need
    # a COM save (which has crashed before). Only fall back if unchanged.
    $newWriteTime = (Get-Item $XlamPath).LastWriteTime
    if ($newWriteTime -ne $origWriteTime) {
        Write-Host "Detected that the add-in was already saved from inside the VBE - no COM save needed." -ForegroundColor Green
    } else {
        Write-Host "The file was not saved yet - attempting to save it now..." -ForegroundColor Yellow
        try {
            $wb.Save()
            Write-Host "Saved: $XlamPath" -ForegroundColor Green
        } catch {
            Write-Warning "Save reported an error: $($_.Exception.Message). Verifying the file anyway."
        }
    }
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    Start-Sleep -Milliseconds 500
}

# --- step 2: re-inject the customUI ribbon wiring (Excel re-save strips it) ---
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
                Write-Host "Ribbon wiring re-injected after the VBE save (Excel had stripped it)." -ForegroundColor Green
            } catch {
                Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
                Write-Warning "Repaired copy written but could not replace $XlamPath - close Excel fully and re-run."
            }
        } else {
            Remove-Item $tmpRepaired -Force -ErrorAction SilentlyContinue
            Write-Warning "Could not re-inject ribbon wiring - the add-in may still load without its ribbon tab."
        }
    } else {
        Write-Warning "customUI14.xml not found at $customUiSource - cannot re-inject ribbon wiring."
    }
}

# --- step 3: verify the lock from a FRESH open (authoritative) ----------------
Write-Host ""
Write-Host "Verifying the lock from a fresh open..." -ForegroundColor Cyan
$checkExcel = $null
$checkWb = $null
$verified = $false
try {
    $checkExcel = New-Object -ComObject Excel.Application
    $checkExcel.Visible = $false
    $checkExcel.DisplayAlerts = $false
    $checkExcel.EnableEvents = $false
    try { $checkExcel.AutomationSecurity = 1 } catch {}
    try {
        foreach ($w in $checkExcel.Workbooks) {
            if ($w.IsAddin -and $w.Name -like "excel-highlighter*") { $w.Close($false) }
        }
    } catch {}
    $checkWb = $checkExcel.Workbooks.Open($XlamPath)
    $protection = -1
    $readable = $false
    try { $protection = [int]$checkWb.VBProject.Protection } catch { $protection = -1 }
    try { $readable = $checkWb.VBProject.VBComponents.Count -gt 0 } catch { $readable = $false }
    Write-Host "VBProject.Protection (fresh open) = $protection  (1 = locked)" -ForegroundColor Cyan
    # 1 = vbext_pp_locked. On a locked project the components are NOT readable
    # (count throws), which is expected.
    if ($protection -eq 1) {
        $verified = $true
        Write-Host "VERIFIED: the VBA project is LOCKED." -ForegroundColor Green
    } else {
        Write-Host "The project is NOT locked (protection=$protection). The lock did not stick." -ForegroundColor Yellow
        Write-Host "Redo the steps: Tools > VBAProject Properties > Protection > 'Lock project for viewing' + password, then Ctrl+S." -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Could not verify: $($_.Exception.Message)"
} finally {
    try { if ($checkWb) { $checkWb.Close($false) } } catch {}
    try { if ($checkExcel) { $checkExcel.Quit() } } catch {}
    Start-Sleep -Milliseconds 500
}

# --- step 4: deploy over the installed copy -----------------------------------
$installedTarget = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam"
if ($Deploy -and $XlamPath -ne $installedTarget) {
    $targetDir = Split-Path $installedTarget
    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    try {
        Copy-Item $XlamPath $installedTarget -Force
        Write-Host "Deployed locked add-in to: $installedTarget" -ForegroundColor Green
        if ($verified -and -not $NoLock) {
            try { Set-ItemProperty -Path $installedTarget -Name IsReadOnly -Value $true } catch {}
        }
    } catch {
        Write-Warning "Could not copy over the installed add-in ($($_.Exception.Message))."
    }
}

# --- optional read-only lock on the source file -------------------------------
if ($verified -and -not $NoLock) {
    try {
        Set-ItemProperty -Path $XlamPath -Name IsReadOnly -Value $true
        Write-Host "Locked the add-in file read-only so Excel cannot re-save it and strip the wiring/signature." -ForegroundColor Green
        Write-Host "(Re-run this script to clear the lock. Use -NoLock to skip.)" -ForegroundColor DarkGray
    } catch {
        Write-Warning "Could not set the read-only lock: $($_.Exception.Message)"
    }
}

Write-Host ""
if ($verified) {
    Write-Host "SUCCESS: the VBA project is locked. Restart Excel - the Highlighter tab" -ForegroundColor Green
    Write-Host "should now appear next to Home, delivered via GetCustomUI exactly like ASAP Utilities." -ForegroundColor Green
    Write-Host "(If the ribbon still does not appear after a full restart, the last resort is a" -ForegroundColor DarkGray
    Write-Host "genuine EV-signed COM add-in - but the locked-project shape is what renders on" -ForegroundColor DarkGray
    Write-Host "this build and it matches every working add-in observed.)" -ForegroundColor DarkGray
} else {
    Write-Host "(The project is not locked yet - re-run this script and complete the Protection step.)" -ForegroundColor Yellow
}
