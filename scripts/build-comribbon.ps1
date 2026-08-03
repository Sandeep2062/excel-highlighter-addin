<#
.SYNOPSIS
    Builds and registers the ExcelHighlighter.Ribbon COM add-in that delivers
    the Highlighter ribbon on Excel 2024 / recent Microsoft 365 builds where
    the VBA ribbon path (customUI part / ThisWorkbook IRibbonExtensibility)
    is silently never queried.

.DESCRIPTION
    Compiles comribbon/Connect.cs into ExcelHighlighterRibbon.dll using the
    .NET Framework csc.exe and the canonical Office interop assemblies from
    the GAC (Extensibility.dll for IDTExtensibility2, OFFICE.DLL for
    IRibbonExtensibility/IRibbonUI/IRibbonControl), deploys the DLL next to
    the .xlam, registers it in HKCU via RegAsm's /regfile (rewritten to
    HKCU\Software\Classes - no admin needed), sets the Excel COM add-in keys
    with LoadBehavior=3 (load at startup), and clears any Resiliency\Disabled
    Items entry for it.

    Idempotent: re-running overwrites the previous DLL and re-imports the
    registry values.

.NOTES
    Requires the .NET Framework 4.x csc.exe/RegAsm.exe (present on every
    Windows 10/11 with .NET Framework) and Office interop assemblies in the
    GAC (present when Office is installed).
#>

[CmdletBinding()]
param(
    [string]$RepoRoot
)

$ErrorActionPreference = "Continue"

if (-not $RepoRoot) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
    if (-not $scriptDir) { $scriptDir = Get-Location }
    $RepoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
}

$sourceFile   = Join-Path $RepoRoot "comribbon\Connect.cs"
$deployDir    = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelHighlighter"
$deployDll    = Join-Path $deployDir "ExcelHighlighterRibbon.dll"

if (-not (Test-Path $sourceFile)) {
    Write-Host "  SKIP: comribbon\Connect.cs not found ($sourceFile) - no COM ribbon add-in." -ForegroundColor Yellow
    exit 0
}
if (-not (Test-Path $deployDir)) {
    New-Item -ItemType Directory -Path $deployDir -Force | Out-Null
}

# --- Locate tooling (framework 4.x csc.exe / RegAsm.exe) ----------------------
$csc = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

$regasm = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $csc -or -not $regasm) {
    Write-Host "  SKIP: .NET Framework 4.x csc.exe/RegAsm.exe not found - no COM ribbon add-in." -ForegroundColor Yellow
    exit 0
}

# --- Locate canonical interop assemblies in the GAC ---------------------------
# Extensibility (IDTExtensibility2) - plain GAC folder; OFFICE.DLL
# (IRibbonExtensibility, IRibbonUI, IRibbonControl) - GAC_MSIL. Both are
# registered into the GAC by every Office install. We match by file name and
# take the first hit, which keeps this working across Office version numbers.
$extGac = Get-ChildItem "C:\Windows\assembly\GAC\Extensibility" -Filter "extensibility.dll" -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
$officeGac = Get-ChildItem "C:\Windows\assembly\GAC_MSIL\office" -Filter "OFFICE.DLL" -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $extGac -or -not $officeGac) {
    Write-Host "  SKIP: Office interop assemblies not found in GAC (Extensibility.dll / OFFICE.DLL) - no COM ribbon add-in." -ForegroundColor Yellow
    exit 0
}

# --- Compile ------------------------------------------------------------------
# The C# code uses System.Drawing + System.Windows.Forms (AxHost) to generate
# the colour swatch icons for the gallery - reference them from the .NET
# Framework directory so no GAC probing is needed.
$fwDir = Split-Path $csc
$sysDrawing = Join-Path $fwDir "System.Drawing.dll"
$sysWinForms = Join-Path $fwDir "System.Windows.Forms.dll"

$tmpDll = Join-Path $env:TEMP "ExcelHighlighterRibbon.dll"
if (Test-Path $tmpDll) { Remove-Item $tmpDll -Force }

Write-Host "  Building ExcelHighlighterRibbon.dll (COM ribbon add-in)..."
& $csc /nologo /target:library "/out:$tmpDll" "/r:$extGac" "/r:$officeGac" "/r:$sysDrawing" "/r:$sysWinForms" $sourceFile
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $tmpDll)) {
    Write-Host "  WARNING: COM add-in compile failed (exit $LASTEXITCODE) - the ribbon will be delivered by the .xlam's own customUI part/interface instead." -ForegroundColor Yellow
    exit 0
}
Copy-Item $tmpDll $deployDll -Force
Remove-Item $tmpDll -Force
Write-Host "  Deployed ExcelHighlighterRibbon.dll ($((Get-Item $deployDll).Length) bytes) to $deployDir" -ForegroundColor Green

# --- Register via RegAsm regfile, rewritten from HKCR to HKCU\Software\Classes --
$regFile  = Join-Path $env:TEMP "eh-ribbon.reg"
$hkcuFile = Join-Path $env:TEMP "eh-ribbon-hkcu.reg"
Remove-Item $regFile, $hkcuFile -Force -ErrorAction SilentlyContinue

& $regasm /regfile:$regFile /codebase $deployDll 2>&1 | Out-Null
if (-not (Test-Path $regFile)) {
    Write-Host "  WARNING: RegAsm could not generate registration - no COM ribbon add-in." -ForegroundColor Yellow
    exit 0
}
$content = (Get-Content $regFile -Raw) -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
[System.IO.File]::WriteAllText($hkcuFile, $content, (New-Object System.Text.ASCIIEncoding))
& reg import $hkcuFile | Out-Null
Write-Host "  Registered COM add-in (HKCU)." -ForegroundColor Green

# --- Excel COM add-in keys + LoadBehavior=3 (load at startup) -----------------
# Register under BOTH the version-independent path and the 16.0 path: some
# Excel builds read one, some the other; the ProgID is the same either way.
foreach ($base in @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins"
)) {
    $k = Join-Path $base "ExcelHighlighter.Ribbon"
    New-Item -Path $k -Force | Out-Null
    Set-ItemProperty -Path $k -Name "FriendlyName" -Value "Excel Highlighter Ribbon"
    Set-ItemProperty -Path $k -Name "Description" -Value "Ribbon delivery for the Excel Highlighter add-in"
    Set-ItemProperty -Path $k -Name "LoadBehavior" -Value 3 -Type DWord
}
Write-Host "  Set Excel COM add-in keys (LoadBehavior=3)." -ForegroundColor Green

# --- Clear any Resiliency\DisabledItems entry for the COM add-in --------------
foreach ($ver in @("16.0", "15.0", "14.0")) {
    $dp = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $dp) {
        Get-ItemProperty $dp -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSObject.Properties } |
            Where-Object { $_.Name -notlike "PS*" -and $_.Value -is [byte[]] } |
            ForEach-Object {
                $t = [System.Text.Encoding]::Unicode.GetString($_.Value)
                if ($t -match "ExcelHighlighter|6E2F4A11") {
                    Remove-ItemProperty -Path $dp -Name $_.Name -Force -ErrorAction SilentlyContinue
                }
            }
    }
}
Write-Host "  Cleared stale Disabled Items entries." -ForegroundColor Green
