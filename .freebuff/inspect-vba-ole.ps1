param(
    [string]$Path
)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Host "=== vbaProject.bin OLE stream inspection ===" -ForegroundColor Cyan
Write-Host "File: $Path"

$z = [System.IO.Compression.ZipFile]::OpenRead($Path)
$vbaEntry = $z.GetEntry("xl/vbaProject.bin")

if (-not $vbaEntry) {
    Write-Host "vbaProject.bin not found!" -ForegroundColor Red
    $z.Dispose()
    exit 1
}

Write-Host "vbaProject.bin size: $($vbaEntry.Length) bytes"

# Extract vbaProject.bin to temp
$tempBin = Join-Path $env:TEMP "vbaProject-check.bin"
$stream = $vbaEntry.Open()
$fileStream = [System.IO.File]::Create($tempBin)
$stream.CopyTo($fileStream)
$fileStream.Close()
$stream.Close()
$z.Dispose()

Write-Host "Extracted to: $tempBin"

# Use structured storage to read OLE streams
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class OleStorage {
    [DllImport("ole32.dll")]
    public static extern int StgOpenStorage(
        [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
        IntPtr pstgPriority,
        uint grfMode,
        IntPtr snbExclude,
        uint reserved,
        out IntPtr ppstgOpen);

    [DllImport("ole32.dll")]
    public static extern int StgIsStorageFile(
        [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
}
"@ -ErrorAction SilentlyContinue
} catch {}

# Simpler approach - use ComObject for structured storage
Write-Host "`n=== Attempting OLE stream enumeration ===" -ForegroundColor Yellow

# Try using the Office tools approach
$excel = $null
$wb = $null
try {
    # Create a temporary xlsx with the vbaProject.bin embedded
    $tempXlsx = Join-Path $env:TEMP "vba-inspect-temp.xlsx"
    Copy-Item $Path $tempXlsx -Force
    
    # Re-open as zip and check for known signature streams
    $z2 = [System.IO.Compression.ZipFile]::OpenRead($tempXlsx)
    $vba2 = $z2.GetEntry("xl/vbaProject.bin")
    $ms = New-Object System.IO.MemoryStream
    $vba2.Open().CopyTo($ms)
    $bytes = $ms.ToArray()
    $ms.Close()
    $z2.Dispose()
    
    # Scan for ASCII strings in the binary
    Write-Host "`n=== Scanning vbaProject.bin for signature-related strings ===" -ForegroundColor Yellow
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
    
    $patterns = @(
        "PKI#X509",
        "MSPST",
        "XmlSignature",
        "Authenticode",
        "DigitalSignature",
        "VBA签",  # Chinese signature marker
        "SignedData",
        "OBJECT SIGNATURE",
        "Contents",
        "_VBA_PROJECT",
        "VBA",
        "PROJECT",
        "PROJECTwm",
        "Module",
        "Class",
        "ThisWorkbook",
        "Sheet1",
        "Signature",
        "Certificate"
    )
    
    foreach ($p in $patterns) {
        $idx = $ascii.IndexOf($p)
        if ($idx -ge 0) {
            $context = $ascii.Substring([Math]::Max(0, $idx - 20), [Math]::Min(60, $ascii.Length - [Math]::Max(0, $idx - 20)))
            $context = $context -replace '[^\x20-\x7E]', '.'
            Write-Host "  FOUND '$p' at offset $idx : ...$context..."
        }
    }
    
    # List all VBA project streams by looking for known stream names
    Write-Host "`n=== Looking for standard VBA streams ===" -ForegroundColor Yellow
    $streamNames = @("_VBA_PROJECT", "PROJECT", "PROJECTwm", "VBA dir", "Module", "Class", "ThisWorkbook", "Sheet1")
    foreach ($s in $streamNames) {
        $idx = $ascii.IndexOf($s)
        if ($idx -ge 0) {
            Write-Host "  Found stream name: '$s' at offset $idx"
        }
    }
    
    # Check file size comparison with a typical clean vbaProject
    Write-Host "`n=== Size analysis ===" -ForegroundColor Yellow
    Write-Host "  Current size: $($bytes.Length) bytes"
    Write-Host "  Typical clean VBA project (13 modules): ~80-150 KB"
    if ($bytes.Length -gt 200000) {
        Write-Host "  WARNING: Size is unusually large - possible embedded signature data" -ForegroundColor Red
    }
    
    # Count OLE header sectors
    Write-Host "`n=== OLE compound header ===" -ForegroundColor Yellow
    $sectorSize = 512
    $minorVer = [BitConverter]::ToUInt16($bytes, 24)
    $majorVer = [BitConverter]::ToUInt16($bytes, 26)
    $byteOrder = [BitConverter]::ToUInt16($bytes, 28)
    $sectorPow = [BitConverter]::ToUInt16($bytes, 30)
    $miniSectorPow = [BitConverter]::ToUInt16($bytes, 32)
    $totalSectors = [BitConverter]::ToUInt32($bytes, 44)
    $totalMiniSectors = [BitConverter]::ToUInt32($bytes, 64)
    Write-Host "  Major version: $majorVer, Minor version: $minorVer"
    Write-Host "  Byte order: $byteOrder (0=LE)"
    Write-Host "  Sector size power: $sectorPow (size=$((1 -shl $sectorPow)))"
    Write-Host "  Total sectors: $totalSectors"
    Write-Host "  Total mini-stream sectors: $totalMiniSectors"
    
    Remove-Item $tempXlsx -Force -ErrorAction SilentlyContinue
    
} finally {
    Remove-Item $tempBin -Force -ErrorAction SilentlyContinue
}

Write-Host "`nDone." -ForegroundColor Cyan
