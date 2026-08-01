#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Installs the Sandeep Khadka code-signing certificate into Trusted Root.
  Must be run as Administrator.
#>

$ErrorActionPreference = "Stop"

# Find the certificate in the personal store
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -like "CN=Sandeep Khadka*" -and $_.NotAfter -gt (Get-Date) } |
    Select-Object -First 1

if (-not $cert) {
    Write-Host "ERROR: No Sandeep Khadka certificate found in your personal store."
    exit 1
}

Write-Host "Found certificate: $($cert.Subject)"
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Expires: $($cert.NotAfter)"

# Check if already in Trusted Root
$existing = Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint }

if ($existing) {
    Write-Host "`nCertificate is ALREADY in Trusted Root store. Nothing to do."
    exit 0
}

# Export to temp file, then import to Trusted Root
$exportPath = Join-Path $env:TEMP "sandeep-khadka.cer"
Export-Certificate -Cert $cert -FilePath $exportPath -Force
Write-Host "Exported to: $exportPath"

Import-Certificate -FilePath $exportPath -CertStoreLocation Cert:\CurrentUser\Root -ErrorAction Stop
Write-Host "`nSUCCESS: Certificate imported to Trusted Root Certification Authorities."

# Verify
$verify = Get-ChildItem Cert:\CurrentUser\Root |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($verify) {
    Write-Host "Verified: $($verify.Subject) is now in Trusted Root."
} else {
    Write-Host "WARNING: Could not verify import."
}

# Clean up
Remove-Item $exportPath -Force -ErrorAction SilentlyContinue
