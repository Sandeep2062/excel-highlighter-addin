$ErrorActionPreference = 'Continue'

Write-Host "=== UI extensibility / add-in security settings ==="
$keys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Common\Security",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Security",
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trust Center",
    "HKLM:\Software\Microsoft\Office\16.0\Common\Security",
    "HKLM:\Software\Microsoft\Office\16.0\Excel\Security",
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Security",
    "HKLM:\Software\Policies\Microsoft\Office\16.0\Common\Security",
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security",
    "HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Security"
)
foreach ($k in $keys) {
    Write-Host "`n--- $k ---"
    if (Test-Path $k) {
        $props = Get-ItemProperty $k -ErrorAction SilentlyContinue
        if ($props) {
            $props.psobject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                Write-Host "  $($_.Name) = $($_.Value)"
            }
        } else {
            Write-Host "  (empty)"
        }
    } else {
        Write-Host "  (not present)"
    }
}

Write-Host "`n=== DisableUIExtensibility anywhere ==="
foreach ($root in @('HKCU:','HKLM:')) {
    foreach ($base in @("$root\Software\Microsoft\Office\16.0\Common", "$root\Software\Microsoft\Office\16.0\Excel", "$root\Software\Policies\Microsoft\Office\16.0\Common", "$root\Software\Policies\Microsoft\Office\16.0\Excel")) {
        if (Test-Path $base) {
            Get-ChildItem $base -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                $item = $_
                try {
                    $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                    if ($props) {
                        $props.psobject.Properties | Where-Object { $_.Name -like '*UI*' -or $_.Name -like '*Addin*' -or $_.Name -like '*Extensib*' } | ForEach-Object {
                            Write-Host "  $($item.PSPath)  $($_.Name) = $($_.Value)"
                        }
                    }
                } catch {}
            }
        }
    }
}
Write-Host "`n=== done ==="
