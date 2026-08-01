param(
    [string]$CustomUIPath,
    [string]$BasPath
)

$ErrorActionPreference = 'Continue'

Write-Host "=== Checking customUI callbacks against VBA module ===" -ForegroundColor Cyan

# Read customUI XML
$customUI = Get-Content $CustomUIPath -Raw

# Extract all callback references (functionName="Module.Function")
$pattern = '(\w+)="(RibbonCallbacks\.(\w+))"'
$matches = [regex]::Matches($customUI, $pattern)

Write-Host "`nFound $($matches.Count) callback references in customUI14.xml:" -ForegroundColor Yellow
$callbackNames = @()
foreach ($m in $matches) {
    $attrName = $m.Groups[1].Value
    $fullRef = $m.Groups[2].Value
    $funcName = $m.Groups[3].Value
    $callbackNames += $funcName
    Write-Host "  $attrName -> $funcName"
}

# Read VBA module
$basContent = Get-Content $BasPath -Raw

# Extract all Public Sub/Function names from the module
$funcPattern = '(?:Public\s+(?:Sub|Function)\s+|Sub\s+|Function\s+)(\w+)'
$funcMatches = [regex]::Matches($basContent, $funcPattern)
$vbaFunctions = @{}
foreach ($fm in $funcMatches) {
    $vbaFunctions[$fm.Groups[1].Value] = $true
}

Write-Host "`nFound $($vbaFunctions.Count) public functions in RibbonCallbacks.bas:" -ForegroundColor Yellow
$vbaFunctions.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

# Check each callback
Write-Host "`n=== Callback verification ===" -ForegroundColor Cyan
$missing = @()
$found = @()
$uniqueCallbacks = $callbackNames | Sort-Object -Unique

foreach ($cb in $uniqueCallbacks) {
    if ($vbaFunctions.ContainsKey($cb)) {
        $found += $cb
        Write-Host "  OK: $cb" -ForegroundColor Green
    } else {
        $missing += $cb
        Write-Host "  MISSING: $cb" -ForegroundColor Red
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "  Found: $($found.Count)" -ForegroundColor Green
Write-Host "  Missing: $($missing.Count)" -ForegroundColor $(if ($missing.Count -gt 0) { "Red" } else { "Green" })

if ($missing.Count -gt 0) {
    Write-Host "`nMissing callbacks:" -ForegroundColor Red
    foreach ($m in $missing) {
        # Find the attribute that references it
        $attrPattern = '(\w+)="RibbonCallbacks\.' + [regex]::Escape($m) + '"'
        $attrMatch = [regex]::Match($customUI, $attrPattern)
        $attrName = if ($attrMatch.Success) { $attrMatch.Groups[1].Value } else { "unknown" }
        Write-Host "  $m (referenced by $attrName attribute)" -ForegroundColor Red
    }
}
