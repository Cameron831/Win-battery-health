# Check for admin
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Please run this script as Administrator."
    exit
}

# Generate  battery report
Write-Host "Generating battery report..."
$reportPath = Join-Path $PSScriptRoot "battery-report.html"
powercfg /batteryreport | Out-Null

# Wait for file to be created (5 second timeout)
$timeout = 0
while (!(Test-Path $reportPath) -and ($timeout -lt 50)) {
    Start-Sleep -Milliseconds 100
    $timeout++
}

if (!(Test-Path $reportPath)) {
    Write-Error "Battery report did not generate in time."
    exit
}

Write-Host "Parsing report..."

# Load the HTML
$html = Get-Content $reportPath -Raw

# Find the matches
$designMatch = [regex]::Match($html, '<span class="label">\s*DESIGN CAPACITY\s*</span></td><td>([\d,]+) mWh', 'IgnoreCase')
$fullMatch   = [regex]::Match($html, '<span class="label">\s*FULL CHARGE CAPACITY\s*</span></td><td>([\d,]+) mWh', 'IgnoreCase')

# Extract numerical value and clean
$designCapacity = $designMatch.Groups[1].Value.Replace(",", "") -as [int]
$fullChargeCapacity = $fullMatch.Groups[1].Value.Replace(",", "") -as [int]

Write-Host "Calculating battery health..."

# Calculate percentage and print results
if ($designCapacity -and $fullChargeCapacity) {
    $healthPercent = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 2)
    $color = if ($healthPercent -ge 80) { 'Green' } else { 'Red' }
    Write-Host ("`nBattery Health: {0}% ({1} / {2} mWh)" -f $healthPercent, $fullChargeCapacity, $designCapacity) -ForegroundColor $color
} else {
    Write-Warning "Could not extract battery capacities from the report."
}

Remove-Item $reportPath -Force
