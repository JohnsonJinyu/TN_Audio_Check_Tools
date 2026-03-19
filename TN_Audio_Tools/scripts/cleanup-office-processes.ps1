param(
  [switch]$IncludeWps,
  [switch]$Force
)

$targets = @('EXCEL', 'et')
if ($IncludeWps) {
  $targets += 'wps'
}

$processes = Get-Process | Where-Object {
  ($targets -contains $_.ProcessName) -and [string]::IsNullOrWhiteSpace($_.MainWindowTitle)
}

$count = @($processes).Count
if ($count -eq 0) {
  Write-Host 'No hidden Office processes found.'
  exit 0
}

if (-not $Force) {
  Write-Host "Found $count hidden Office processes. Re-run with -Force to terminate."
  $processes | Select-Object ProcessName, Id, StartTime | Format-Table -AutoSize
  exit 0
}

$processes | Stop-Process -Force
Write-Host "Killed $count hidden Office processes."
