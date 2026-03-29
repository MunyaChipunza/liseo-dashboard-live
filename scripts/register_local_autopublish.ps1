[CmdletBinding()]
param(
    [string]$TaskName = "Liseo Dashboard Auto Publish",
    [string]$WorkbookPath = "..\\..\\..\\Liseo Assemblies Master Sheet.xlsm"
)

$ErrorActionPreference = "Stop"

$scriptPath = (Resolve-Path (Join-Path $PSScriptRoot "publish_dashboard_data.py")).Path
$workbookFullPath = (Resolve-Path (Join-Path $PSScriptRoot $WorkbookPath)).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$startTime = (Get-Date).AddMinutes(1).ToString("HH:mm")
$taskCommand = "`"$pythonPath`" `"$scriptPath`" --workbook `"$workbookFullPath`""

schtasks /Create /F /SC MINUTE /MO 1 /ST $startTime /TN $TaskName /TR $taskCommand | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Could not create the scheduled task."
}

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $workbookFullPath"
Write-Host "  Runs every minute on this PC."
