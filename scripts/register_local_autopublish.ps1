[CmdletBinding()]
param(
    [string]$TaskName = "Liseo Dashboard Auto Publish",
    [string]$WorkbookPath = "..\\..\\..\\Liseo Assemblies Master Sheet.xlsm"
)

$ErrorActionPreference = "Stop"

$scriptPath = (Resolve-Path (Join-Path $PSScriptRoot "publish_dashboard_data.py")).Path
$workbookFullPath = (Resolve-Path (Join-Path $PSScriptRoot $WorkbookPath)).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$triggerTime = (Get-Date).AddMinutes(1)
$taskArgs = '"' + $scriptPath + '" --workbook "' + $workbookFullPath + '"'
$action = New-ScheduledTaskAction -Execute $pythonPath -Argument $taskArgs
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Publishes the Liseo dashboard from the local workbook every minute." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $workbookFullPath"
Write-Host "  Runs every minute on this PC."
