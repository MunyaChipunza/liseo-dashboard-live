param(
    [string]$TaskName = "Liseo Dashboard Auto Publish"
)

schtasks /Delete /F /TN $TaskName | Out-Null
Write-Host "Scheduled task removed: $TaskName"
