$ErrorActionPreference = 'Stop'

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RestartScript = Join-Path $BaseDir 'restart_napcat.ps1'
$TaskName = 'NotesDepartmentNapCatRestart'

if (-not (Test-Path $RestartScript)) {
    Write-Host "[ERROR] Missing file: $RestartScript" -ForegroundColor Red
    exit 2
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principalCheck = New-Object Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principalCheck.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host '[INFO] Administrator permission is required once to install the scheduled task.'
    $quotedScript = '"' + $PSCommandPath + '"'
    Start-Process -FilePath 'powershell.exe' `
        -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File $quotedScript"
    exit 0
}

$userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$escapedRestart = '"' + $RestartScript + '"'

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File $escapedRestart"

$taskPrincipal = New-ScheduledTaskPrincipal `
    -UserId $userId `
    -LogonType Interactive `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Principal $taskPrincipal `
    -Settings $settings `
    -Force | Out-Null

Write-Host ''
Write-Host '[OK] Scheduled task installed.' -ForegroundColor Green
Write-Host "Task name: $TaskName"
Write-Host "User: $userId"
Write-Host 'The watchdog can now trigger this task without showing a UAC prompt each time.'
