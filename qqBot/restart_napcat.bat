@echo off
setlocal

set "TASK_NAME=NotesDepartmentNapCatRestart"

schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    echo [Watchdog] Scheduled task is missing.
    echo [Watchdog] Run install_napcat_restart_task.ps1 once first.
    exit /b 5
)

schtasks /Run /TN "%TASK_NAME%" >nul 2>&1
if errorlevel 1 (
    echo [Watchdog] Failed to start scheduled task.
    exit /b 6
)

echo [Watchdog] NapCat restart task triggered.
exit /b 0
