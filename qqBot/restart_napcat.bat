@echo off
setlocal
chcp 65001 >nul

set "NAPCAT_DIR=C:\NapCat.Shell"
set "NAPCAT_LAUNCHER=%NAPCAT_DIR%\launcher-win10.bat"

rem 默认不结束 QQ.exe，避免误杀你电脑上的其他 QQ。
rem 如果这台电脑只运行 Bot 的 QQ，并且“不杀 QQ 就无法恢复”，可改为 1。
set "KILL_QQ=0"

if not exist "%NAPCAT_LAUNCHER%" (
    echo [Watchdog] 找不到 NapCat 启动脚本：%NAPCAT_LAUNCHER%
    exit /b 2
)

if "%KILL_QQ%"=="1" (
    echo [Watchdog] 正在结束 QQ.exe ...
    taskkill /F /IM QQ.exe >nul 2>&1
    timeout /t 3 /nobreak >nul
)

echo [Watchdog] 正在启动 NapCat：%NAPCAT_LAUNCHER%
start "NapCat" /D "%NAPCAT_DIR%" cmd /c ""%NAPCAT_LAUNCHER%""
exit /b 0
