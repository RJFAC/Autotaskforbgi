@echo off
title 強制停止 AutoTask
color 0C
echo 正在呼叫 PowerShell 執行清理...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\AutoTask\Scripts\StopAll.ps1"

echo.
pause