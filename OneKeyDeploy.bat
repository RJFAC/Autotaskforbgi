@echo off
title AutoTask GitHub 部署工具
color 0B
echo.
echo ========================================================
echo        正在啟動 AutoTask 雲端部署/更新程序...
echo ========================================================
echo.

:: 強制以管理員身分執行 PowerShell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0DeployFromGit.ps1""' -Verb RunAs"