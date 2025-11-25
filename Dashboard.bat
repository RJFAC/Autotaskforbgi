@echo off
title 自動化任務儀表板 (AutoTask Dashboard)
mode con: cols=100 lines=40
color 0B
echo 正在啟動儀表板...
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\AutoTask\Scripts\Dashboard.ps1"