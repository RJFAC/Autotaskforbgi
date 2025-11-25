@echo off
:: 使用 start /min 讓黑窗一開始就是最小化狀態，隨後 PowerShell 會將其完全隱藏
start "" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\AutoTask\Scripts\Verify-VersionDates.ps1"