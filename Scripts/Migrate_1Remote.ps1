# =============================================================================
# AutoTask 1Remote 遷移工具
# =============================================================================
$SourceDir = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64"
$DestDir   = "C:\AutoTask\1Remote"
$EnvConf   = "C:\AutoTask\Configs\EnvConfig.json"

Write-Host "正在遷移 1Remote..." -ForegroundColor Cyan

# 1. 檢查來源
if (-not (Test-Path "$SourceDir\1Remote.exe")) {
    Write-Error "找不到來源 1Remote，請確認路徑: $SourceDir"
    pause; exit
}

# 2. 停止正在運行的 1Remote
Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
Start-Sleep 2

# 3. 複製檔案 (使用 Robocopy 以保留屬性並處理大量檔案)
if (-not (Test-Path $DestDir)) { New-Item -Path $DestDir -ItemType Directory | Out-Null }
Robocopy $SourceDir $DestDir /E /MOVE /NFL /NDL /NJH /NJS

# 4. 確保是 Portable 模式 (檢查資料庫是否存在)
# 1Remote 的資料庫通常是 1Remote.db。如果之前是在 AppData，可能需要手動複製過來。
# 但既然您說它是 Zip 解壓版，通常資料庫就在目錄下。
if (-not (Test-Path "$DestDir\1Remote.db")) {
    Write-Warning "警告：在目錄中未發現 1Remote.db。"
    Write-Warning "若您的設定檔在 AppData，請手動將其複製到 $DestDir 以啟用可攜帶模式。"
}

# 5. 更新 AutoTask 環境設定 (EnvConfig.json)
if (Test-Path $EnvConf) {
    $json = Get-Content $EnvConf -Raw -Encoding UTF8 | ConvertFrom-Json
    # 更新路徑
    $json | Add-Member -Name "Path1Remote" -Value "$DestDir\1Remote.exe" -MemberType NoteProperty -Force
    $json | ConvertTo-Json -Depth 4 | Set-Content $EnvConf -Encoding UTF8
    Write-Host "已更新 EnvConfig.json 指向新位置。" -ForegroundColor Green
}

Write-Host "遷移完成！新路徑: $DestDir" -ForegroundColor Green
pause
