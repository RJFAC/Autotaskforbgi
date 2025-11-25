# =============================================================================
# AutoTask 專案淨化與發布工具 V2.2 (含部署腳本)
# =============================================================================
$SourceDir = "C:\AutoTask"
$DestDir   = "C:\AutoTask_Public"
$MyUser    = [System.Environment]::UserName
$DateStr   = Get-Date -Format "yyyy-MM-dd HH:mm"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=== AutoTask 自動發布工具 ===" -ForegroundColor Cyan
Write-Host "來源: $SourceDir"
Write-Host "目標: $DestDir"
Write-Host ""

# --- 1. 準備目標目錄 ---
if (-not (Test-Path $DestDir)) {
    New-Item -Path $DestDir -ItemType Directory -Force | Out-Null
} else {
    Get-ChildItem -Path $DestDir -Exclude ".git" | Remove-Item -Recurse -Force
}

# --- 2. 複製檔案 ---
Write-Host "[複製] 專案核心檔案..." -ForegroundColor Green

# 複製根目錄的 .bat (包含 Dashboard.bat, STOP_ALL.bat, OneKeyDeploy.bat)
Copy-Item "$SourceDir\*.bat" "$DestDir" -Force

# [關鍵新增] 複製部署用的 PowerShell 腳本
if (Test-Path "$SourceDir\DeployFromGit.ps1") {
    Copy-Item "$SourceDir\DeployFromGit.ps1" "$DestDir" -Force
    Write-Host "  - 已加入部署腳本 (DeployFromGit.ps1)" -ForegroundColor Gray
}

# 複製 Scripts
New-Item "$DestDir\Scripts" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Scripts\*.ps1" "$DestDir\Scripts" -Recurse

# 複製 Configs
New-Item "$DestDir\Configs" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Configs\WeeklyConfig.json" "$DestDir\Configs"
"https://discord.com/api/webhooks/YOUR_ID/YOUR_TOKEN" | Set-Content "$DestDir\Configs\Webhook.url" -Encoding UTF8

# --- 3. 敏感資料脫敏 ---
Write-Host "[淨化] 移除敏感個資..." -ForegroundColor Green

$FilesToClean = Get-ChildItem "$DestDir\Scripts\*.ps1"
if (Test-Path "$DestDir\DeployFromGit.ps1") { $FilesToClean += Get-Item "$DestDir\DeployFromGit.ps1" }

foreach ($file in $FilesToClean) {
    $content = Get-Content $file.FullName -Raw -Encoding UTF8
    $isModified = $false

    if ($content -match [regex]::Escape("C:\Users\$MyUser")) {
        $content = $content -replace [regex]::Escape("C:\Users\$MyUser"), "%USERPROFILE%"
        $isModified = $true
    }
    
    # 如果 DeployFromGit.ps1 裡面有寫死您的 Repo URL，這裡可以選擇是否替換
    # 但通常那是公開 URL，所以保留即可

    if ($isModified) {
        Set-Content $file.FullName -Value $content -Encoding UTF8
    }
}

# --- 4. 建立 .gitignore ---
$GitIgnore = @"
Logs/
LogBackups/
Flags/
Configs/*.log
Configs/DateConfig.map
Configs/TaskStatus.json
Configs/Webhook.url
"@
Set-Content "$DestDir\.gitignore" -Value $GitIgnore -Encoding UTF8

# --- 5. Git 同步 ---
Write-Host "[同步] 推送至 GitHub..." -ForegroundColor Cyan
Set-Location $DestDir

if (-not (Test-Path "$DestDir\.git")) {
    git init
    git branch -M main
}

$CheckName = git config user.name
if ([string]::IsNullOrWhiteSpace($CheckName)) {
    git config user.name "AutoTask Bot"
    git config user.email "bot@autotask.local"
}

$CurrentRemote = git remote get-url origin 2>$null
if (-not $CurrentRemote) {
    $RemoteUrl = Read-Host "請輸入 GitHub 公開倉庫網址"
    if (-not [string]::IsNullOrWhiteSpace($RemoteUrl)) {
        git remote add origin $RemoteUrl
    }
}

try {
    git add .
    git commit -m "Update $DateStr (With Deploy Scripts)"
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "推送失敗，嘗試強制覆蓋..." -ForegroundColor Yellow
        git push -u origin main -f
    }
    Write-Host "✅ 發布成功！" -ForegroundColor Green
} catch {
    Write-Host "發生錯誤: $_" -ForegroundColor Red
}

Write-Host "`n作業結束。"
Read-Host "按 Enter 關閉..."