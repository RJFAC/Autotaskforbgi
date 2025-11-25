# =============================================================================
# AutoTask 專案淨化與發布工具 V2.1 (修復 Git 身分與上傳問題)
# =============================================================================
$SourceDir = "C:\AutoTask"
$DestDir   = "C:\AutoTask_Public"
$MyUser    = [System.Environment]::UserName
$DateStr   = Get-Date -Format "yyyy-MM-dd HH:mm"

# 確保 console 編碼支援中文
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=== AutoTask 自動發布工具 ===" -ForegroundColor Cyan
Write-Host "來源: $SourceDir"
Write-Host "目標: $DestDir"
Write-Host "偵測敏感使用者名稱: $MyUser"
Write-Host ""

# --- 1. 準備目標目錄 ---
if (-not (Test-Path $DestDir)) {
    New-Item -Path $DestDir -ItemType Directory -Force | Out-Null
    Write-Host "[1/5] 建立目標目錄..." -ForegroundColor Green
} else {
    Write-Host "[1/5] 清理舊檔案 (保留 .git)..." -ForegroundColor Green
    Get-ChildItem -Path $DestDir -Exclude ".git" | Remove-Item -Recurse -Force
}

# --- 2. 複製檔案 (排除清單) ---
Write-Host "[2/5] 複製專案檔案..." -ForegroundColor Green

# 複製根目錄檔案
Copy-Item "$SourceDir\*.bat" "$DestDir" -Force

# 複製 Scripts
New-Item "$DestDir\Scripts" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Scripts\*.ps1" "$DestDir\Scripts" -Recurse

# 複製 Configs (部分)
New-Item "$DestDir\Configs" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Configs\WeeklyConfig.json" "$DestDir\Configs"
# 建立假 Webhook
"https://discord.com/api/webhooks/YOUR_ID/YOUR_TOKEN" | Set-Content "$DestDir\Configs\Webhook.url" -Encoding UTF8

# --- 3. 敏感資料脫敏 ---
Write-Host "[3/5] 執行代碼脫敏 (移除個資)..." -ForegroundColor Green

$ScriptFiles = Get-ChildItem "$DestDir\Scripts\*.ps1"
foreach ($file in $ScriptFiles) {
    $content = Get-Content $file.FullName -Raw -Encoding UTF8
    $isModified = $false

    # 替換使用者名稱路徑
    if ($content -match [regex]::Escape("C:\Users\$MyUser")) {
        $content = $content -replace [regex]::Escape("C:\Users\$MyUser"), "%USERPROFILE%"
        $isModified = $true
    }

    if ($isModified) {
        Set-Content $file.FullName -Value $content -Encoding UTF8
        Write-Host "  - 已淨化: $($file.Name)" -ForegroundColor Gray
    }
}

# --- 4. 建立 .gitignore ---
Write-Host "[4/5] 產生 .gitignore..." -ForegroundColor Green
$GitIgnore = @"
# Logs & Status
Logs/
LogBackups/
Flags/
Configs/*.log
Configs/DateConfig.map
Configs/TaskStatus.json

# User Configs
Configs/Webhook.url
"@
Set-Content "$DestDir\.gitignore" -Value $GitIgnore -Encoding UTF8

# --- 5. Git 同步 ---
Write-Host "[5/5] 執行 Git 同步..." -ForegroundColor Cyan
Set-Location $DestDir

# 檢查是否已初始化
if (-not (Test-Path "$DestDir\.git")) {
    Write-Host "偵測到首次發布，正在初始化 Git..." -ForegroundColor Yellow
    git init
    git branch -M main
}

# [關鍵修正] 檢查並設定 Git 身分 (僅針對此倉庫設定，不影響全域)
$CheckName = git config user.name
$CheckEmail = git config user.email

if ([string]::IsNullOrWhiteSpace($CheckName) -or [string]::IsNullOrWhiteSpace($CheckEmail)) {
    Write-Host "⚠️  Git 尚未設定使用者身分 (導致 Commit 失敗的原因)。" -ForegroundColor Yellow
    Write-Host "請輸入您要在 GitHub 上顯示的名字與 Email (僅用於此專案紀錄)"
    
    $InputName = Read-Host "請輸入 Name (直接按 Enter 使用 'AutoTask Bot')"
    if ([string]::IsNullOrWhiteSpace($InputName)) { $InputName = "AutoTask Bot" }
    
    $InputEmail = Read-Host "請輸入 Email (直接按 Enter 使用 'bot@autotask.local')"
    if ([string]::IsNullOrWhiteSpace($InputEmail)) { $InputEmail = "bot@autotask.local" }

    git config user.name "$InputName"
    git config user.email "$InputEmail"
    Write-Host "✅ 已設定 Git 身分: $InputName <$InputEmail>" -ForegroundColor Green
}

# 詢問或確認 Remote URL
$CurrentRemote = git remote get-url origin 2>$null
if (-not $CurrentRemote) {
    Write-Host "尚未設定 GitHub 倉庫網址。" -ForegroundColor Yellow
    $RemoteUrl = Read-Host "請輸入 GitHub 公開倉庫網址 (https://...)"
    if (-not [string]::IsNullOrWhiteSpace($RemoteUrl)) {
        git remote add origin $RemoteUrl
    } else {
        Write-Host "⚠️ 未輸入網址，僅完成本地打包。" -ForegroundColor Red
        Read-Host "按 Enter 結束"; exit
    }
} else {
    Write-Host "使用現有倉庫: $CurrentRemote" -ForegroundColor Gray
}

# 執行 Git 操作
try {
    Write-Host "正在加入檔案 (git add)..."
    git add .
    
    Write-Host "正在提交變更 (git commit)..."
    git commit -m "Release Update $DateStr"
    
    Write-Host "正在推送至 GitHub (git push)..." -ForegroundColor Cyan
    # 嘗試標準推送
    $pushOutput = git push -u origin main 2>&1
    
    # 檢查是否因為非空倉庫被拒絕 (error: failed to push some refs)
    if ($LASTEXITCODE -ne 0) {
        Write-Host "⚠️ 推送被拒絕 (可能是因為倉庫非空或歷史衝突)。" -ForegroundColor Yellow
        $force = Read-Host "是否嘗試強制覆蓋遠端倉庫？(Y/N)"
        if ($force -eq "Y") {
            Write-Host "正在執行強制推送 (git push -f)..." -ForegroundColor Magenta
            git push -u origin main -f
        } else {
            Write-Host "推送已取消。請手動解決衝突。" -ForegroundColor Red
        }
    } else {
        Write-Host "✅ 發布成功！" -ForegroundColor Green
    }

} catch {
    Write-Host "發生未預期的錯誤: $_" -ForegroundColor Red
}

Write-Host "`n作業結束。"
Read-Host "請按 Enter 鍵關閉視窗 (請檢查上方是否有錯誤訊息)..."