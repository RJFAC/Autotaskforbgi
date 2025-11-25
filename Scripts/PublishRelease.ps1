# =============================================================================
# AutoTask 專案淨化與發布工具 V2.3 (新增版本雜湊紀錄)
# =============================================================================
$SourceDir = "C:\AutoTask"
$DestDir   = "C:\AutoTask_Public"
$ConfigsDir = "$SourceDir\Configs"
$HashFile   = "$ConfigsDir\ScriptHash.txt"
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
Write-Host "[2/5] 複製專案檔案..." -ForegroundColor Green
Copy-Item "$SourceDir\*.bat" "$DestDir" -Force
if (Test-Path "$SourceDir\DeployFromGit.ps1") { Copy-Item "$SourceDir\DeployFromGit.ps1" "$DestDir" -Force }

New-Item "$DestDir\Scripts" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Scripts\*.ps1" "$DestDir\Scripts" -Recurse

New-Item "$DestDir\Configs" -ItemType Directory -Force | Out-Null
Copy-Item "$SourceDir\Configs\WeeklyConfig.json" "$DestDir\Configs"
"https://discord.com/api/webhooks/YOUR_ID/YOUR_TOKEN" | Set-Content "$DestDir\Configs\Webhook.url" -Encoding UTF8

# --- 3. 敏感資料脫敏 ---
Write-Host "[3/5] 執行代碼脫敏..." -ForegroundColor Green
$ScriptFiles = Get-ChildItem "$DestDir\Scripts\*.ps1"
foreach ($file in $ScriptFiles) {
    $content = Get-Content $file.FullName -Raw -Encoding UTF8
    $isModified = $false
    if ($content -match [regex]::Escape("C:\Users\$MyUser")) {
        $content = $content -replace [regex]::Escape("C:\Users\$MyUser"), "%USERPROFILE%"
        $isModified = $true
    }
    if ($isModified) { Set-Content $file.FullName -Value $content -Encoding UTF8 }
}

# --- 4. 建立 .gitignore ---
$GitIgnore = @"
Logs/
LogBackups/
Flags/
Configs/*.log
Configs/*.map
Configs/*.json
Configs/Webhook.url
Configs/ScriptHash.txt
"@
Set-Content "$DestDir\.gitignore" -Value $GitIgnore -Encoding UTF8

# --- 5. Git 同步 ---
Write-Host "[5/5] 執行 Git 同步..." -ForegroundColor Cyan
Set-Location $DestDir

if (-not (Test-Path "$DestDir\.git")) { git init; git branch -M main }

$CheckName = git config user.name
if ([string]::IsNullOrWhiteSpace($CheckName)) { git config user.name "AutoTask Bot"; git config user.email "bot@autotask.local" }

$CurrentRemote = git remote get-url origin 2>$null
if (-not $CurrentRemote) {
    $RemoteUrl = Read-Host "請輸入 GitHub 公開倉庫網址"
    if (-not [string]::IsNullOrWhiteSpace($RemoteUrl)) { git remote add origin $RemoteUrl }
}

try {
    git add .
    git commit -m "Update $DateStr"
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "推送失敗，嘗試強制覆蓋..." -ForegroundColor Yellow
        git push -u origin main -f
    }
    Write-Host "✅ 發布成功！" -ForegroundColor Green
    
    # [關鍵新增] 更新本地的版本雜湊紀錄
    Write-Host "正在更新版本雜湊紀錄..."
    $CurrentHash = ""
    Get-ChildItem "$SourceDir\Scripts" -Include "*.ps1", "*.bat" -Recurse | Sort-Object Name | ForEach-Object { 
        $CurrentHash += (Get-FileHash $_.FullName).Hash 
    }
    Set-Content -Path $HashFile -Value $CurrentHash -Force
    Write-Host "雜湊已儲存至: $HashFile" -ForegroundColor Gray

} catch {
    Write-Host "發生錯誤: $_" -ForegroundColor Red
}

Write-Host "`n作業結束。"
Read-Host "按 Enter 關閉..."