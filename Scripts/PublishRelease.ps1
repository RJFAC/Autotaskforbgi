# =============================================================================
# AutoTask 專案淨化與發布工具 V2.4 (新增 URL 生成與存檔)
# =============================================================================
$SourceDir = "C:\AutoTask"
$DestDir   = "C:\AutoTask_Public"
$ConfigsDir = "$SourceDir\Configs"
$HashFile   = "$ConfigsDir\ScriptHash.txt"
$UrlLogFile = "$ConfigsDir\GitHub_Raw_Links.txt" # [新增] 網址存檔路徑
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

# --- 3. 敏感資料脫敏 (符合 SSOT Guideline Sec 6) ---
Write-Host "[3/5] 執行深度隱私淨化..." -ForegroundColor Green

# 定義敏感資料清單 (請根據實際情況擴充)
$SensitiveMap = @{
    [regex]::Escape("C:\Users\$MyUser") = "%USERPROFILE%"
    "[USER_NAME]"                           = "[USER_NAME]"
    "[MACHINE_ID]"                  = "[MACHINE_ID]"
    # 若有特定的 SID 也應加入
    # "S-1-5-21-..."                   = "[SID_REMOVED]"
}

$ScriptFiles = Get-ChildItem "$DestDir\Scripts\*.ps1", "$DestDir\Configs\*.json"
# 若未來加入 XML 複製，請取消註解下方
# $ScriptFiles += Get-ChildItem "$DestDir\*.xml"

foreach ($file in $ScriptFiles) {
    $content = Get-Content $file.FullName -Raw -Encoding UTF8
    $isModified = $false
    
    foreach ($key in $SensitiveMap.Keys) {
        if ($content -match $key) {
            # 針對 Regex 類型的 Key 不需要再 Escape，但普通字串需要注意
            # 這裡簡化處理，假設 Key 已經是 Regex Safe 或是純文字
            $content = $content -replace $key, $SensitiveMap[$key]
            $isModified = $true
            Write-Host "  [SEC] 在 $($file.Name) 中淨化了敏感字串: $key" -ForegroundColor Yellow
        }
    }
    
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
Configs/*.map
Configs/*.json
Configs/Webhook.url
Configs/ScriptHash.txt
Configs/GitHub_Raw_Links.txt
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
    
    # --- [新增] 自動生成並存檔 Raw 網址 ---
    Write-Host "`n[INFO] 正在生成全腳本 Raw 網址..." -ForegroundColor Cyan
    $UrlList = @()
    $Remote = git remote get-url origin
    if ($Remote -match "github\.com[:/](?<U>.+?)/(?<R>.+?)(\.git)?$") {
        $User = $Matches.U; $Repo = $Matches.R
        $Sha = git rev-parse HEAD
        
        $Header = "=== GitHub Raw Links (版本: $($Sha.Substring(0,7)) | 時間: $DateStr) ==="
        $UrlList += $Header
        Write-Host $Header -ForegroundColor Yellow

        # 掃描目前的目錄 ($DestDir) 獲取所有 PS1
        Get-ChildItem -Path . -Filter "*.ps1" -Recurse | Sort-Object Name | ForEach-Object {
            $RelPath = $_.FullName.Substring($PWD.Path.Length + 1).Replace("\", "/")
            $Url = "https://raw.githubusercontent.com/$User/$Repo/$Sha/$RelPath"
            
            $Entry = "$RelPath`n$Url"
            $UrlList += $Entry
            Write-Host $Entry
            $UrlList += "----------------------------------------"
        }
        
        # 寫入檔案
        $UrlList | Set-Content $UrlLogFile -Encoding UTF8
        Write-Host "`n📄 網址清單已儲存至: $UrlLogFile" -ForegroundColor Green
    }

    # 更新版本雜湊紀錄 (防止 Dashboard 報錯)
    Write-Host "正在更新版本雜湊紀錄..."
    $CurrentHash = ""
    # 注意：這裡指回 SourceDir 確保計算的是原始腳本
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

