# =============================================================================
# AutoTask 雲端部署與升級工具 (DeployFromGit.ps1)
# 功能：從 GitHub 下載最新版，保留本地設定，自動適配路徑
# =============================================================================

$RepoUrl = "https://github.com/RJFAC/Autotaskforbgi.git"
$ZipUrl  = "https://github.com/RJFAC/Autotaskforbgi/archive/refs/heads/main.zip"
$InstallDir = "C:\AutoTask"
$TempDir = "$env:TEMP\AutoTask_Update_$(Get-Random)"

Write-Host "=== AutoTask 部署精靈 ===" -ForegroundColor Cyan

# --- 1. 下載原始碼 (Git 或 ZIP) ---
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -Path $TempDir -ItemType Directory | Out-Null

Write-Host "`n[1/5] 正在下載最新版本..." -ForegroundColor Yellow

if (Get-Command "git" -ErrorAction SilentlyContinue) {
    Write-Host "使用 Git Clone..." -ForegroundColor Gray
    git clone $RepoUrl "$TempDir\Repo" | Out-Null
    $SourceRoot = "$TempDir\Repo"
} else {
    Write-Host "未偵測到 Git，改用 ZIP 下載..." -ForegroundColor Gray
    $ZipFile = "$TempDir\source.zip"
    try {
        Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipFile
        Expand-Archive -Path $ZipFile -DestinationPath $TempDir
        $SourceRoot = "$TempDir\Autotaskforbgi-main"
    } catch {
        Write-Error "下載失敗，請檢查網路連線。"
        pause; exit
    }
}

if (-not (Test-Path "$SourceRoot\Scripts\Master.ps1")) {
    Write-Error "下載的檔案結構不正確，找不到 Master.ps1。"
    pause; exit
}

# --- 2. 讀取舊設定 (如果是更新模式) ---
$IsUpdate = Test-Path $InstallDir
$LocalPaths = @{}

if ($IsUpdate) {
    Write-Host "`n[2/5] 偵測到現有安裝，正在備份本地路徑設定..." -ForegroundColor Cyan
    
    # 嘗試從舊的 Master.ps1 讀取使用者設定的路徑
    $OldMaster = "$InstallDir\Scripts\Master.ps1"
    if (Test-Path $OldMaster) {
        $Content = Get-Content $OldMaster -Raw -Encoding UTF8
        if ($Content -match '\$1RemoteDir\s*=\s*"(.*?)"') { $LocalPaths["1Remote"] = $matches[1] }
    }
    
    # 嘗試從舊的 Payload.ps1 讀取
    $OldPayload = "$InstallDir\Scripts\Payload.ps1"
    if (Test-Path $OldPayload) {
        $Content = Get-Content $OldPayload -Raw -Encoding UTF8
        if ($Content -match '\$BettergiDir\s*=\s*"(.*?)"') { $LocalPaths["BetterGI"] = $matches[1] }
    }
    
    if ($LocalPaths.Count -gt 0) {
        Write-Host "  - 保留 1Remote 路徑: $($LocalPaths['1Remote'])" -ForegroundColor Gray
        Write-Host "  - 保留 BetterGI 路徑: $($LocalPaths['BetterGI'])" -ForegroundColor Gray
    }
} else {
    Write-Host "`n[2/5] 全新安裝模式" -ForegroundColor Cyan
}

# --- 3. 部署檔案 (智慧覆蓋) ---
Write-Host "`n[3/5] 正在部署檔案..." -ForegroundColor Yellow

# 建立結構
$Dirs = @("Scripts", "Configs", "Flags", "Logs", "LogBackups")
foreach ($d in $Dirs) {
    $p = "$InstallDir\$d"
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
}

# (A) 覆蓋腳本 (Scripts/*.ps1)
Copy-Item "$SourceRoot\Scripts\*.ps1" "$InstallDir\Scripts" -Force
# (B) 覆蓋根目錄工具 (*.bat)
Copy-Item "$SourceRoot\*.bat" "$InstallDir" -Force

# (C) 設定檔 (Configs) - 僅在不存在時複製 (保留使用者設定)
$ConfigFiles = Get-ChildItem "$SourceRoot\Configs"
foreach ($file in $ConfigFiles) {
    $DestPath = "$InstallDir\Configs\$($file.Name)"
    if (-not (Test-Path $DestPath)) {
        Copy-Item $file.FullName $DestPath
        Write-Host "  - 初始化設定: $($file.Name)" -ForegroundColor Gray
    } else {
        # 特例：如果是 Webhook.url.example，總是複製過去當參考
        if ($file.Name -like "*.example") { Copy-Item $file.FullName $DestPath -Force }
    }
}

# --- 4. 還原路徑設定 & 詢問新設定 ---
Write-Host "`n[4/5] 設定環境變數..." -ForegroundColor Yellow

# 如果是全新安裝，或讀取不到舊路徑，則詢問使用者
if (-not $LocalPaths.ContainsKey("1Remote")) {
    $Input1R = Read-Host "請輸入 1Remote.exe 所在資料夾 (例如 C:\Tools\1Remote)"
    if (-not [string]::IsNullOrWhiteSpace($Input1R)) { $LocalPaths["1Remote"] = $Input1R.Trim('\') }
}

if (-not $LocalPaths.ContainsKey("BetterGI")) {
    $InputBG = Read-Host "請輸入 BetterGI 資料夾 (例如 C:\Program Files\BetterGI)"
    if (-not [string]::IsNullOrWhiteSpace($InputBG)) { $LocalPaths["BetterGI"] = $InputBG.Trim('\') }
}

$RemoteUser = Read-Host "請輸入遠端帳戶名稱 (預設: Remote)"
if ([string]::IsNullOrWhiteSpace($RemoteUser)) { $RemoteUser = "Remote" }

# 寫入路徑到新腳本中
$ScriptsToUpdate = @("$InstallDir\Scripts\Master.ps1", "$InstallDir\Scripts\Payload.ps1", "$InstallDir\Scripts\Monitor.ps1", "$InstallDir\Scripts\Dashboard.ps1")

foreach ($Script in $ScriptsToUpdate) {
    if (Test-Path $Script) {
        $Content = Get-Content $Script -Raw -Encoding UTF8
        $Modified = $false
        
        if ($LocalPaths["1Remote"]) {
            $Path1R = $LocalPaths["1Remote"].Replace("\", "\\") # Regex Escape
            $Content = $Content -replace '\$1RemoteDir\s*=\s*".*?"', "`$1RemoteDir     = `"$($LocalPaths["1Remote"])`""
            $Modified = $true
        }
        if ($LocalPaths["BetterGI"]) {
            $PathBG = $LocalPaths["BetterGI"].Replace("\", "\\")
            $UserDir = "$($LocalPaths["BetterGI"])\User\OneDragon"
            $Content = $Content -replace '\$BettergiDir\s*=\s*".*?"', "`$BettergiDir    = `"$($LocalPaths["BetterGI"])`""
            $Content = $Content -replace '\$BetterGI_UserDir\s*=\s*".*?"', "`$BetterGI_UserDir = `"$UserDir`""
            $Modified = $true
        }
        
        # 更新 Remote 帳戶名稱
        $Content = $Content -replace '-r Remote', "-r $RemoteUser"
        $Content = $Content -replace 'Sessions -match "Remote"', "Sessions -match `"\b$RemoteUser\b`""
        
        if ($Modified) { Set-Content $Script -Value $Content -Encoding UTF8 }
    }
}

# --- 5. 系統註冊與權限 ---
Write-Host "`n[5/5] 註冊系統設定..." -ForegroundColor Yellow

# 權限
icacls "$InstallDir" /grant Everyone:(OI)(CI)F /T /C | Out-Null

# 註冊表 (RDP 最小化)
reg add "HKEY_LOCAL_MACHINE\Software\Microsoft\Terminal Server Client" /v "RemoteDesktop_SuppressWhenMinimized" /t REG_DWORD /d 2 /f | Out-Null

# 註冊工作排程 (Master)
Unregister-ScheduledTask -TaskName "Auto_1Remote_Master" -Confirm:$false -ErrorAction SilentlyContinue
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$InstallDir\Scripts\Master.ps1`""
$Trigger1 = New-ScheduledTaskTrigger -AtStartup
$Trigger2 = New-ScheduledTaskTrigger -Daily -At "03:55"
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2) -Priority 0
Register-ScheduledTask -TaskName "Auto_1Remote_Master" -Action $Action -Trigger @($Trigger1, $Trigger2) -Settings $Settings -RunLevel Highest -Force | Out-Null

# 產生 Remote 端安裝檔
$RemoteBat = "$InstallDir\Setup_Remote_Task.bat"
@"
@echo off
echo 正在註冊 Remote ($RemoteUser) 的自動化任務...
schtasks /create /tn "Auto_BetterGI_Payload" /tr "powershell.exe -ExecutionPolicy Bypass -File \"$InstallDir\Scripts\Payload.ps1\"" /sc ONLOGON /rl HIGHEST /f
echo.
echo 請手動檢查工作排程器，建議補上「連線時」觸發條件。
pause
"@ | Set-Content $RemoteBat -Encoding ASCII

# 清理暫存
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n========================================================" -ForegroundColor Green
Write-Host "  部署/更新完成！"
Write-Host "========================================================"
Write-Host "請記得："
Write-Host "1. 若是新電腦，請登入 [$RemoteUser] 帳戶執行 $RemoteBat"
Write-Host "2. 若有 Discord Webhook，請檢查 Configs\Webhook.url"
Write-Host ""
pause