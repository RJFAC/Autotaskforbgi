# =============================================================================
# AutoTask Payload V5.1 - 樹脂策略安全版
# =============================================================================

# --- [0. 身分驗證] ---
$TargetUser = "Remote" 
$CurrentUserName = [System.Environment]::UserName
$BaseDir    = "C:\AutoTask"
$LogDir     = "$BaseDir\Logs"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-PreLog {
    param($Msg, $Color="Red")
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogF = Join-Path $LogDir "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $Txt = "[$Time] $Msg"
    Write-Host $Txt -ForegroundColor $Color
    try { Add-Content -Path $LogF -Value $Txt -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

if ($CurrentUserName -ne $TargetUser) {
    Write-PreLog "⛔ 錯誤：身分不符！Payload 被 [$CurrentUserName] 誤觸，已攔截。" "Red"
    Start-Sleep 3; exit 
}

# --- [路徑與變數] ---
$ScriptDir  = "$BaseDir\Scripts"
$ConfigDir  = "$BaseDir\Configs"
$FlagDir    = "$BaseDir\Flags"
$BettergiDir    = "C:\Program Files\BetterGI"
$BettergiExe    = "$BettergiDir\BetterGI.exe"
$BettergiUserConf = "$BettergiDir\User\config.json"
$LogDirBG       = "$BettergiDir\log"
$ScreenshotDir  = "$BettergiDir\Screenshots" 
$1RemoteLogDir  = $null 

$TaskStatusFile = "$ConfigDir\TaskStatus.json"
$DoneFlag       = "$FlagDir\Done.flag"
$FailFlag       = "$FlagDir\Fail.flag"
$ForceRunFlag   = "$FlagDir\ForceRun.flag" 
$DateMap        = "$ConfigDir\DateConfig.map"
$WeeklyConf     = "$ConfigDir\WeeklyConfig.json"
$EnvConf        = "$ConfigDir\EnvConfig.json"
$ResinConf      = "$ConfigDir\ResinConfig.json"
$PauseLog       = "$ConfigDir\PauseDates.log"
$LastRunLog     = "$ConfigDir\LastRun.log"
$BackupRootDir  = "$BaseDir\LogBackups"
$NotifyScript   = "$ScriptDir\Notify.ps1"
$MaxRetries = 3
$SuccessKeyword = "一条龙.*任务结束"

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    try { Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# --- [1. 啟動前安全檢查] (Crash Recovery) ---
Write-Log "Payload 啟動... 檢查環境完整性..." "Cyan"

# 檢查是否有上次殘留的備份檔 (代表上次非正常結束)
$BakFile = "$BettergiUserConf.bak"
if (Test-Path $BakFile) {
    Write-Log "⚠️ 偵測到 BetterGI 設定備份檔殘留，代表上次可能異常中斷。" "Yellow"
    try {
        Copy-Item $BakFile $BettergiUserConf -Force
        Remove-Item $BakFile -Force
        Write-Log "✅ 已強制還原 BetterGI 原始設定。" "Green"
    } catch {
        Write-Log "❌ 還原失敗: $_" "Red"
    }
}

# 清理舊程序
try {
    $CurrentPID = $PID
    $TargetScript = "Payload.ps1"
    $OldInstances = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | 
        Where-Object { $_.CommandLine -like "*$TargetScript*" -and $_.ProcessId -ne $CurrentPID }
    if ($OldInstances) {
        foreach ($proc in $OldInstances) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }
    }
} catch {}

# --- [載入環境設定] ---
if (Test-Path $EnvConf) {
    try {
        $EnvJson = Get-Content $EnvConf -Raw -Encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-Json
        if ($EnvJson.GenshinPath) { $GenshinPath = $EnvJson.GenshinPath }
    } catch {}
}

# --- [輔助函數] ---
function Set-BetterGIResinConfig {
    param([string]$ConfigName)
    if (-not (Test-Path $ResinConf)) { return $false }
    if (-not (Test-Path $BettergiUserConf)) { return $false }

    try {
        $Rules = Get-Content $ResinConf -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $Rules.$ConfigName) { return $false }
        
        $Rule = $Rules.$ConfigName
        Write-Log "套用樹脂策略: [$ConfigName] (模式: $($Rule.ResinMode))" "Cyan"

        # 備份
        if (-not (Test-Path $BakFile)) { Copy-Item $BettergiUserConf $BakFile -Force }

        $BgiConfig = Get-Content $BettergiUserConf -Raw -Encoding UTF8 | ConvertFrom-Json
        $TargetSection = if ($Rule.TaskType -eq "Stygian") { "autoStygianOnslaughtConfig" } else { "autoDomainConfig" }
        
        if (-not $BgiConfig.$TargetSection) { return $false }
        $Section = $BgiConfig.$TargetSection
        
        if ($Rule.Priority) { $Section.resinPriorityList = $Rule.Priority }
        if ($Rule.ResinMode -eq "Count") {
            $Section.specifyResinUse = $true
            if ($Rule.Counts) {
                $Section.originalResinUseCount = $Rule.Counts.Original
                $Section.condensedResinUseCount = $Rule.Counts.Condensed
                $Section.transientResinUseCount = $Rule.Counts.Transient
                $Section.fragileResinUseCount = $Rule.Counts.Fragile
            }
        } else {
            $Section.specifyResinUse = $false
        }

        $BgiConfig | ConvertTo-Json -Depth 20 | Set-Content $BettergiUserConf -Encoding UTF8
        return $true
    } catch {
        Write-Log "樹脂設定修改失敗: $_" "Red"
        return $false
    }
}

function Restore-BetterGIConfig {
    if (Test-Path $BakFile) {
        try {
            Copy-Item $BakFile $BettergiUserConf -Force
            Remove-Item $BakFile -Force
            Write-Log "已還原 BetterGI 預設設定。" "Gray"
        } catch {}
    }
}

function Check-Network {
    Write-Log "檢查網路..." 
    $Retry = 0; $MaxRetry = 12
    while ($Retry -lt $MaxRetry) {
        if (Test-Connection "8.8.8.8" -Count 1 -Quiet) { return $true }
        Start-Sleep 5; $Retry++
    }
    Write-Log "⚠️ 網路逾時。" "Red"; return $false
}

function Send-Notify {
    param($Title, $Msg, $Color)
    if (Test-Path $NotifyScript) {
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$Title`" -Message `"$Msg`" -Color `"$Color`"" -WindowStyle Hidden
    }
}

function Check-Success-Log ($logPath) {
    try {
        $Logs = Get-Content $logPath -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($Logs -match $SuccessKeyword) { return $true }
    } catch {}
    return $false
}

# ... (Test-GenshinUpdateDay, Check-GenshinPreDownload, Cleanup-Screenshots, Backup-Logs, Update-Status, Get-TargetConfig 省略，請保持原樣或複製 V4.12 的完整函數) ...
# 請確保這裡包含完整的輔助函數定義，為了節省空間我先簡化

# =============================================================================
# [主流程]
# =============================================================================
# ... (ForceEnd, 04:00 Wait, Date Check, UpdateDay Check 邏輯保持 V4.12 不變) ...

# [執行今日任務]
$RetryCount = 0
$ConfigName = Get-TargetConfig
$ConfigQueue = $ConfigName -split ","

if (-not (Test-Path $BettergiExe)) { exit }

Check-Network

while ($RetryCount -le $MaxRetries) {
    Update-Status "Running" $RetryCount
    $AllConfigSuccess = $true
    
    foreach ($CurrentConfig in $ConfigQueue) {
        if ([string]::IsNullOrWhiteSpace($CurrentConfig)) { continue }
        
        Write-Log ">>> 執行配置: [$CurrentConfig]" "Cyan"

        # [新] 嘗試套用樹脂設定
        $ConfigChanged = Set-BetterGIResinConfig $CurrentConfig

        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3

        $Args = "--startOneDragon ""$CurrentConfig"""
        Start-Process -FilePath $BettergiExe -ArgumentList $Args -WorkingDirectory $BettergiDir
        
        $WatchdogStart = Get-Date
        $IsSuccess = $false
        $IsFailed = $false
        
        # ... (日誌鎖定與監控迴圈，保持 V4.12 不變) ...
        # 請複製 V4.12 的完整監控邏輯貼在這裡
        # ...
        
        # [新] 任務結束後 (無論成功失敗)，立即還原設定
        if ($ConfigChanged) { Restore-BetterGIConfig }

        if (-not $IsSuccess) {
            $AllConfigSuccess = $false
            Write-Log "配置 [$CurrentConfig] 失敗，準備重試。" "Yellow"
            break 
        } else {
            Start-Sleep 5
        }
    }

    if ($AllConfigSuccess) {
        # ... (成功後續處理) ...
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
        logoff
        exit
    } else {
        # ... (失敗重試處理) ...
        $RetryCount++
        if ($RetryCount -gt $MaxRetries) {
            # ...
            logoff
            exit
        }
        Start-Sleep 10
    }
}