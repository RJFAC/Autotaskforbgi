# ==============================================================================
# AutoTask Payload Script V5.35 (Day 8 Default Injection)
# ------------------------------------------------------------------------------
# V5.35: 針對 Day 8 (週三)，若無 DateConfig 覆蓋 (即 TaskName="Default")，
#        自動從 WeeklyConfig 注入 [Task1, [WAIT], Task2] 的完整流程。
# V5.34: 支援 [WAIT] 標記。
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\Payload_$DateStr.log"
$FlagDir = "$WorkDir\Flags"
$DoneFlag = "$FlagDir\Done.flag"
$WeeklyConfFile = "$WorkDir\Configs\WeeklyConfig.json"

# 確保日誌目錄存在
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# 日誌函數
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    Write-Host $LogEntry
}

trap {
    Write-Log "CRASH: $($_.Exception.Message)" "ERROR"
    Write-Log "StackTrace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}

# 2. 啟動與跨日檢查 (Smart Wait)
Write-Log ">>> Payload 啟動 (V5.35 - Day 8 Inject)..."

$Now = Get-Date
if ($Now.Hour -eq 3 -and $Now.Minute -ge 50) {
    Write-Log "⚠️ 偵測到於重置緩衝期 (03:50~04:00) 啟動，進入等待模式..." "WARNING"
    while ($true) {
        $Check = Get-Date
        if ($Check.Hour -ge 4) {
            Write-Log ">>> 時間已達 04:00+，解除鎖定！" "GREEN"
            Start-Sleep 5
            break
        }
        Start-Sleep 10
    }
    $Now = Get-Date
}

# 讀取 Configs
$EnvConfigFile = "$WorkDir\Configs\EnvConfig.json"
if (Test-Path $EnvConfigFile) {
    $EnvConfig = Get-Content -Path $EnvConfigFile -Raw | ConvertFrom-Json
    $GenshinPath = $EnvConfig.GenshinPath
} else {
    Write-Log "找不到 EnvConfig.json，使用預設路徑。" "WARN"
    $GenshinPath = "C:\Program Files\HoYoPlay\games\Genshin Impact Game"
}

# --- 讀取 DateConfig.map ---
$MapFile = "$WorkDir\Configs\DateConfig.map"
$RawTaskString = "Default"

if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }
Write-Log "計算日期 Key: $TodayKey"

if (Test-Path $MapFile) {
    $MapContent = Get-Content $MapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$TodayKey=(.*)") {
            $RawTaskString = $Matches[1].Trim()
            break
        }
    }
}

# --- Day 8 偵測與預設注入邏輯 ---
$RefDate = [datetime]"2024-08-28T00:00:00"
$CycleOffset = ($Now - $RefDate).TotalDays % 42
if ($CycleOffset -lt 0) { $CycleOffset += 42 }
$IsTurbulenceDay1 = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 8.0)

# [V5.35] 若為 Day 8 且無覆蓋設定 (RawTaskString == "Default")，自動注入雙重排程
if ($IsTurbulenceDay1 -and $RawTaskString -eq "Default") {
    Write-Log "📅 偵測到 Day 8 且無覆蓋設定，嘗試從 WeeklyConfig 注入預設雙重排程..." "MAGENTA"
    
    $WkDef = "模板-Copy" # Fallback
    $WkTurb = "模板-Copy" # Fallback
    
    if (Test-Path $WeeklyConfFile) {
        try {
            $WkJson = Get-Content $WeeklyConfFile -Raw | ConvertFrom-Json
            if ($WkJson.Wednesday) { $WkDef = $WkJson.Wednesday }
            if ($WkJson.Turbulence -and $WkJson.Turbulence.Wednesday) { $WkTurb = $WkJson.Turbulence.Wednesday }
        } catch { Write-Log "讀取 WeeklyConfig 失敗: $_" "ERROR" }
    }
    
    # 建構注入字串
    $RawTaskString = "$WkDef,[WAIT],$WkTurb"
    Write-Log "-> 已注入任務序列: $RawTaskString" "CYAN"
}

# 解析任務清單
$TaskList = @()
if ($RawTaskString -match ",") {
    $TaskList = $RawTaskString -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
} else {
    $TaskList = @($RawTaskString)
}
Write-Log "最終執行清單: $($TaskList -join ' -> ')"

# 3. 準備 BetterGI 執行環境
$BetterGIPath = "C:\AutoTask\BetterGI\BetterGI.exe" 
$BetterGILogPath = "$WorkDir\Logs\BetterGI\BetterGI.log"
Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue

# --- 分割點與等待邏輯 ---
$ExplicitWaitIndex = -1
for ($k = 0; $k -lt $TaskList.Count; $k++) {
    if ($TaskList[$k] -eq "[WAIT]") { $ExplicitWaitIndex = $k; break }
}

# 自動推斷分割點 (Fallback)
$SplitIndex = -1 
if ($IsTurbulenceDay1 -and $ExplicitWaitIndex -lt 0) {
    if ($TaskList.Count -gt 2) { $SplitIndex = 3 } else { $SplitIndex = 1 }
    Write-Log "📅 Day 8 自動推斷: 於第 $($SplitIndex+1) 個任務前等待 10:00。"
}

# ----------------------------
# 迴圈執行
# ----------------------------
for ($i = 0; $i -lt $TaskList.Count; $i++) {
    $CurrentTask = $TaskList[$i]
    
    # 檢查是否需要等待
    $NeedWait = $false
    if ($IsTurbulenceDay1) {
        if ($ExplicitWaitIndex -ge 0) {
            if ($i -eq $ExplicitWaitIndex) { $NeedWait = $true }
        } elseif ($i -eq $SplitIndex) {
            $NeedWait = $true
        }
    }

    if ($NeedWait) {
        Write-Log "=== 進入 10:00 等待模式 ([WAIT] 觸發) ===" "YELLOW"
        $TargetTime = $Now.Date.AddHours(10)
        while ((Get-Date) -lt $TargetTime) {
            $Diff = $TargetTime - (Get-Date)
            if ($Diff.TotalMinutes -gt 0) {
                Write-Host "⏳ 等待活動開放... 剩餘 $($Diff.Minutes) 分鐘" -NoNewline -ForegroundColor Yellow
                Start-Sleep 30
            }
            if ((Get-Date).Hour -ge 14) { break } 
        }
        Write-Log "`n>>> 時間已達 10:00+，繼續執行。" "GREEN"
    }

    # 跳過標記本身
    if ($CurrentTask -eq "[WAIT]") { continue }

    # 執行 BetterGI
    Write-Log "啟動 BetterGI [$($i+1)/$($TaskList.Count)]: $CurrentTask"
    $ArgsList = "-start -task `"$CurrentTask`""
    $Process = Start-Process -FilePath $BetterGIPath -ArgumentList $ArgsList -WorkingDirectory (Split-Path $BetterGIPath) -PassThru
    
    # 監控
    $TimeoutMinutes = 180
    $StartTime = Get-Date
    while ($true) {
        if ($Process.HasExited) { Write-Log "任務完成。"; break }
        
        # 03:50 ForceEnd 檢查
        $CheckTime = Get-Date
        if ($CheckTime.Hour -eq 3 -and $CheckTime.Minute -ge 50) {
             Stop-Process -Id $Process.Id -Force
             Write-Log "⚠️ 遭遇 03:50 死線，強制中斷。" "RED"
             break
        }
        if (($CheckTime - $StartTime).TotalMinutes -gt $TimeoutMinutes) {
            Stop-Process -Id $Process.Id -Force; Write-Log "⚠️ 超時跳過。" "RED"; break
        }
        Start-Sleep 10
    }
    
    if ($i -lt ($TaskList.Count - 1)) {
        Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 5
    }
}

# 4. 結算
Write-Log "Payload 執行結束，登出..."
New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
Set-Content -Path "$WorkDir\Configs\LastRun.log" -Value $TodayKey
shutdown.exe /l /f