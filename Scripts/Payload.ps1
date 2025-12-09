# ==============================================================================
# AutoTask Payload Script V5.40 (Status Update Fix)
# ------------------------------------------------------------------------------
# V5.40: 新增啟動時更新 TaskStatus.json 為 "Running" 的邏輯，修復 Dashboard 卡在 Preparing 的問題。
# V5.39: 強制將所有任務視為「一條龍配置組」，使用 --startOneDragon 參數啟動。
# V5.38: 修正 BetterGI 日誌路徑為 "C:\Program Files\BetterGI\log\"。
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\Payload_$DateStr.log"
$FlagDir = "$WorkDir\Flags"
$DoneFlag = "$FlagDir\Done.flag"
$WeeklyConfFile = "$WorkDir\Configs\WeeklyConfig.json"
$TaskStatusFile = "$WorkDir\Configs\TaskStatus.json" # [V5.40] 新增狀態檔路徑

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
    # [V5.40] 發生崩潰時嘗試更新狀態為 Failed
    try {
        if (Test-Path $TaskStatusFile) {
            $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($Json.Date -eq $DateStr) {
                $Json.Status = "Failed"
                $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8
            }
        }
    } catch {}
    exit 1
}

# 2. 啟動與跨日檢查 (Smart Wait)
Write-Log ">>> Payload 啟動 (V5.40 - Status Fix)..."

# [V5.40] 啟動時立即更新狀態為 Running
if (Test-Path $TaskStatusFile) {
    try {
        $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
        # 只有當日期匹配且狀態是 Preparing 時才接手更新
        if ($Json.Date -eq $DateStr) {
            $Json.Status = "Running"
            $Json.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8
            Write-Log "狀態同步: TaskStatus 已更新為 'Running'"
        }
    } catch {
        Write-Log "更新 TaskStatus 失敗: $_" "WARN"
    }
}

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

# Day 8 預設注入
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
$BetterGIPath = "C:\Program Files\BetterGI\BetterGI.exe"

# 路徑驗證
if (-not (Test-Path $BetterGIPath)) {
    Write-Log "❌ 致命錯誤: 找不到 BetterGI 執行檔！路徑: $BetterGIPath" "ERROR"
    exit 1
}

$BGIDir = Split-Path $BetterGIPath -Parent
if (-not (Test-Path $BGIDir)) {
    Write-Log "❌ 致命錯誤: WorkingDirectory 不存在: $BGIDir" "ERROR"
    exit 1
}

# [V5.38 Fix] 修正日誌路徑為 "log" (小寫)
$BGILogsDir = Join-Path $BGIDir "log"
$BetterGILogPath = "" 
if (Test-Path $BGILogsDir) {
    # 嘗試抓取最新的 better-genshin-impact*.log (根據 BGI 命名慣例)
    $LatestLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($LatestLog) {
        $BetterGILogPath = $LatestLog.FullName
        Write-Log "鎖定最新 BGI 日誌: $($LatestLog.Name)"
    } else {
         Write-Log "⚠️ 在 $BGILogsDir 中找不到任何 .log 檔案。" "WARN"
    }
} else {
    Write-Log "⚠️ 找不到 BGI log 目錄: $BGILogsDir" "WARN"
}

Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue

# --- 分割點與等待邏輯 ---
$ExplicitWaitIndex = -1
for ($k = 0; $k -lt $TaskList.Count; $k++) {
    if ($TaskList[$k] -eq "[WAIT]") { $ExplicitWaitIndex = $k; break }
}

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

    if ($CurrentTask -eq "[WAIT]") { continue }

    Write-Log "啟動 BetterGI [$($i+1)/$($TaskList.Count)]: $CurrentTask"
    
    # [V5.39] 強制使用 --startOneDragon 參數
    # 根據定義，所有 payload 啟動的任務均為一條龍配置組
    $ArgsList = "--startOneDragon `"$CurrentTask`""
    
    $Process = Start-Process -FilePath $BetterGIPath -ArgumentList $ArgsList -WorkingDirectory $BGIDir -PassThru
    
    # 監控
    $TimeoutMinutes = 180
    $StartTime = Get-Date
    while ($true) {
        if ($Process.HasExited) { Write-Log "任務完成。"; break }
        
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

# [V5.40] 任務完成時更新狀態為 Success
if (Test-Path $TaskStatusFile) {
    try {
        $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($Json.Date -eq $DateStr) {
            $Json.Status = "Success"
            $Json.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8
        }
    } catch {}
}

shutdown.exe /l /f