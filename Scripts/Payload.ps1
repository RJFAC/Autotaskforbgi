# ==============================================================================
# AutoTask Payload Script V5.43 (Logic Fix)
# ------------------------------------------------------------------------------
# V5.43: 修復 WeeklyConfig 讀取邏輯。
#        解決 V5.42 僅針對 Day 8 處理，導致一般日期無法解析 "Default" 而回退錯誤的問題。
#        現在會完整判斷「紊亂期」與「一般期」並正確讀取週配置。
# V5.42: Idempotency Check (重複執行防護)。
# V5.41: Log Liveness Monitor.
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\Payload_$DateStr.log"
$FlagDir = "$WorkDir\Flags"
$DoneFlag = "$FlagDir\Done.flag"
$WeeklyConfFile = "$WorkDir\Configs\WeeklyConfig.json"
$TaskStatusFile = "$WorkDir\Configs\TaskStatus.json"
$LastRunFile = "$WorkDir\Configs\LastRun.log"
$ForceRunFlag = "$FlagDir\ForceRun.flag"

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

# 2. 啟動檢查 (Pre-flight Checks)
Write-Log ">>> Payload 啟動 (V5.43 - Logic Fix)..."

# 計算今日 Key (04:00 界線)
$Now = Get-Date
if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }

# [V5.42] 重複執行防護：檢查 LastRun.log
if (Test-Path $LastRunFile) {
    try {
        $LastRunDate = (Get-Content $LastRunFile -Raw).Trim()
        $IsForceRun = Test-Path $ForceRunFlag
        
        if ($LastRunDate -eq $TodayKey) {
            if ($IsForceRun) {
                Write-Log "⚠️ 檢測到今日任務已完成 ($LastRunDate)，但存在 ForceRun 標記，強制重跑。" "YELLOW"
                # 移除 ForceRun 防止下次誤判
                Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "✅ 今日任務已標記為完成 ($TodayKey)。" "GREEN"
                Write-Log "   觸發原因推測: 使用者登入檢查或排程重複觸發。" "GRAY"
                Write-Log "   Payload 將自動退出 (Idempotency Guard)。" "GRAY"
                Start-Sleep 3
                exit 0
            }
        }
    } catch {
        Write-Log "讀取 LastRun.log 發生錯誤，將繼續執行: $_" "WARN"
    }
}

# 狀態同步：立即更新為 Running
if (Test-Path $TaskStatusFile) {
    try {
        $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($Json.Date -eq $DateStr) {
            $Json.Status = "Running"
            $Json.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8
            Write-Log "狀態同步: TaskStatus 已更新為 'Running'"
        }
    } catch { Write-Log "更新 TaskStatus 失敗: $_" "WARN" }
}

# 03:50 等待邏輯
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
    # 重新計算時間與 Key
    $Now = Get-Date
    if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }
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

# --- 讀取 DateConfig.map (優先順序 1) ---
$MapFile = "$WorkDir\Configs\DateConfig.map"
$RawTaskString = "Default"
Write-Log "計算日期 Key: $TodayKey"

if (Test-Path $MapFile) {
    $MapContent = Get-Content $MapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$TodayKey=(.*)") {
            $RawTaskString = $Matches[1].Trim()
            Write-Log "📅 命中 DateConfig.map 指定配置: $RawTaskString" "CYAN"
            break
        }
    }
}

# --- 讀取 WeeklyConfig.json (優先順序 2) ---
# [Fix V5.43] 補完 Day 8 以外日期的讀取邏輯
$RefDate = [datetime]"2024-08-28T00:00:00"
$CycleOffset = ($Now - $RefDate).TotalDays % 42
if ($CycleOffset -lt 0) { $CycleOffset += 42 }
$IsTurbulenceDay1 = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 8.0)

if ($RawTaskString -eq "Default") {
    if (Test-Path $WeeklyConfFile) {
        try {
            $WkJson = Get-Content $WeeklyConfFile -Raw | ConvertFrom-Json
            $WeekKey = $Now.DayOfWeek.ToString() # e.g., "Thursday"
            
            # 定義紊亂期範圍 (Day 8 ~ Day 18)
            # Day 8 starts at offset 7.x
            # Day 18 ends at offset 17.x (Saturday 03:59)
            $IsTurbulencePeriod = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 17.2)
            
            if ($IsTurbulenceDay1) {
                # Day 8 (週三) 特殊處理：注入 [WAIT]
                $WkDef = if ($WkJson.$WeekKey) { $WkJson.$WeekKey } else { "Default" }
                $WkTurb = if ($WkJson.Turbulence -and $WkJson.Turbulence.$WeekKey) { $WkJson.Turbulence.$WeekKey } else { "Default" }
                $RawTaskString = "$WkDef,[WAIT],$WkTurb"
                Write-Log "📅 偵測到 Day 8，注入雙重排程: $RawTaskString" "MAGENTA"
            } elseif ($IsTurbulencePeriod) {
                # 紊亂期其他天 (Day 9 - 17)
                if ($WkJson.Turbulence -and $WkJson.Turbulence.$WeekKey) {
                    $RawTaskString = $WkJson.Turbulence.$WeekKey
                    Write-Log "🔥 偵測到紊亂期 ($WeekKey)，使用紊亂配置: $RawTaskString" "MAGENTA"
                } else {
                    # 若無紊亂配置，回退到一般配置
                    if ($WkJson.$WeekKey) { 
                        $RawTaskString = $WkJson.$WeekKey 
                        Write-Log "🔥 紊亂期 ($WeekKey) 但無專屬配置，使用一般配置: $RawTaskString"
                    }
                }
            } else {
                # 一般期間 (非紊亂期)
                if ($WkJson.$WeekKey) {
                    $RawTaskString = $WkJson.$WeekKey
                    Write-Log "📅 使用一般每週配置 ($WeekKey): $RawTaskString"
                }
            }
        } catch { Write-Log "讀取 WeeklyConfig 失敗: $_" "ERROR" }
    }
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
if (-not (Test-Path $BetterGIPath)) { Write-Log "❌ 致命錯誤: 找不到 BetterGI: $BetterGIPath" "ERROR"; exit 1 }
$BGIDir = Split-Path $BetterGIPath -Parent
$BGILogsDir = Join-Path $BGIDir "log"

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
        } elseif ($i -eq $SplitIndex) { $NeedWait = $true }
    }

    if ($NeedWait) {
        Write-Log "=== 進入 10:00 等待模式 ([WAIT] 觸發) ===" "YELLOW"
        $TargetTime = $Now.Date.AddHours(10)
        while ((Get-Date) -lt $TargetTime) {
            $Diff = $TargetTime - (Get-Date); if ($Diff.TotalMinutes -gt 0) { Write-Host "⏳ 等待... $($Diff.Minutes) 分" -NoNewline -ForegroundColor Yellow; Start-Sleep 30 }
            if ((Get-Date).Hour -ge 14) { break }
        }
        Write-Log "`n>>> 時間已達 10:00+，繼續執行。" "GREEN"
    }

    if ($CurrentTask -eq "[WAIT]") { continue }

    Write-Log "啟動 BetterGI [$($i+1)/$($TaskList.Count)]: $CurrentTask"
    $ArgsList = "--startOneDragon `"$CurrentTask`""
    $Process = Start-Process -FilePath $BetterGIPath -ArgumentList $ArgsList -WorkingDirectory $BGIDir -PassThru
    
    # [V5.41] 啟動後等待並鎖定最新日誌
    Start-Sleep 5 
    $CurrentBGILogPath = ""
    if (Test-Path $BGILogsDir) {
        $LatestLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($LatestLog) {
            $CurrentBGILogPath = $LatestLog.FullName
            Write-Log "鎖定日誌: $($LatestLog.Name)" "Cyan"
        }
    }

    # 監控迴圈 (Log Liveness Mode)
    $StuckThresholdMinutes = 15 # 定義：15分鐘無寫入=卡死
    
    while ($true) {
        if ($Process.HasExited) { Write-Log "任務完成。"; break }
        
        $CheckTime = Get-Date
        
        # 1. 死線檢查 (03:50)
        if ($CheckTime.Hour -eq 3 -and $CheckTime.Minute -ge 50) {
             Stop-Process -Id $Process.Id -Force; Write-Log "⚠️ 遭遇 03:50 死線，強制中斷。" "RED"; break
        }

        # 2. 日誌活躍度檢查
        if (-not [string]::IsNullOrWhiteSpace($CurrentBGILogPath) -and (Test-Path $CurrentBGILogPath)) {
            $LogFileItem = Get-Item $CurrentBGILogPath
            $SilenceMinutes = ($CheckTime - $LogFileItem.LastWriteTime).TotalMinutes
            
            if ($SilenceMinutes -gt $StuckThresholdMinutes) {
                Stop-Process -Id $Process.Id -Force
                Write-Log "⛔ 日誌靜止超過 $StuckThresholdMinutes 分鐘，判定卡死，強制跳過。" "RED"
                break
            }
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
Set-Content -Path $LastRunFile -Value $TodayKey

# 更新狀態為 Success
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