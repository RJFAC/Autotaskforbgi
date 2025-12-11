# ==============================================================================
# AutoTask Payload Script V5.44 (Retry & Recheck Logic)
# ------------------------------------------------------------------------------
# V5.44: 
#   1. 啟動後等待時間延長至 20秒，防止鎖定舊日誌。
#   2. 新增「日誌雙重確認」機制：超時 15 分鐘時，再次檢查是否有新日誌產生。
#   3. 新增「任務重試」機制：判定卡死後，嘗試重啟當前配置 (Max 3次)。
#   4. 若重試失敗，標記 TaskStatus 為 Failed 並退出。
# V5.43: WeeklyConfig Logic Fix.
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

function Update-TaskStatus {
    param ([string]$Status)
    try {
        if (Test-Path $TaskStatusFile) {
            $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($Json.Date -eq $DateStr) {
                $Json.Status = $Status
                $Json.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8
            }
        }
    } catch { Write-Log "更新 TaskStatus 失敗: $_" "WARN" }
}

trap {
    Write-Log "CRASH: $($_.Exception.Message)" "ERROR"
    Write-Log "StackTrace: $($_.ScriptStackTrace)" "ERROR"
    Update-TaskStatus "Failed"
    exit 1
}

# 2. 啟動檢查 (Pre-flight Checks)
Write-Log ">>> Payload 啟動 (V5.44 - Retry & Recheck)..."

# 計算今日 Key (04:00 界線)
$Now = Get-Date
if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }

# 重複執行防護：檢查 LastRun.log
if (Test-Path $LastRunFile) {
    try {
        $LastRunDate = (Get-Content $LastRunFile -Raw).Trim()
        $IsForceRun = Test-Path $ForceRunFlag
        
        if ($LastRunDate -eq $TodayKey) {
            if ($IsForceRun) {
                Write-Log "⚠️ 檢測到今日任務已完成，但存在 ForceRun 標記，強制重跑。" "YELLOW"
                Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "✅ 今日任務已標記為完成 ($TodayKey)。Payload 自動退出。" "GREEN"
                Start-Sleep 3
                exit 0
            }
        }
    } catch {
        Write-Log "讀取 LastRun.log 發生錯誤，將繼續執行: $_" "WARN"
    }
}

# 狀態同步：立即更新為 Running
Update-TaskStatus "Running"

# 03:50 等待邏輯
if ($Now.Hour -eq 3 -and $Now.Minute -ge 50) {
    Write-Log "⚠️ 偵測到於重置緩衝期 (03:50~04:00) 啟動，進入等待模式..." "WARNING"
    while ($true) {
        if ((Get-Date).Hour -ge 4) {
            Write-Log ">>> 時間已達 04:00+，解除鎖定！" "GREEN"
            Start-Sleep 5
            break
        }
        Start-Sleep 10
    }
    $Now = Get-Date
    if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }
}

# --- 配置讀取邏輯 (DateConfig -> Day 8 -> WeeklyConfig) ---
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

$RefDate = [datetime]"2024-08-28T00:00:00"
$CycleOffset = ($Now - $RefDate).TotalDays % 42
if ($CycleOffset -lt 0) { $CycleOffset += 42 }
$IsTurbulenceDay1 = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 8.0)

if ($RawTaskString -eq "Default") {
    if (Test-Path $WeeklyConfFile) {
        try {
            $WkJson = Get-Content $WeeklyConfFile -Raw | ConvertFrom-Json
            $WeekKey = $Now.DayOfWeek.ToString() 
            
            $IsTurbulencePeriod = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 17.2)
            
            if ($IsTurbulenceDay1) {
                $WkDef = if ($WkJson.$WeekKey) { $WkJson.$WeekKey } else { "Default" }
                $WkTurb = if ($WkJson.Turbulence -and $WkJson.Turbulence.$WeekKey) { $WkJson.Turbulence.$WeekKey } else { "Default" }
                $RawTaskString = "$WkDef,[WAIT],$WkTurb"
                Write-Log "📅 偵測到 Day 8，注入雙重排程: $RawTaskString" "MAGENTA"
            } elseif ($IsTurbulencePeriod) {
                if ($WkJson.Turbulence -and $WkJson.Turbulence.$WeekKey) {
                    $RawTaskString = $WkJson.Turbulence.$WeekKey
                    Write-Log "🔥 偵測到紊亂期 ($WeekKey)，使用紊亂配置: $RawTaskString" "MAGENTA"
                } else {
                    if ($WkJson.$WeekKey) { $RawTaskString = $WkJson.$WeekKey }
                }
            } else {
                if ($WkJson.$WeekKey) { 
                    $RawTaskString = $WkJson.$WeekKey 
                    Write-Log "📅 使用一般每週配置 ($WeekKey): $RawTaskString"
                }
            }
        } catch { Write-Log "讀取 WeeklyConfig 失敗: $_" "ERROR" }
    }
}

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
$BGILogsDir = Join-Path $BGIDir "log" # BetterGI log dir is 'log' not 'Logs'

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
# 迴圈執行 (含重試機制)
# ----------------------------
$MaxTaskRetries = 3 # 每個任務最多重試 3 次

for ($i = 0; $i -lt $TaskList.Count; $i++) {
    $CurrentTask = $TaskList[$i]
    
    # 處理 WAIT 邏輯
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

    # === 任務重試迴圈 ===
    $RetryCount = 0
    $TaskSuccess = $false

    while ($RetryCount -lt $MaxTaskRetries -and -not $TaskSuccess) {
        
        Write-Log "啟動 BetterGI [$($i+1)/$($TaskList.Count)]: $CurrentTask (Attempt $($RetryCount + 1)/$MaxTaskRetries)"
        
        # 確保環境乾淨
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        
        $ArgsList = "--startOneDragon `"$CurrentTask`""
        $Process = Start-Process -FilePath $BetterGIPath -ArgumentList $ArgsList -WorkingDirectory $BGIDir -PassThru
        
        # [Fix 1] 延長等待時間至 20 秒，確保新日誌已生成
        Write-Log "等待 20 秒以鎖定日誌..."
        Start-Sleep 20 
        
        $CurrentBGILogPath = ""
        if (Test-Path $BGILogsDir) {
            $LatestLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($LatestLog) {
                $CurrentBGILogPath = $LatestLog.FullName
                Write-Log "鎖定日誌: $($LatestLog.Name) (Time: $($LatestLog.LastWriteTime))" "Cyan"
            }
        }

        # 監控迴圈
        $StuckThresholdMinutes = 15
        
        while ($true) {
            if ($Process.HasExited) { 
                Write-Log "✅ BetterGI 進程正常結束。" "GREEN"
                $TaskSuccess = $true
                break 
            }
            
            $CheckTime = Get-Date
            
            # 死線檢查 (03:50)
            if ($CheckTime.Hour -eq 3 -and $CheckTime.Minute -ge 50) {
                 Stop-Process -Id $Process.Id -Force
                 Write-Log "⚠️ 遭遇 03:50 死線，強制中斷所有任務。" "RED"
                 Update-TaskStatus "ForceEnd"
                 exit 0 # 視為正常結束，避免重試
            }

            # 日誌活躍度檢查
            if (-not [string]::IsNullOrWhiteSpace($CurrentBGILogPath) -and (Test-Path $CurrentBGILogPath)) {
                $LogFileItem = Get-Item $CurrentBGILogPath
                $SilenceMinutes = ($CheckTime - $LogFileItem.LastWriteTime).TotalMinutes
                
                if ($SilenceMinutes -gt $StuckThresholdMinutes) {
                    Write-Log "⚠️ 警告：日誌 ($($LogFileItem.Name)) 已靜止 $StuckThresholdMinutes 分鐘。" "YELLOW"
                    
                    # [Fix 2] 雙重確認機制：檢查是否有更新的日誌
                    Write-Log "🔍 正在重新掃描日誌目錄，檢查是否有更新的日誌..." "CYAN"
                    $ReCheckLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    
                    if ($ReCheckLog -and $ReCheckLog.FullName -ne $CurrentBGILogPath) {
                        Write-Log "♻️ 發現更新的日誌！切換鎖定目標 -> $($ReCheckLog.Name)" "GREEN"
                        $CurrentBGILogPath = $ReCheckLog.FullName
                        # 重置靜止時間，繼續監控
                        continue 
                    } else {
                        Write-Log "⛔ 確認無新日誌，判定為真卡死 (True Freeze)。" "RED"
                        Stop-Process -Id $Process.Id -Force
                        
                        # [Fix 3] 觸發重試機制
                        $RetryCount++
                        if ($RetryCount -lt $MaxTaskRetries) {
                            Write-Log "🔄 準備重試當前任務 ($RetryCount/$MaxTaskRetries)..." "YELLOW"
                            Start-Sleep 5
                            break # 跳出監控迴圈，回到 while retry 迴圈
                        } else {
                            Write-Log "❌ 任務 $CurrentTask 重試次數耗盡，宣告任務失敗。" "RED"
                            Update-TaskStatus "Failed"
                            # 發送失敗信號並退出
                            New-Item -ItemType File -Path "$FlagDir\Fail.flag" -Force | Out-Null
                            exit 1
                        }
                    }
                }
            }
            Start-Sleep 10
        } # End Monitor While

        if ($TaskSuccess) { break }

    } # End Retry While

    # 如果重試完畢仍未成功 (理論上 Retry Loop 內會 exit，此為雙重保險)
    if (-not $TaskSuccess) {
        Write-Log "❌ 任務異常終止: $CurrentTask" "RED"
        Update-TaskStatus "Failed"
        exit 1
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

Update-TaskStatus "Success"

shutdown.exe /l /f