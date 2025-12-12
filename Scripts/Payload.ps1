# ==============================================================================
# AutoTask Payload Script V5.45 (Full Notification & Robust Status)
# ------------------------------------------------------------------------------
# V5.45:
#   1. [Fix] Update-TaskStatus 加入重試與強制日期更新，解決 Dashboard 狀態卡死問題。
#   2. [Add] 整合 Lib_Discord.ps1，實現全流程狀態通知 (啟動/異常/結束)。
#   3. [Mod] 優化日誌與錯誤處理流程。
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
$ScriptDir = "$WorkDir\Scripts"

# 確保日誌目錄存在
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# 載入 Discord 模組
if (Test-Path "$ScriptDir\Lib_Discord.ps1") { . "$ScriptDir\Lib_Discord.ps1" }

# 日誌函數
function Write-Log {
    param ([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    Write-Host $LogEntry
}

# 狀態更新函數 (增強版：重試 + 強制更新)
function Update-TaskStatus {
    param ([string]$Status)
    $MaxRetries = 5
    $Retry = 0
    $Success = $false
    
    while (-not $Success -and $Retry -lt $MaxRetries) {
        try {
            if (Test-Path $TaskStatusFile) {
                # 這裡不再檢查舊日期，直接讀取並覆蓋為今日日期，確保 Dashboard 顯示正確
                $Json = Get-Content $TaskStatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
                
                $Json.Date = $DateStr
                $Json.Status = $Status
                $Json.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                # 保留或重置重試計數
                if (-not $Json.RetryCount) { $Json | Add-Member -Name "RetryCount" -Value 0 -MemberType NoteProperty }
                
                $Json | ConvertTo-Json -Depth 5 | Set-Content $TaskStatusFile -Encoding UTF8 -Force
                $Success = $true
            }
        } catch {
            Write-Log "更新 TaskStatus 失敗 ($($Retry+1)/$MaxRetries): $_" "WARN"
            Start-Sleep -Milliseconds 500
            $Retry++
        }
    }
}

function Notify {
    param ([string]$Title, [string]$Msg, [string]$Color="Blue")
    if (Get-Command Send-DiscordNotification -ErrorAction SilentlyContinue) {
        Send-DiscordNotification -Title $Title -Message $Msg -Color $Color
    }
}

trap {
    $Err = $_.Exception.Message
    Write-Log "CRASH: $Err" "ERROR"
    Update-TaskStatus "Failed"
    Notify "❌ Payload 腳本崩潰 (Trap)" "錯誤訊息: $Err`nStackTrace: $($_.ScriptStackTrace)" "Red"
    exit 1
}

# 2. 啟動檢查 (Pre-flight Checks)
Write-Log ">>> Payload 啟動 (V5.45)..."

# 計算今日 Key
$Now = Get-Date
if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }

# 重複執行防護
if (Test-Path $LastRunFile) {
    try {
        $LastRunDate = (Get-Content $LastRunFile -Raw).Trim()
        $IsForceRun = Test-Path $ForceRunFlag
        if ($LastRunDate -eq $TodayKey) {
            if ($IsForceRun) {
                Write-Log "⚠️ 存在 ForceRun 標記，強制重跑。" "YELLOW"
                Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "✅ 今日任務已完成 ($TodayKey)。退出。" "GREEN"
                Notify "⚠️ Payload 重複啟動" "檢測到今日任務已完成，自動略過。" "Yellow"
                Start-Sleep 3; exit 0
            }
        }
    } catch {}
}

# 狀態同步：Running
Update-TaskStatus "Running"

# 03:50 等待邏輯
if ($Now.Hour -eq 3 -and $Now.Minute -ge 50) {
    Write-Log "⚠️ 處於重置緩衝期，進入等待..." "WARNING"
    Notify "⏳ 進入跨日等待" "現在時間 03:50+，Payload 將暫停直到 04:00。" "Yellow"
    while ((Get-Date).Hour -ne 4) { Start-Sleep 10 }
    Write-Log ">>> 解除鎖定！" "GREEN"
    $Now = Get-Date
    if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }
}

# --- 配置讀取邏輯 ---
$MapFile = "$WorkDir\Configs\DateConfig.map"
$RawTaskString = "Default"

if (Test-Path $MapFile) {
    $MapContent = Get-Content $MapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$TodayKey=(.*)") {
            $RawTaskString = $Matches[1].Trim()
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
            } elseif ($IsTurbulencePeriod) {
                 if ($WkJson.Turbulence -and $WkJson.Turbulence.$WeekKey) { $RawTaskString = $WkJson.Turbulence.$WeekKey }
                 else { if ($WkJson.$WeekKey) { $RawTaskString = $WkJson.$WeekKey } }
            } else {
                if ($WkJson.$WeekKey) { $RawTaskString = $WkJson.$WeekKey }
            }
        } catch { Write-Log "讀取 WeeklyConfig 失敗: $_" "ERROR" }
    }
}

# 發送正式啟動通知
Notify "🚀 Payload 任務啟動" "日期: $TodayKey`n配置: $RawTaskString" "Blue"

$TaskList = @()
if ($RawTaskString -match ",") { $TaskList = $RawTaskString -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } } else { $TaskList = @($RawTaskString) }

# 3. 準備 BetterGI
$BetterGIPath = "C:\Program Files\BetterGI\BetterGI.exe"
if (-not (Test-Path $BetterGIPath)) { 
    Notify "❌ 致命錯誤" "找不到 BetterGI 執行檔！" "Red"
    exit 1 
}
$BGIDir = Split-Path $BetterGIPath -Parent
$BGILogsDir = Join-Path $BGIDir "log"
Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue

# 推斷 Day 8 等待點
$ExplicitWaitIndex = -1
for ($k = 0; $k -lt $TaskList.Count; $k++) { if ($TaskList[$k] -eq "[WAIT]") { $ExplicitWaitIndex = $k; break } }
$SplitIndex = -1 
if ($IsTurbulenceDay1 -and $ExplicitWaitIndex -lt 0) { if ($TaskList.Count -gt 2) { $SplitIndex = 3 } else { $SplitIndex = 1 } }

# ----------------------------
# 執行迴圈
# ----------------------------
$MaxTaskRetries = 3 

for ($i = 0; $i -lt $TaskList.Count; $i++) {
    $CurrentTask = $TaskList[$i]
    
    # WAIT 處理
    $NeedWait = $false
    if ($IsTurbulenceDay1) {
        if ($ExplicitWaitIndex -ge 0) { if ($i -eq $ExplicitWaitIndex) { $NeedWait = $true } } elseif ($i -eq $SplitIndex) { $NeedWait = $true }
    }
    if ($NeedWait) {
        Write-Log "=== 進入 10:00 等待模式 ===" "YELLOW"
        Notify "⏳ 暫停任務" "正在等待時間到達 10:00 (Day 8 機制)..." "Yellow"
        $TargetTime = $Now.Date.AddHours(10)
        while ((Get-Date) -lt $TargetTime) {
            if ((Get-Date).Hour -ge 14) { break }
            Start-Sleep 30 
        }
        Notify "▶️ 恢復任務" "時間已達，繼續執行後續配置。" "Green"
    }

    if ($CurrentTask -eq "[WAIT]") { continue }

    # 重試迴圈
    $RetryCount = 0
    $TaskSuccess = $false

    while ($RetryCount -lt $MaxTaskRetries -and -not $TaskSuccess) {
        Write-Log "啟動 BetterGI: $CurrentTask (Attempt $($RetryCount + 1))"
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        
        $Process = Start-Process -FilePath $BetterGIPath -ArgumentList "--startOneDragon `"$CurrentTask`"" -WorkingDirectory $BGIDir -PassThru
        Start-Sleep 20 
        
        $CurrentBGILogPath = ""
        if (Test-Path $BGILogsDir) {
            $LatestLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($LatestLog) { $CurrentBGILogPath = $LatestLog.FullName }
        }

        # 監控
        $StuckThresholdMinutes = 15
        while ($true) {
            if ($Process.HasExited) { $TaskSuccess = $true; break }
            $CheckTime = Get-Date
            
            # 死線
            if ($CheckTime.Hour -eq 3 -and $CheckTime.Minute -ge 50) {
                 Stop-Process -Id $Process.Id -Force
                 Write-Log "⚠️ 03:50 死線觸發。" "RED"
                 Update-TaskStatus "ForceEnd"
                 Notify "⛔ 強制中止" "觸發 03:50 死線，為防止跨日重置，強制停止任務。" "Red"
                 exit 0 
            }

            # 卡死偵測
            if ($CurrentBGILogPath -and (Test-Path $CurrentBGILogPath)) {
                $LogFileItem = Get-Item $CurrentBGILogPath
                if (($CheckTime - $LogFileItem.LastWriteTime).TotalMinutes -gt $StuckThresholdMinutes) {
                    # 雙重確認
                    $ReCheckLog = Get-ChildItem $BGILogsDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($ReCheckLog -and $ReCheckLog.FullName -ne $CurrentBGILogPath) {
                        Write-Log "切換日誌目標 -> $($ReCheckLog.Name)" "GREEN"
                        $CurrentBGILogPath = $ReCheckLog.FullName
                        continue 
                    } else {
                        Write-Log "判定真卡死。" "RED"
                        Stop-Process -Id $Process.Id -Force
                        $RetryCount++
                        if ($RetryCount -lt $MaxTaskRetries) {
                            Notify "🔄 任務卡死重試" "任務: $CurrentTask`n嘗試重啟 ($RetryCount/$MaxTaskRetries)..." "Orange"
                            break # 重試
                        } else {
                            Notify "❌ 任務失敗" "任務: $CurrentTask 已達最大重試次數，放棄執行。" "Red"
                            Update-TaskStatus "Failed"
                            New-Item -ItemType File -Path "$FlagDir\Fail.flag" -Force | Out-Null
                            exit 1
                        }
                    }
                }
            }
            Start-Sleep 10
        }
        if ($TaskSuccess) { break }
    }

    if (-not $TaskSuccess) {
        Notify "❌ 任務異常終止" "Payload 內部錯誤: 任務 $CurrentTask 未能成功完成。" "Red"
        Update-TaskStatus "Failed"
        exit 1
    }
    if ($i -lt ($TaskList.Count - 1)) {
        Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 5
    }
}

# 4. 結算
Write-Log "Payload 執行結束。"
New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
Set-Content -Path $LastRunFile -Value $TodayKey
Update-TaskStatus "Success"
shutdown.exe /l /f