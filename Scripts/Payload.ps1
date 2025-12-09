# ==============================================================================
# AutoTask Payload Script V5.32 (Multi-Config Support)
# ------------------------------------------------------------------------------
# V5.32: 支援讀取 DateConfig.map 中的多重配置 (以逗號分隔)。
#        Day 8 邏輯更新: Task 1 取第一個配置，Task 2 取第二個配置 (若無則回退預設)。
# V5.31: 新增紊亂期 Day 1 (週三) 的雙重排程邏輯。
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

# 錯誤捕捉 Wrapper
trap {
    Write-Log "CRASH: $($_.Exception.Message)" "ERROR"
    Write-Log "StackTrace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}

# 2. 啟動與跨日檢查 (Smart Wait)
Write-Log ">>> Payload 啟動 (V5.32 - Multi-Config)..."

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

# --- [V5.32] 讀取 DateConfig.map 並解析多重任務 ---
$MapFile = "$WorkDir\Configs\DateConfig.map"
$TaskName = "Default"
$Task2Override = $null

if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") } else { $TodayKey = $Now.ToString("yyyyMMdd") }
Write-Log "計算日期 Key: $TodayKey"

if (Test-Path $MapFile) {
    $MapContent = Get-Content $MapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$TodayKey=(.*)") {
            $RawVal = $Matches[1]
            if ($RawVal -match ",") {
                # 偵測到多重配置 (例如 "TaskA,TaskB")
                $Parts = $RawVal -split ","
                $TaskName = $Parts[0].Trim()
                if ($Parts.Count -gt 1) { $Task2Override = $Parts[1].Trim() }
                Write-Log "偵測到多重配置: Task1=[$TaskName], Task2=[$Task2Override]"
            } else {
                $TaskName = $RawVal.Trim()
            }
            break
        }
    }
}
Write-Log "Task 1 (Primary): [$TaskName]"

# 3. 準備 BetterGI 執行環境
$BetterGIPath = "C:\AutoTask\BetterGI\BetterGI.exe" 
$BetterGILogPath = "$WorkDir\Logs\BetterGI\BetterGI.log"
Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue

# --- [雙重排程邏輯] ---
# 判斷是否為紊亂期 Day 1 (Cycle Offset 7.0 ~ 8.0)
$RefDate = [datetime]"2024-08-28T00:00:00"
$CycleOffset = ($Now - $RefDate).TotalDays % 42
if ($CycleOffset -lt 0) { $CycleOffset += 42 }

$IsTurbulenceDay1 = ($CycleOffset -ge 7.0 -and $CycleOffset -lt 8.0)
if ($IsTurbulenceDay1) { Write-Log "📅 偵測到紊亂期首日 (Day 1 - Wednesday)，啟用雙重排程機制。" "MAGENTA" }

# ----------------------------
# 執行 Task 1 (Primary)
# ----------------------------
Write-Log "啟動 BetterGI (Task 1): $TaskName"
$Args1 = "-start -task `"$TaskName`""
$Process1 = Start-Process -FilePath $BetterGIPath -ArgumentList $Args1 -WorkingDirectory (Split-Path $BetterGIPath) -PassThru

# 監控 Loop (Task 1)
$TimeoutMinutes = 180
$StartTime = Get-Date
while ($true) {
    if ($Process1.HasExited) { Write-Log "Task 1 執行程序已結束。"; break }
    if ((Get-Date) - $StartTime).TotalMinutes -gt $TimeoutMinutes {
        Stop-Process -Id $Process1.Id -Force -ErrorAction SilentlyContinue; break
    }
    Start-Sleep 10
}
Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue # Task 1 結束後清理遊戲

# ----------------------------
# 執行 Task 2 (Secondary - if Day 1)
# ----------------------------
if ($IsTurbulenceDay1) {
    Write-Log "準備執行 Task 2 (10:00+)..." 
    
    # A. 等待至 10:00
    $TargetTime = $Now.Date.AddHours(10) # 當天 10:00
    while ((Get-Date) -lt $TargetTime) {
        $Diff = $TargetTime - (Get-Date)
        Write-Host "⏳ 等待活動開放 (10:00)... 剩餘 $($Diff.Minutes) 分鐘" -NoNewline -ForegroundColor Yellow
        Start-Sleep 30
        if ((Get-Date).Hour -ge 14) { break } 
    }
    Write-Log "`n時間已達 10:00，準備啟動 Task 2。" "GREEN"

    # B. 決定 Task 2 配置 (優先使用 DateConfig 的第二參數，否則使用 WeeklyConfig)
    $Task2Name = $null
    
    if ($Task2Override) {
        $Task2Name = $Task2Override
        Write-Log "使用 DateConfig 指定的 Task 2: [$Task2Name]"
    } else {
        if (Test-Path $WeeklyConfFile) {
            try {
                $WkJson = Get-Content $WeeklyConfFile -Raw | ConvertFrom-Json
                if ($WkJson.Turbulence -and $WkJson.Turbulence.Wednesday) {
                    $Task2Name = $WkJson.Turbulence.Wednesday
                    Write-Log "使用 WeeklyConfig 預設的 Task 2: [$Task2Name]"
                }
            } catch { Write-Log "讀取 WeeklyConfig 失敗: $_" "ERROR" }
        }
    }

    if ($Task2Name) {
        Write-Log "啟動 BetterGI (Task 2): $Task2Name"
        $Args2 = "-start -task `"$Task2Name`""
        $Process2 = Start-Process -FilePath $BetterGIPath -ArgumentList $Args2 -WorkingDirectory (Split-Path $BetterGIPath) -PassThru
        
        # 監控 Loop (Task 2)
        $StartTime2 = Get-Date
        while ($true) {
            if ($Process2.HasExited) { Write-Log "Task 2 執行程序已結束。"; break }
            if ((Get-Date) - $StartTime2).TotalMinutes -gt $TimeoutMinutes {
                Stop-Process -Id $Process2.Id -Force -ErrorAction SilentlyContinue; break
            }
            Start-Sleep 10
        }
        Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
    } else {
        Write-Log "⚠️ 無法獲取 Task 2 配置名稱，跳過執行。" "WARN"
    }
}

# 4. 寫入完成並登出
Write-Log "Payload 執行結束 (Tasks Completed)，建立標記並登出..."
New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
Set-Content -Path "$WorkDir\Configs\LastRun.log" -Value $TodayKey
shutdown.exe /l /f