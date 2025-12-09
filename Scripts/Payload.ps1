# ==============================================================================
# AutoTask Payload Script V5.29 (ForceEnd with 'forceend' Task)
# ------------------------------------------------------------------------------
# 職責: 在 RDP 遠端桌面會話中運行，負責調度 BetterGI 執行遊戲自動化。
# V5.29: 優化跨日保護，03:50 觸發時執行 "forceend" 任務進行優雅收尾。
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\Payload_$DateStr.log"
$FlagDir = "$WorkDir\Flags"
$DoneFlag = "$FlagDir\Done.flag"

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

# 2. 檢查是否需要執行
Write-Log ">>> Payload 啟動 (V5.29 - ForceEnd Task)..."

# 讀取 EnvConfig
$EnvConfigFile = "$WorkDir\Configs\EnvConfig.json"
if (Test-Path $EnvConfigFile) {
    $EnvConfig = Get-Content -Path $EnvConfigFile -Raw | ConvertFrom-Json
    $GenshinPath = $EnvConfig.GenshinPath
} else {
    Write-Log "找不到 EnvConfig.json，使用預設路徑。" "WARN"
    $GenshinPath = "C:\Program Files\HoYoPlay\games\Genshin Impact Game"
}

# 讀取 DateConfig.map
$MapFile = "$WorkDir\Configs\DateConfig.map"
$TaskName = "Default"

# 計算今日 (原神 04:00 換日邏輯)
# 若現在是 00:00 - 03:59，則視為「前一天」
$Now = Get-Date
if ($Now.Hour -lt 4) {
    $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd")
} else {
    $TodayKey = $Now.ToString("yyyyMMdd")
}

Write-Log "計算日期 Key: $TodayKey (當前時間: $($Now.ToString('HH:mm')))"

if (Test-Path $MapFile) {
    $MapContent = Get-Content $MapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$TodayKey=(.*)") {
            $TaskName = $Matches[1]
            break
        }
    }
}

Write-Log "今日任務目標: [$TaskName]"

# 3. 啟動 BetterGI
$BetterGIPath = "C:\AutoTask\BetterGI\BetterGI.exe" 
$BetterGILogPath = "$WorkDir\Logs\BetterGI\BetterGI.log"

# 殺死殘留進程
Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue

# 啟動參數
$Args = "-start -task `"$TaskName`""
Write-Log "啟動 BetterGI: $Args"

try {
    Start-Process -FilePath $BetterGIPath -ArgumentList $Args -WorkingDirectory (Split-Path $BetterGIPath)
} catch {
    Write-Log "無法啟動 BetterGI: $($_.Exception.Message)" "ERROR"
    exit 1
}

# 4. 監控迴圈 (Monitor Loop)
$TimeoutMinutes = 180 # 3小時超時
$StartTime = Get-Date

Write-Log "進入監控模式..."

while ($true) {
    $CurrentTime = Get-Date
    
    # --------------------------------------------------------------------------
    # [CRITICAL UPDATE] 03:50 跨日收尾流程 (ForceEnd Protocol)
    # --------------------------------------------------------------------------
    # 只要時間進入 03:50 ~ 03:59 區間，啟動 "forceend" 任務進行收尾。
    if ($CurrentTime.Hour -eq 3 -and $CurrentTime.Minute -ge 50) {
        Write-Log "⚠️ [ForceEnd] 時間已達 03:50 ($($CurrentTime.ToString('HH:mm:ss')))，啟動 'forceend' 收尾流程..." "WARNING"

        # 4.1 停止當前正在運行的主要任務 (釋放資源)
        Write-Log "中止當前任務，準備切換..."
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        # 4.2 啟動 ForceEnd 任務
        $ForceEndTask = "forceend"
        $ForceEndArgs = "-start -task `"$ForceEndTask`""
        Write-Log "啟動 BetterGI 收尾任務: $ForceEndTask (預計耗時 5 分鐘)"
        
        try {
            Start-Process -FilePath $BetterGIPath -ArgumentList $ForceEndArgs -WorkingDirectory (Split-Path $BetterGIPath)
        } catch {
            Write-Log "無法啟動 ForceEnd: $($_.Exception.Message)" "ERROR"
            # 若無法啟動，直接跳去登出
        }

        # 4.3 進入 ForceEnd 專用監控迴圈 (直到 03:59:30 或任務結束)
        # 設定硬性死線：03:59:30 (保留30秒登出緩衝)
        $ForceEndHardLimit = Get-Date -Hour 3 -Minute 59 -Second 30
        
        Write-Log "等待收尾任務完成 (硬性截止時間: 03:59:30)..."
        
        while ($true) {
            $SubTime = Get-Date
            
            # (A) 硬性死線檢查
            if ($SubTime -ge $ForceEndHardLimit) {
                Write-Log "⚠️ [ForceEnd] 已達硬性截止時間 (03:59:30)，強制中斷收尾！" "WARNING"
                break
            }

            # (B) 檢查 BetterGI 是否自行結束 (視為任務完成)
            $BGI = Get-Process -Name "BetterGI" -ErrorAction SilentlyContinue
            if (!$BGI) {
                Write-Log "[ForceEnd] BetterGI 進程已結束，視為收尾完成。"
                break
            }
            
            # (C) 檢查日誌是否顯示完成 (如果 BGI 沒關閉)
            if (Test-Path $BetterGILogPath) {
                 # 嘗試讀取最後 30 行
                 $LastLogs = Get-Content $BetterGILogPath -Tail 30 -ErrorAction SilentlyContinue
                 if ($LastLogs -match "全部任务已结束") {
                     Write-Log "[ForceEnd] 偵測到日誌: '全部任务已结束'。"
                     break
                 }
            }
            
            Start-Sleep -Seconds 5
        }

        # 4.4 最終清理與登出
        Write-Log "執行最終清理與登出 (Logoff)..."
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "YuanShen" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        
        # 不建立 Done.flag，因為這不算完成今日目標，只是收尾。
        # 04:05 Master 再次喚醒時，將會執行新的一天真正的任務。
        shutdown.exe /l /f
        exit
    }
    # --------------------------------------------------------------------------

    # 一般任務監控邏輯
    $BGIProcess = Get-Process -Name "BetterGI" -ErrorAction SilentlyContinue
    if (!$BGIProcess) {
        Write-Log "BetterGI 進程已結束。"
        if (Test-Path $BetterGILogPath) {
            $LastLogs = Get-Content $BetterGILogPath -Tail 20
            if ($LastLogs -match "全部任务已结束") {
                Write-Log "檢測到任務成功完成。"
                New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
                Set-Content -Path "$WorkDir\Configs\LastRun.log" -Value $TodayKey
            } else {
                Write-Log "BetterGI 異常退出 (未見成功訊息)。" "ERROR"
            }
        }
        break
    }

    # 檢查超時
    if (($CurrentTime - $StartTime).TotalMinutes -gt $TimeoutMinutes) {
        Write-Log "任務執行超時 ($TimeoutMinutes 分鐘)，強制終止。" "ERROR"
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "YuanShen" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        break
    }

    Start-Sleep -Seconds 10
}

# 5. 結束與登出
Write-Log "Payload 執行結束，執行登出..."
shutdown.exe /l /f