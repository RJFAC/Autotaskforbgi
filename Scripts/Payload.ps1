# ==============================================================================
# AutoTask Payload Script V5.48 (ForceEnd Sequence Logic)
# ------------------------------------------------------------------------------
# V5.48:
#   1. [Feature] 實作智慧收尾邏輯：
#      - 03:45 若未完成 -> 啟動 "forceend" 配置。
#      - 03:55 若 "forceend" 未完成 -> 強制殺進程。
#   2. [Refactor] 移除舊版單純的 03:50 死線檢查。
#   3. [Fix] 補全 Log Watchdog 完整邏輯 (雙重確認與原地重試)。
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\Payload_$DateStr.log"
$FlagDir = "$WorkDir\Flags"
$DoneFlag = "$FlagDir\Done.flag"
$FailFlag = "$FlagDir\Fail.flag"
$WeeklyConfFile = "$WorkDir\Configs\WeeklyConfig.json"
$TaskStatusFile = "$WorkDir\Configs\TaskStatus.json"
$LastRunFile = "$WorkDir\Configs\LastRun.log"
$ForceRunFlag = "$FlagDir\ForceRun.flag"
$ScriptDir = "$WorkDir\Scripts"

# 確保日誌目錄存在
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# 載入 Discord 模組 (兼容相對路徑)
$LibPath = "$ScriptDir\Lib_Discord.ps1"
if (Test-Path $LibPath) { 
    . $LibPath 
} else {
    function Write-Log { param($Msg, $Color="Cyan") Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $Msg" -ForegroundColor $Color }
    function Send-DiscordNotification { param($Title, $Message, $Color) Write-Host "[Mock Discord] $Title - $Message" }
}

# 重新導向輸出至 Log
Start-Transcript -Path $LogFile -Append -Force

Write-Log ">>> Payload 啟動 (V5.48 - ForceEnd Sequence)..." "Green"

# ------------------------------------------------------------------------------
# 函數: 更新狀態
# ------------------------------------------------------------------------------
function Update-TaskStatus {
    param([string]$Status)
    $Data = @{
        Date = $DateStr
        Status = $Status
        LastUpdate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    try {
        $Data | ConvertTo-Json | Set-Content $TaskStatusFile -Encoding UTF8
    } catch {
        Write-Log "⚠️ 更新狀態檔失敗: $_" "Yellow"
    }
}

# ------------------------------------------------------------------------------
# 函數: 執行 BetterGI 任務
# ------------------------------------------------------------------------------
function Run-BetterGI {
    param([string]$ConfigName)
    
    $BGIPath = "C:\Program Files\BetterGI\BetterGenshinImpact.exe"
    if (-not (Test-Path $BGIPath)) {
        Write-Log "❌ 找不到 BetterGI: $BGIPath" "Red"
        return $false
    }

    $ArgList = "--startOneDragon `"$ConfigName`""
    Write-Log "啟動 BetterGI: $ConfigName" "Cyan"
    
    $Proc = Start-Process -FilePath $BGIPath -ArgumentList $ArgList -PassThru
    return $Proc
}

# ------------------------------------------------------------------------------
# 函數: 測試 JSON 檔案 (SSOT V3.7 要求)
# ------------------------------------------------------------------------------
function Test-JsonFile {
    param($Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $Content = Get-Content $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($Content)) { return $false }
        $null = $Content | ConvertFrom-Json -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# 2. 啟動檢查
Update-TaskStatus "Running"
New-Item -Path "$FlagDir\Payload.flag" -ItemType File -Force | Out-Null

# 檢查重複執行
if (Test-Path $LastRunFile) {
    $LastRunDate = (Get-Content $LastRunFile -ErrorAction SilentlyContinue).Trim()
    if ($LastRunDate -eq $DateStr -and -not (Test-Path $ForceRunFlag)) {
        Write-Log "今日任務已完成 (LastRun match)，且無強制標記。退出。" "Yellow"
        Send-DiscordNotification -Title "⚠️ 任務跳過" -Message "今日 ($DateStr) 任務紀錄已存在。" -Color "Yellow"
        Stop-Transcript
        exit 0
    }
}
if (Test-Path $ForceRunFlag) { Remove-Item $ForceRunFlag -Force }

# 3. 任務解析
$TaskList = @()
$DateConfigPath = "$WorkDir\Configs\DateConfig.map"

# A. 優先讀取 DateConfig.map
if (Test-Path $DateConfigPath) {
    $MapContent = Get-Content $DateConfigPath
    foreach ($line in $MapContent) {
        if ($line -match "^$DateStr=(.*)") {
            $TaskList = $Matches[1].Split(",")
            Write-Log "使用 DateConfig 配置。" "Cyan"
            break
        }
    }
}

# B. Fallback: WeeklyConfig.json
if ($TaskList.Count -eq 0) {
    if (Test-JsonFile $WeeklyConfFile) {
        try {
            $WeeklyJson = Get-Content $WeeklyConfFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $DayOfWeek = (Get-Date).DayOfWeek.ToString()
            if ($WeeklyJson.$DayOfWeek) {
                $TaskList = $WeeklyJson.$DayOfWeek.Split(",")
                Write-Log "使用 WeeklyConfig ($DayOfWeek) 配置。" "Cyan"
                
                # Day 8 紊亂期演算法 (Turbulence)
                $BaseDate = Get-Date -Date "2024-11-20"
                $DiffDays = ((Get-Date) - $BaseDate).Days
                if ($DiffDays % 42 -eq 8) {
                    Write-Log "🌊 偵測到紊亂爆發期 (Day 8)，注入 [WAIT] 與 Turbulence 任務。" "Magenta"
                    # 注入邏輯：一般任務 -> 等待 -> 紊亂任務
                    $TurbulenceTask = if ($WeeklyJson.Turbulence.$DayOfWeek) { $WeeklyJson.Turbulence.$DayOfWeek } else { "每日任務" }
                    $TaskList = $TaskList + @("[WAIT]", $TurbulenceTask)
                }
            }
        } catch {
            Write-Log "WeeklyConfig 解析失敗，使用預設值。" "Red"
        }
    }
}

# C. Default
if ($TaskList.Count -eq 0) { $TaskList = @("每日任務") }

Write-Log "今日任務清單: $($TaskList -join ', ')" "Cyan"
Send-DiscordNotification -Title "🚀 任務啟動" -Message "配置: $($TaskList -join ', ')" -Color "Blue"

# 4. 執行迴圈 (含 ForceEnd 邏輯)
$AllSuccess = $true

foreach ($TaskName in $TaskList) {
    
    # === [特殊標記] WAIT 模式 ===
    if ($TaskName -eq "[WAIT]") {
        Write-Log "遇到 [WAIT] 標記，檢查時間..." "Yellow"
        $WaitUntil = (Get-Date).Date.AddHours(10).AddMinutes(5) # 10:05
        if ((Get-Date) -lt $WaitUntil) {
            Write-Log "時間早於 10:05，進入等待模式..." "Yellow"
            Send-DiscordNotification -Title "⏳ 進入等待" -Message "等待伺服器刷新 (10:05)..." -Color "Orange"
            
            while ((Get-Date) -lt $WaitUntil) {
                Start-Sleep 60
                # 等待期間仍需檢查 ForceEnd
                if ((Get-Date).Hour -eq 3 -and (Get-Date).Minute -ge 45) { break } 
            }
            Send-DiscordNotification -Title "▶️ 恢復執行" -Message "等待結束，繼續任務。" -Color "Green"
        } else {
            Write-Log "時間已過 10:05，略過等待。" "Cyan"
        }
        continue
    }

    # === [關鍵修改] 03:45 ForceEnd 檢查點 ===
    $Now = Get-Date
    if ($Now.Hour -eq 3 -and $Now.Minute -ge 45) {
        Write-Log "⚠️ [ForceEnd] 時間已達 03:45，前序任務未完成。" "Yellow"
        Write-Log "🛑 觸發收尾流程：停止當前任務，轉為執行 'forceend'。" "Yellow"
        Send-DiscordNotification -Title "🛑 觸發強制收尾" -Message "時間 03:45，切換至 forceend 配置。" -Color "Orange"

        # 1. 強制關閉當前所有遊戲相關進程
        Stop-Process -Name "BetterGenshinImpact", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 5

        # 2. 啟動 forceend 配置
        $ForceProc = Run-BetterGI "forceend"
        
        # 3. 監控 forceend 直到 03:55
        while ($true) {
            if ($ForceProc.HasExited) {
                Write-Log "✅ 'forceend' 配置執行完畢。" "Green"
                break
            }

            # 檢查是否到達 03:55 死線
            $CheckTime = Get-Date
            if ($CheckTime.Hour -eq 3 -and $CheckTime.Minute -ge 55) {
                Write-Log "⏰ [Deadline] 時間已達 03:55，'forceend' 未能完成。" "Red"
                Write-Log "🛑 強制終止所有程序以保護隔日排程。" "Red"
                Stop-Process -Id $ForceProc.Id -Force -ErrorAction SilentlyContinue
                Send-DiscordNotification -Title "❌ 收尾逾時" -Message "forceend 在 03:55 前未能完成，已強制終止。" -Color "Red"
                break
            }
            Start-Sleep 5
        }

        # 4. 最終清理
        Stop-Process -Name "BetterGenshinImpact", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
        
        # 5. 退出並標記 (視為本次流程結束，等待明天)
        Update-TaskStatus "ForceEnded"
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        $DateStr | Set-Content $LastRunFile -Encoding UTF8
        
        Write-Log ">>> Payload 結束 (ForceEnd Mode)。" "Magenta"
        Stop-Transcript
        exit 0
    }
    # ========================================

    # 正常任務執行 (含 Watchdog 邏輯)
    Write-Log "執行子任務: $TaskName"
    $Proc = Run-BetterGI $TaskName
    $TaskStartTime = Get-Date
    
    # [Watchdog] 初始鎖定日誌
    Start-Sleep 20 # 等待 BGI 生成 Log
    $CurrentBGILogPath = $null
    $BGILogDir = "C:\Program Files\BetterGI\log"
    try {
        $LatestLog = Get-ChildItem "$BGILogDir\log_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($LatestLog) { 
            $CurrentBGILogPath = $LatestLog.FullName
            Write-Log "已鎖定監控日誌: $($LatestLog.Name)" "Cyan"
        }
    } catch {}

    # [Watchdog] 迴圈
    $RetryCount = 0
    $MaxTaskRetries = 3
    $LogStuckThreshold = 15 # 分鐘

    while (-not $Proc.HasExited) {
        if ($CurrentBGILogPath) {
            $LogItem = Get-Item $CurrentBGILogPath
            $IdleMinutes = ((Get-Date) - $LogItem.LastWriteTime).TotalMinutes
            
            if ($IdleMinutes -ge $LogStuckThreshold) {
                Write-Log "⚠️ 警告: 日誌已靜止 $IdleMinutes 分鐘，執行雙重確認..." "Orange"
                
                # 雙重確認: 檢查是否有更新的 Log 檔產生 (Log Rotation)
                $ReCheckLog = Get-ChildItem "$BGILogDir\log_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($ReCheckLog.LastWriteTime -gt $LogItem.LastWriteTime) {
                    Write-Log "發現新日誌文件 $($ReCheckLog.Name)，切換監控目標。" "Green"
                    $CurrentBGILogPath = $ReCheckLog.FullName
                } else {
                    Write-Log "❌ 判定真卡死 (Stuck)。執行原地重試 ($($RetryCount+1)/$MaxTaskRetries)..." "Red"
                    Stop-Process -Id $Proc.Id -Force -ErrorAction SilentlyContinue
                    Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
                    
                    if ($RetryCount -lt $MaxTaskRetries) {
                        $RetryCount++
                        Send-DiscordNotification -Title "🔄 任務重試" -Message "任務 $TaskName 卡死，第 $RetryCount 次重啟。" -Color "Orange"
                        Start-Sleep 5
                        $Proc = Run-BetterGI $TaskName # 重啟
                        # 重置 Log 鎖定
                        Start-Sleep 20
                        try {
                            $LatestLog = Get-ChildItem "$BGILogDir\log_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($LatestLog) { $CurrentBGILogPath = $LatestLog.FullName }
                        } catch {}
                        continue
                    } else {
                        Write-Log "已達最大重試次數，放棄此任務。" "Red"
                        Send-DiscordNotification -Title "❌ 任務失敗" -Message "任務 $TaskName 重試失敗。" -Color "Red"
                        $AllSuccess = $false
                        break 
                    }
                }
            }
        }
        Start-Sleep 10
    }
    
    # 任務間清理
    Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
    Start-Sleep 5
}

# 5. 正常結算
if ($AllSuccess) {
    Write-Log "所有任務執行完畢。" "Green"
    Update-TaskStatus "Success"
} else {
    Write-Log "任務執行完畢，但有部分失敗。" "Red"
    Update-TaskStatus "Failed" # 這裡不設 FailFlag，讓 Master 仍能執行收尾，但知道有錯
}

# 無論如何都建立 DoneFlag 讓 Master 收尾 (ForceEnd 以外的正常結束流程)
New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
$DateStr | Set-Content $LastRunFile -Encoding UTF8

# 6. 登出 (Logoff)
Write-Log "執行登出..."
shutdown /l

Stop-Transcript