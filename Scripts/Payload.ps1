# --- [路徑定義] ---
$BaseDir    = "C:\AutoTask"
$ScriptDir  = "$BaseDir\Scripts"
$ConfigDir  = "$BaseDir\Configs"
$FlagDir    = "$BaseDir\Flags"
$LogDir     = "$BaseDir\Logs"

# --- [全域日誌設定] ---
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8
}

# --- [清理舊日誌] ---
Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue

# =============================================================================
# [優先處理] 清理舊的 Payload 實例
# =============================================================================
try {
    $CurrentPID = $PID
    $TargetScript = "Payload.ps1"
    $OldInstances = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
        Where-Object { $_.CommandLine -like "*$TargetScript*" -and $_.ProcessId -ne $CurrentPID }
    foreach ($proc in $OldInstances) { 
        Write-Log "發現舊實例 (PID: $($proc.ProcessId))，強制終止。" "Yellow"
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue 
    }
} catch {}

# --- [變數定義] ---
$BettergiDir    = "C:\Program Files\BetterGI"
$BettergiExe    = "$BettergiDir\BetterGI.exe"
$LogDirBG       = "$BettergiDir\log"
$ScreenshotDir  = "$BettergiDir\Screenshots" 
$1RemoteLogDir  = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64\.logs"

$TaskStatusFile = "$ConfigDir\TaskStatus.json"
$DoneFlag       = "$FlagDir\Done.flag"
$FailFlag       = "$FlagDir\Fail.flag"
$ForceRunFlag   = "$FlagDir\ForceRun.flag" 
$DateMap        = "$ConfigDir\DateConfig.map"
$WeeklyConf     = "$ConfigDir\WeeklyConfig.json"
$PauseLog       = "$ConfigDir\PauseDates.log"
$LastRunLog     = "$ConfigDir\LastRun.log"
$BackupRootDir  = "$BaseDir\LogBackups"
$NotifyScript   = "$ScriptDir\Notify.ps1"

$MaxRetries = 3
$SuccessKeyword = "一条龙.*任务结束"

Write-Log "Payload 啟動..." "Cyan"

# --- [輔助函數] ---
function Check-Network {
    Write-Log "正在檢查網路連線..." "Cyan"
    $Retry = 0
    $MaxRetry = 12
    while ($Retry -lt $MaxRetry) {
        try {
            $ping = Test-Connection -ComputerName "8.8.8.8" -Count 1 -ErrorAction Stop
            if ($ping.StatusCode -eq 0) {
                Write-Log "網路連線正常。" "Green"
                return $true
            }
        } catch {}
        Write-Log "網路未就緒，等待 5 秒... ($($Retry+1)/$MaxRetry)" "Yellow"
        Start-Sleep 5
        $Retry++
    }
    Write-Log "⚠️ 警告：網路連線逾時 (60秒)，遊戲可能無法登入。" "Red"
    return $false
}

function Send-Notify {
    param([string]$Title, [string]$Msg, [string]$Color)
    if (Test-Path $NotifyScript) {
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$Title`" -Message `"$Msg`" -Color `"$Color`"" -WindowStyle Hidden
    }
}

function Cleanup-Screenshots {
    if (Test-Path $ScreenshotDir) {
        $LimitDate = (Get-Date).AddDays(-30)
        $OldFiles = Get-ChildItem -Path $ScreenshotDir -Include "*.png", "*.jpg", "*.bmp" -Recurse | Where-Object { $_.LastWriteTime -lt $LimitDate }
        if ($OldFiles) {
            $Count = $OldFiles.Count
            $OldFiles | Remove-Item -Force
            Write-Log "已清理 $Count 張過期截圖 (超過30天)。" "Gray"
        }
    }
}

function Backup-Logs {
    Write-Log "正在執行失敗現場日誌備份..." "Magenta"
    try {
        $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $TargetDir = Join-Path $BackupRootDir "Failed_$Timestamp"
        New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$LogDir\*" -Destination (New-Item -Path "$TargetDir\AutoTask_Logs" -ItemType Directory) -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -Path "$LogDirBG\*" -Destination (New-Item -Path "$TargetDir\BetterGI_Logs" -ItemType Directory) -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $1RemoteLogDir) { Copy-Item -Path "$1RemoteLogDir\*" -Destination (New-Item -Path "$TargetDir\1Remote_Logs" -ItemType Directory) -Recurse -Force -ErrorAction SilentlyContinue }
        Write-Log "日誌備份完成。路徑: $TargetDir" "Green"
    } catch { Write-Log "日誌備份失敗: $_" "Red" }
}

function Update-Status ($status, $retry) {
    $obj = @{
        Date = (Get-Date).AddHours(-3).ToString("yyyyMMdd")
        Status = $status
        RetryCount = $retry
        LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
    $obj | ConvertTo-Json | Set-Content $TaskStatusFile
}

function Check-Success-Log ($logPath) {
    try {
        $Logs = Get-Content $logPath -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($Logs -match $SuccessKeyword) { return $true }
    } catch {}
    return $false
}

# --- [核心演算法] ---
# 1. 計算最近的版本更新日
function Get-LastUpdateDate ($CheckDate) {
    $RefDate = [datetime]"2024-08-28" # 5.0 基準日
    $DaysDiff = ($CheckDate.Date - $RefDate).Days
    $Cycles = [math]::Floor($DaysDiff / 42)
    return $RefDate.AddDays($Cycles * 42)
}

# 2. 判斷是否為版本更新日
function Test-GenshinUpdateDay ($CheckDate) {
    $LastUpdate = Get-LastUpdateDate $CheckDate
    return ($LastUpdate.Date -eq $CheckDate.Date)
}

# 3. 判斷是否為紊亂爆發期 (Update + 8 ~ + 18)
function Test-TurbulencePeriod ($CheckDate) {
    $LastUpdate = Get-LastUpdateDate $CheckDate
    $Start = $LastUpdate.AddDays(8)
    $End = $LastUpdate.AddDays(18)
    # 檢查日期是否在區間內
    return ($CheckDate.Date -ge $Start.Date -and $CheckDate.Date -le $End.Date)
}

# 4. 決定配置 (整合所有規則)
function Get-TargetConfig {
    $today = (Get-Date).AddHours(-3)
    $dStr = $today.ToString("yyyyMMdd")
    $weekDay = $today.DayOfWeek.ToString()
    
    # A. 指定日期 (最高優先)
    if (Test-Path $DateMap) {
        $map = Get-Content $DateMap
        foreach ($line in $map) { if ($line -match "^$dStr=(.+)$") { return $matches[1] } }
    }
    
    $wk = $null
    if (Test-Path $WeeklyConf) {
        try { $wk = Get-Content $WeeklyConf -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue } catch {}
    }
    
    if ($wk) {
        # B. 紊亂爆發期 (次優先)
        if (Test-TurbulencePeriod $today) {
             if ($wk.Turbulence -and $wk.Turbulence.$weekDay) {
                 Write-Log "今日為紊亂爆發期 ($weekDay)，使用特殊配置。" "Cyan"
                 return $wk.Turbulence.$weekDay
             }
        }
        # C. 一般每週配置
        return $wk.$weekDay
    }
    return "day"
}

# =============================================================================
# [前日殘留處理] (03:50 ~ 04:00)
# =============================================================================
$CurrentTime = Get-Date
$TodayLimit = $CurrentTime.Date.AddHours(4) 
if ($CurrentTime.Hour -ge 4) { $TodayLimit = $TodayLimit.AddDays(1) }

if ($CurrentTime -lt $TodayLimit -and $CurrentTime.Hour -lt 4) {
    Check-Network

    $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue
    $GiProc = Get-Process "GenshinImpact" -ErrorAction SilentlyContinue

    if ($BgProc -or $GiProc) {
        Write-Log "偵測到前一日殘留，執行 forceend..." "Yellow"
        Send-Notify -Title "系統維護" -Msg "偵測到前日殘留，正在執行 ForceEnd 清理..." -Color "Yellow"
        
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3

        $ForceArgs = "--startOneDragon ""forceend"""
        Write-Log "啟動 ForceEnd 配置..."
        $ForceProc = Start-Process -FilePath $BettergiExe -ArgumentList $ForceArgs -WorkingDirectory $BettergiDir -PassThru

        while (-not $ForceProc.HasExited) {
            if ((Get-Date) -ge $TodayLimit) {
                Write-Log "時間到 04:00，強制殺掉 forceend 進程！" "Red"
                $ForceProc.Kill()
                break
            }
            Start-Sleep 2
            $ForceProc.Refresh()
        }
        Write-Log "ForceEnd 階段結束，執行最終清理。"
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 2
    }
}

# =============================================================================
# [等待 04:00]
# =============================================================================
$TargetTime = (Get-Date).Date.AddHours(4)
if ($TargetTime -lt (Get-Date)) { $TargetTime = $TargetTime.AddDays(1) }

if ((Get-Date).Hour -lt 4) {
    while ((Get-Date) -lt $TargetTime) {
        $WaitSpan = $TargetTime - (Get-Date)
        Write-Host "[PAYLOAD] 等待 04:00... 剩餘 $($WaitSpan.Minutes) 分 $($WaitSpan.Seconds) 秒"
        Start-Sleep 30
    }
}

# =============================================================================
# [今日排程檢查]
# =============================================================================
$CurrentDateObj = (Get-Date).AddHours(-3)
$CurrentDateStr = $CurrentDateObj.ToString("yyyyMMdd")
Write-Log "今日日期 (計算結果): $CurrentDateStr"

Cleanup-Screenshots

$IsForceRun = $false
if (Test-Path $ForceRunFlag) {
    Write-Log "偵測到 ForceRun 旗標，將無視暫停與重複執行檢查。" "Magenta"
    $IsForceRun = $true
    Remove-Item $ForceRunFlag -Force
}

# [新功能] 版本更新日特別邏輯
$IsUpdateDay = Test-GenshinUpdateDay $CurrentDateObj
$UpdateResumeTime = $CurrentDateObj.Date.AddHours(11).AddMinutes(30) # 11:30

if ($IsUpdateDay -and -not $IsForceRun) {
    if ((Get-Date) -lt $UpdateResumeTime) {
        Write-Log "⚠️ 今日為版本更新日！進入維護待機模式。" "Magenta"
        Send-Notify -Title "版本更新" -Msg "系統進入待機，預計 11:30 恢復運行。" -Color "Yellow"
        
        # [新功能] 預下載啟動 (Req 4)
        # 嘗試尋找啟動器 (假設在上一層目錄的 Launcher 下)
        $LauncherPath = Join-Path (Split-Path (Split-Path $BettergiDir)) "Genshin Impact Game\YuanShen.exe" 
        # 修正：通常啟動器是 launcher.exe，遊戲本體是 YuanShen.exe
        # 這裡直接嘗試啟動遊戲本體，通常會觸發更新檢查
        # 或者 BetterGI 啟動時會自動處理？假設這裡我們只做等待。
        # 為了滿足「提早自動完成預下載」，我們嘗試啟動 BetterGI 一次，讓它去撞更新
        Write-Log "嘗試啟動 BetterGI 以觸發預下載..."
        Start-Process -FilePath $BettergiExe -WorkingDirectory $BettergiDir
        Start-Sleep 300 # 給它 5 分鐘下載
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue

        # 進入長等待
        while ((Get-Date) -lt $UpdateResumeTime) {
            $Diff = $UpdateResumeTime - (Get-Date)
            Write-Host "等待維護結束... 剩餘 $($Diff.Hours)時$($Diff.Minutes)分"
            Start-Sleep 60
        }
        Write-Log "維護時間已過，準備執行。" "Green"
    }
}

# 一般暫停檢查
if (-not $IsForceRun -and (Test-Path $PauseLog)) {
    if ((Get-Content $PauseLog) -contains $CurrentDateStr) {
        Write-Log "今日排程暫停。執行清理並登出。" "Yellow"
        Update-Status "Paused" 0
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        Start-Sleep 3
        logoff
        exit
    }
}

if (-not $IsForceRun -and (Test-Path $LastRunLog)) {
    if ((Get-Content $LastRunLog) -eq $CurrentDateStr) {
        Update-Status "Success" 0
        Write-Log "今日任務已標記為完成，無需重複執行。"
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        Start-Sleep 3
        logoff
        exit
    }
}

# =============================================================================
# [執行今日任務]
# =============================================================================
$RetryCount = 0
$ConfigName = Get-TargetConfig
$ConfigQueue = $ConfigName -split ","

if (-not (Test-Path $BettergiExe)) {
    Write-Log "嚴重錯誤：找不到 BetterGI 執行檔 ($BettergiExe)" "Red"
    Send-Notify -Title "執行失敗" -Msg "找不到 BetterGI 執行檔！" -Color "Red"
    exit
}

Check-Network

while ($RetryCount -le $MaxRetries) {
    Update-Status "Running" $RetryCount
    $AllConfigSuccess = $true
    $CurrentQueueIndex = 0
    $ConfigFrequency = @{}
    $CurrentConfigRunCount = @{} 
    foreach ($c in $ConfigQueue) { if (-not [string]::IsNullOrWhiteSpace($c)) { $ConfigFrequency[$c]++ } }

    foreach ($CurrentConfig in $ConfigQueue) {
        if ([string]::IsNullOrWhiteSpace($CurrentConfig)) { continue }
        $CurrentQueueIndex++
        $CurrentConfigRunCount[$CurrentConfig]++
        
        $LogMsg = ">>> 執行配置: [$CurrentConfig] (進度: $CurrentQueueIndex/$($ConfigQueue.Count))"
        if ($ConfigFrequency[$CurrentConfig] -gt 1) { $LogMsg += " (第 $($CurrentConfigRunCount[$CurrentConfig]) 次)" }
        Write-Log $LogMsg "Cyan"

        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3

        $Args = "--startOneDragon ""$CurrentConfig"""
        Start-Process -FilePath $BettergiExe -ArgumentList $Args -WorkingDirectory $BettergiDir
        
        $WatchdogStart = Get-Date
        $IsSuccess = $false
        $IsFailed = $false
        
        # [修正] 心跳超時設定
        # 若是更新日，放寬到 60 分鐘 (Req 3)
        $HeartbeatLimit = if ($IsUpdateDay) { 60 } else { 15 }
        
        $LogFile = $null
        for ($i=0; $i -lt 90; $i++) {
            $Candidate = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($Candidate -and $Candidate.LastWriteTime -gt (Get-Date).AddMinutes(-5)) {
                $LogFile = $Candidate
                break 
            }
            Start-Sleep 1
        }

        if (-not $LogFile) {
            Write-Log "錯誤：BetterGI 啟動失敗 (日誌超時)" "Red"
            $IsFailed = $true
        } else {
            Write-Log "鎖定日誌: $($LogFile.Name)"
            while (-not $IsSuccess -and -not $IsFailed) {
                Start-Sleep 5
                
                # 1. 檢查成功
                if (Check-Success-Log $LogFile.FullName) {
                    Write-Log "配置 [$CurrentConfig] 執行完成！" "Green"
                    $IsSuccess = $true
                    break
                }
                
                # 2. 檢查進程 + 二次確認
                if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) { 
                    Write-Log "BetterGI 進程消失，等待 3 秒後確認..." "Yellow"
                    Start-Sleep 3
                    if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                        if (Check-Success-Log $LogFile.FullName) {
                            Write-Log "確認配置 [$CurrentConfig] 已完成 (進程結束後)。" "Green"
                            $IsSuccess = $true
                        } else {
                            Write-Log "配置 [$CurrentConfig] 意外中斷！" "Red"
                            $IsFailed = $true
                        }
                        break 
                    }
                }
                
                # 3. 心跳檢查
                $LogFile.Refresh()
                if (((Get-Date) - $LogFile.LastWriteTime).TotalMinutes -gt $HeartbeatLimit) { 
                    if (Check-Success-Log $LogFile.FullName) { $IsSuccess = $true; break }
                    Write-Log "⚠️ 警報：日誌超過 $HeartbeatLimit 分鐘未更新，判定卡死！" "Red"
                    $IsFailed = $true 
                }
            }
        }

        if (-not $IsSuccess) {
            $AllConfigSuccess = $false
            Write-Log "配置 [$CurrentConfig] 失敗，準備重試整個佇列。" "Yellow"
            break 
        } else {
            Start-Sleep 5
        }
    }

    if ($AllConfigSuccess) {
        $Duration = New-TimeSpan -Start $StartTime -End (Get-Date)
        Write-Log ">>> 所有排程完成。總耗時: {0:hh}時{0:mm}分" -f $Duration "Green"
        Send-Notify -Title "任務成功" -Msg "配置 [$ConfigName] 已完成。" -Color "Green"

        Update-Status "Success" $RetryCount
        Set-Content $LastRunLog -Value $CurrentDateStr
        while (Get-Process "GenshinImpact" -ErrorAction SilentlyContinue) { Start-Sleep 5 }
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        
        if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
        logoff
        exit
    } else {
        $RetryCount++
        if ($RetryCount -gt $MaxRetries) {
            Write-Log ">>> 已達最大重試次數，放棄。" "Red"
            Backup-Logs
            Update-Status "Failed" $RetryCount
            Send-Notify -Title "任務失敗" -Msg "已達最大重試次數。" -Color "Red"
            New-Item -Path $FailFlag -ItemType File -Force | Out-Null
            if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
            logoff
            exit
        }
    }
}
