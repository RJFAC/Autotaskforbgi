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

        $DestAuto = Join-Path $TargetDir "AutoTask_Logs"
        New-Item -Path $DestAuto -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$LogDir\*" -Destination $DestAuto -Recurse -Force -ErrorAction SilentlyContinue

        $DestBG = Join-Path $TargetDir "BetterGI_Logs"
        New-Item -Path $DestBG -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$LogDirBG\*" -Destination $DestBG -Recurse -Force -ErrorAction SilentlyContinue

        if (Test-Path $1RemoteLogDir) {
            $Dest1R = Join-Path $TargetDir "1Remote_Logs"
            New-Item -Path $Dest1R -ItemType Directory -Force | Out-Null
            Copy-Item -Path "$1RemoteLogDir\*" -Destination $Dest1R -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Log "日誌備份完成。路徑: $TargetDir" "Green"
        return $TargetDir
    } catch { 
        Write-Log "日誌備份失敗: $_" "Red"
        return $null
    }
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

function Get-TargetConfig {
    $today = (Get-Date).AddHours(-3)
    $dStr = $today.ToString("yyyyMMdd")
    if (Test-Path $DateMap) {
        $map = Get-Content $DateMap
        foreach ($line in $map) { if ($line -match "^$dStr=(.+)$") { return $matches[1] } }
    }
    if (Test-Path $WeeklyConf) {
        try {
            $wk = Get-Content $WeeklyConf -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($wk) { return $wk.$($today.DayOfWeek.ToString()) }
        } catch {}
    }
    return "day"
}

function Check-Success-Log ($logPath) {
    try {
        $Logs = Get-Content $logPath -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($Logs -match $SuccessKeyword) { return $true }
    } catch {}
    return $false
}

# =============================================================================
# [前日殘留處理]
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
$CurrentDateStr = (Get-Date).AddHours(-3).ToString("yyyyMMdd")
Write-Log "今日日期 (計算結果): $CurrentDateStr"

Cleanup-Screenshots

$IsForceRun = $false
if (Test-Path $ForceRunFlag) {
    Write-Log "偵測到 ForceRun 旗標，將無視暫停與重複執行檢查。" "Magenta"
    $IsForceRun = $true
    Remove-Item $ForceRunFlag -Force
}

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
$StartTime = Get-Date

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
        
        $LogMsg = ">>> 執行進度 [$CurrentQueueIndex/$($ConfigQueue.Count)]: 配置 [$CurrentConfig]"
        if ($ConfigFrequency[$CurrentConfig] -gt 1) { $LogMsg += " (本日第 $($CurrentConfigRunCount[$CurrentConfig])/$($ConfigFrequency[$CurrentConfig]) 次)" }
        $LogMsg += " (重試: $RetryCount/$MaxRetries)"
        Write-Log $LogMsg "Cyan"

        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
        Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3

        $Args = "--startOneDragon ""$CurrentConfig"""
        Start-Process -FilePath $BettergiExe -ArgumentList $Args -WorkingDirectory $BettergiDir
        
        $WatchdogStart = Get-Date
        $IsSuccess = $false
        $IsFailed = $false
        
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
                
                if (Check-Success-Log $LogFile.FullName) {
                    Write-Log "配置 [$CurrentConfig] 執行完成！" "Green"
                    $IsSuccess = $true
                    break
                }
                
                # [修正] 二次確認機制 (Double Check)
                if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) { 
                    Write-Log "⚠️ 偵測到 BetterGI 進程消失，等待 3 秒後進行二次確認..." "Yellow"
                    Start-Sleep 3
                    
                    if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                        Write-Log "確認 BetterGI 已完全停止，檢查是否為正常結束..."
                        if (Check-Success-Log $LogFile.FullName) {
                            Write-Log "配置 [$CurrentConfig] 執行完成！(進程結束後確認)" "Green"
                            $IsSuccess = $true
                        } else {
                            Write-Log "配置 [$CurrentConfig] 意外中斷且無成功訊號！" "Red"
                            $IsFailed = $true
                        }
                        break
                    } else {
                        Write-Log "BetterGI 仍在運行 (二次確認成功)，繼續監控。" "Green"
                        continue
                    }
                }
                
                $LogFile.Refresh()
                if (((Get-Date) - $LogFile.LastWriteTime).TotalMinutes -gt 15) { 
                    if (Check-Success-Log $LogFile.FullName) { $IsSuccess = $true; break }
                    Write-Log "⚠️ 警報：配置 [$CurrentConfig] 日誌停滯，判定卡死！" "Red"
                    $IsFailed = $true 
                }
            }
        }

        if (-not $IsSuccess) {
            $AllConfigSuccess = $false
            Write-Log "配置 [$CurrentConfig] 失敗，將觸發整體重試。" "Yellow"
            break 
        } else {
            Start-Sleep 5
        }
    }

    if ($AllConfigSuccess) {
        $Duration = New-TimeSpan -Start $StartTime -End (Get-Date)
        $DurationStr = "{0:hh}時{0:mm}分" -f $Duration
        Write-Log ">>> 所有排程配置皆已完成。總耗時: $DurationStr" "Green"
        
        Send-Notify -Title "任務成功" -Msg "配置 [$ConfigName] 已完成。`n耗時: $DurationStr" -Color "Green"

        Update-Status "Success" $RetryCount
        Set-Content $LastRunLog -Value $CurrentDateStr
        while (Get-Process "GenshinImpact" -ErrorAction SilentlyContinue) { Start-Sleep 5 }
        New-Item -Path $DoneFlag -ItemType File -Force | Out-Null
        
        if ($IsForceRun) {
            Write-Log "手動執行完成。10 秒後登出..."
            Start-Sleep 10
        } else {
            Start-Sleep 3
        }
        logoff
        exit
    } else {
        Write-Log ">>> 任務鏈中斷，準備重試..." "Yellow"
        $RetryCount++
        if ($RetryCount -gt $MaxRetries) {
            Write-Log ">>> 已達最大重試次數，放棄。" "Red"
            
            $BackupPath = Backup-Logs
            Send-Notify -Title "任務失敗" -Msg "配置 [$ConfigName] 嚴重失敗 (重試3次)。`n日誌已備份至: $BackupPath" -Color "Red"

            Update-Status "Failed" $RetryCount
            New-Item -Path $FailFlag -ItemType File -Force | Out-Null
            if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
            logoff
            exit
        } else {
            Start-Sleep 10
        }
    }
}
