# =============================================================================
# AutoTask Master V5.34 - Fix WMI Crash on Sleep/Wake
# =============================================================================
# V5.34:
#   1. [Fix] 修正 Get-CimInstance 在系統睡眠/喚醒時報錯 (0x80041033) 導致 Master
#      誤判 Monitor 已死並錯誤重啟的問題。加入 Try-Catch 容錯機制。
# V5.33:
#   1. [Fix] 修復手動觸發邏輯順序錯誤。
#   2. [Fix] 修正 Write-Log 顏色錯誤。
# =============================================================================

# --- [0. 權限自我檢查] ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Set-Location $PSScriptRoot

# --- [路徑定義] ---
$BaseDir    = "C:\AutoTask"
$ScriptDir  = "$BaseDir\Scripts"
$ConfigDir  = "$BaseDir\Configs"
$FlagDir    = "$BaseDir\Flags"
$LogDir     = "$BaseDir\Logs"
$LogFileName = "Master_$(Get-Date -Format 'yyyyMMdd').log"
$LogFile     = Join-Path $LogDir $LogFileName

# 關鍵旗標路徑
$RunFlag           = "$FlagDir\Run.flag"
$ManualTriggerFlag = "$FlagDir\ManualTrigger.flag"
$ForceRunFlag      = "$FlagDir\ForceRun.flag"
$DoneFlag          = "$FlagDir\Done.flag"
$FailFlag          = "$FlagDir\Fail.flag"

# 載入 Discord 模組
if (Test-Path "$ScriptDir\Lib_Discord.ps1") { . "$ScriptDir\Lib_Discord.ps1" }

if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogContent = "[$TimeStamp] $Message"
    Write-Host $LogContent -ForegroundColor $Color
    $LogContent | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

# --- [1. 啟動初始化] ---
Write-Log ">>> Master 啟動 (Admin Mode - V5.34)..." "Cyan"

# --- [絕對禁區檢查 (03:55 ~ 04:05)] ---
$Now = Get-Date
$DeadZoneStart = $Now.Date.AddHours(3).AddMinutes(55) # 03:55
$DeadZoneEnd   = $Now.Date.AddHours(4).AddMinutes(5)  # 04:05

if ($Now -ge $DeadZoneStart -and $Now -lt $DeadZoneEnd) {
    $WaitSeconds = [math]::Ceiling(($DeadZoneEnd - $Now).TotalSeconds)
    Write-Log "⛔ 當前時間 ($($Now.ToString("HH:mm:ss"))) 位於 [絕對禁區] (03:55~04:05)。" "Red"
    Write-Log "系統將強制鎖定並等待 $WaitSeconds 秒..." "Yellow"
    Start-Sleep -Seconds $WaitSeconds
    $Now = Get-Date
    Write-Log "等待結束，解除鎖定。" "Green"
}

# --- [執行靜音與勿擾設定] ---
# 條件: 1. 非手動觸發 AND 2. 位於排程時間窗 (03:35~04:30)
$SilenceStart = $Now.Date.AddHours(3).AddMinutes(35)
$SilenceEnd   = $Now.Date.AddHours(4).AddMinutes(30)

if (-not (Test-Path $ManualTriggerFlag)) {
    if ($Now -ge $SilenceStart -and $Now -le $SilenceEnd) {
        if (Test-Path "$ScriptDir\Set-Silence.ps1") {
            Write-Log "偵測到排程時段啟動，執行系統靜音..." "Cyan"
            & "$ScriptDir\Set-Silence.ps1"
        }
    } else {
        Write-Log "非排程時段啟動 ($($Now.ToString("HH:mm")))，跳過靜音設定。" "Gray"
    }
}

# 重置 TaskStatus
$TaskStatusFile = "$ConfigDir\TaskStatus.json"
try {
    $InitialStatus = @{ "Date" = Get-Date -Format "yyyy/MM/dd"; "Status" = "Preparing"; "Message" = "Master Initializing..."; "LastUpdate" = Get-Date -Format "HH:mm:ss" }
    $InitialStatus | ConvertTo-Json -Depth 2 | Set-Content -Path $TaskStatusFile -Encoding UTF8
} catch {}

# 清理舊日誌
Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force

# 1.2 判斷啟動模式
$IsResume = $false

if ((Test-Path $RunFlag) -and (Get-Process "1Remote" -ErrorAction SilentlyContinue)) {
    Write-Log "偵測到 Run.flag 與 1Remote 進程，判定為 [熱重載] 模式。" "Yellow"
    $IsResume = $true
} else {
    # --- 全新啟動流程 ---
    
    # [V5.33 Fix] 先捕捉手動觸發狀態，再進行全域 Flag 清理
    $IsManualTrigger = Test-Path $ManualTriggerFlag
    
    # 重置 Flags (清理舊狀態)
    Get-ChildItem $FlagDir -Filter "*.flag" | Remove-Item -Force
    
    # [檢查手動觸發]
    if ($IsManualTrigger) {
        Write-Log "偵測到手動觸發 (ManualTrigger)，執行強制清理與啟動..." "Magenta"
        Send-DiscordNotification -Title "🚀 手動啟動" -Message "使用者強制啟動任務。" -Color "Blue"
        
        # 建立強制執行標記 (在清理後建立)
        New-Item -ItemType File -Path $ForceRunFlag -Force | Out-Null
        
        Get-Process | Where-Object { $_.Name -match "1Remote|Monitor" } | Stop-Process -Force -ErrorAction SilentlyContinue
        $SessionInfo = qwinsta 2>$null | Select-String "\bRemote\b"
        if ($SessionInfo) { cmd /c "logoff $((($SessionInfo.ToString().Trim() -replace "\s+", " ").Split(" "))[2])" 2>$null; Start-Sleep 2 }
    } else {
        # --- [排程啟動邏輯] ---
        
        # 1. 檢查暫停
        $PauseLog = "$ConfigDir\PauseDates.log"
        $TodayStr = Get-Date -Format "yyyy/MM/dd"
        if (Test-Path $PauseLog) {
            if ((Get-Content $PauseLog) -contains $TodayStr) {
                Write-Log "今日暫停，任務取消。" "Yellow"
                Send-DiscordNotification -Title "⏸️ 今日暫停" -Message "任務已取消。" -Color "Orange"
                exit
            }
        }

        # 2. 檢查時間窗 (03:35 ~ 04:30)
        if ($Now -lt $SilenceStart -or $Now -gt $SilenceEnd) {
            Write-Log "當前時間 ($($Now.ToString("HH:mm"))) 不在任務執行窗口 (03:35-04:30)，退出。" "Red"
            exit
        }

        # 3. 檢查 LastRun (03:45~04:10 啟動緩衝豁免)
        $BufferStart = $Now.Date.AddHours(3).AddMinutes(45)
        $BufferEnd   = $Now.Date.AddHours(4).AddMinutes(10)
        
        if ($Now -ge $BufferStart -and $Now -le $BufferEnd) {
            Write-Log "處於啟動緩衝區 (03:45~04:10)，跳過 LastRun 重複檢查。" "Cyan"
        } else {
            $LastRunLog = "$ConfigDir\LastRun.log"
            if (Test-Path $LastRunLog) {
                if ((Get-Content $LastRunLog -Raw).Trim() -eq $TodayStr) {
                    Write-Log "今日任務已完成，退出。" "Yellow"
                    exit
                }
            }
        }
        
        Send-DiscordNotification -Title "⏰ 排程啟動" -Message "Master 開始執行任務。" -Color "Blue"
    }

    # 建立 Run.flag (正式 Run)
    New-Item -ItemType File -Path $RunFlag -Force | Out-Null
    
    # 網路檢查
    $NetRetry = 0
    while ($true) {
        if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { break }
        $NetRetry++; if ($NetRetry -ge 12) { Send-DiscordNotification -Title "❌ 網路錯誤" -Message "無網路連線。" -Color "Red"; exit }
        Start-Sleep 5
    }
    
    # 啟動環境
    $1RemotePath = "C:\AutoTask\1Remote\1Remote.exe"
    if (Test-Path "$ConfigDir\EnvConfig.json") { try { $env = Get-Content "$ConfigDir\EnvConfig.json" -Raw | ConvertFrom-Json; if ($env.Path1Remote) { $1RemotePath = $env.Path1Remote } } catch {} }
    
    # 分段啟動 1Remote
    Write-Log "啟動 1Remote 主程式..."
    Start-Process -FilePath $1RemotePath -WindowStyle Minimized
    Write-Log "等待 5 秒讓 1Remote 就緒..."
    Start-Sleep 5
    Write-Log "發送連線指令 (目標: Remote)..."
    Start-Process -FilePath $1RemotePath -ArgumentList "-r Remote" -WindowStyle Minimized
    
    Write-Log "啟動 Monitor..."
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptDir\Monitor.ps1`"" -WindowStyle Minimized
}

# --- [2. 監督迴圈] ---
Write-Log ">>> 進入監督模式" "Green"
$PayloadLaunched = $IsResume
$SupervisorStart = Get-Date

while ($true) {
    Start-Sleep 5
    
    if (Test-Path $DoneFlag) {
        Write-Log "任務成功 (Done)！" "Green"
        Send-AutoTaskReport "Success" $LogFile
        break
    }

    if (Test-Path $FailFlag) {
        Write-Log "任務失敗 (Fail)。" "Red"
        Send-AutoTaskReport "Error" $LogFile
        Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
        Remove-Item $RunFlag -Force
        exit
    }

    # Monitor 存活檢查 [V5.34: 加入 WMI 容錯]
    $MonitorRunning = $true # 預設為真，防止 WMI 錯誤導致誤殺
    try {
        $ProcList = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop
        $MonitorProc = $ProcList | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
        if (-not $MonitorProc) { $MonitorRunning = $false }
    } catch {
        # 若 WMI 失敗 (如休眠中/關機中)，假定 Monitor 還活著，避免重啟
        # Write-Log "WMI 查詢異常 (可能正在休眠)，跳過 Monitor 檢查。" "Gray" 
    }
    
    if (-not $MonitorRunning) {
        Write-Log "警告: Monitor 已消失，正在重啟..." "Yellow"
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptDir\Monitor.ps1`"" -WindowStyle Minimized
    }

    if (-not $PayloadLaunched -and ((Get-Date) - $SupervisorStart).TotalMinutes -lt 15) {
        # [V5.34: 加入 WMI 容錯]
        try {
            $ProcList = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop
            $PayloadProc = $ProcList | Where-Object { $_.CommandLine -like "*Payload.ps1*" }
            
            if ($PayloadProc) { 
                $PayloadLaunched = $true 
            } else {
                 if (((Get-Date) - $SupervisorStart).TotalMinutes -gt 2) {
                     Write-Log "Payload 逾時未啟動，嘗試重送連線指令..." "DarkYellow"
                     Start-Process -FilePath $1RemotePath -ArgumentList "-r Remote" -WindowStyle Minimized
                     Start-Sleep 10
                 }
            }
        } catch {
            # WMI 失敗時忽略 Payload 檢查
        }
    } elseif (-not $PayloadLaunched) {
        New-Item -ItemType File -Path $FailFlag -Force | Out-Null
    }
}

# --- [清理與結束] ---
Remove-Item $RunFlag -Force
Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
Get-Process | Where-Object { $_.CommandLine -like "*Monitor.ps1*" } | Stop-Process -Force
$SessionInfo = qwinsta 2>$null | Select-String "\bRemote\b"
if ($SessionInfo) { cmd /c "logoff $((($SessionInfo.ToString().Trim() -replace "\s+", " ").Split(" "))[2])" 2>$null }

# --- [關機檢查 (V5.32 優化)] ---
$Shut = $true
$TodayStr = Get-Date -Format "yyyy/MM/dd"

# 1. 優先檢查 NoShutdown.log (手動指定覆蓋)
if (Test-Path "$ConfigDir\NoShutdown.log") {
    $List = Get-Content "$ConfigDir\NoShutdown.log"
    if ($List -contains $TodayStr) {
        Write-Log "NoShutdown.log 指定今日 [不關機]。" "Green"
        $Shut = $false
    }
}

# 2. 若 Log 未指定不關機，則檢查 WeeklyConfig (每週預設)
if ($Shut -and (Test-Path "$ConfigDir\WeeklyConfig.json")) {
    try {
        $Weekly = Get-Content "$ConfigDir\WeeklyConfig.json" -Raw | ConvertFrom-Json
        $DayOfWeek = (Get-Date).DayOfWeek.ToString()
        
        # 紊亂期計算
        $RefDate = Get-Date "2024-08-28"
        $DiffDays = ((Get-Date) - $RefDate).TotalDays
        $CycleDay = $DiffDays % 42
        if ($CycleDay -lt 0) { $CycleDay += 42 }
        
        $IsNoShut = $false
        
        # 判斷是否為紊亂期
        if ($CycleDay -ge 7.4 -and $CycleDay -le 17.2) {
            # 優先讀取 Turbulence.NoShutdown
            if ($Weekly.Turbulence.NoShutdown.$DayOfWeek -eq $true) {
                Write-Log "紊亂期每週設定 ($DayOfWeek) 為 [不關機]。" "Cyan"
                $IsNoShut = $true
            }
        } else {
            # 讀取一般 NoShutdown
            if ($Weekly.NoShutdown.$DayOfWeek -eq $true) {
                Write-Log "一般每週設定 ($DayOfWeek) 為 [不關機]。" "Cyan"
                $IsNoShut = $true
            }
        }
        
        if ($IsNoShut) { $Shut = $false }
        
    } catch {
        Write-Log "讀取 WeeklyConfig 失敗，維持預設關機。" "Red"
    }
}

if ($Shut) {
    if (Test-Path "$ScriptDir\Shutdown-Countdown.ps1") {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptDir\Shutdown-Countdown.ps1`""
    } else { Stop-Computer -Force }
} else {
    Write-Log "今日設定為不關機，任務結束。" "Green"
}