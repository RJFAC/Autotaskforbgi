# =============================================================================
# AutoTask Master V5.17 - Notification & Status Fix
# =============================================================================
# V5.17: 
#   1. 強制所有啟動（含排程）皆發送 Discord 通知。
#   2. 強制所有啟動皆重置 TaskStatus 為 "Preparing"，修復 Dashboard 顯示滯後問題。
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

# 載入 Discord 模組
if (Test-Path "$ScriptDir\Lib_Discord.ps1") { . "$ScriptDir\Lib_Discord.ps1" }

if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $FlagDir)) { New-Item -Path $FlagDir -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    try { Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# --- [清理舊日誌] ---
try { Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}

# --- [變數對映] ---
$RunFlag        = "$FlagDir\Run.flag"
$DoneFlag       = "$FlagDir\Done.flag"
$FailFlag       = "$FlagDir\Fail.flag"
$ManualFlag     = "$FlagDir\ManualTrigger.flag"
$ForceRunFlag   = "$FlagDir\ForceRun.flag" 
$MonitorScript  = "$ScriptDir\Monitor.ps1"
$PauseLog       = "$ConfigDir\PauseDates.log"
$NoShutdownLog  = "$ConfigDir\NoShutdown.log"
$CountdownScript= "$ScriptDir\Shutdown-Countdown.ps1"
$TaskStatus     = "$ConfigDir\TaskStatus.json"
$LastRunLog     = "$ConfigDir\LastRun.log"

# 讀取 1Remote 路徑
$1RemoteDir = "C:\AutoTask\1Remote"
$1RemoteExe = "$1RemoteDir\1Remote.exe"
if (Test-Path "$ConfigDir\EnvConfig.json") {
    try {
        $env = Get-Content "$ConfigDir\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env.Path1Remote) { 
            $1RemoteExe = $env.Path1Remote
            $1RemoteDir = Split-Path $1RemoteExe -Parent 
        }
    } catch {}
}

function Check-Network {
    Write-Log "檢查網路..." 
    $Retry = 0; $MaxRetry = 12
    while ($Retry -lt $MaxRetry) {
        try { if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -ErrorAction Stop) { Write-Log "網路正常。" "Green"; return $true } } catch {}
        Write-Log "網路未就緒... ($($Retry+1))" "Yellow"; Start-Sleep 5; $Retry++
    }
    Write-Log "⚠️ 網路連線逾時。" "Red"; return $false
}

Write-Log ">>> Master 啟動 (Admin Mode - V5.17 + Notify Fix)..." "Cyan"

# =============================================================================
# [核心邏輯] 判斷是「全新啟動」還是「接手續跑」
# =============================================================================
$IsResume = $false
if (Test-Path $RunFlag) {
    $P1 = Get-Process "1Remote" -ErrorAction SilentlyContinue
    if ($P1) {
        Write-Log "偵測到 Run.flag 與 1Remote，判定為 [熱重載/斷點續接]。" "Magenta"
        $IsResume = $true
    }
}

if (-not $IsResume) {
    # --- [全新啟動流程] ---
    
    # [Fix] 無論是手動還是排程，只要是新啟動，都重置狀態並通知
    # 這樣 Dashboard 就能立刻顯示 "Preparing"，使用者也會收到 Discord 通知
    try {
        $ResetStatus = @{ 
            Date = (Get-Date).AddHours(-4).ToString("yyyyMMdd");
            Status = "Preparing"; 
            RetryCount = 0; 
            LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
        }
        $ResetStatus | ConvertTo-Json | Set-Content $TaskStatus -Encoding UTF8 -Force
        
        $StartType = if (Test-Path $ManualFlag) { "手動觸發 (Manual)" } else { "排程啟動 (Scheduled)" }
        
        # 發送啟動通知
        if (Get-Command Send-DiscordNotification -ErrorAction SilentlyContinue) {
            Send-DiscordNotification -Title "⚙️ Master 系統啟動" -Message "類型: $StartType`n狀態: 正在初始化環境並準備執行任務..." -Color "Blue"
        }
    } catch {
        Write-Log "狀態重置或通知發送失敗: $_" "Red"
    }

    if (-not (Test-Path $ManualFlag)) {
        $Now = Get-Date
        $CheckDateStr = $Now.AddHours(-4).ToString("yyyyMMdd")
        
        if (Test-Path $PauseLog) {
            if ((Get-Content $PauseLog) -contains $CheckDateStr) { 
                Write-Log "今日暫停 ($CheckDateStr)。" "Yellow"
                Send-DiscordNotification -Title "⏸️ 今日暫停" -Message "檢測到暫停設定，Master 將停止執行。" -Color "Orange"
                exit 
            }
        }
        
        # 時間窗檢查 (03:35 ~ 04:25)
        $Target = (Get-Date).Date.AddHours(3).AddMinutes(55)
        if ($Now -lt $Target.AddMinutes(-20) -or $Now -gt $Target.AddMinutes(30)) {
            if (-not (Test-Path $RunFlag)) { Write-Log "非任務時間，退出。" "Gray"; exit }
        }
        
        # LastRun 檢查 (含緩衝區間)
        $InResetBuffer = ($Now.Hour -eq 3 -and $Now.Minute -ge 45) -or ($Now.Hour -eq 4 -and $Now.Minute -lt 5)
        if (-not $InResetBuffer) {
             if (Test-Path $LastRunLog) {
                if ((Get-Content $LastRunLog) -eq $CheckDateStr) { 
                    Write-Log "今日任務已完成。" "Green"; 
                    # 避免重複通知，這裡可選擇不發或發送簡單提示
                    exit 
                }
             }
        } else {
             Write-Log "處於換日緩衝期 (03:45~04:05)，跳過 LastRun 檢查。" "Yellow"
        }

    } else {
        Write-Log "手動觸發，執行淨化..." "Magenta"
        # 手動觸發的清理邏輯
        Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
        $MyPID = $PID
        try {
            Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
            Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        } catch {}
        
        # 強制登出 Remote
        $SessionOutput = qwinsta 2>$null | Select-String "\bRemote\b"
        if ($SessionOutput) {
            $Line = $SessionOutput.ToString().Trim() -replace "\s+", " "; $Parts = $Line.Split(" "); $SessionID = $null; foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }
            if ($SessionID) {
                cmd /c "logoff $SessionID"
                Start-Sleep 2
            }
        }
        if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }
        Remove-Item $ManualFlag -Force
        New-Item -Path $ForceRunFlag -ItemType File -Force | Out-Null
        Write-Log "已建立 ForceRun 標記。" "Cyan"
    }

    # 重置 Flags
    if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }
    if (Test-Path $DoneFlag) { Remove-Item $DoneFlag -Force }
    if (Test-Path $FailFlag) { Remove-Item $FailFlag -Force }
    New-Item -Path $RunFlag -ItemType File -Force | Out-Null

    Check-Network

    Write-Log "啟動 1Remote..."
    Start-Process -FilePath $1RemoteExe -WorkingDirectory $1RemoteDir
    Start-Sleep 5
    Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
    Start-Sleep 2
}

# 確保 Monitor 運行
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if (-not $MonitorProc) {
    Write-Log "啟動 Monitor..."
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""
}

# --- [Master 監督迴圈] ---
$SupervisorStart = Get-Date
Write-Log ">>> 進入監督模式" "Green"
$PayloadLaunched = $false
if ($IsResume) { $PayloadLaunched = $true }

$RapidRestartCount = 0
$LastRestartTime = Get-Date

while ($true) {
    Start-Sleep 5

    if (Test-Path $DoneFlag) { 
        Write-Log "任務成功 (Done)！" "Green"
        if (Get-Command Send-AutoTaskReport -ErrorAction SilentlyContinue) {
             Send-AutoTaskReport -Status "Success" -LogFile $LogFile
        }
        break 
    }
    if (Test-Path $FailFlag) { 
        Write-Log "任務失敗 (Fail)！" "Red"
        if (Get-Command Send-AutoTaskReport -ErrorAction SilentlyContinue) {
            Send-AutoTaskReport -Status "Error" -LogFile $LogFile
        }
        Remove-Item $RunFlag -Force; Stop-Process -Name "1Remote" -Force
        exit 
    }

    $MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
    if (-not $MonitorProc) {
        Write-Log "⚠️ Monitor 消失，重啟..." "Red"
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""
    }

    $PayloadProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Payload.ps1*" }
    
    if (-not $PayloadProc) {
        if ($PayloadLaunched) {
            $Now = Get-Date
            $Mem = Get-CimInstance Win32_OperatingSystem | Select-Object @{Name="FreeGB";Expression={$_.FreePhysicalMemory/1MB}}
            Write-Log "⚠️ Payload 消失！(系統剩餘記憶體: $([math]::Round($Mem.FreeGB, 2)) GB)" "Red"
            
            # 發送異常通知
            if (Get-Command Send-DiscordNotification -ErrorAction SilentlyContinue) {
                Send-DiscordNotification -Title "⚠️ Payload 進程異常消失" -Message "正在嘗試救援重啟... (重試次數: $($RapidRestartCount + 1))" -Color "Orange"
            }

            if (($Now - $LastRestartTime).TotalSeconds -lt 60) { $RapidRestartCount++ } else { $RapidRestartCount = 1 }
            $LastRestartTime = $Now

            if ($RapidRestartCount -gt 5) {
                Write-Log "⛔ Payload 連續閃退超過 5 次，停止救援！" "Red"
                New-Item -Path $FailFlag -ItemType File -Force | Out-Null
                Remove-Item $RunFlag -Force; Stop-Process -Name "1Remote" -Force
                exit
            }

            Write-Log "⚠️ 正在執行救援重啟 (嘗試 $RapidRestartCount)..." "Yellow"
            $LogOutput = schtasks /run /tn "Auto_BetterGI_Payload" 2>&1
            Start-Sleep 10
        } else {
            if ((Get-Date) -gt $SupervisorStart.AddMinutes(15)) {
                 Write-Log "Payload 啟動超時 (15分鐘)，RDP 重試..." "Yellow"
                 Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
                 $SupervisorStart = Get-Date
            }
        }
    } else {
        if (-not $PayloadLaunched) {
             Write-Log "偵測到 Payload 運作中 (PID: $($PayloadProc.ProcessId))" "Cyan"
        }
        $PayloadLaunched = $true; $SupervisorStart = Get-Date 
    }
}

# --- [清理與結束] ---
Write-Log "任務結束，清理中..."
Remove-Item $RunFlag -Force
Start-Sleep 5
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if ($MonitorProc) { Stop-Process -Id $MonitorProc.ProcessId -Force }

Write-Log "等待 Remote 登出..."
$Timeout = 0; $MaxTimeout = 60
while ($true) {
    $SessionInfo = qwinsta 2>$null | Select-String "\bRemote\b"
    if (-not $SessionInfo) { Write-Log "Remote 已登出。" "Green"; break }
    if ($Timeout -ge $MaxTimeout) { 
        try {
            $Line = $SessionInfo.ToString().Trim() -replace "\s+", " "; $Parts = $Line.Split(" "); $SessionID = $null; foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }
            if ($SessionID) { cmd /c "logoff $SessionID" }
        } catch {}
        break 
    }
    Start-Sleep 3; $Timeout++
}
Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
Remove-Item $DoneFlag -Force

$CurrentDateStr = (Get-Date).AddHours(-4).ToString("yyyyMMdd")
if (Test-Path $NoShutdownLog) {
    if ((Get-Content $NoShutdownLog) -contains $CurrentDateStr) { Write-Log "今日不關機。" "Cyan"; exit }
}
if (Test-Path $CountdownScript) { Start-Process powershell.exe -ExecutionPolicy Bypass -File "$CountdownScript" }
exit