# =============================================================================
# AutoTask Master V5.7 - 跨日修正版 (03:55 啟動邏輯優化)
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

if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $FlagDir)) { New-Item -Path $FlagDir -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Master_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
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

Write-Log ">>> Master 啟動 (Admin Mode)..." "Cyan"

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
    if (-not (Test-Path $ManualFlag)) {
        
        # [核心修正] 日期判斷 (AddHours -4)
        $Now = Get-Date
        # 如果現在是 03:50 ~ 04:00 之間，視為「即將到來的今天」
        # 此時 -4h 會得到昨天，如果我們用它去查 PauseLog，可能會查到昨天的紀錄。
        # 但因為 04:00 才是換日點，理論上 03:55 啟動時，它的任務目標是「今天的 04:00」。
        # 所以我們應該稍微寬容一點：如果現在接近 04:00，我們查的是「明天(相對於-4h)」的設定嗎？
        # 不，最簡單的邏輯是：03:55 啟動時，PauseLog 應該還沒寫入今天的暫停 (除非手動預設)。
        # 如果使用者在昨天設了暫停 (停在昨天)，那麼今天 03:55 檢查 -4h (昨天) 會被暫停擋住。
        # 解決方案：如果現在 > 03:50，我們暫時假設它是為了「新的一天」而跑，
        # 所以我們檢查 PauseLog 時，應該檢查 `AddHours(-4)` 之後的日期 (如果昨天暫停，今天應該跑)。
        # 讓我們先保持 -4h 邏輯，但如果發現是「昨天」已完成 (`LastRunLog`)，而現在時間接近 04:00，則允許執行。

        $CheckDateStr = $Now.AddHours(-4).ToString("yyyyMMdd")
        
        if (Test-Path $PauseLog) {
            if ((Get-Content $PauseLog) -contains $CheckDateStr) { Write-Log "今日暫停 ($CheckDateStr)。" "Yellow"; exit }
        }
        
        # 時間窗檢查 (03:35 ~ 04:25)
        $Target = (Get-Date).Date.AddHours(3).AddMinutes(55)
        if ($Now -lt $Target.AddMinutes(-20) -or $Now -gt $Target.AddMinutes(30)) {
            if (-not (Test-Path $RunFlag)) { Write-Log "非任務時間，退出。" "Gray"; exit }
        }
        
        # [核心修正] 如果 LastRun 紀錄的是昨天，且現在是 03:55，代表是新任務，不要退出
        # 原本邏輯：檢查 LastRun == CurrentDateStr。
        # 如果昨天跑過，LastRun=昨天。現在03:55，CurrentDateStr=昨天。-> 相等 -> 退出！ (Bug)
        # 修正：如果現在是 03:50~03:59，我們跳過 LastRun 檢查，交給 Payload 去擋 (Payload 會等到 04:00 後再檢查)
        if ($Now.Hour -ne 3 -or $Now.Minute -lt 50) {
             if (Test-Path $LastRunLog) {
                if ((Get-Content $LastRunLog) -eq $CheckDateStr) { Write-Log "今日任務已完成。" "Green"; exit }
             }
        }

    } else {
        Write-Log "手動觸發，執行淨化..." "Magenta"
        try {
            $ResetStatus = @{ Date=(Get-Date).AddHours(-4).ToString("yyyyMMdd"); Status="Preparing"; RetryCount=0; LastUpdate=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
            $ResetStatus | ConvertTo-Json | Set-Content $TaskStatus -Encoding UTF8
        } catch {}

        Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
        $MyPID = $PID
        try {
            Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        } catch {}
        
        $SessionOutput = qwinsta 2>$null | Select-String "\bRemote\b"
        if ($SessionOutput) {
            $Line = $SessionOutput.ToString().Trim() -replace "\s+", " "; $Parts = $Line.Split(" "); $SessionID = $null; foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }
            if ($SessionID) {
                Write-Log "強制登出 Remote (ID: $SessionID)..." "Yellow"
                cmd /c "logoff $SessionID"
                $Wait=0; while($true){ if(-not(qwinsta 2>$null|Select-String "\bRemote\b")){break}; if($Wait-ge 20){break}; Start-Sleep 1; $Wait++ }
            }
        }
        if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }
        Remove-Item $ManualFlag -Force
        New-Item -Path $ForceRunFlag -ItemType File -Force | Out-Null
    }

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

# 3. 確保 Monitor 運行
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if (-not $MonitorProc) {
    Write-Log "啟動 Monitor..."
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""
} else {
    Write-Log "Monitor 已在運行。"
}

# --- [Master 監督迴圈] ---
$SupervisorStart = Get-Date
Write-Log ">>> 進入監督模式" "Green"
$PayloadLaunched = $false
if ($IsResume) { $PayloadLaunched = $true }

# [新] 連續重啟計數器
$RapidRestartCount = 0
$LastRestartTime = Get-Date

while ($true) {
    Start-Sleep 5

    if (Test-Path $DoneFlag) { Write-Log "任務成功 (Done)！" "Green"; break }
    if (Test-Path $FailFlag) { Write-Log "任務失敗 (Fail)！" "Red"; Remove-Item $RunFlag -Force; Stop-Process -Name "1Remote" -Force; exit }

    $MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
    if (-not $MonitorProc) {
        Write-Log "⚠️ Monitor 消失，重啟..." "Red"
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""
    }

    $PayloadProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Payload.ps1*" }
    if (-not $PayloadProc) {
        if ($PayloadLaunched) {
            $Now = Get-Date
            if (($Now - $LastRestartTime).TotalSeconds -lt 60) { $RapidRestartCount++ } else { $RapidRestartCount = 1 }
            $LastRestartTime = $Now

            if ($RapidRestartCount -gt 5) {
                Write-Log "⛔ Payload 連續閃退超過 5 次，停止救援！" "Red"
                New-Item -Path $FailFlag -ItemType File -Force | Out-Null
                Remove-Item $RunFlag -Force; Stop-Process -Name "1Remote" -Force
                exit
            }

            Write-Log "⚠️ Payload 消失 (嘗試 $RapidRestartCount)，重啟..." "Red"
            schtasks /run /tn "Auto_BetterGI_Payload"
            Start-Sleep 10
        } else {
            if ((Get-Date) -gt $SupervisorStart.AddMinutes(15)) {
                 Write-Log "Payload 啟動超時，重試 RDP..." "Yellow"
                 Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
                 $SupervisorStart = Get-Date
            }
        }
    } else {
        $PayloadLaunched = $true; $SupervisorStart = Get-Date 
    }
}

# --- [清理] ---
Write-Log "任務結束，清理中..."
Remove-Item $RunFlag -Force
Start-Sleep 5
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if ($MonitorProc) { Stop-Process -Id $MonitorProc.ProcessId -Force }

Write-Log "等待 Remote 登出..."
$Timeout=0
while ($true) {
    if (-not (qwinsta 2>$null | Select-String "\bRemote\b")) { Write-Log "Remote 已登出。" "Green"; break }
    if ($Timeout -ge 60) { Write-Log "登出逾時。" "Yellow"; break }
    Start-Sleep 3; $Timeout++
}

Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
Remove-Item $DoneFlag -Force

# [修正] 不關機檢查也改為 -4
$CurrentDateStr = (Get-Date).AddHours(-4).ToString("yyyyMMdd")
if (Test-Path $NoShutdownLog) {
    if ((Get-Content $NoShutdownLog) -contains $CurrentDateStr) { Write-Log "今日不關機。" "Cyan"; exit }
}

if (Test-Path $CountdownScript) { Start-Process powershell.exe -ExecutionPolicy Bypass -File "$CountdownScript" }
exit