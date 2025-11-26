# =============================================================================
# AutoTask Master V4.7 - 強制提權與路徑修正版
# =============================================================================

# --- [0. 權限自我檢查 (Self-Elevation)] ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # 重新以管理員身分啟動自己
    $process = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -PassThru
    exit # 舊的低權限進程退出
}

# 設定工作目錄為腳本所在目錄，避免相對路徑錯誤
Set-Location $PSScriptRoot

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
    $LogFileName = "Master_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8
}

# --- [清理舊日誌] ---
Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue

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

# 嘗試讀取 EnvConfig
$1RemoteDir = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64"
if (Test-Path "$ConfigDir\EnvConfig.json") {
    try {
        $env = Get-Content "$ConfigDir\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env.Path1Remote) { $1RemoteDir = Split-Path $env.Path1Remote -Parent }
    } catch {}
}
$1RemoteExe = "$1RemoteDir\1Remote.exe"

# --- [輔助函數：網路檢查] ---
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
    Write-Log "⚠️ 警告：網路連線逾時 (60秒)。" "Red"
    return $false
}

Write-Log ">>> Master 啟動程序開始 (Admin Mode)..." "Cyan"

# --- [主流程] ---

# 1. 檢查觸發條件
if (-not (Test-Path $ManualFlag)) {
    # [自動模式]
    if (Test-Path $PauseLog) {
        $CheckDateStr = (Get-Date).AddHours(-3).ToString("yyyyMMdd")
        if ((Get-Content $PauseLog) -contains $CheckDateStr) { 
            Write-Log "偵測到今日暫停 ($CheckDateStr)，Master 退出。" "Yellow"
            exit 
        }
    }
    $Now = Get-Date
    $Target = (Get-Date).Date.AddHours(3).AddMinutes(55)
    if ($Now -lt $Target.AddMinutes(-20) -or $Now -gt $Target.AddMinutes(30)) {
        if (-not (Test-Path $RunFlag)) { 
            Write-Log "非任務時間且無恢復旗標，Master 退出。" "Gray"
            exit 
        }
    }
} else {
    # [手動模式] - 執行深度淨化
    Write-Log "偵測到手動觸發旗標，執行環境深度淨化..." "Magenta"
    
    # A. 停止 1Remote
    Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
    
    # B. 停止舊的 Monitor 和 Master
    $MyPID = $PID
    try {
        Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
            Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID } |
            ForEach-Object { 
                Write-Log "終止殘留腳本 PID: $($_.ProcessId)"
                Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue 
            }
    } catch {}
    
    # C. [核心] 強制登出 Remote
    Write-Log "正在檢查 Remote 帳戶狀態..."
    $SessionOutput = qwinsta 2>$null | Select-String "\bRemote\b"
    if ($SessionOutput) {
        $Line = $SessionOutput.ToString().Trim() -replace "\s+", " "
        $Parts = $Line.Split(" ")
        $SessionID = $null
        foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }

        if ($SessionID) {
            Write-Log "發現殘留的 Remote Session (ID: $SessionID)，強制登出..." "Yellow"
            # 使用 cmd /c 確保 logoff 指令執行
            cmd /c "logoff $SessionID"
            
            $LogoutWait = 0
            while ($true) {
                if (-not (qwinsta 2>$null | Select-String "\bRemote\b")) {
                    Write-Log "✅ Remote 已完全登出。" "Green"
                    break
                }
                if ($LogoutWait -ge 20) {
                    Write-Log "⚠️ 等待登出逾時，強制繼續。" "Red"
                    break
                }
                Start-Sleep 1
                $LogoutWait++
            }
        }
    } else {
        Write-Log "Remote 目前未登入。" "Gray"
    }
    
    if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }
    Remove-Item $ManualFlag -Force
    New-Item -Path $ForceRunFlag -ItemType File -Force | Out-Null
    Write-Log "已建立 ForceRun 旗標，準備啟動。"
}

# 2. 初始化環境
if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }
if (Test-Path $DoneFlag) { Remove-Item $DoneFlag -Force }
if (Test-Path $FailFlag) { Remove-Item $FailFlag -Force }
New-Item -Path $RunFlag -ItemType File -Force | Out-Null
Write-Log "初始化環境完成，Run.flag 已建立。"

# 3. 網路預檢與啟動
Check-Network

Write-Log "正在啟動 1Remote..."
Start-Process -FilePath $1RemoteExe -WorkingDirectory $1RemoteDir
Start-Sleep 5
Write-Log "正在建立連線 (-r Remote)..."
Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
Start-Sleep 2

Write-Log "正在啟動監控腳本 (Monitor.ps1)..."
Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""

# 4. --- [Master 監督迴圈] ---
$SupervisorStart = Get-Date
Write-Log ">>> Master 監督模式已啟動" "Green"

while ($true) {
    Start-Sleep 5

    if (Test-Path $DoneFlag) {
        Write-Log "偵測到任務成功 (Done.flag)！" "Green"
        break
    }

    if (Test-Path $FailFlag) {
        Write-Log "偵測到任務嚴重失敗 (Fail.flag)！" "Red"
        $host.UI.RawUI.BackgroundColor = "DarkRed"
        Clear-Host
        Write-Host "`n========================================"
        Write-Host "      今日一條龍任務執行失敗！"
        Write-Host "========================================"
        Remove-Item $RunFlag -Force
        Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
        Read-Host "按 Enter 結束 Master (不關機)..."
        exit
    }

    $StatusObj = $null
    if (Test-Path $TaskStatus) {
        try {
            $StatusObj = Get-Content $TaskStatus -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
        } catch { $StatusObj = $null }
    }

    if (-not $StatusObj) {
        if ((Get-Date) -gt $SupervisorStart.AddMinutes(15)) {
             Write-Log "警告：Payload 似乎未啟動 (超時 15 分)，嘗試重連 RDP..." "Yellow"
             Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
             $SupervisorStart = Get-Date
        }
    }
}

# 5. --- [清理與登出確認] ---
Write-Log "任務完成，開始執行清理程序..."
Remove-Item $RunFlag -Force
Write-Log "已停止 Monitor。"

Write-Log "正在確認 Remote 帳戶是否已完全登出..."
$Timeout = 0
while ($true) {
    $Sessions = qwinsta 2>$null | Out-String
    if ($Sessions -match "\bRemote\b") {
        if ($Timeout -ge 60) {
            Write-Log "等待登出逾時 (3分鐘)，強制執行後續清理。" "Yellow"
            break
        }
        Write-Host "." -NoNewline
        Start-Sleep 3
        $Timeout++
    } else {
        Write-Host ""
        Write-Log "確認 Remote 已完全登出。" "Green"
        break
    }
}

Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
Write-Log "已關閉 1Remote 程序。"

Remove-Item $DoneFlag -Force
Write-Log "清理完畢。"

# 檢查不關機
$CurrentDateStr = (Get-Date).AddHours(-3).ToString("yyyyMMdd")
if (Test-Path $NoShutdownLog) {
    if ((Get-Content $NoShutdownLog) -contains $CurrentDateStr) {
        Write-Log "今日排定 [不關機]，Master 結束。" "Cyan"
        exit
    }
}

Write-Log "準備啟動倒數關機程序..."
if (Test-Path $CountdownScript) {
    Start-Process powershell.exe -ExecutionPolicy Bypass -File "$CountdownScript"
} else {
    Write-Log "錯誤：找不到倒數腳本 ($CountdownScript)！" "Red"
}
exit
