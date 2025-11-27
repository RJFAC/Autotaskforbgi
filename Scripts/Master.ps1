# =============================================================================
# AutoTask Master V5.2 - 增強日誌與狀態追蹤版
# =============================================================================

# ... (前段權限檢查、路徑定義保持不變) ...
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Set-Location $PSScriptRoot
$BaseDir = "C:\AutoTask"; $ScriptDir = "$BaseDir\Scripts"; $ConfigDir = "$BaseDir\Configs"; $FlagDir = "$BaseDir\Flags"; $LogDir = "$BaseDir\Logs"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $FlagDir)) { New-Item -Path $FlagDir -ItemType Directory -Force | Out-Null }
function Write-Log { param($Message, $Color="White"); $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"; $LogFile = Join-Path $LogDir "Master_$(Get-Date -Format 'yyyyMMdd').log"; Write-Host "[$TimeStamp] $Message" -ForegroundColor $Color; try { Add-Content -Path $LogFile -Value "[$TimeStamp] $Message" -Encoding UTF8 -ErrorAction SilentlyContinue } catch {} }
try { Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}

$RunFlag="$FlagDir\Run.flag"; $DoneFlag="$FlagDir\Done.flag"; $FailFlag="$FlagDir\Fail.flag"; $ManualFlag="$FlagDir\ManualTrigger.flag"; $ForceRunFlag="$FlagDir\ForceRun.flag"; $MonitorScript="$ScriptDir\Monitor.ps1"; $PauseLog="$ConfigDir\PauseDates.log"; $NoShutdownLog="$ConfigDir\NoShutdown.log"; $CountdownScript="$ScriptDir\Shutdown-Countdown.ps1"; $TaskStatus="$ConfigDir\TaskStatus.json"

# 讀取 1Remote
$1RemoteDir = "C:\AutoTask\1Remote"; $1RemoteExe = "$1RemoteDir\1Remote.exe"
if (Test-Path "$ConfigDir\EnvConfig.json") { try { $env = Get-Content "$ConfigDir\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json; if ($env.Path1Remote) { $1RemoteExe = $env.Path1Remote; $1RemoteDir = Split-Path $1RemoteExe -Parent } } catch {} }

function Check-Network {
    Write-Log "檢查網路..." 
    $Retry = 0; $MaxRetry = 12
    while ($Retry -lt $MaxRetry) {
        try { if (Test-Connection "8.8.8.8" -Count 1 -ErrorAction Stop) { Write-Log "網路正常。" "Green"; return $true } } catch {}
        Write-Log "網路未就緒... ($($Retry+1))" "Yellow"; Start-Sleep 5; $Retry++
    }
    Write-Log "⚠️ 網路連線逾時。" "Red"; return $false
}

Write-Log ">>> Master 啟動 (Admin Mode)..." "Cyan"

# --- [啟動判斷] ---
$IsResume = $false
if (Test-Path $RunFlag) {
    $P1 = Get-Process "1Remote" -ErrorAction SilentlyContinue
    if ($P1) {
        Write-Log "偵測到 Run.flag 與 1Remote (PID: $($P1.Id))，判定為 [熱重載/斷點續接]。" "Magenta"
        $IsResume = $true
    }
}

if (-not $IsResume) {
    # 全新啟動
    if (-not (Test-Path $ManualFlag)) {
        if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains (Get-Date).AddHours(-3).ToString("yyyyMMdd")) { Write-Log "今日暫停。" "Yellow"; exit } }
        $Now = Get-Date; $Target = (Get-Date).Date.AddHours(3).AddMinutes(55)
        if ($Now -lt $Target.AddMinutes(-20) -or $Now -gt $Target.AddMinutes(30)) { if (-not (Test-Path $RunFlag)) { Write-Log "非任務時間，退出。" "Gray"; exit } }
    } else {
        Write-Log "手動觸發，執行淨化..." "Magenta"
        Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
        try { Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $PID } | ForEach { Stop-Process -Id $_.ProcessId -Force } } catch {}
        $SessionOutput = qwinsta 2>$null | Select-String "\bRemote\b"
        if ($SessionOutput) {
            $Line = $SessionOutput.ToString().Trim() -replace "\s+", " "; $Parts = $Line.Split(" "); $SessionID = $null; foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }
            if ($SessionID) {
                Write-Log "強制登出 Remote (ID: $SessionID)..." "Yellow"
                cmd /c "logoff $SessionID"; $Wait=0; while($true){ if(-not(qwinsta 2>$null|Select-String "\bRemote\b")){break}; if($Wait-ge 20){break}; Start-Sleep 1; $Wait++ }
            }
        }
        if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }; Remove-Item $ManualFlag -Force; New-Item -Path $ForceRunFlag -ItemType File -Force | Out-Null
    }

    if (Test-Path $RunFlag) { Remove-Item $RunFlag -Force }; if (Test-Path $DoneFlag) { Remove-Item $DoneFlag -Force }; if (Test-Path $FailFlag) { Remove-Item $FailFlag -Force }
    New-Item -Path $RunFlag -ItemType File -Force | Out-Null

    Check-Network
    Write-Log "啟動 1Remote..."
    Start-Process -FilePath $1RemoteExe -WorkingDirectory $1RemoteDir; Start-Sleep 5
    Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir; Start-Sleep 2
}

# 監控啟動
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if (-not $MonitorProc) {
    Write-Log "啟動 Monitor..."
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`""
} else { Write-Log "Monitor 已在運行 (PID: $($MonitorProc.ProcessId))。" }

# --- [Master 監督迴圈] ---
$SupervisorStart = Get-Date; Write-Log ">>> 進入監督模式" "Green"; $PayloadLaunched = $false; if ($IsResume) { $PayloadLaunched = $true }

while ($true) {
    Start-Sleep 5
    if (Test-Path $DoneFlag) { Write-Log "任務成功 (Done)！" "Green"; break }
    if (Test-Path $FailFlag) { Write-Log "任務失敗 (Fail)！" "Red"; Remove-Item $RunFlag -Force; Stop-Process -Name "1Remote" -Force; exit }

    $MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
    if (-not $MonitorProc) { Write-Log "⚠️ Monitor 消失，重啟..." "Red"; Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MonitorScript`"" }

    $PayloadProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Payload.ps1*" }
    if (-not $PayloadProc) {
        if ($PayloadLaunched) { Write-Log "⚠️ Payload 消失，重啟..." "Red"; schtasks /run /tn "Auto_BetterGI_Payload"; Start-Sleep 10 }
        else {
            if ((Get-Date) -gt $SupervisorStart.AddMinutes(15)) { Write-Log "Payload 啟動超時，重試 RDP..." "Yellow"; Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir; $SupervisorStart = Get-Date }
        }
    } else { $PayloadLaunched = $true; $SupervisorStart = Get-Date }
}

# --- [清理] ---
Write-Log "清理中..."
Remove-Item $RunFlag -Force
Start-Sleep 5
$MonitorProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Monitor.ps1*" }
if ($MonitorProc) { Stop-Process -Id $MonitorProc.ProcessId -Force }

Write-Log "等待 Remote 登出..."
$Timeout=0; while ($true) { if (-not (qwinsta 2>$null | Select-String "\bRemote\b")) { Write-Log "Remote 已登出。" "Green"; break }; if ($Timeout -ge 60) { Write-Log "登出逾時。" "Yellow"; break }; Start-Sleep 3; $Timeout++ }

Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue; Remove-Item $DoneFlag -Force

if (Test-Path $NoShutdownLog) { if ((Get-Content $NoShutdownLog) -contains (Get-Date).AddHours(-3).ToString("yyyyMMdd")) { Write-Log "今日不關機。" "Cyan"; exit } }
if (Test-Path $CountdownScript) { Start-Process powershell.exe -ExecutionPolicy Bypass -File "$CountdownScript" }
exit