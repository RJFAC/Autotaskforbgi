# =============================================================================
# AutoTask Payload V5.13 - 變數修復與防崩潰版
# =============================================================================
$ErrorActionPreference = "Stop"
trap {
    $ErrInfo = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] CRASH: $_ `nStackTrace: $($_.ScriptStackTrace)"
    Add-Content -Path "C:\AutoTask\PAYLOAD_CRASH.log" -Value $ErrInfo -Force
    exit 1
}

# --- [0. 身分驗證] ---
$TargetUser = "Remote" 
$CurrentUserName = [System.Environment]::UserName
$BaseDir    = "C:\AutoTask"
$LogDir     = "$BaseDir\Logs"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-PreLog {
    param($Msg, $Color="Red")
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogF = Join-Path $LogDir "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $Txt = "[$Time] $Msg"
    Write-Host $Txt -ForegroundColor $Color
    try { Add-Content -Path $LogF -Value $Txt -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

if ($CurrentUserName -ne $TargetUser) {
    Write-PreLog "⛔ 錯誤：身分不符！Payload 被 [$CurrentUserName] 誤觸，已攔截。" "Red"
    Start-Sleep 3; exit 
}

# --- [變數定義] ---
$ScriptDir  = "$BaseDir\Scripts"
$ConfigDir  = "$BaseDir\Configs"
$FlagDir    = "$BaseDir\Flags"
$BettergiDir    = "C:\Program Files\BetterGI"
$BettergiExe    = "$BettergiDir\BetterGI.exe"
$BettergiUserConf = "$BettergiDir\User\config.json"
$LogDirBG       = "$BettergiDir\log"
$ScreenshotDir  = "$BettergiDir\Screenshots" 
$1RemoteLogDir  = $null 

$TaskStatusFile = "$ConfigDir\TaskStatus.json"
$DoneFlag       = "$FlagDir\Done.flag"
$FailFlag       = "$FlagDir\Fail.flag"
$ForceRunFlag   = "$FlagDir\ForceRun.flag" 
$DateMap        = "$ConfigDir\DateConfig.map"
$WeeklyConf     = "$ConfigDir\WeeklyConfig.json"
$EnvConf        = "$ConfigDir\EnvConfig.json"
$ResinConf      = "$ConfigDir\ResinConfig.json"
$PauseLog       = "$ConfigDir\PauseDates.log"
$LastRunLog     = "$ConfigDir\LastRun.log"
$BackupRootDir  = "$BaseDir\LogBackups"
$NotifyScript   = "$ScriptDir\Notify.ps1"
$MaxRetries = 3
$SuccessKeyword = "一条龙.*任务结束"
$SelfPath = $PSCommandPath
$InitialWriteTime = (Get-Item $SelfPath).LastWriteTime

# [重要] 定義開始時間變數
$StartTime = Get-Date

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    try { Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# --- [1. 啟動前安全檢查] ---
Write-Log "Payload 啟動 (V5.13)..." "Cyan"

try {
    $CurrentPID = $PID
    $TargetScript = "Payload.ps1"
    $OldInstances = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | 
        Where-Object { $_.CommandLine -like "*$TargetScript*" -and $_.ProcessId -ne $CurrentPID }
    if ($OldInstances) {
        foreach ($proc in $OldInstances) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }
    }
} catch {}

$BakFile = "$BettergiUserConf.bak"

if (Test-Path $EnvConf) {
    try {
        $EnvJson = Get-Content $EnvConf -Raw -Encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-Json
        if ($EnvJson.GenshinPath) { $GenshinPath = $EnvJson.GenshinPath }
        if ($EnvJson.Path1Remote) { 
            $1RemoteDir = Split-Path $EnvJson.Path1Remote -Parent
            $1RemoteLogDir = "$1RemoteDir\.logs"
        }
    } catch {}
}

# --------------------------------------------------
# 輔助函數 (與 V5.11/V5.9 相同)
# --------------------------------------------------
function Get-Error-Diagnosis { param($e,$p); $r="未知";$f="All Log"; switch($e){"LogLockFail"{$r="BGI啟動超時";$f="Payload.log"} "ProcessCrash"{$r="BGI閃退";$f="bg.log,pl.log"} "HeartbeatTimeout"{$r="卡死";$f="bg.log"} "NetworkError"{$r="斷網";$f="pl.log"}}; return @{Reason=$r;Files=$f} }
function Send-Notify-With-Diagnosis { param($t,$m,$c,$d); $f=@{"Err"=$d.Reason;"File"=$d.Files;"Bak"=$m}; if(Test-Path $NotifyScript){try{& $NotifyScript -Title $t -Message "Error" -Color $c -Fields $f}catch{}} }
function Set-BetterGIResinConfig { param($c); if(-not(Test-Path $ResinConf)){return $false}; try{ $r=Get-Content $ResinConf -Raw|ConvertFrom-Json; if(-not $r.$c){return $false}; Write-Log "Resin: $c" "Cyan"; if(-not(Test-Path $BakFile)){Copy-Item $BettergiUserConf $BakFile -Force}; $b=Get-Content $BettergiUserConf -Raw|ConvertFrom-Json; $s=if($r.$c.TaskType-eq"Stygian"){$b.autoStygianOnslaughtConfig}else{$b.autoDomainConfig}; if($r.$c.Priority){$s.resinPriorityList=$r.$c.Priority}; if($r.$c.ResinMode-eq"Count"){$s.specifyResinUse=$true;$s.originalResinUseCount=$r.$c.Counts.Original}else{$s.specifyResinUse=$false}; $b|ConvertTo-Json -Depth 20|Set-Content $BettergiUserConf -Enc UTF8; return $true }catch{return $false} }
function Restore-BetterGIConfig { if(Test-Path $BakFile){try{Copy-Item $BakFile $BettergiUserConf -Force;Remove-Item $BakFile -Force;Write-Log "Resin Restored" "Gray"}catch{}} }
function Check-Network { $r=0; while($r-lt 12){if(Test-Connection "8.8.8.8" -Count 1 -Quiet){return $true};Start-Sleep 5;$r++};Write-Log "Net Fail" "Red";return $false }
function Send-Notify { param($t,$m,$c); if(Test-Path $NotifyScript){Start-Process powershell -Arg "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$t`" -Message `"$m`" -Color `"$c`"" -WindowStyle Hidden} }
function Check-Success-Log { param($p); try{$l=Get-Content $p -Tail 50 -Enc UTF8 -EA SilentlyContinue;if($l-match "一条龙.*任务结束"){return $true}}catch{};return $false }
function Backup-Logs { $t=Get-Date -Format "yyyyMMdd_HHmmss";$d="$BackupRootDir\Failed_$t";New-Item $d -ItemType Directory -Force|Out-Null;$l=(Get-Date).AddHours(-24);$da=New-Item "$d\AutoTask_Logs" -ItemType Directory;Get-ChildItem $LogDir -Filter "*.log"|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $da -Force;$db=New-Item "$d\BetterGI_Logs" -ItemType Directory;Get-ChildItem $LogDirBG -Filter "*.log"|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $db -Force;if($1RemoteLogDir-and(Test-Path $1RemoteLogDir)){$dr=New-Item "$d\1Remote_Logs" -ItemType Directory;Get-ChildItem $1RemoteLogDir -Include "*.md","*.log" -Recurse|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $dr -Force};return $d }
function Update-Status { param($s,$r); try{$o=@{Date=(Get-Date).AddHours(-4).ToString("yyyyMMdd");Status=$s;RetryCount=$r;LastUpdate=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")};$o|ConvertTo-Json|Set-Content $TaskStatusFile}catch{} }
function Get-TargetConfig { $t=(Get-Date).AddHours(-4);$ds=$t.ToString("yyyyMMdd");if(Test-Path $DateMap){try{$m=Get-Content $DateMap;foreach($l in $m){if($l-match"^$ds=(.+)$"){return $matches[1]}}}catch{}};if(Test-Path $WeeklyConf){try{$w=Get-Content $WeeklyConf -Raw|ConvertFrom-Json;if($w){if($w.IT_Period_Days-gt 0 -and $t.Day-ge 1 -and $t.Day-le $w.IT_Period_Days){return $w.IT_Period_Config};return $w.$($t.DayOfWeek.ToString())}}catch{}};return "day" }
function Test-GenshinUpdateDay ($d) { $r=[datetime]"2024-08-28";$diff=($d.Date-$r).Days;if($diff-ge 0 -and $diff%42-eq 0){return $true}return $false }
function Check-GenshinPreDownload { if(-not $Global:GenshinPath){return};$r=[datetime]"2024-08-28";$d=((Get-Date).Date-$r).Days%42;if($d-ne 40-and $d-ne 41){return};Write-Log "Check PreDL" "Gray" }
function Cleanup-Screenshots { if(Test-Path $ScreenshotDir){try{Get-ChildItem $ScreenshotDir -Recurse|Where{$_.LastWriteTime-lt(Get-Date).AddDays(-30)}|Remove-Item -Force}catch{}} }
# --------------------------------------------------

$CurrentDateObj = (Get-Date).AddHours(-4)
$CurrentDateStr = $CurrentDateObj.ToString("yyyyMMdd")

# ForceEnd (03:30~04:00)
$CurrentTime = Get-Date
$TodayLimit = $CurrentTime.Date.AddHours(4) 
if ($CurrentTime.Hour -ge 4) { $TodayLimit = $TodayLimit.AddDays(1) }
$ForceEndStart = $TodayLimit.AddMinutes(-30)

if ($CurrentTime -ge $ForceEndStart -and $CurrentTime -lt $TodayLimit) {
    Check-Network
    $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue; $GiProc = Get-Process "GenshinImpact" -ErrorAction SilentlyContinue
    if ($BgProc -or $GiProc) {
        Write-Log "執行 ForceEnd..." "Yellow"
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue; Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3
        $AppliedForce = Set-BetterGIResinConfig "forceend"
        $ForceProc = Start-Process -FilePath $BettergiExe -ArgumentList "--startOneDragon ""forceend""" -WorkingDirectory $BettergiDir -PassThru
        while (-not $ForceProc.HasExited) { if ((Get-Date) -ge $TodayLimit) { $ForceProc.Kill(); break }; Start-Sleep 2; $ForceProc.Refresh() }
        if ($AppliedForce) { Restore-BetterGIConfig }; Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue; Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 2
    }
}

# Wait 04:00
if ($CurrentTime.Hour -eq 3) { 
    while ((Get-Date) -lt $TargetTime) { 
        $Span = $TargetTime - (Get-Date)
        if ($Span.Seconds % 60 -eq 0) { Write-Log "等待 04:00... 剩餘 $($Span.Minutes) 分" "Gray" }
        Start-Sleep 1 
    } 
    $CurrentDateObj = (Get-Date).AddHours(-4)
    $CurrentDateStr = $CurrentDateObj.ToString("yyyyMMdd")
    Write-Log "已過 04:00，重新計算日期: $CurrentDateStr" "Cyan"
}

Write-Log "日期: $CurrentDateStr"
Cleanup-Screenshots

$IsForceRun = $false
if (Test-Path $ForceRunFlag) { Write-Log "偵測到 ForceRun" "Magenta"; $IsForceRun = $true; Remove-Item $ForceRunFlag -Force }

if (-not $IsForceRun) {
    if ((Test-Path $PauseLog) -and ((Get-Content $PauseLog) -contains $CurrentDateStr)) { Write-Log "今日暫停。"; Update-Status "Paused" 0; Stop-Process -Name "BetterGI" -Force; Stop-Process -Name "GenshinImpact" -Force; New-Item $DoneFlag -Force; logoff; exit }
    if ((Test-Path $LastRunLog) -and ((Get-Content $LastRunLog) -eq $CurrentDateStr)) { Update-Status "Success" 0; New-Item $DoneFlag -Force; logoff; exit }
}

$IsUpdateDay = Test-GenshinUpdateDay $CurrentDateObj
$UpdateResumeTime = $CurrentDateObj.Date.AddHours(11).AddMinutes(30)
if ($IsUpdateDay -and -not $IsForceRun) {
    if ((Get-Date) -lt $UpdateResumeTime) {
        Write-Log "⚠️ 版本更新日！待機至 11:30" "Magenta"; Send-Notify -Title "版本更新" -Msg "系統待機" -Color "Yellow"
        Check-GenshinPreDownload
        while ((Get-Date) -lt $UpdateResumeTime) { Start-Sleep 60 }; Write-Log "維護結束。" "Green"
    }
}

$RetryCount = 0
$ConfigName = Get-TargetConfig
$ConfigQueue = $ConfigName -split ","

if (-not (Test-Path $BettergiExe)) {
    Write-Log "⛔ 錯誤：找不到 BetterGI ($BettergiExe)！" "Red"
    Send-Notify "執行失敗" "找不到 BetterGI" "Red"
    exit
}

Check-Network
if (-not (Check-Network)) { $ErrorType = "NetworkError" }

while ($RetryCount -le $MaxRetries) {
    Update-Status "Running" $RetryCount
    $AllConfigSuccess = $true
    $ErrorType = ""
    
    foreach ($CurrentConfig in $ConfigQueue) {
        if ([string]::IsNullOrWhiteSpace($CurrentConfig)) { continue }
        if ((Get-Item $SelfPath).LastWriteTime -ne $InitialWriteTime) { Write-Log "♻️ 腳本更新，重啟..." "Magenta"; Restore-BetterGIConfig; exit }

        Write-Log ">>> 執行: [$CurrentConfig]" "Cyan"
        
        $SkipStart = $false
        $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue
        if ($BgProc) {
            $RecentLog = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort LastWriteTime -Desc | Select -First 1
            if ($RecentLog -and $RecentLog.LastWriteTime -gt (Get-Date).AddMinutes(-5)) { Write-Log "熱修復: 接手監控..." "Yellow"; $SkipStart = $true }
            else { Write-Log "BetterGI 殭屍程序，清理。" "Red"; Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue }
        }

        if (-not $SkipStart) {
            $ConfigChanged = Set-BetterGIResinConfig $CurrentConfig
            Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
            if ($RetryCount -gt 0) { Write-Log "重試: 重啟遊戲..." "Yellow"; Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue }
            Start-Sleep 3
            Start-Process -FilePath $BettergiExe -ArgumentList "--startOneDragon ""$CurrentConfig""" -WorkingDirectory $BettergiDir
        }

        $WatchdogStart = Get-Date
        $IsSuccess = $false; $IsFailed = $false
        $HeartbeatLimit = if ($IsUpdateDay) { 60 } else { 15 }
        
        $LogFile = $null
        for ($i=0; $i -lt 90; $i++) {
            $Candidate = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort LastWriteTime -Desc | Select -First 1
            if ($Candidate -and $Candidate.LastWriteTime -gt (Get-Date).AddMinutes(-5)) { $LogFile = $Candidate; break }
            Start-Sleep 1
        }

        if (-not $LogFile) { 
            Write-Log "錯誤：日誌鎖定失敗！" "Red"; $IsFailed = $true 
            $ErrorType = "LogLockFail"
        } else {
            Write-Log "鎖定日誌: $($LogFile.Name)"
            $CurrentSize = 0; try { $CurrentSize = (Get-Item $LogFile.FullName).Length } catch {}
            while (-not $IsSuccess -and -not $IsFailed) {
                Start-Sleep 5
                if (Check-Success-Log $LogFile.FullName) { Write-Log "完成！" "Green"; $IsSuccess=$true; break }
                if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                    Start-Sleep 3
                    if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                         if (Check-Success-Log $LogFile.FullName) { Write-Log "完成(進程結束)。" "Green"; $IsSuccess=$true }
                         else { Write-Log "意外退出！" "Red"; $IsFailed=$true; $ErrorType = "ProcessCrash" }
                         break
                    }
                }
                $LogFile.Refresh()
                if (((Get-Date) - $LogFile.LastWriteTime).TotalMinutes -gt $HeartbeatLimit) { Write-Log "卡死判定！" "Red"; $IsFailed=$true; $ErrorType = "HeartbeatTimeout" }
            }
        }

        if ($ConfigChanged -or $SkipStart) { Restore-BetterGIConfig }
        if (-not $IsSuccess) { $AllConfigSuccess = $false; break }
        Start-Sleep 5
    }

    if ($AllConfigSuccess) {
        # [變數修復] 加入 try-catch 防止計算耗時崩潰
        try {
            if ($StartTime) {
                $Duration = New-TimeSpan -Start $StartTime -End (Get-Date)
                $DurStr = "{0:hh}時{0:mm}分" -f $Duration
            } else { $DurStr = "未知" }
        } catch { $DurStr = "未知(CalcErr)" }
        
        Write-Log ">>> 全部完成。耗時: $DurStr" "Green"
        Send-Notify "Success" "Config: $ConfigName" "Green"
        Update-Status "Success" $RetryCount
        Set-Content $LastRunLog -Value $CurrentDateStr
        while (Get-Process "GenshinImpact" -ErrorAction SilentlyContinue) { Stop-Process -Name "GenshinImpact" -Force; Start-Sleep 5 }
        New-Item $DoneFlag -Force | Out-Null
        if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
        logoff
        exit
    } else {
        $RetryCount++
        if ($RetryCount -gt $MaxRetries) {
            $BackupPath = Backup-Logs
            $Diagnosis = Get-Error-Diagnosis $ErrorType $LogFile
            Send-Notify-With-Diagnosis "Fail" $BackupPath "Red" $Diagnosis
            Update-Status "Failed" $RetryCount
            New-Item $FailFlag -Force | Out-Null
            if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
            logoff
            exit
        }
        Start-Sleep 10
    }
}