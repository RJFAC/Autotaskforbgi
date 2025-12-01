# =============================================================================
# AutoTask Payload V5.21 - ForceRun 邏輯修正與診斷增強
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
Write-Log "Payload 啟動 (V5.21) PID: $PID..." "Cyan"

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

# --- [輔助函數] ---
function Get-Error-Diagnosis { param($e,$p); $r="未知";$f="All Log"; switch($e){"LogLockFail"{$r="BGI啟動超時";$f="Payload.log"} "ProcessCrash"{$r="BGI閃退";$f="bg.log,pl.log"} "HeartbeatTimeout"{$r="卡死";$f="bg.log"} "NetworkError"{$r="斷網";$f="pl.log"}}; return @{Reason=$r;Files=$f;LogPath=$p} }
function Send-Notify-With-Diagnosis { 
    param($t,$m,$c,$d); 
    $f=@{"Err"=$d.Reason;"File"=$d.Files;"Bak"=$m}; 
    $targetLog = if($d.LogPath){$d.LogPath.FullName}else{""}
    if(Test-Path $NotifyScript){
        try{
            & $NotifyScript -Title $t -Message "Error: $($d.Reason)" -Color $c -Fields $f -LogPath $targetLog -Mention $true
        }catch{}
    } 
}
function Set-BetterGIResinConfig { param($c); if(-not(Test-Path $ResinConf)){return $false}; try{ $r=Get-Content $ResinConf -Raw|ConvertFrom-Json; if(-not $r.$c){return $false}; Write-Log "Resin: $c" "Cyan"; if(-not(Test-Path $BakFile)){Copy-Item $BettergiUserConf $BakFile -Force}; $b=Get-Content $BettergiUserConf -Raw|ConvertFrom-Json; $s=if($r.$c.TaskType-eq"Stygian"){$b.autoStygianOnslaughtConfig}else{$b.autoDomainConfig}; if($r.$c.Priority){$s.resinPriorityList=$r.$c.Priority}; if($r.$c.ResinMode-eq"Count"){$s.specifyResinUse=$true;$s.originalResinUseCount=$r.$c.Counts.Original}else{$s.specifyResinUse=$false}; $b|ConvertTo-Json -Depth 20|Set-Content $BettergiUserConf -Enc UTF8; return $true }catch{return $false} }
function Restore-BetterGIConfig { if(Test-Path $BakFile){try{Copy-Item $BakFile $BettergiUserConf -Force;Remove-Item $BakFile -Force;Write-Log "Resin Restored" "Gray"}catch{}} }
function Check-Network { $r=0; while($r-lt 12){if(Test-Connection "8.8.8.8" -Count 1 -Quiet){return $true};Start-Sleep 5;$r++};Write-Log "Net Fail" "Red";return $false }
function Send-Notify { param($t,$m,$c); if(Test-Path $NotifyScript){Start-Process powershell -Arg "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$t`" -Message `"$m`" -Color `"$c`"" -WindowStyle Hidden} }
function Backup-Logs { $t=Get-Date -Format "yyyyMMdd_HHmmss";$d="$BackupRootDir\Failed_$t";New-Item $d -ItemType Directory -Force|Out-Null;$l=(Get-Date).AddHours(-24);$da=New-Item "$d\AutoTask_Logs" -ItemType Directory;Get-ChildItem $LogDir -Filter "*.log"|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $da -Force;$db=New-Item "$d\BetterGI_Logs" -ItemType Directory;Get-ChildItem $LogDirBG -Filter "*.log"|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $db -Force;if($1RemoteLogDir-and(Test-Path $1RemoteLogDir)){$dr=New-Item "$d\1Remote_Logs" -ItemType Directory;Get-ChildItem $1RemoteLogDir -Include "*.md","*.log" -Recurse|Where{$_.LastWriteTime-gt$l}|Copy-Item -Dest $dr -Force};return $d }
function Update-Status { param($s,$r); try{$o=@{Date=(Get-Date).AddHours(-4).ToString("yyyyMMdd");Status=$s;RetryCount=$r;LastUpdate=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")};$o|ConvertTo-Json|Set-Content $TaskStatusFile}catch{} }
function Get-TargetConfig { $t=(Get-Date).AddHours(-4);$ds=$t.ToString("yyyyMMdd");if(Test-Path $DateMap){try{$m=Get-Content $DateMap;foreach($l in $m){if($l-match"^$ds=(.+)$"){return $matches[1]}}}catch{}};if(Test-Path $WeeklyConf){try{$w=Get-Content $WeeklyConf -Raw|ConvertFrom-Json;if($w){if($w.IT_Period_Days-gt 0 -and $t.Day-ge 1 -and $t.Day-le $w.IT_Period_Days){return $w.IT_Period_Config};return $w.$($t.DayOfWeek.ToString())}}catch{}};return "day" }
function Test-GenshinUpdateDay ($d) { $r=[datetime]"2024-08-28";$diff=($d.Date-$r).Days;if($diff-ge 0 -and $diff%42-eq 0){return $true}return $false }
function Check-GenshinPreDownload { if(-not $Global:GenshinPath){return};$r=[datetime]"2024-08-28";$d=((Get-Date).Date-$r).Days%42;if($d-ne 40-and $d-ne 41){return};Write-Log "Check PreDL" "Gray" }
function Cleanup-Screenshots { if(Test-Path $ScreenshotDir){try{Get-ChildItem $ScreenshotDir -Recurse|Where{$_.LastWriteTime-lt(Get-Date).AddDays(-30)}|Remove-Item -Force}catch{}} }

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
$TargetTime = $TodayLimit 
if ($CurrentTime.Hour -eq 3) { 
    while ((Get-Date) -lt $TargetTime) { 
        $Span = $TargetTime - (Get-Date)
        if ($Span.Seconds % 60 -eq 0) { 
            Write-Log "等待 04:00... 剩餘 $($Span.Minutes) 分" "Gray" 
            Update-Status "Waiting(04:00)" 0
        }
        Start-Sleep 1 
    } 
    $CurrentDateObj = (Get-Date).AddHours(-4)
    $CurrentDateStr = $CurrentDateObj.ToString("yyyyMMdd")
    Write-Log "已過 04:00，重新計算日期: $CurrentDateStr" "Cyan"
}

Write-Log "日期: $CurrentDateStr"
Cleanup-Screenshots

# --- [修正: ForceRun 標記檢查] ---
# 改動: 讀取標記，但不立即刪除，以防腳本崩潰後重啟無法識別
$IsForceRun = $false
if (Test-Path $ForceRunFlag) { 
    Write-Log "偵測到 ForceRun (保留標記以支援救援重啟)" "Magenta"
    $IsForceRun = $true 
    # Remove-Item $ForceRunFlag -Force # 移除此行
}

if (-not $IsForceRun) {
    if ((Test-Path $PauseLog) -and ((Get-Content $PauseLog) -contains $CurrentDateStr)) { 
        Write-Log "今日設定為暫停。腳本結束 (保留連線)。" "Yellow"
        Update-Status "Paused" 0
        New-Item $DoneFlag -Force | Out-Null
        exit 
    }
    
    if ((Test-Path $LastRunLog) -and ((Get-Content $LastRunLog) -eq $CurrentDateStr)) { 
        Write-Log "今日任務已完成。" "Green"
        Update-Status "Success" 0
        New-Item $DoneFlag -Force | Out-Null
        
        Write-Log "任務目標已達成，執行快速登出..." "Yellow"
        Start-Sleep 2
        logoff
        exit 
    }
} else {
    Write-Log "診斷: 處於 ForceRun 模式，強制跳過 LastRun 檢查。" "Gray"
}

$IsUpdateDay = Test-GenshinUpdateDay $CurrentDateObj
$UpdateResumeTime = $CurrentDateObj.Date.AddHours(11).AddMinutes(30)
if ($IsUpdateDay -and -not $IsForceRun) {
    if ((Get-Date) -lt $UpdateResumeTime) {
        Write-Log "⚠️ 版本更新日！待機至 11:30" "Magenta"; Send-Notify -Title "版本更新" -Msg "系統待機" -Color "Yellow"
        Check-GenshinPreDownload
        while ((Get-Date) -lt $UpdateResumeTime) { 
            Start-Sleep 60 
            Update-Status "Maintenance" 0
        }
        Write-Log "維護結束。" "Green"
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
    
    $PreCheckBG = Get-Process "BetterGI" -ErrorAction SilentlyContinue
    if ($PreCheckBG) { Write-Log "診斷: 啟動前偵測到殘留 BetterGI (PID: $($PreCheckBG.Id))" "Gray" }
    
    foreach ($CurrentConfig in $ConfigQueue) {
        if ([string]::IsNullOrWhiteSpace($CurrentConfig)) { continue }
        if ((Get-Item $SelfPath).LastWriteTime -ne $InitialWriteTime) { Write-Log "♻️ 腳本更新，重啟..." "Magenta"; Restore-BetterGIConfig; exit }

        Write-Log ">>> 執行: [$CurrentConfig]" "Cyan"
        
        $SkipStart = $false
        $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue
        if ($BgProc) {
            $RecentLog = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort LastWriteTime -Desc | Select -First 1
            if ($RecentLog -and $RecentLog.LastWriteTime -gt (Get-Date).AddMinutes(-5)) {
                Write-Log "熱修復: 接手監控..." "Yellow"; $SkipStart = $true
            } else {
                Write-Log "BetterGI 殭屍程序，清理。" "Red"; Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
            }
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
            Write-Log "錯誤：日誌鎖定失敗！" "Red"
            $IsFailed = $true 
            $ErrorType = "LogLockFail"
            Write-Log "--- [診斷資訊: 檔案列表] ---" "Gray"
            try {
                $DebugFiles = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log"
                if ($DebugFiles) {
                    foreach ($f in $DebugFiles) {
                        $TimeDiff = ((Get-Date) - $f.LastWriteTime).TotalMinutes
                        Write-Log "發現檔案: $($f.Name) | 時間: $($f.LastWriteTime.ToString('HH:mm:ss')) | 距今: $([math]::Round($TimeDiff, 1)) 分" "Gray"
                    }
                } else {
                    Write-Log "目錄內無任何符合 'better-genshin-impact*.log' 的檔案。" "Yellow"
                }
            } catch {
                Write-Log "無法讀取目錄: $_" "Red"
            }
            Write-Log "----------------------------" "Gray"
        } else {
            $LogPath = $LogFile.FullName
            $StartOffset = 0
            try { $StartOffset = (Get-Item $LogPath).Length } catch {}
            Write-Log "鎖定日誌: $($LogFile.Name) (初始 Offset: $StartOffset)" "Cyan"
            
            while (-not $IsSuccess -and -not $IsFailed) {
                Start-Sleep 5
                
                $NewContent = ""
                try {
                    $LogFile.Refresh()
                    $CurrentSize = $LogFile.Length
                    if ($CurrentSize -gt $StartOffset) {
                        $Stream = [System.IO.File]::Open($LogPath, 'Open', 'Read', 'ReadWrite')
                        $Reader = New-Object System.IO.StreamReader($Stream, [System.Text.Encoding]::UTF8)
                        $null = $Reader.BaseStream.Seek($StartOffset, [System.IO.SeekOrigin]::Begin)
                        $NewContent = $Reader.ReadToEnd()
                        $Reader.Close(); $Stream.Close()
                        $StartOffset = $CurrentSize
                    }
                } catch { Write-Log "讀取日誌警告: $_" "Yellow" }

                if ($NewContent -match "$SuccessKeyword|全部完成") {
                     Write-Log "偵測到完成訊號！(觸發: '$($matches[0])')" "Green"
                     $IsSuccess = $true; break 
                }
                
                if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                    Start-Sleep 3
                    if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                         Write-Log "BetterGI 進程已結束，檢查最後狀態..." "Yellow"
                         if ($NewContent -match "$SuccessKeyword|全部完成") {
                             Write-Log "完成(進程結束)。" "Green"; $IsSuccess=$true
                         } else {
                             Write-Log "意外退出！(未偵測到成功訊號)" "Red"; $IsFailed=$true; $ErrorType = "ProcessCrash"
                         }
                         break
                    }
                }
                
                $LogFile.Refresh()
                if (((Get-Date) - $LogFile.LastWriteTime).TotalMinutes -gt $HeartbeatLimit) { Write-Log "卡死判定！(日誌靜止超過 $HeartbeatLimit 分)" "Red"; $IsFailed=$true; $ErrorType = "HeartbeatTimeout" }
            }
        }

        if ($ConfigChanged -or $SkipStart) { Restore-BetterGIConfig }
        if (-not $IsSuccess) { $AllConfigSuccess = $false; break }
        Start-Sleep 5
    }

    if ($AllConfigSuccess) {
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
        
        # [修正] 任務成功後才清除 ForceRun 標記
        if ($IsForceRun) { 
            Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
            Write-Log "任務成功，清理 ForceRun 標記。" "Gray"
        }

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
            
            # [修正] 任務最終失敗後清除標記，避免無限循環
            if ($IsForceRun) { 
                Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
                Write-Log "任務失敗 (達最大重試)，清理 ForceRun 標記。" "Gray"
            }

            if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
            logoff
            exit
        }
        Start-Sleep 10
    }
}