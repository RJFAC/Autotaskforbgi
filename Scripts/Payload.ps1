# =============================================================================
# AutoTask Payload V5.3 - 遊戲駐留優化版 (不殺後台遊戲)
# =============================================================================

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

# --- [路徑與變數] ---
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
Write-Log "Payload 啟動 (V5.3 - 遊戲駐留)..." "Cyan"

# 檢查自我清理
try {
    $CurrentPID = $PID
    $TargetScript = "Payload.ps1"
    $OldInstances = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | 
        Where-Object { $_.CommandLine -like "*$TargetScript*" -and $_.ProcessId -ne $CurrentPID }
    if ($OldInstances) {
        foreach ($proc in $OldInstances) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }
    }
} catch {}

# 檢查殘留備份
$BakFile = "$BettergiUserConf.bak"

# --- [載入環境設定] ---
if (Test-Path $EnvConf) {
    try {
        $EnvJson = Get-Content $EnvConf -Raw -Encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-Json
        if ($EnvJson.GenshinPath) { $GenshinPath = $EnvJson.GenshinPath }
    } catch {}
}

# --- [輔助函數] ---
function Set-BetterGIResinConfig {
    param([string]$ConfigName)
    if (-not (Test-Path $ResinConf)) { return $false }
    if (-not (Test-Path $BettergiUserConf)) { return $false }
    try {
        $Rules = Get-Content $ResinConf -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $Rules.$ConfigName) { return $false }
        $Rule = $Rules.$ConfigName
        Write-Log "套用樹脂策略: [$ConfigName]" "Cyan"
        if (-not (Test-Path $BakFile)) { Copy-Item $BettergiUserConf $BakFile -Force }
        $BgiConfig = Get-Content $BettergiUserConf -Raw -Encoding UTF8 | ConvertFrom-Json
        $TargetSection = if ($Rule.TaskType -eq "Stygian") { "autoStygianOnslaughtConfig" } else { "autoDomainConfig" }
        if (-not $BgiConfig.$TargetSection) { return $false }
        $Section = $BgiConfig.$TargetSection
        if ($Rule.Priority) { $Section.resinPriorityList = $Rule.Priority }
        if ($Rule.ResinMode -eq "Count") {
            $Section.specifyResinUse = $true
            if ($Rule.Counts) {
                $Section.originalResinUseCount = $Rule.Counts.Original
                $Section.condensedResinUseCount = $Rule.Counts.Condensed
                $Section.transientResinUseCount = $Rule.Counts.Transient
                $Section.fragileResinUseCount = $Rule.Counts.Fragile
            }
        } else {
            $Section.specifyResinUse = $false
        }
        $BgiConfig | ConvertTo-Json -Depth 20 | Set-Content $BettergiUserConf -Encoding UTF8
        return $true
    } catch { return $false }
}
function Restore-BetterGIConfig {
    if (Test-Path $BakFile) { try { Copy-Item $BakFile $BettergiUserConf -Force; Remove-Item $BakFile -Force; Write-Log "已還原 BetterGI 預設設定。" "Gray" } catch {} }
}
function Check-Network { $r=0; while($r-lt 12){ if(Test-Connection "8.8.8.8" -Count 1 -Quiet){return $true}; Start-Sleep 5; $r++ }; Write-Log "網路逾時" "Red"; return $false }
function Send-Notify { param($Title,$Msg,$Color); if(Test-Path $NotifyScript){ Start-Process powershell -Arg "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$Title`" -Message `"$Msg`" -Color `"$Color`"" -WindowStyle Hidden } }
function Check-Success-Log { param($p); try{ $l=Get-Content $p -Tail 50 -Enc UTF8 -EA SilentlyContinue; if($l-match "一条龙.*任务结束"){return $true} }catch{}; return $false }
function Backup-Logs { $t=Get-Date -Format "yyyyMMdd_HHmmss"; $d="$BackupRootDir\Failed_$t"; New-Item $d -ItemType Directory -Force|Out-Null; Copy-Item "$LogDir\*" "$d\AutoTask_Logs" -Recurse -Force -EA SilentlyContinue; Copy-Item "$LogDirBG\*" "$d\BetterGI_Logs" -Recurse -Force -EA SilentlyContinue; return $d }
function Update-Status { param($s,$r); try{ $o=@{Date=(Get-Date).AddHours(-3).ToString("yyyyMMdd");Status=$s;RetryCount=$r;LastUpdate=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}; $o|ConvertTo-Json|Set-Content $TaskStatusFile }catch{} }
function Get-TargetConfig { 
    $t=(Get-Date).AddHours(-3); $ds=$t.ToString("yyyyMMdd"); 
    if(Test-Path $DateMap){try{$m=Get-Content $DateMap; foreach($l in $m){if($l-match"^$ds=(.+)$"){return $matches[1]}}}catch{}}
    if(Test-Path $WeeklyConf){try{$w=Get-Content $WeeklyConf -Raw|ConvertFrom-Json; if($w){if($w.IT_Period_Days-gt 0 -and $t.Day-ge 1 -and $t.Day-le $w.IT_Period_Days){return $w.IT_Period_Config}; return $w.$($t.DayOfWeek.ToString())}}catch{}}
    return "day" 
}
function Check-Failure-Log ($logPath) {
    try {
        # 讀取最後 50 行
        $Logs = Get-Content $logPath -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue
        
        if ($Logs) {
            # 1. 檢查嚴重錯誤標記 [ERR]
            if ($Logs -match "\[ERR\]") {
                # [排除] 忽略 "鋤地規劃統計異常" (這是假警報，不影響任務)
                if ($Logs -match "锄地规划统计异常") { 
                    # 可以在這裡寫個 Log 記錄但不回傳錯誤
                    # Write-Log "忽略非致命錯誤: 鋤地規劃統計異常" "Gray"
                    return $false 
                }

                # [確認] 其他 ERR 視為真失敗
                if ($Logs -match "模拟键鼠消息发送失败") { return "InputError" }
                if ($Logs -match "Hosting failed to start") { return "CrashError" }
                if ($Logs -match "版本号要求") { return "VersionError" }
                
                # 其他未列出的 ERR 也視為失敗 (寧殺錯不放過，除了已知的排除項)
                return "GeneralError"
            }

            # 2. 檢查軟性卡死 (連續截圖失敗)
            # 如果最後 50 行裡面有超過 15 行都是截圖失敗，判定為黑屏或被遮擋
            $ScreenshotFailCount = ($Logs | Select-String "截图失败").Count
            if ($ScreenshotFailCount -gt 15) {
                return "ScreenshotError"
            }
        }

    } catch {}
    return $false
}
function Test-GenshinUpdateDay ($d) { $ref=[datetime]"2024-08-28"; $diff=($d.Date-$ref).Days; if($diff-ge 0 -and $diff%42-eq 0){return $true} return $false }
function Check-GenshinPreDownload {
    if (-not (Test-Path $Global:GenshinPath)) { return }
    $ref=[datetime]"2024-08-28"; $diff=((Get-Date).Date-$ref).Days%42
    if ($diff -eq 40 -or $diff -eq 41) {
         Write-Log "檢查預下載..." "Cyan"
         # (此處保留預下載詳細邏輯，為確保腳本可執行，此處省略，請保持 V4.9 內容)
    }
}
function Cleanup-Screenshots { if(Test-Path $ScreenshotDir){try{Get-ChildItem $ScreenshotDir -Recurse|Where{$_.LastWriteTime-lt(Get-Date).AddDays(-30)}|Remove-Item -Force}catch{}} }

# =============================================================================
# [主流程]
# =============================================================================
$CurrentDateObj = (Get-Date).AddHours(-3)
$CurrentDateStr = $CurrentDateObj.ToString("yyyyMMdd")

# ForceEnd 邏輯 (03:50-04:00)
$CurrentTime = Get-Date; $TodayLimit = $CurrentTime.Date.AddHours(4); if ($CurrentTime.Hour -ge 4) { $TodayLimit = $TodayLimit.AddDays(1) }
if ($CurrentTime -lt $TodayLimit -and $CurrentTime.Hour -lt 4) {
    Check-Network
    $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue; $GiProc = Get-Process "GenshinImpact" -ErrorAction SilentlyContinue
    if ($BgProc -or $GiProc) {
        Write-Log "偵測到前一日殘留，執行 forceend..." "Yellow"
        Send-Notify -Title "系統維護" -Msg "偵測到前日殘留，執行 ForceEnd..." -Color "Yellow"
        Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue; Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 3
        $ForceArgs = "--startOneDragon ""forceend"""
        $AppliedForce = Set-BetterGIResinConfig "forceend"
        $ForceProc = Start-Process -FilePath $BettergiExe -ArgumentList $ForceArgs -WorkingDirectory $BettergiDir -PassThru
        while (-not $ForceProc.HasExited) { if ((Get-Date) -ge $TodayLimit) { $ForceProc.Kill(); break }; Start-Sleep 2; $ForceProc.Refresh() }
        Write-Log "ForceEnd 結束。"; if ($AppliedForce) { Restore-BetterGIConfig }; Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue; Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 2
    }
}

# 等待 04:00
$TargetTime = (Get-Date).Date.AddHours(4); if ($TargetTime -lt (Get-Date)) { $TargetTime = $TargetTime.AddDays(1) }
if ((Get-Date).Hour -lt 4) { while ((Get-Date) -lt $TargetTime) { Start-Sleep 30 } }

Write-Log "今日日期: $CurrentDateStr"
Cleanup-Screenshots

$IsForceRun = $false
if (Test-Path $ForceRunFlag) { Write-Log "偵測到 ForceRun" "Magenta"; $IsForceRun = $true; Remove-Item $ForceRunFlag -Force }

# 檢查暫停與完成 (略，保持 V4.12)
if (-not $IsForceRun) {
    if ((Test-Path $PauseLog) -and ((Get-Content $PauseLog) -contains $CurrentDateStr)) { Write-Log "今日暫停。"; Update-Status "Paused" 0; Stop-Process -Name "BetterGI" -Force; Stop-Process -Name "GenshinImpact" -Force; New-Item $DoneFlag -Force; logoff; exit }
    if ((Test-Path $LastRunLog) -and ((Get-Content $LastRunLog) -eq $CurrentDateStr)) { Update-Status "Success" 0; New-Item $DoneFlag -Force; logoff; exit }
}

# 更新日待機
$IsUpdateDay = Test-GenshinUpdateDay $CurrentDateObj
$UpdateResumeTime = $CurrentDateObj.Date.AddHours(11).AddMinutes(30)
if ($IsUpdateDay -and -not $IsForceRun) {
    if ((Get-Date) -lt $UpdateResumeTime) {
        Write-Log "⚠️ 版本更新日！" "Magenta"; Send-Notify -Title "版本更新" -Msg "系統待機至 11:30" -Color "Yellow"
        Check-GenshinPreDownload
        while ((Get-Date) -lt $UpdateResumeTime) { Start-Sleep 60 }; Write-Log "維護結束。" "Green"
    }
}

$RetryCount = 0
$ConfigName = Get-TargetConfig
$ConfigQueue = $ConfigName -split ","
if (-not (Test-Path $BettergiExe)) { exit }

Check-Network

while ($RetryCount -le $MaxRetries) {
    Update-Status "Running" $RetryCount
    $AllConfigSuccess = $true
    
    foreach ($CurrentConfig in $ConfigQueue) {
        if ([string]::IsNullOrWhiteSpace($CurrentConfig)) { continue }
        
        Write-Log ">>> 準備執行: [$CurrentConfig]" "Cyan"
        
        # 熱修復偵測
        $SkipStart = $false
        $BgProc = Get-Process "BetterGI" -ErrorAction SilentlyContinue
        if ($BgProc) {
            $RecentLog = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort LastWriteTime -Desc | Select -First 1
            if ($RecentLog -and $RecentLog.LastWriteTime -gt (Get-Date).AddMinutes(-5)) {
                Write-Log "偵測到 BetterGI 活躍，嘗試熱修復..." "Yellow"; $SkipStart = $true
            } else {
                Write-Log "BetterGI 殭屍程序，清理。" "Red"; Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
            }
        }

        if (-not $SkipStart) {
            $ConfigChanged = Set-BetterGIResinConfig $CurrentConfig
            Stop-Process -Name "BetterGI" -Force -ErrorAction SilentlyContinue
            
            # [核心修正] 遊戲駐留邏輯
            # 只有在「重試模式」下才強制殺死遊戲，否則保留遊戲以節省時間
            if ($RetryCount -gt 0) {
                 Write-Log "重試模式：強制重啟遊戲..." "Yellow"
                 Stop-Process -Name "GenshinImpact" -Force -ErrorAction SilentlyContinue
            }
            
            Start-Sleep 3
            $Args = "--startOneDragon ""$CurrentConfig"""
            Start-Process -FilePath $BettergiExe -ArgumentList $Args -WorkingDirectory $BettergiDir
        } else {
            Write-Log "跳過啟動步驟，直接鎖定。"
        }

        # 監控邏輯
        $WatchdogStart = Get-Date
        $IsSuccess = $false; $IsFailed = $false
        $HeartbeatLimit = if ($IsUpdateDay) { 60 } else { 15 }
        
        $LogFile = $null
        for ($i=0; $i -lt 90; $i++) {
            $Candidate = Get-ChildItem $LogDirBG -Filter "better-genshin-impact*.log" | Sort LastWriteTime -Desc | Select -First 1
            if ($Candidate -and $Candidate.LastWriteTime -gt (Get-Date).AddMinutes(-5)) { $LogFile = $Candidate; break }
            Start-Sleep 1
        }

        if (-not $LogFile) { Write-Log "錯誤：日誌鎖定失敗！" "Red"; $IsFailed = $true }
        else {
            Write-Log "鎖定日誌: $($LogFile.Name)"
            while (-not $IsSuccess -and -not $IsFailed) {
                Start-Sleep 5
		# ... (在 while (-not $IsSuccess -and -not $IsFailed) 迴圈內部) ...

        	# [新增] 檢查日誌是否包含失敗訊號
        	$ErrorStatus = Check-Failure-Log $LogFile.FullName
 		if ($ErrorStatus) {
   	        Write-Log "偵測到任務失敗訊號！類型: $ErrorStatus" "Red"
        	$IsFailed = $true
    	        # 這裡會跳出迴圈，觸發外層的重試機制 (Retry)
        	break
        }
        
        # ... (原本的 Check-Success-Log 和 Get-Process 檢查保持不變) ...
                if (Check-Success-Log $LogFile.FullName) { Write-Log "配置完成！" "Green"; $IsSuccess = $true; break }
                if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                    Start-Sleep 3
                    if (-not (Get-Process "BetterGI" -ErrorAction SilentlyContinue)) {
                         if (Check-Success-Log $LogFile.FullName) { Write-Log "確認完成。" "Green"; $IsSuccess=$true }
                         else { Write-Log "意外退出！" "Red"; $IsFailed=$true }
                         break
                    }
                }
                $LogFile.Refresh()
                if (((Get-Date) - $LogFile.LastWriteTime).TotalMinutes -gt $HeartbeatLimit) { Write-Log "卡死判定！" "Red"; $IsFailed=$true }
            }
        }

        if ($ConfigChanged -or $SkipStart) { Restore-BetterGIConfig }

        if (-not $IsSuccess) { $AllConfigSuccess = $false; break }
        Start-Sleep 5
    }

    if ($AllConfigSuccess) {
        $Duration = New-TimeSpan -Start $StartTime -End (Get-Date)
        Write-Log ">>> 全部完成。耗時: {0:hh}時{0:mm}分" -f $Duration "Green"
        Send-Notify -Title "任務成功" -Msg "配置 [$ConfigName] 已完成。" -Color "Green"
        Update-Status "Success" $RetryCount
        Set-Content $LastRunLog -Value $CurrentDateStr
        
        # 任務全部完成後，才關閉遊戲 (或者不關)
        # 根據需求：任務結束後關閉遊戲
        while (Get-Process "GenshinImpact" -ErrorAction SilentlyContinue) { Stop-Process -Name "GenshinImpact" -Force; Start-Sleep 5 }
        
        New-Item $DoneFlag -Force | Out-Null
        if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
        logoff
        exit
    } else {
        $RetryCount++
        if ($RetryCount -gt $MaxRetries) {
            Backup-Logs
            Update-Status "Failed" $RetryCount
            Send-Notify -Title "任務失敗" -Msg "已達最大重試次數。" -Color "Red"
            New-Item $FailFlag -Force | Out-Null
            if ($IsForceRun) { Start-Sleep 10 } else { Start-Sleep 3 }
            logoff
            exit
        }
    }
}