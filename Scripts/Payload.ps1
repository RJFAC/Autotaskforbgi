# ==============================================================================
# AutoTask Payload Script V5.53 (Notify Fix)
# ------------------------------------------------------------------------------
# V5.53:
#   1. [Critical] 修復 Notify 函式未定義導致的崩潰問題。新增本地 Notify 包裝函式。
# V5.52:
#   1. [Startup] 03:45~03:55 間啟動，視為新的一天前置，不執行 ForceEnd，直接 Done。
#   2. [Runtime] 僅在 "執行中" 跨越 03:45 時，才觸發 ForceEnd 攔截與收尾。
# ==============================================================================

# 1. 初始化與環境設定
$WorkDir = "C:\AutoTask"
$Script:LogDir = "$WorkDir\\Logs"
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\\Payload_$DateStr.log"
$FlagDir = "$WorkDir\\Flags"
$DoneFlag = "$FlagDir\\Done.flag"
$WeeklyConfFile = "$WorkDir\\Configs\\WeeklyConfig.json"
$TaskStatusFile = "$WorkDir\\Configs\\TaskStatus.json"
$LastRunFile = "$WorkDir\\Configs\\LastRun.log"
$ForceRunFlag = "$FlagDir\\ForceRun.flag"
$ScriptDir = "$WorkDir\\Scripts"

if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# 載入 Discord 模組並定義 Notify
if (Test-Path "$ScriptDir\\Lib_Discord.ps1") { . "$ScriptDir\\Lib_Discord.ps1" } 

# [V5.53 Fix] 定義 Notify 轉接函式，確保代碼相容性
function Notify {
    param(
        [string]$Title, 
        [string]$Message, 
        [string]$Color="Blue"
    )
    # 若 Send-DiscordNotification 存在 (已載入 Lib)，則呼叫它
    if (Get-Command Send-DiscordNotification -ErrorAction SilentlyContinue) {
        Send-DiscordNotification -Title $Title -Message $Message -Color $Color
    } else {
        # 若 Lib 不存在，僅輸出到 Console (Dummy)
        Write-Host "[$Title] $Message" -ForegroundColor $Color
    }
}

function Write-Log {
    param([string]$Msg, [string]$Color="White")
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$Time] [$Color] $Msg" -ForegroundColor $Color
    "[$Time] [$Color] $Msg" | Out-File -Append -FilePath $LogFile -Encoding UTF8
}

function Update-TaskStatus {
    param([string]$Status)
    try {
        $JsonData = @{ "Date" = Get-Date -Format "yyyy/MM/dd"; "Status" = $Status; "LastUpdate" = Get-Date -Format "HH:mm:ss" }
        $JsonData | ConvertTo-Json -Depth 2 | Set-Content -Path $TaskStatusFile -Encoding UTF8
    } catch { Write-Log "更新狀態失敗: $_" "Red" }
}

Write-Log ">>> Payload 啟動 (V5.53)..." "Cyan"

# --- [1. 啟動時時序檢查 (Startup Check)] ---
$Now = Get-Date
$ForceEndStart = $Now.Date.AddHours(3).AddMinutes(45)
$ForceEndDeadline = $Now.Date.AddHours(3).AddMinutes(55)

# [情境 A] 03:45 ~ 03:55 之間啟動 (Fresh Start / Restart)
if ($Now -ge $ForceEndStart -and $Now -lt $ForceEndDeadline) {
    Write-Log "啟動於 ForceEnd 緩衝區 (03:45~03:55)。" "Yellow"
    Write-Log "判定為新啟動，跳過 ForceEnd 配置組，直接執行收尾。" "Green"
    Notify "🧹 自動收尾" "系統於緩衝區間啟動，執行清理並等待換日。" "Green"
    
    Stop-Process -Name "BetterGI", "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
    New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
    shutdown /l
    exit
}

# --- [2. 日期與 LastRun 檢查] ---
$TodayKey = $Now.ToString("yyyyMMdd")
if ($Now.Hour -lt 4) { $TodayKey = $Now.AddDays(-1).ToString("yyyyMMdd") }

if (-not (Test-Path $ForceRunFlag)) {
    if (Test-Path $LastRunFile) {
        $LastDate = Get-Content $LastRunFile -Raw
        if ($LastDate.Trim() -eq $TodayKey) {
            Write-Log "✅ 今日任務已完成 ($TodayKey)。退出。" "Green"
            exit
        }
    }
}

# --- [3. 配置讀取 (正常模式)] ---
$TaskList = @()
$ConfigName = "Default"

# 3.1 優先讀取 DateConfig.map
if (Test-Path "$WorkDir\Configs\DateConfig.map") {
    $MapContent = Get-Content "$WorkDir\Configs\DateConfig.map"
    foreach ($line in $MapContent) {
        if ($line -match "^$TodayKey=(.+)$") {
            $ConfigName = $Matches[1]
            Write-Log "使用指定日期配置: $ConfigName" "Cyan"
            break
        }
    }
}

# 3.2 讀取 WeeklyConfig
if ($ConfigName -eq "Default") {
    if (Test-Path $WeeklyConfFile) {
        try {
            $Weekly = Get-Content $WeeklyConfFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $DayOfWeek = $Now.DayOfWeek.ToString()
            if ($Now.Hour -lt 4) { $DayOfWeek = $Now.AddDays(-1).DayOfWeek.ToString() }
            
            $RefDate = Get-Date "2024-08-28"
            $DiffDays = ($Now - $RefDate).TotalDays
            $CycleDay = $DiffDays % 42
            if ($CycleDay -lt 0) { $CycleDay += 42 }
            
            if ($CycleDay -ge 7.4 -and $CycleDay -le 17.2) {
                Write-Log "偵測到紊亂期 (Day $([math]::Round($CycleDay, 1)))" "Magenta"
                if ($Weekly.Turbulence.$DayOfWeek) {
                    $ConfigName = $Weekly.Turbulence.$DayOfWeek
                    Write-Log "使用紊亂期配置: $ConfigName" "Cyan"
                } else {
                    $ConfigName = $Weekly.$DayOfWeek
                }
            } else {
                if ($Weekly.$DayOfWeek) { $ConfigName = $Weekly.$DayOfWeek }
            }
        } catch {
            Write-Log "讀取 WeeklyConfig 失敗，使用預設值。" "Red"
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ConfigName) -or $ConfigName -eq "Default") {
    Write-Log "未設定配置，任務結束。" "Yellow"
    exit
}
$TaskList = $ConfigName -split ","

# --- [4. 執行任務迴圈] ---
Update-TaskStatus "Running"
$BetterGIPath = "C:\Program Files\BetterGI\BetterGI.exe"
$MaxTaskRetries = 3

for ($i = 0; $i -lt $TaskList.Count; $i++) {
    $CurrentTask = $TaskList[$i]
    
    if ($CurrentTask -eq "[WAIT]") {
        $WaitTarget = $Now.Date.AddHours(10)
        if ($Now.Hour -lt 4) { 
            Write-Log "遇到 [WAIT] 標記，但已過目標時間 (補跑昨日)，跳過。" "Gray"
        } elseif ($Now -lt $WaitTarget) {
            $WaitSec = [math]::Ceiling(($WaitTarget - $Now).TotalSeconds)
            Write-Log "遇到 [WAIT] 標記，暫停腳本直到 10:00 (剩餘 $WaitSec 秒)..." "Cyan"
            Notify "⏳ 暫停執行" "系統進入等待模式，將於 10:00 繼續。" "Blue"
            Start-Sleep $WaitSec
            $Now = Get-Date
        }
        continue 
    }
    
    if ($CurrentTask -eq "PAUSE") {
        Write-Log "遇到 PAUSE 標記，暫停執行。" "Yellow"
        break
    }

    Write-Log "執行配置: $CurrentTask" "Cyan"
    Notify "▶️ 開始執行" "配置: $CurrentTask" "Blue"
    
    $RetryCount = 0
    $TaskSuccess = $false
    
    while ($RetryCount -lt $MaxTaskRetries) {
        $Proc = Start-Process -FilePath $BetterGIPath -ArgumentList "startOneDragon `"$CurrentTask`"" -PassThru
        
        while (-not $Proc.HasExited) {
            $CurrentTime = Get-Date
            
            # A. 死線檢查 (03:55)
            if ($CurrentTime.Hour -eq 3 -and $CurrentTime.Minute -ge 55) {
                Write-Log "⛔ 時間已達 03:55 (絕對死線)，強制終止！" "Red"
                Stop-Process -Id $Proc.Id -Force
                Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
                New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
                shutdown /l
                exit
            }
            
            # B. ForceEnd 觸發檢查 (03:45)
            if ($CurrentTime.Hour -eq 3 -and $CurrentTime.Minute -ge 45 -and $CurrentTask -ne "forceend") {
                Write-Log "⚠️ 執行中遇到 03:45，中斷當前任務，轉為 ForceEnd。" "Orange"
                Notify "🧹 切換模式" "任務超時，切換至 ForceEnd 收尾。" "Orange"
                
                Stop-Process -Id $Proc.Id -Force
                Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
                Start-Sleep 5
                
                Write-Log "啟動 forceend 配置..."
                $Proc = Start-Process -FilePath $BetterGIPath -ArgumentList "startOneDragon `"forceend`"" -PassThru
                $CurrentTask = "forceend" 
                $i = $TaskList.Count 
                continue 
            }
            Start-Sleep 5
        }
        
        if ($Proc.ExitCode -eq 0) {
            $TaskSuccess = $true
            break
        } else {
            $RetryCount++
            Write-Log "任務異常退出 (Code: $($Proc.ExitCode))，重試 $RetryCount/$MaxTaskRetries..." "Red"
            Start-Sleep 5
        }
    }
    
    if (-not $TaskSuccess) {
        Write-Log "任務 $CurrentTask 失敗，已達最大重試次數。" "Red"
        if ($CurrentTask -eq "forceend") { break }
        Update-TaskStatus "Failed"
        New-Item -ItemType File -Path "$FlagDir\Fail.flag" -Force | Out-Null
        exit
    }
    
    if ($i -lt ($TaskList.Count - 1)) {
        Stop-Process -Name "YuanShen", "GenshinImpact" -Force -ErrorAction SilentlyContinue
        Start-Sleep 5
    }
}

# --- [5. 結算] ---
if ($CurrentTask -ne "forceend") {
    Write-Log "所有任務完成。" "Green"
    New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
    $TodayKey | Set-Content -Path $LastRunFile -Encoding UTF8
    Update-TaskStatus "Success"
} else {
    Write-Log "ForceEnd 作業結束。" "Green"
    New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
}

shutdown /l