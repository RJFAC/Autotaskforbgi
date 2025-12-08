<#
    .SYNOPSIS
    AutoTask Payload Script V5.27 (ForceEnd Fix)
    .DESCRIPTION
    遠端 Session 執行腳本。
    負責啟動 BetterGI，監控執行狀態，並在任務完成或時間截止(03:50)時執行智慧登出。
    
    V5.27 Update:
    - [Fix] 將 ForceEnd (03:50) 檢查移入 Start-BetterGI 的等待迴圈內，解決無限等待導致無法觸發的問題。
#>

$ErrorActionPreference = "Stop"

# --- [設定與路徑] ---
$WorkDir = "C:\AutoTask"
$ScriptsDir = "$WorkDir\Scripts"
$FlagsDir = "$WorkDir\Flags"
$LogsDir = "$WorkDir\Logs"
$ConfigsDir = "$WorkDir\Configs"

$LastRunFile = "$LogsDir\LastRun.log"
$ForceRunFlag = "$FlagsDir\ForceRun.flag"
$DoneFlag = "$FlagsDir\Done.flag"

# 確保目錄存在
if (-not (Test-Path $LogsDir)) { New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null }
if (-not (Test-Path $FlagsDir)) { New-Item -ItemType Directory -Path $FlagsDir -Force | Out-Null }

# 定義 Log 檔案
$DateStr = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogsDir\Payload_$DateStr.log"

# --- [輔助函數] ---

function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMsg = "[$Time] $Message"
    Write-Host $LogMsg -ForegroundColor Cyan
    Add-Content -Path $LogFile -Value $LogMsg -Encoding UTF8
}

function Send-Notify {
    param(
        [string]$Title,
        [string]$Msg,
        [string]$Color="Blue"
    )
    $NotifyScript = "$ScriptsDir\Notify.ps1"
    if (Test-Path $NotifyScript) {
        # 呼叫 Notify.ps1，傳遞 Log 路徑以便發生錯誤時附上摘要
        & $NotifyScript -Title $Title -Message $Msg -Color $Color -LogPath $LogFile
    }
}

# 檢查 BetterGI Log 是否出現結束訊號
function Check-BetterGIEnded {
    param([string]$BGILogPath)
    
    if (-not (Test-Path $BGILogPath)) { return $false }

    try {
        # 讀取最後 50 行，避免讀取過慢
        $Content = Get-Content $BGILogPath -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue
        if (-not $Content) { return $false }

        # 關鍵字定義 (根據 BetterGI 的輸出)
        $Keywords = @(
            "全部任務已結束",
            "停止調度",
            "任务执行完毕",
            "全部完成"
        )

        foreach ($Line in $Content) {
            foreach ($Key in $Keywords) {
                if ($Line -match $Key) {
                    Write-Log "偵測到完成訊號！(觸發: '$Key')"
                    return $true
                }
            }
        }
    } catch {
        Write-Log "讀取 BetterGI Log 發生錯誤: $_"
    }
    return $false
}

function Start-BetterGI {
    param(
        [string]$TaskName
    )

    Write-Log ">>> 執行: [$TaskName]"

    # 1. 讀取設定檔取得路徑
    $EnvConfigPath = "$ConfigsDir\EnvConfig.json"
    if (Test-Path $EnvConfigPath) {
        $EnvConfig = Get-Content $EnvConfigPath -Raw | ConvertFrom-Json
        $BGIPath = "$($EnvConfig.BetterGIPath)\BetterGI.exe"
        $BGIDir = $EnvConfig.BetterGIPath
    } else {
        # Fallback 預設路徑 (如果 Config 讀不到)
        $BGIPath = "C:\AutoTask\BetterGI\BetterGI.exe"
        $BGIDir = "C:\AutoTask\BetterGI"
    }

    if (-not (Test-Path $BGIPath)) {
        Write-Log "錯誤: 找不到 BetterGI.exe ($BGIPath)"
        Send-Notify -Title "執行失敗" -Msg "找不到 BetterGI 執行檔。" -Color "Red"
        return
    }

    # 2. 啟動程序
    $Arguments = "start -n `"$TaskName`""
    
    try {
        $Process = Start-Process -FilePath $BGIPath -ArgumentList $Arguments -WorkingDirectory $BGIDir -PassThru -ErrorAction Stop
        Write-Log "BetterGI 已啟動 (PID: $($Process.Id))，開始監控執行狀態..."

        # 取得 BetterGI 今日 Log 路徑 (假設格式為 better-genshin-impactYYYYMMDD.log)
        # 注意: BetterGI Log 檔名格式可能隨版本變動，此處依慣例設定
        $BGILogFile = Join-Path $BGIDir "Logs\better-genshin-impact$DateStr.log"
        Write-Log "鎖定日誌: $BGILogFile"

        # 3. 監控迴圈 (Busy Wait Fix applied here)
        $TimeoutCounter = 0
        $MaxTimeout = 14400 # 4小時超時保護 (4 * 60 * 60 / 5)

        while (-not $Process.HasExited) {
            
            # --- [V5.27 Fix] 03:50 ForceEnd 檢查 ---
            $Now = Get-Date
            if ($Now.Hour -eq 3 -and $Now.Minute -ge 50) {
                Write-Log ⚠️ [ForceEnd] 時間已達 03:50，觸發跨日強制中斷保護！"
                Send-Notify -Title "強制中斷" -Msg "時間已達 03:50，強制結束任務以執行每日重置。" -Color "Orange"
                
                Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
                Write-Log "已強制終止 BetterGI 程序。"
                break
            }
            # -------------------------------------

            # A. 檢查 Log 是否顯示任務完成
            if (Check-BetterGIEnded $BGILogFile) {
                Write-Log "偵測到 Log 顯示任務已完成，準備終止程序..."
                Start-Sleep -Seconds 5
                if (-not $Process.HasExited) {
                    Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
                }
                break
            }

            # B. 超時保護 (WatchDog)
            $TimeoutCounter++
            if ($TimeoutCounter -gt $MaxTimeout) {
                Write-Log "❌ 錯誤: 任務執行超過 4 小時，強制終止。"
                Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
                Send-Notify -Title "任務超時" -Msg "BetterGI 執行超過 4 小時，已強制終止。" -Color "Red"
                break
            }

            Start-Sleep -Seconds 5
        }

        Write-Log "BetterGI 程序已結束。"

    } catch {
        Write-Log "啟動或監控 BetterGI 時發生例外: $_"
        Send-Notify -Title "執行錯誤" -Msg "Payload 發生例外: $_" -Color "Red"
    }
}

# --- [主要邏輯] ---

Write-Log "Payload 啟動 (V5.27) PID: $PID..."
Write-Log "日期: $DateStr"

# 1. 檢查是否為今日首次執行 (且無 ForceRun)
$SkipCheck = $false
if (Test-Path $ForceRunFlag) {
    Write-Log "偵測到 ForceRun.flag，跳過日期檢查，強制執行任務。"
    $SkipCheck = $true
    Remove-Item $ForceRunFlag -Force -ErrorAction SilentlyContinue
}

if (-not $SkipCheck -and (Test-Path $LastRunFile)) {
    $LastDate = Get-Content $LastRunFile -Raw
    if ($LastDate.Trim() -eq $DateStr) {
        Write-Log "今日 ($DateStr) 任務已執行過。目前為使用者手動登入模式。"
        Write-Log "Payload 將退出但不登出。"
        exit
    }
}

# 2. 決定今日任務模板
$TaskName = "Default" # 預設
$DateMapFile = "$ConfigsDir\DateConfig.map"

if (Test-Path $DateMapFile) {
    # 讀取對應表: 格式 20251209=TaskName
    $MapContent = Get-Content $DateMapFile
    foreach ($Line in $MapContent) {
        if ($Line -match "^$DateStr=(.+)") {
            $TaskName = $Matches[1]
            break
        }
    }
}

# 3. 執行任務
Send-Notify -Title "任務開始" -Msg "Payload 已啟動，執行模板: [$TaskName]" -Color "Blue"

Start-BetterGI -TaskName $TaskName

# 4. 任務結束後處理
Write-Log "任務流程結束。"

# 記錄 LastRun
Set-Content -Path $LastRunFile -Value $DateStr -Encoding UTF8

# 建立 Done.flag 通知 Master
New-Item -ItemType File -Path $DoneFlag -Force | Out-Null
Write-Log "已建立 Done.flag。"

Send-Notify -Title "任務完成" -Msg "Payload 執行完畢，系統將在 10 秒後登出。" -Color "Green"

# 5. 智慧登出 (Smart Logoff)
Write-Log "準備登出..."
Start-Sleep -Seconds 10
logoff