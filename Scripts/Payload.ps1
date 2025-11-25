# ... (前段代碼與 V4.8 相同，請複製 V4.8 的前段，或直接使用下方完整代碼) ...

# --- [完整代碼請見下文，包含 API 下載邏輯] ---
# 為了確保功能完整，我將提供包含 BITS 下載功能的完整 Payload

# --- [路徑定義] ---
$BaseDir    = "C:\AutoTask"
$ScriptDir  = "$BaseDir\Scripts"
$ConfigDir  = "$BaseDir\Configs"
$FlagDir    = "$BaseDir\Flags"
$LogDir     = "$BaseDir\Logs"

if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Payload_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8
}

Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue

try {
    $CurrentPID = $PID
    $TargetScript = "Payload.ps1"
    $OldInstances = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$TargetScript*" -and $_.ProcessId -ne $CurrentPID }
    foreach ($proc in $OldInstances) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }
} catch {}

$BettergiDir    = "C:\Program Files\BetterGI"
$BettergiExe    = "$BettergiDir\BetterGI.exe"
$LogDirBG       = "$BettergiDir\log"
$ScreenshotDir  = "$BettergiDir\Screenshots" 
$1RemoteLogDir  = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64\.logs"

$TaskStatusFile = "$ConfigDir\TaskStatus.json"
$DoneFlag       = "$FlagDir\Done.flag"
$FailFlag       = "$FlagDir\Fail.flag"
$ForceRunFlag   = "$FlagDir\ForceRun.flag" 
$DateMap        = "$ConfigDir\DateConfig.map"
$WeeklyConf     = "$ConfigDir\WeeklyConfig.json"
$PauseLog       = "$ConfigDir\PauseDates.log"
$LastRunLog     = "$ConfigDir\LastRun.log"
$BackupRootDir  = "$BaseDir\LogBackups"
$NotifyScript   = "$ScriptDir\Notify.ps1"

$MaxRetries = 3
$SuccessKeyword = "一条龙.*任务结束"

Write-Log "Payload 啟動..." "Cyan"

function Send-Notify {
    param([string]$Title, [string]$Msg, [string]$Color)
    if (Test-Path $NotifyScript) {
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$Title`" -Message `"$Msg`" -Color `"$Color`"" -WindowStyle Hidden
    }
}

# --- [新功能] 預下載檢查 ---
function Check-GenshinPreDownload {
    # 1. 讀取設定檔中的遊戲路徑
    $GamePath = ""
    if (Test-Path $WeeklyConf) {
        try {
            $json = Get-Content $WeeklyConf -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($json.GenshinPath) { $GamePath = $json.GenshinPath }
        } catch {}
    }
    
    if (-not (Test-Path $GamePath)) {
        Write-Log "跳過預下載檢查：未設定有效的原神遊戲路徑。" "Gray"
        return
    }

    # 2. 檢查是否為預下載期間 (更新日前 2 天)
    $RefDate = [datetime]"2024-08-28"
    $DiffDays = ((Get-Date).Date - $RefDate).Days % 42
    # 42-2=40 (週一), 41 (週二)
    if ($DiffDays -ne 40 -and $DiffDays -ne 41) { return }

    Write-Log "偵測到預下載期間，正在查詢官方 API..." "Cyan"
    
    # 3. 查詢 API
    try {
        $ApiUrl = "https://sdk-os-static.mihoyo.com/hk4e_global/mdk/launcher/api/resource?key=gcStgarh&launcher_id=10"
        $Response = Invoke-RestMethod -Uri $ApiUrl -Method Get
        $GameData = $Response.data.game.latest
        
        if ($GameData.pre_download_game) {
            $PreVer = $GameData.pre_download_game.version
            $PreUrl = $GameData.pre_download_game.path
            $PreName = $GameData.pre_download_game.name
            $PreSize = $GameData.pre_download_game.size
            $PreHash = $GameData.pre_download_game.md5
            
            $DestFile = Join-Path $GamePath $PreName
            
            # 檢查是否已下載
            if (Test-Path $DestFile) {
                $LocalSize = (Get-Item $DestFile).Length
                if ($LocalSize -eq $PreSize) {
                    Write-Log "預下載檔案已存在且大小相符 ($PreName)，跳過。" "Green"
                    return
                } else {
                    Write-Log "發現未完成的下載檔，嘗試續傳..." "Yellow"
                }
            }

            Write-Log "開始下載版本 $PreVer 更新檔..." "Cyan"
            Send-Notify -Title "開始預下載" -Msg "正在下載原神 $PreVer 版本更新檔..." -Color "Blue"
            
            # 使用 BITS 下載 (支援後台、續傳)
            try {
                Import-Module BitsTransfer
                Start-BitsTransfer -Source $PreUrl -Destination $DestFile -DisplayName "Genshin Pre-Download" -Priority Normal
                Write-Log "預下載完成！" "Green"
                Send-Notify -Title "預下載完成" -Msg "版本 $PreVer 更新檔已下載完畢。" -Color "Green"
            } catch {
                Write-Log "下載失敗: $_" "Red"
                Send-Notify -Title "預下載失敗" -Msg "BITS 傳輸發生錯誤。" -Color "Red"
            }
        } else {
            Write-Log "官方尚未釋出預下載資源。" "Gray"
        }
    } catch {
        Write-Log "API 查詢失敗: $_" "Red"
    }
}

# ... (Cleanup-Screenshots, Backup-Logs, Update-Status, Get-TargetConfig, Check-Success-Log, Test-GenshinUpdateDay 函數與 V4.8 相同，省略以節省篇幅，請務必保留) ...
# [請在此處插入 V4.8 的所有輔助函數]
# ...

# =============================================================================
# [主流程]
# =============================================================================
# (ForceEnd, 等待04:00, 排程檢查 邏輯與 V4.8 相同，省略)
# ...

# 在執行今日任務前，插入預下載檢查
Check-GenshinPreDownload

# ... (執行今日任務 While 迴圈與 V4.8 相同) ...
