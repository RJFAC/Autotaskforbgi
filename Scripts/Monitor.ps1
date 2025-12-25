# =============================================================================
# AutoTask Monitor V3.9 - WMI 容錯增強版
# =============================================================================
# V3.9:
#   1. [Fix] 修正 Get-CimInstance 在系統休眠/關機時報錯導致誤判 Master 已死的問題。
#      加入 Try-Catch 保護，當 WMI 查詢失敗時，假設 Master 仍存活。
# =============================================================================

# --- [定義路徑] ---
$BaseDir    = "C:\AutoTask"
$ScriptDir  = "$BaseDir\Scripts"
$FlagDir    = "$BaseDir\Flags"
$LogDir     = "$BaseDir\Logs"
$RunFlag    = "$FlagDir\Run.flag"
$DoneFlag   = "$FlagDir\Done.flag"
$MasterScript = "$ScriptDir\Master.ps1"

# 嘗試讀取 EnvConfig
$1RemoteDir = "C:\Users\[USER_NAME]\Downloads\1Remote-1.2.1-net9-x64"
if (Test-Path "$BaseDir\Configs\EnvConfig.json") {
    try {
        $env = Get-Content "$BaseDir\Configs\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env.Path1Remote) { $1RemoteDir = Split-Path $env.Path1Remote -Parent }
    } catch {}
}
$1RemoteExe = "$1RemoteDir\1Remote.exe"
$1RemoteArgs= "-r Remote"

# --- [全域日誌設定] ---
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color="White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFileName = "Monitor_$(Get-Date -Format 'yyyyMMdd').log"
    $LogFile = Join-Path $LogDir $LogFileName
    $FormattedMsg = "[$TimeStamp] $Message"
    Write-Host $FormattedMsg -ForegroundColor $Color
    try { Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# 清理舊日誌
try { Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}

Write-Log "Monitor 啟動 (V3.9 - WMI 容錯)..." "Cyan"

# --- [狀態變數] ---
$CurrentLogFile = $null
$LastSize = 0
$MyPID = $PID

# 記錄啟動時的檔案時間，用於偵測更新
$SelfPath = $PSCommandPath
$InitialWriteTime = (Get-Item $SelfPath).LastWriteTime
$LastNetCheckTime = Get-Date

while ($true) {
    # 1. 檢查自身存活條件 (新增防抖動容錯機制)
    if (-not (Test-Path $RunFlag)) {
        # 第一次偵測到消失，等待緩衝
        Start-Sleep 2
        if (-not (Test-Path $RunFlag)) {
            # 第二次確認
            Start-Sleep 2
            if (-not (Test-Path $RunFlag)) {
                # 第三次確認消失，才判定為真正結束
                Write-Log "Run.flag 確認消失，Monitor 正常停止。" "Gray"
                break
            } else {
                Write-Log "Run.flag 短暫消失後恢復 (忽略抖動)。" "Gray"
            }
        }
    }

    # 2. 檢查自我更新
    try {
        if ((Get-Item $SelfPath).LastWriteTime -ne $InitialWriteTime) {
            Write-Log "♻️ 偵測到 Monitor 腳本更新，正在重啟以應用變更..." "Magenta"
            exit 
        }
    } catch {}

    # 3. 監督 Master 是否活著 [V3.9: 加入 WMI 容錯]
    $MasterAlive = $true # 預設為真，避免 WMI 錯誤導致誤殺
    try {
        $ProcList = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop
        $MasterProc = $ProcList | Where-Object { $_.CommandLine -like "*Master.ps1*" }
        if (-not $MasterProc) { $MasterAlive = $false }
    } catch {
        # 若 WMI 失敗 (如休眠中)，假定 Master 還活著
        # Write-Log "WMI 查詢異常，跳過 Master 檢查。" "Gray"
    }
    
    if (-not $MasterAlive) {
        if (-not (Test-Path $DoneFlag)) {
            Write-Log "⚠️ 警報：Master 意外消失且任務未完成！正在重啟 Master..." "Red"
            Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`"" -Verb RunAs
            Start-Sleep 5 
        } else {
            Write-Log "Master 已結束且任務完成，Monitor 跟隨退出。" "Green"
            break
        }
    }

    # 4. 鎖定或更新日誌檔
    $DateStr = Get-Date -Format "yyyyMMdd"
    $LatestLog = Get-ChildItem "$1RemoteDir\.logs\1Remote.log_$DateStr*.md" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($LatestLog) {
        if ($null -eq $CurrentLogFile -or $LatestLog.FullName -ne $CurrentLogFile.FullName) {
            Write-Log "鎖定新日誌: $($LatestLog.Name)" "Cyan"
            $CurrentLogFile = $LatestLog
            $LastSize = $CurrentLogFile.Length 
        }

        # 5. 讀取新增內容
        try {
            $CurrentLogFile.Refresh()
            $CurrentSize = $CurrentLogFile.Length
            
            if ($CurrentSize -gt $LastSize) {
                $Stream = [System.IO.File]::Open($CurrentLogFile.FullName, 'Open', 'Read', 'ReadWrite')
                $Reader = New-Object System.IO.StreamReader($Stream)
                $null = $Reader.BaseStream.Seek($LastSize, [System.IO.SeekOrigin]::Begin)
                $NewContent = $Reader.ReadToEnd()
                $Reader.Close(); $Stream.Close()
                
                $LastSize = $CurrentSize
                
                if ($NewContent -match "exit with error code") {
                    Write-Log "⚠️ 偵測到 RDP 斷線訊號 (Exit code)！" "Red"
                    
                    $RdpProc = Get-Process "1Remote" -ErrorAction SilentlyContinue
                    if ($RdpProc) {
                        Write-Log "診斷: 1Remote 進程仍在執行 (PID: $($RdpProc.Id))，但連線已斷開。" "Yellow"
                    } else {
                        Write-Log "診斷: 1Remote 進程已完全消失。" "Red"
                    }

                    Write-Log "正在重啟 1Remote..." "Yellow"
                    Stop-Process -Name "1Remote" -Force -ErrorAction SilentlyContinue
                    Start-Sleep 2
                    Start-Process -FilePath $1RemoteExe -WorkingDirectory $1RemoteDir
                    Start-Sleep 5
                    Start-Process -FilePath $1RemoteExe -ArgumentList $1RemoteArgs -WorkingDirectory $1RemoteDir
                    Write-Log "重連指令已發送。" "Green"
                    Start-Sleep 10
                }
            }
        } catch {
            Write-Log "讀取日誌錯誤: $_" "Red"
        }
    }
    
    # 6. TCP 連線診斷
    if (((Get-Date) - $LastNetCheckTime).TotalSeconds -ge 30) {
        $LastNetCheckTime = Get-Date 
        
        $RdpProc = Get-Process "1Remote" -ErrorAction SilentlyContinue
        if ($RdpProc) {
            try {
                $NetStat = Get-NetTCPConnection -OwningProcess $RdpProc.Id -State Established -ErrorAction SilentlyContinue
                if (-not $NetStat) {
                    Write-Log "⚠️ 診斷警報: 1Remote 進程存在，但無 ESTABLISHED 連線 (可能已斷線或假死)。" "Yellow"
                }
            } catch {}
        }
    }

    Start-Sleep -Seconds 2
}
