# --- [路徑定義] ---
$BaseDir    = "C:\AutoTask"
$FlagDir    = "$BaseDir\Flags"
$LogDir     = "$BaseDir\Logs"
$RunFlag    = "$FlagDir\Run.flag"

$1RemoteDir = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64"
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
    Add-Content -Path $LogFile -Value $FormattedMsg -Encoding UTF8
}

Get-ChildItem -Path $LogDir -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force -ErrorAction SilentlyContinue

Write-Log "Monitor 啟動..." "Cyan"

$CurrentLogFile = $null
$StreamReader = $null

while ($true) {
    if (-not (Test-Path $RunFlag)) {
        Write-Log "Run.flag 已消失，Monitor 停止。" "Gray"
        break
    }

    if (-not (Get-Process "1Remote" -ErrorAction SilentlyContinue)) {
        Write-Log "⚠️ 警報：偵測到 1Remote 進程消失 (崩潰)！立即執行重啟..." "Red"
        Start-Process -FilePath $1RemoteExe -WorkingDirectory $1RemoteDir
        Start-Sleep 5
        Start-Process -FilePath $1RemoteExe -ArgumentList $1RemoteArgs -WorkingDirectory $1RemoteDir
        Write-Log "✅ 1Remote 已發送重啟指令。" "Green"
        if ($StreamReader) { $StreamReader.Close(); $StreamReader = $null }
        Start-Sleep 10
        continue
    }

    $LatestLog = Get-ChildItem "$1RemoteDir\.logs\1Remote.log_*.md" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($LatestLog -and ($null -eq $CurrentLogFile -or $LatestLog.FullName -ne $CurrentLogFile.FullName)) {
        Write-Log "鎖定 1Remote 日誌: $($LatestLog.Name)" "Cyan"
        if ($StreamReader) { $StreamReader.Close() }
        $CurrentLogFile = $LatestLog
        $FileStream = [System.IO.File]::Open($CurrentLogFile.FullName, 'Open', 'Read', 'ReadWrite')
        $StreamReader = New-Object System.IO.StreamReader($FileStream)
        $StreamReader.BaseStream.Seek(0, [System.IO.SeekOrigin]::End)
    }

    if ($StreamReader) {
        while (-not $StreamReader.EndOfStream) {
            $line = $StreamReader.ReadLine()
            if ($line -match "error code") {
                Write-Log "⚠️ 偵測到斷線日誌: $line" "Yellow"
                Write-Log "執行重連..." 
                Start-Process -FilePath $1RemoteExe -ArgumentList $1RemoteArgs -WorkingDirectory $1RemoteDir
                Start-Sleep 5
            }
        }
    }
    Start-Sleep 2
}
if ($StreamReader) { $StreamReader.Close() }
