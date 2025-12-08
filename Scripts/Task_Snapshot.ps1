<#
    .SYNOPSIS
    AutoTask Snapshot Generator V2.3 (Full Source Edition)
    .DESCRIPTION
    蒐集系統關鍵檔案、腳本與日誌，打包成 Zip 供 AI 分析或備份。
    
    V2.3 Update (2025-12-09):
    - [Source] 完整捕獲: 強制擷取 Scripts 下所有 .ps1 與根目錄下所有 .bat，禁止截斷。
    - [Structure] 資料夾優化: 將原始碼與日誌分開存放，方便 AI 讀取。
    - [Log] 維持 V2.1/2.2 的智能截斷 (僅針對 .log 檔案)。
#>

$ErrorActionPreference = "SilentlyContinue"

# --- [路徑設定] ---
$WorkDir = "C:\AutoTask"
$ScriptsDir = "$WorkDir\Scripts"
$LogsDir = "$WorkDir\Logs"
$ConfigsDir = "$WorkDir\Configs"
$RemoteLogsDir = "$WorkDir\1Remote\.logs"
$BetterGILogsDir = "C:\Program Files\BetterGI\log"

# 產生時間戳記
$DateStr = Get-Date -Format "yyyyMMdd_HHmmss"
$SnapshotName = "AutoTask_Snapshot_$DateStr"
$TempDir = "$WorkDir\Snapshot_Temp_$DateStr"
$ZipPath = "$WorkDir\$SnapshotName.zip"

# 建立暫存目錄結構
New-Item -ItemType Directory -Path "$TempDir\0_Root_Bats" -Force | Out-Null
New-Item -ItemType Directory -Path "$TempDir\1_Full_Scripts" -Force | Out-Null
New-Item -ItemType Directory -Path "$TempDir\2_Configs" -Force | Out-Null
New-Item -ItemType Directory -Path "$TempDir\3_Logs_Summary" -Force | Out-Null

Write-Host "=== AutoTask Snapshot Generator V2.3 ===" -ForegroundColor Cyan
Write-Host "正在建立快照: $SnapshotName" -ForegroundColor Gray

# ==============================================================================
# 1. 核心腳本與執行檔 (完整複製，禁止截斷)
# ==============================================================================
Write-Host " -> 1/4 正在備份原始碼 (完整模式)..."

# 複製所有 Scripts 目錄下的 .ps1 (完整)
Get-ChildItem -Path $ScriptsDir -Filter "*.ps1" | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination "$TempDir\1_Full_Scripts" -Force
}

# 複製根目錄下的 .bat (完整)
Get-ChildItem -Path $WorkDir -Filter "*.bat" | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination "$TempDir\0_Root_Bats" -Force
}

# 為了方便 AI 閱讀，同時生成一個合併的文本文件 (Full_Source_Text.txt)
# 這樣 AI 可以一次讀取所有代碼，不用切換多個檔案
$SourceOutFile = "$TempDir\1_Full_Scripts\All_Source_Combined.txt"
Add-Content -Path $SourceOutFile -Value "=== [ AutoTask Full Source Code Dump ] ===`r`n" -Encoding UTF8

# 寫入 .bat 內容
Get-ChildItem -Path $WorkDir -Filter "*.bat" | ForEach-Object {
    $Header = "`r`n======================================================================`r`n" +
              "FILE: $($_.Name) | PATH: $($_.FullName)`r`n" +
              "======================================================================`r`n"
    Add-Content -Path $SourceOutFile -Value $Header -Encoding UTF8
    $Content = Get-Content -Path $_.FullName -Raw -Encoding UTF8
    Add-Content -Path $SourceOutFile -Value $Content -Encoding UTF8
}

# 寫入 .ps1 內容
Get-ChildItem -Path $ScriptsDir -Filter "*.ps1" | ForEach-Object {
    $Header = "`r`n======================================================================`r`n" +
              "FILE: $($_.Name) | PATH: $($_.FullName)`r`n" +
              "======================================================================`r`n"
    Add-Content -Path $SourceOutFile -Value $Header -Encoding UTF8
    # 使用 -Raw 確保讀取完整內容，不進行任何處理
    $Content = Get-Content -Path $_.FullName -Raw -Encoding UTF8
    Add-Content -Path $SourceOutFile -Value $Content -Encoding UTF8
}

# ==============================================================================
# 2. 設定檔 (完整複製)
# ==============================================================================
Write-Host " -> 2/4 正在備份設定檔..."
Copy-Item "$ConfigsDir\*.json" -Destination "$TempDir\2_Configs" -Force
Copy-Item "$ConfigsDir\*.map" -Destination "$TempDir\2_Configs" -Force
Copy-Item "$ConfigsDir\*.log" -Destination "$TempDir\2_Configs" -Force # 包含週曆紀錄等
Copy-Item "$WorkDir\0_File_Structure.txt" -Destination "$TempDir" -Force # 如果有結構檔的話

# ==============================================================================
# 3. 日誌檔案 (智能截斷: Head 500 + Tail 2000)
# ==============================================================================
Write-Host " -> 3/4 正在處理日誌 (智能截斷模式)..."

# 定義智能讀取函數 (僅用於 Log)
function Get-SmartLogContent {
    param ( [string]$FilePath )
    try {
        $TotalLines = (Get-Content $FilePath).Count
        if ($TotalLines -le 2500) {
            return (Get-Content $FilePath -Raw -Encoding UTF8)
        } else {
            $Head = Get-Content $FilePath -Head 500 -Encoding UTF8
            $Tail = Get-Content $FilePath -Tail 2000 -Encoding UTF8
            return ($Head + "`r`n... (省略中間 $($TotalLines - 2500) 行) ...`r`n" + $Tail) | Out-String
        }
    } catch {
        return "Error reading file: $($_.Exception.Message)"
    }
}

# 定義日誌獲取函數 (24小時過濾 + 回退機制)
function Get-RecentLogs {
    param ($Path)
    if (-not (Test-Path $Path)) { return @() }
    
    # 嘗試獲取 24 小時內的日誌
    $Logs = Get-ChildItem -Path "$Path\*.log" | Where-Object { $_.LastWriteTime -ge (Get-Date).AddHours(-24) } | Sort-Object LastWriteTime -Descending
    
    # 如果 24 小時內沒有日誌，則獲取最新的一份 (回退機制)
    if ($Logs.Count -eq 0) {
        $Logs = Get-ChildItem -Path "$Path\*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    return $Logs
}

function Get-RecentRemoteLogs {
    param ($Path)
    if (-not (Test-Path $Path)) { return @() }
    # 1Remote 的日誌通常是 markdown 格式
    $Logs = Get-ChildItem -Path "$Path\*.md" | Where-Object { $_.LastWriteTime -ge (Get-Date).AddHours(-24) } | Sort-Object LastWriteTime -Descending
    if ($Logs.Count -eq 0) {
        $Logs = Get-ChildItem -Path "$Path\*.md" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    return $Logs
}

$LogsOutFile = "$TempDir\3_Logs_Summary\Recent_Logs_Summary.txt"
Add-Content -Path $LogsOutFile -Value "=== [ Recent Logs Summary (Last 24h + Fallback) ] ===`r`n" -Encoding UTF8

$LogTargets = @()
$LogTargets += Get-RecentLogs -Path $LogsDir
$LogTargets += Get-RecentRemoteLogs -Path $RemoteLogsDir
$LogTargets += Get-RecentLogs -Path $BetterGILogsDir

foreach ($File in $LogTargets) {
    if ($null -eq $File) { continue }
    
    $Header = "`r`n======================================================================`r`n" +
              "FILE: $($File.Name)`r`n" +
              "PATH: $($File.FullName)`r`n" +
              "TIME: $($File.LastWriteTime)`r`n" +
              "SIZE: $([math]::Round($File.Length / 1MB, 2)) MB`r`n" +
              "======================================================================`r`n"
    Add-Content -Path $LogsOutFile -Value $Header -Encoding UTF8

    # 使用智能截斷讀取日誌內容
    $Content = Get-SmartLogContent -FilePath $File.FullName
    $Content | Out-File -FilePath $LogsOutFile -Append -Encoding UTF8
}

# ==============================================================================
# 4. 壓縮與清理
# ==============================================================================
Write-Host " -> 4/4 正在壓縮打包..."
Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force

# 清理暫存
Remove-Item -Path $TempDir -Recurse -Force

Write-Host "`r`n[完成] 快照已儲存至: $ZipPath" -ForegroundColor Green
Write-Host "請將此 Zip 檔案上傳給 AI 進行分析。" -ForegroundColor Yellow

# 如果是在 Dashboard 中執行，暫停一下讓使用者看到結果
if ($Host.Name -eq "ConsoleHost") {
    Start-Sleep -Seconds 3
}