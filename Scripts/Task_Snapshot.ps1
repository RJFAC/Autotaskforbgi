<#
    .SYNOPSIS
    AutoTask Snapshot Generator V2.2 (24h Filter Edition)
    .DESCRIPTION
    蒐集系統關鍵檔案、腳本與日誌，打包成 Zip 供 AI 分析或備份。
    V2.2 Update:
    - [Log] 強制過濾: 預設僅擷取「過去 24 小時內」的日誌檔。
    - [Log] 智能回退: 若 24 小時內無日誌，則擷取最近的 1 份以供參考。
    - [Log] 維持 V2.1 的智能截斷 (Head 500 + Tail 2000)。
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

# 建立暫存目錄
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

Write-Host ">>> 正在建立快照: $SnapshotName" -ForegroundColor Cyan

# --- [輔助函數] 智能讀取日誌 ---
function Get-SmartLogContent {
    param([string]$FilePath)
    
    try {
        $FileItem = Get-Item $FilePath
        # 門檻: 2MB
        if ($FileItem.Length -lt 2MB) {
            return Get-Content -Path $FilePath -Raw -Encoding UTF8 -ErrorAction Stop
        }
        else {
            # 大檔案策略: Head 500 + Tail 2000
            $Head = Get-Content -Path $FilePath -Head 500 -Encoding UTF8
            $Tail = Get-Content -Path $FilePath -Tail 2000 -Encoding UTF8
            
            $Separator = "`r`n`r`n... [Log Truncated: Skipped Middle Content] ...`r`n... [Showing First 500 lines & Last 2000 lines] ...`r`n`r`n"
            
            return ($Head -join "`r`n") + $Separator + ($Tail -join "`r`n")
        }
    }
    catch {
        return "[Error Reading File: $_]"
    }
}

# --- [輔助函數] 獲取近期日誌 (24h 邏輯) ---
function Get-RecentLogs {
    param([string]$Path, [int]$MaxCount=5)
    if (-not (Test-Path $Path)) { return @() }
    
    $Now = Get-Date
    $OneDayAgo = $Now.AddHours(-24)

    # 1. 先嘗試抓取 24 小時內的檔案
    $Logs = Get-ChildItem -Path $Path -Filter "*.log" -File | 
            Where-Object { $_.LastWriteTime -ge $OneDayAgo } | 
            Sort-Object LastWriteTime -Descending

    # 2. 如果 24 小時內沒有檔案 (例如剛過午夜)，則抓取最近的 1 個檔案作為參考
    if ($Logs.Count -eq 0) {
        $Logs = Get-ChildItem -Path $Path -Filter "*.log" -File | 
                Sort-Object LastWriteTime -Descending | 
                Select-Object -First 1
    } else {
        # 如果有檔案，則限制最大數量
        $Logs = $Logs | Select-Object -First $MaxCount
    }
    
    return $Logs
}

# 針對 1Remote 的特殊處理 (副檔名 .md)
function Get-RecentRemoteLogs {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    
    $OneDayAgo = (Get-Date).AddHours(-24)
    
    $Logs = Get-ChildItem -Path $Path -Filter "*.md" -File | 
            Where-Object { $_.LastWriteTime -ge $OneDayAgo } | 
            Sort-Object LastWriteTime -Descending
            
    if ($Logs.Count -eq 0) {
        $Logs = Get-ChildItem -Path $Path -Filter "*.md" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    return $Logs | Select-Object -First 3
}

# ==============================================================================
# 1. 輸出檔案結構
# ==============================================================================
Write-Host " -> 掃描檔案結構..."
$StructureFile = "$TempDir\0_File_Structure.txt"
$StructureContent = Get-ChildItem -Path $WorkDir -Recurse | Select-Object FullName, Length, LastWriteTime | Out-String -Width 120
Set-Content -Path $StructureFile -Value $StructureContent -Encoding UTF8

# ==============================================================================
# 2. 輸出完整腳本內容 (1_Full_Scripts.txt)
# ==============================================================================
Write-Host " -> 彙整腳本原始碼 (強制完整讀取文件)..."
$ScriptsOutFile = "$TempDir\1_Full_Scripts.txt"
Add-Content -Path $ScriptsOutFile -Value "=== [ AutoTask Scripts & Docs Snapshot ] ===`r`n" -Encoding UTF8

$ScriptFiles = Get-ChildItem -Path $ScriptsDir, $WorkDir -Include *.ps1, *.bat, *.md, *.xml -File -Recurse | 
               Where-Object { $_.FullName -notmatch "Snapshot_Temp" -and $_.FullName -notmatch ".zip" }

foreach ($File in $ScriptFiles) {
    $Header = "`r`n======================================================================`r`n" +
              "FILE: $($File.Name) | PATH: $($File.FullName)`r`n" +
              "======================================================================`r`n"
    Add-Content -Path $ScriptsOutFile -Value $Header -Encoding UTF8

    if ($File.Name -eq "AutoTask_Core_Analysis.md" -or $File.Length -lt 2MB) {
        try {
            $Content = Get-Content -Path $File.FullName -Raw -Encoding UTF8 -ErrorAction Stop
            Add-Content -Path $ScriptsOutFile -Value $Content -Encoding UTF8
        } catch {
            Add-Content -Path $ScriptsOutFile -Value "[Error Reading File: $_]" -Encoding UTF8
        }
    } else {
        Add-Content -Path $ScriptsOutFile -Value "[File too large (>2MB), truncated.]" -Encoding UTF8
    }
}

# ==============================================================================
# 3. 輸出設定檔 (2_Configs.txt)
# ==============================================================================
Write-Host " -> 彙整設定檔..."
$ConfigsOutFile = "$TempDir\2_Configs.txt"
Add-Content -Path $ConfigsOutFile -Value "=== [ AutoTask Configs ] ===`r`n" -Encoding UTF8

$ConfigFiles = Get-ChildItem -Path $ConfigsDir -File
foreach ($File in $ConfigFiles) {
    $Header = "`r`n======================================================================`r`n" +
              "FILE: $($File.Name)`r`n" 
    Add-Content -Path $ConfigsOutFile -Value $Header -Encoding UTF8
    
    try {
        $Content = Get-Content -Path $File.FullName -Raw -Encoding UTF8
        Add-Content -Path $ConfigsOutFile -Value $Content -Encoding UTF8
    } catch {
        Add-Content -Path $ConfigsOutFile -Value "[Error reading config]"
    }
}

# ==============================================================================
# 4. 輸出近期日誌 (3_Recent_Logs.txt) - 智能截斷 + 24H 過濾
# ==============================================================================
Write-Host " -> 彙整近期日誌 (24h 限制 + 智能 Head/Tail)..."
$LogsOutFile = "$TempDir\3_Recent_Logs.txt"
Add-Content -Path $LogsOutFile -Value "=== [ Recent Logs Summary (Last 24h) ] ===`r`n" -Encoding UTF8

# 收集目標 (使用新的過濾函數)
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

    $Content = Get-SmartLogContent -FilePath $File.FullName
    $Content | Out-File -FilePath $LogsOutFile -Append -Encoding UTF8
}

# ==============================================================================
# 5. 壓縮與清理
# ==============================================================================
Write-Host " -> 正在壓縮..."
Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force
Remove-Item -Path $TempDir -Recurse -Force

Write-Host ">>> 快照建立完成: $ZipPath" -ForegroundColor Green
Write-Host "提示: 您現在可以在 Dashboard 上傳此 Zip 檔案供 AI 分析。" -ForegroundColor Yellow

Invoke-Item $WorkDir