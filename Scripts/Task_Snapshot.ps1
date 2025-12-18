<#
.SYNOPSIS
    AutoTask Snapshot Tool V2.14 (Self-Backup Enabled & Safe Strings)
    用於打包所有腳本、設定檔與最近日誌，供 AI 進行除錯分析。
    
    V2.14 更新:
    1. [Fix] 修正 EOF 標記寫入方式，使用變數拼接，防止編輯器解析錯誤或截斷。
    2. [Mod] 修改過濾邏輯，現在會將 Task_Snapshot.ps1 自身包含在快照中。
#>

$SnapshotVersion = "V2.14"
$SourceDir = "C:\AutoTask"
$OutputDir = "C:\AutoTask_Snapshots"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ZipName = "AutoTask_Snapshot_$Timestamp.zip"
$ZipPath = Join-Path $OutputDir $ZipName

# --- [防閃退機制] ---
trap {
    Write-Host "`n[CRITICAL ERROR] 發生嚴重錯誤：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "位置: $($_.InvocationInfo.ScriptLineNumber) 行" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
}

# 確保輸出目錄存在
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# 建立暫存目錄結構
$TempDir = Join-Path $OutputDir "Temp_$Timestamp"
$DirSrc  = Join-Path $TempDir "1_Source_Code"
$DirConf = Join-Path $TempDir "2_Configs"
$DirLogs = Join-Path $TempDir "3_Logs_Summary"

New-Item -ItemType Directory -Path $DirSrc -Force | Out-Null
New-Item -ItemType Directory -Path $DirConf -Force | Out-Null
New-Item -ItemType Directory -Path $DirLogs -Force | Out-Null

Write-Host "=== AutoTask 快照工具 $SnapshotVersion ===" -ForegroundColor Cyan
Write-Host "來源: $SourceDir"
Write-Host "目標: $ZipPath"

# ==============================================================================
# 定義輸出檔案 (分流策略)
# ==============================================================================
$File_Logic  = "$DirSrc\Source_Logic_Only.txt"   # 只放代碼 (.ps1)
$File_Config = "$DirConf\Configs_Only.txt"       # 只放設定 (.json, .map, etc)
$File_Docs   = "$DirSrc\Docs_Reference.txt"      # 只放文件 (.md, .txt)

# 標頭寫入函數
function Add-FileHeader {
    param($OutFile, $FileName)
    $Line = "`n--- FILE: $FileName ---`n" 
    Add-Content -Path $OutFile -Value $Line -Encoding UTF8
}

# 安全讀取內容 (處理空檔案)
function Get-SafeContent {
    param($Path)
    $Content = Get-Content -Path $Path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrEmpty($Content)) { return "[EMPTY FILE]" }
    return $Content
}

# 智能日誌讀取 (防止鎖定與過大)
function Get-SmartLogContent {
    param($FilePath)
    $LimitSize = 2MB 
    $Item = Get-Item $FilePath
    if ($Item.Length -gt $LimitSize) {
        return "[SYSTEM] 日誌過大 ($([math]::Round($Item.Length/1MB,2)) MB)，僅截取最後 2000 行...`r`n" + 
               (Get-Content $FilePath -Tail 2000 -Encoding UTF8 | Out-String)
    } else {
        # 使用 FileShare.ReadWrite 防止鎖定
        $Stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $Reader = New-Object System.IO.StreamReader($Stream, [System.Text.Encoding]::UTF8)
        $Text = $Reader.ReadToEnd()
        $Reader.Close(); $Stream.Close()
        return $Text
    }
}

# [重要修正] 定義安全的 EOF 標記
# 使用拼接方式建立字串，避免 AI 介面誤判為 Markdown 結束
$EOF_Marker = "`n" + "``" + "`" + "eof"

# ==============================================================================
# 1. 處理邏輯代碼 (.ps1)
# ==============================================================================
Write-Host " -> 1/4 提取核心代碼 (Logic)..."
Add-Content -Path $File_Logic -Value "=== AutoTask Logic Source ($SnapshotVersion) ===`n" -Encoding UTF8

# 取得根目錄與 Scripts 下的 .ps1 (移除對自身的排除，僅排除 Temp 目錄)
$LogicFiles = Get-ChildItem -Path $SourceDir -Recurse -Filter "*.ps1" -ErrorAction SilentlyContinue | 
              Where-Object { $_.FullName -notmatch "Temp_" }

foreach ($File in $LogicFiles) {
    Add-FileHeader -OutFile $File_Logic -FileName $File.Name
    $Content = Get-SafeContent -Path $File.FullName
    Add-Content -Path $File_Logic -Value $Content -Encoding UTF8
}
# 使用變數寫入標記，避免語法錯誤
Add-Content -Path $File_Logic -Value $EOF_Marker -Encoding UTF8

# ==============================================================================
# 2. 處理設定檔 (Configs)
# ==============================================================================
Write-Host " -> 2/4 提取設定檔 (Configs)..."
Add-Content -Path $File_Config -Value "=== AutoTask Configs Dump ($SnapshotVersion) ===`n" -Encoding UTF8

$ConfigExtensions = @("*.json", "*.map", "*.xml", "*.txt", "*.log", "*.url")
$ConfigFiles = Get-ChildItem -Path "$SourceDir\Configs" -Recurse -Include $ConfigExtensions -ErrorAction SilentlyContinue

foreach ($File in $ConfigFiles) {
    Add-FileHeader -OutFile $File_Config -FileName $File.Name
    $Content = Get-SafeContent -Path $File.FullName
    Add-Content -Path $File_Config -Value $Content -Encoding UTF8
}
Add-Content -Path $File_Config -Value $EOF_Marker -Encoding UTF8

# ==============================================================================
# 3. 處理文件 (Docs)
# ==============================================================================
Write-Host " -> 3/4 提取文件 (Docs)..."
Add-Content -Path $File_Docs -Value "=== AutoTask Documentation ($SnapshotVersion) ===`n" -Encoding UTF8

$DocFiles = Get-ChildItem -Path $SourceDir -Recurse -Include "*.md", "*.txt" -ErrorAction SilentlyContinue | 
            Where-Object { $_.FullName -notmatch "Configs" -and $_.FullName -notmatch "Logs" -and $_.FullName -notmatch "Temp_" }

foreach ($File in $DocFiles) {
    Add-FileHeader -OutFile $File_Docs -FileName $File.Name
    $Content = Get-SafeContent -Path $File.FullName
    Add-Content -Path $File_Docs -Value $Content -Encoding UTF8
}
Add-Content -Path $File_Docs -Value $EOF_Marker -Encoding UTF8

# ==============================================================================
# 4. 處理日誌摘要 (Logs)
# ==============================================================================
Write-Host " -> 4/4 提取日誌摘要..."
$LogsOutFile = "$DirLogs\Recent_Logs_Summary.txt"
Add-Content -Path $LogsOutFile -Value "=== [ Recent Logs Summary ] ===`n" -Encoding UTF8

# 4.1 Master Log
$MasterLog = Get-ChildItem "$SourceDir\Logs\Master_*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$LogTargets = @($MasterLog)

# 4.2 BetterGI Log
$BetterGILogsDir = "C:\Program Files\BetterGI\log"
if (Test-Path $BetterGILogsDir) {
    $LogTargets += Get-ChildItem "$BetterGILogsDir\log_*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 2 
}

foreach ($File in $LogTargets) {
    if ($null -eq $File) { continue }
    Add-FileHeader -OutFile $LogsOutFile -FileName $File.Name
    $Content = Get-SmartLogContent -FilePath $File.FullName
    Add-Content -Path $LogsOutFile -Value $Content -Encoding UTF8
}

# ==============================================================================
# 5. 壓縮打包
# ==============================================================================
Write-Host " -> 5/5 正在壓縮打包..."
try {
    Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force -ErrorAction Stop
} catch {
    Write-Host "壓縮失敗！可能原因：檔案被鎖定或權限不足。這不影響上述檔案生成。" -ForegroundColor Yellow
    Write-Host "錯誤訊息: $($_.Exception.Message)" -ForegroundColor Yellow
}

# 清理暫存
Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "`n[OK] 快照已生成: $ZipPath" -ForegroundColor Green

# UX: 自動選取檔案
if (Test-Path $ZipPath) {
    $arg = "/select,`"$ZipPath`""
    Start-Process explorer.exe -ArgumentList $arg
}

Write-Host "請將該 Zip 檔案上傳給 AI 進行分析。" -ForegroundColor Yellow
Read-Host "按 Enter 鍵結束..."