# =======================================================
# 檔案名稱: Task_Snapshot.ps1
# 功能: 製作 AutoTask 完整系統快照 (策略 v3.0 - 大容量日誌版)
# =======================================================
$ErrorActionPreference = "SilentlyContinue"
$BaseDir = "C:\AutoTask"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SnapshotDir = "$BaseDir\Snapshot_Temp_$TimeStamp"
$ZipFile = "$BaseDir\AutoTask_Snapshot_$TimeStamp.zip"

# --- [設定] 日誌截斷參數 (大幅放寬限制) ---
# 原因：文字檔壓縮率極高 (5MB -> ~300KB zip)，為了除錯應盡量保留完整日誌。
$LogSizeThreshold = 5MB   # 閾值提升至 5MB (確保 99% 的日誌都能完整讀取)
$LogHeadLines     = 200   # 若必須截斷，保留前 200 行 (確保包含所有啟動參數與環境檢查)
$LogTailLines     = 5000  # 若必須截斷，保留後 5000 行 (約 500KB，涵蓋完整錯誤堆疊與上下文)

Write-Host ">>> [AutoTask] 正在製作系統診斷快照..." -ForegroundColor Cyan
New-Item -Path $SnapshotDir -ItemType Directory -Force | Out-Null

# 1. 檔案結構
Write-Host " -> 掃描檔案結構..."
Get-ChildItem -Path $BaseDir -Recurse | Select-Object FullName | Out-String | Set-Content "$SnapshotDir\0_File_Structure.txt" -Encoding UTF8

# 2. 腳本檔案 (完整內容)
Write-Host " -> 備份腳本內容..."
$ScriptContent = "=== [ AutoTask Scripts Snapshot ] ===`r`n"
$ScriptFiles = Get-ChildItem -Path "$BaseDir\Scripts", "$BaseDir" -Include *.ps1, *.bat -Recurse
foreach ($File in $ScriptFiles) {
    if ($File.Length -lt 5000000) { 
        $ScriptContent += "`r`n" + ("=" * 70) + "`r`n"
        $ScriptContent += "FILE: $($File.Name) | PATH: $($File.FullName)`r`n"
        $ScriptContent += ("=" * 70) + "`r`n"
        try { $ScriptContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n" } catch { $ScriptContent += "[Error Reading File]`r`n" }
    }
}
Set-Content -Path "$SnapshotDir\1_Full_Scripts.txt" -Value $ScriptContent -Encoding UTF8

# 3. 設定檔
Write-Host " -> 備份設定檔..."
$ConfigContent = "=== [ AutoTask Configs ] ===`r`n"
$ConfigFiles = Get-ChildItem -Path "$BaseDir\Configs" -Include *.json, *.xml, *.map -Recurse
foreach ($File in $ConfigFiles) {
    $ConfigContent += "`r`n" + ("=" * 70) + "`r`n"
    $ConfigContent += "FILE: $($File.Name)`r`n"
    $ConfigContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
}
Set-Content -Path "$SnapshotDir\2_Configs.txt" -Value $ConfigContent -Encoding UTF8

# 4. 近期日誌 (策略優化)
Write-Host " -> 擷取近期日誌 (閾值: 5MB)..."
$LogContent = "=== [ Recent Logs Summary ] ===`r`n"
# 抓取 AutoTask Logs, 1Remote Logs 以及 BetterGI Logs
$BetterGILogDir = "C:\Program Files\BetterGI\log"
$LogSearchPaths = @("$BaseDir\Logs", "$BaseDir\1Remote\.logs")
if (Test-Path $BetterGILogDir) { $LogSearchPaths += $BetterGILogDir }

# 增加擷取數量至 20 個，確保包含 Master, Monitor, Payload 以及 BetterGI 的日誌
$LogFiles = Get-ChildItem -Path $LogSearchPaths -Include *.log, *.md -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 20

foreach ($File in $LogFiles) {
    $SizeMB = [math]::Round($File.Length/1MB, 2)
    $LogContent += "`r`n" + ("=" * 70) + "`r`n"
    $LogContent += "FILE: $($File.Name)`r`nPATH: $($File.FullName)`r`nTIME: $($File.LastWriteTime)`r`nSIZE: $SizeMB MB`r`n"
    $LogContent += ("=" * 70) + "`r`n"
    
    try {
        if ($File.Length -gt $LogSizeThreshold) {
            Write-Host "    [截斷警告] $($File.Name) 大小 ($SizeMB MB) 超過閾值，僅保留頭尾..." -ForegroundColor Yellow
            $Head = Get-Content $File.FullName -Head $LogHeadLines -Encoding UTF8
            $Tail = Get-Content $File.FullName -Tail $LogTailLines -Encoding UTF8
            
            $LogContent += ($Head -join "`r`n") + "`r`n"
            $LogContent += "`r`n... [FILE TOO LARGE - CONTENT TRUNCATED] ...`r`n`r`n"
            $LogContent += ($Tail -join "`r`n") + "`r`n"
        } else {
            # 5MB 以下直接完整讀取
            $LogContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
        }
    } catch {
        $LogContent += "[讀取檔案時發生錯誤: $_]`r`n"
    }
}
Set-Content -Path "$SnapshotDir\3_Recent_Logs.txt" -Value $LogContent -Encoding UTF8

# 5. 壓縮
Write-Host " -> 正在壓縮..."
Compress-Archive -Path "$SnapshotDir\*" -DestinationPath $ZipFile -Force
Remove-Item -Path $SnapshotDir -Recurse -Force

Write-Host ">>> 快照製作完成！" -ForegroundColor Green
Write-Host "檔案: $ZipFile" -ForegroundColor Yellow
Invoke-Item $ZipFile
Start-Sleep -Seconds 2