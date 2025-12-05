# =======================================================
# 檔案名稱: Task_Snapshot.ps1
# 功能: 製作 AutoTask 完整系統快照 (GUI 整合修復版 v2)
# 更新: 修復 BetterGI 日誌未被包含的問題、加入大檔案截斷保護
# =======================================================
$ErrorActionPreference = "SilentlyContinue"
$BaseDir = "C:\AutoTask"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SnapshotDir = "$BaseDir\Snapshot_Temp_$TimeStamp"
$ZipFile = "$BaseDir\AutoTask_Snapshot_$TimeStamp.zip"

# --- [設定] 日誌參數 ---
$LogSizeThreshold = 5MB   # 超過 5MB 則截斷
$LogHeadLines     = 200   # 保留前 200 行
$LogTailLines     = 5000  # 保留後 5000 行

Write-Host ">>> [AutoTask] 正在製作系統診斷快照..." -ForegroundColor Cyan
New-Item -Path $SnapshotDir -ItemType Directory -Force | Out-Null

# 1. 檔案結構
Write-Host " -> [1/5] 掃描檔案結構..."
Get-ChildItem -Path $BaseDir -Recurse | Select-Object FullName | Out-String | Set-Content "$SnapshotDir\0_File_Structure.txt" -Encoding UTF8

# 2. 腳本檔案 (完整內容)
Write-Host " -> [2/5] 備份腳本內容..."
$ScriptContent = "=== [ AutoTask Scripts Snapshot ] ===`r`n"
$ScriptFiles = Get-ChildItem -Path "$BaseDir\Scripts", "$BaseDir" -Include *.ps1, *.bat -Recurse
foreach ($File in $ScriptFiles) {
    if ($File.Length -lt 5000000) { 
        $ScriptContent += "`r`n" + ("=" * 70) + "`r`n"
        $ScriptContent += "FILE: $($File.Name) | PATH: $($File.FullName)`r`n"
        $ScriptContent += ("=" * 70) + "`r`n"
        try { 
            $ScriptContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n" 
        } catch { 
            $ScriptContent += "[Error Reading File]`r`n" 
        }
    }
}
Set-Content -Path "$SnapshotDir\1_Full_Scripts.txt" -Value $ScriptContent -Encoding UTF8

# 3. 設定檔
Write-Host " -> [3/5] 備份設定檔..."
$ConfigContent = "=== [ AutoTask Configs ] ===`r`n"
$ConfigFiles = Get-ChildItem -Path "$BaseDir\Configs" -Include *.json, *.xml, *.map -Recurse
foreach ($File in $ConfigFiles) {
    $ConfigContent += "`r`n" + ("=" * 70) + "`r`n"
    $ConfigContent += "FILE: $($File.Name)`r`n"
    $ConfigContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
}
Set-Content -Path "$SnapshotDir\2_Configs.txt" -Value $ConfigContent -Encoding UTF8

# 4. 近期日誌 (包含 BetterGI)
Write-Host " -> [4/5] 擷取近期日誌 (含 BetterGI)..."
$LogContent = "=== [ Recent Logs Summary ] ===`r`n"

# 定義日誌來源路徑
$LogSearchPaths = @("$BaseDir\Logs", "$BaseDir\1Remote\.logs")

# [修復] 加入 BetterGI 日誌路徑偵測
$BetterGIDefault = "C:\Program Files\BetterGI\log"
# 嘗試從設定檔讀取自定義路徑 (如果有的話)
if (Test-Path "$BaseDir\Configs\EnvConfig.json") {
    try {
        $EnvParams = Get-Content "$BaseDir\Configs\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        # 如果未來有在 EnvConfig 定義 BetterGIPath，可以在這裡擴充
    } catch {}
}

if (Test-Path $BetterGIDefault) { 
    Write-Host "    - 偵測到 BetterGI 日誌目錄，已加入掃描。" -ForegroundColor Gray
    $LogSearchPaths += $BetterGIDefault 
} else {
    Write-Host "    -⚠️ 未偵測到預設 BetterGI 日誌目錄 ($BetterGIDefault)" -ForegroundColor Yellow
}

# 增加擷取數量至 25 個，並包含 .log, .md
$LogFiles = Get-ChildItem -Path $LogSearchPaths -Include *.log, *.md -Recurse -ErrorAction SilentlyContinue | 
    Sort-Object LastWriteTime -Descending | Select-Object -First 25

foreach ($File in $LogFiles) {
    $SizeMB = [math]::Round($File.Length/1MB, 2)
    $LogContent += "`r`n" + ("=" * 70) + "`r`n"
    $LogContent += "FILE: $($File.Name)`r`nPATH: $($File.FullName)`r`nTIME: $($File.LastWriteTime)`r`nSIZE: $SizeMB MB`r`n"
    $LogContent += ("=" * 70) + "`r`n"
    
    try {
        if ($File.Length -gt $LogSizeThreshold) {
            # 大檔案截斷處理
            $LogContent += "[⚠️ FILE TOO LARGE ($SizeMB MB) - TRUNCATED MODE]`r`n`r`n"
            $Head = Get-Content $File.FullName -Head $LogHeadLines -Encoding UTF8
            $Tail = Get-Content $File.FullName -Tail $LogTailLines -Encoding UTF8
            
            $LogContent += ($Head -join "`r`n") + "`r`n"
            $LogContent += "`r`n... [ Content Truncated / 略過中間內容 ] ...`r`n`r`n"
            $LogContent += ($Tail -join "`r`n") + "`r`n"
        } else {
            # 一般讀取
            $LogContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
        }
    } catch {
        $LogContent += "[讀取檔案時發生錯誤: $_]`r`n"
    }
}
Set-Content -Path "$SnapshotDir\3_Recent_Logs.txt" -Value $LogContent -Encoding UTF8

# 5. 壓縮
Write-Host " -> [5/5] 正在壓縮..."
Compress-Archive -Path "$SnapshotDir\*" -DestinationPath $ZipFile -Force
Remove-Item -Path $SnapshotDir -Recurse -Force

Write-Host ">>> 快照製作完成！" -ForegroundColor Green
Write-Host "檔案: $ZipFile" -ForegroundColor Yellow

# 自動開啟檔案位置
Invoke-Item $ZipFile
Start-Sleep -Seconds 2