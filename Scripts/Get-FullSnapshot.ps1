# =======================================================
# 檔案名稱: Get-FullSnapshot.ps1
# 功能: 完整擷取 AutoTask 系統快照 (無截斷模式)
# 說明: 
#   1. 強制擷取所有 Scripts 與 Configs 的完整內容
#   2. Logs 僅擷取最後 100 行以節省空間
#   3. 自動打包為 ZIP 方便上傳
# =======================================================

$ErrorActionPreference = "SilentlyContinue"
$BaseDir = "C:\AutoTask"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SnapshotDir = "$BaseDir\Snapshot_Temp_$TimeStamp"
$ZipFile = "$BaseDir\AutoTask_FullSnapshot_$TimeStamp.zip"

# 建立暫存目錄
New-Item -Path $SnapshotDir -ItemType Directory -Force | Out-Null

Write-Host ">>> 開始製作完整系統快照..." -ForegroundColor Cyan

# -------------------------------------------------------
# 1. 檔案結構 (File Structure)
# -------------------------------------------------------
Write-Host "1. 掃描檔案結構..."
Get-ChildItem -Path $BaseDir -Recurse | Select-Object FullName | Out-String | Set-Content "$SnapshotDir\0_File_Structure.txt" -Encoding UTF8

# -------------------------------------------------------
# 2. 腳本檔案 (Scripts) - !!關鍵修改: 強制讀取全文!!
# -------------------------------------------------------
Write-Host "2. 擷取腳本檔案 (完整內容)..."
$ScriptContent = "=== [ AutoTask PowerShell & Batch Scripts (FULL CONTENT) ] ===`r`n"
$ScriptFiles = Get-ChildItem -Path "$BaseDir\Scripts", "$BaseDir" -Include *.ps1, *.bat -Recurse

foreach ($File in $ScriptFiles) {
    # 跳過備份檔或過大的檔案 (超過 5MB 跳過，避免記憶體溢位，但腳本通常很小)
    if ($File.Length -lt 5000000) {
        $ScriptContent += "`r`n" + ("=" * 70) + "`r`n"
        $ScriptContent += "FILE: $($File.Name)  |  PATH: $($File.FullName)`r`n"
        $ScriptContent += "SIZE: $([math]::Round($File.Length / 1KB, 2)) KB`r`n"
        $ScriptContent += ("=" * 70) + "`r`n"
        
        # 讀取原始內容 (Raw)
        try {
            $Content = Get-Content $File.FullName -Raw -Encoding UTF8
            $ScriptContent += $Content + "`r`n"
        } catch {
            $ScriptContent += "[Error Reading File: $_]`r`n"
        }
    }
}
Set-Content -Path "$SnapshotDir\1_Full_Scripts.txt" -Value $ScriptContent -Encoding UTF8

# -------------------------------------------------------
# 3. 設定檔 (Configs)
# -------------------------------------------------------
Write-Host "3. 擷取設定檔..."
$ConfigContent = "=== [ AutoTask Configs ] ===`r`n"
$ConfigFiles = Get-ChildItem -Path "$BaseDir\Configs" -Include *.json, *.xml, *.map -Recurse

foreach ($File in $ConfigFiles) {
    $ConfigContent += "`r`n" + ("=" * 70) + "`r`n"
    $ConfigContent += "FILE: $($File.Name)`r`n"
    $ConfigContent += ("=" * 70) + "`r`n"
    $ConfigContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
}
Set-Content -Path "$SnapshotDir\2_Configs.txt" -Value $ConfigContent -Encoding UTF8

# -------------------------------------------------------
# 4. 系統日誌 (Logs) - 限制行數以免過大
# -------------------------------------------------------
Write-Host "4. 擷取近期日誌 (Last 100 Lines)..."
$LogContent = "=== [ Recent Logs Summary ] ===`r`n"
$LogFiles = Get-ChildItem -Path "$BaseDir\Logs", "$BaseDir\1Remote\.logs" -Include *.log, *.md -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 10

foreach ($File in $LogFiles) {
    $LogContent += "`r`n" + ("=" * 70) + "`r`n"
    $LogContent += "FILE: $($File.Name)  |  TIME: $($File.LastWriteTime)`r`n"
    $LogContent += ("=" * 70) + "`r`n"
    $LogContent += (Get-Content $File.FullName -Tail 100 -Encoding UTF8 | Out-String) + "`r`n"
}
Set-Content -Path "$SnapshotDir\3_Recent_Logs.txt" -Value $LogContent -Encoding UTF8

# -------------------------------------------------------
# 5. 壓縮與清理
# -------------------------------------------------------
Write-Host "5. 壓縮檔案中..."
Compress-Archive -Path "$SnapshotDir\*" -DestinationPath $ZipFile -Force

# 清理暫存資料夾
Remove-Item -Path $SnapshotDir -Recurse -Force

Write-Host "`n>>> 快照完成！" -ForegroundColor Green
Write-Host "檔案位置: $ZipFile" -ForegroundColor Yellow
Write-Host "請將此 ZIP 檔案上傳給我。"