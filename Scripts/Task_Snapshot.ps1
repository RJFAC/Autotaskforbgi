# =======================================================
# 檔案名稱: Task_Snapshot.ps1
# 功能: 製作 AutoTask 完整系統快照 (GUI 整合版)
# =======================================================
$ErrorActionPreference = "SilentlyContinue"
$BaseDir = "C:\AutoTask"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SnapshotDir = "$BaseDir\Snapshot_Temp_$TimeStamp"
$ZipFile = "$BaseDir\AutoTask_Snapshot_$TimeStamp.zip"

Write-Host ">>> [AutoTask] 正在製作系統診斷快照..." -ForegroundColor Cyan
New-Item -Path $SnapshotDir -ItemType Directory -Force | Out-Null

Write-Host " -> 掃描檔案結構..."
Get-ChildItem -Path $BaseDir -Recurse | Select-Object FullName | Out-String | Set-Content "$SnapshotDir\0_File_Structure.txt" -Encoding UTF8

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

Write-Host " -> 備份設定檔..."
$ConfigContent = "=== [ AutoTask Configs ] ===`r`n"
$ConfigFiles = Get-ChildItem -Path "$BaseDir\Configs" -Include *.json, *.xml, *.map -Recurse
foreach ($File in $ConfigFiles) {
    $ConfigContent += "`r`n" + ("=" * 70) + "`r`n"
    $ConfigContent += "FILE: $($File.Name)`r`n"
    $ConfigContent += (Get-Content $File.FullName -Raw -Encoding UTF8) + "`r`n"
}
Set-Content -Path "$SnapshotDir\2_Configs.txt" -Value $ConfigContent -Encoding UTF8

Write-Host " -> 擷取近期日誌..."
$LogContent = "=== [ Recent Logs Summary ] ===`r`n"
$LogFiles = Get-ChildItem -Path "$BaseDir\Logs", "$BaseDir\1Remote\.logs" -Include *.log, *.md -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 10
foreach ($File in $LogFiles) {
    $LogContent += "`r`n" + ("=" * 70) + "`r`n"
    $LogContent += "FILE: $($File.Name) | TIME: $($File.LastWriteTime)`r`n"
    $LogContent += ("=" * 70) + "`r`n"
    $LogContent += (Get-Content $File.FullName -Tail 100 -Encoding UTF8 | Out-String) + "`r`n"
}
Set-Content -Path "$SnapshotDir\3_Recent_Logs.txt" -Value $LogContent -Encoding UTF8

Write-Host " -> 正在壓縮..."
Compress-Archive -Path "$SnapshotDir\*" -DestinationPath $ZipFile -Force
Remove-Item -Path $SnapshotDir -Recurse -Force

Write-Host ">>> 快照製作完成！" -ForegroundColor Green
Write-Host "檔案: $ZipFile" -ForegroundColor Yellow
Invoke-Item $ZipFile
Start-Sleep -Seconds 2