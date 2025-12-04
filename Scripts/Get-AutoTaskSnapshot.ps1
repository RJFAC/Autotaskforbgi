# =============================================================================
# Get-AutoTaskSnapshot_Merged_v5.ps1 - 系統資訊合併打包工具 (智慧截斷版)
# =============================================================================
# 版本：V5 (優化大檔案處理策略)
# 功能：
# 1. 收集 AutoTask 架構、腳本、設定、XML 排程。
# 2. 智慧合併日誌：
#    - 小檔案 (<10MB)：完整保留。
#    - 大檔案 (>10MB)：保留 "前100行(啟動資訊)" + "最後20000行(關鍵錯誤)"。
# 3. 輸出至 C:\AutoTask\，檔名包含時間戳記。

$ErrorActionPreference = "SilentlyContinue"
$BaseDir = "C:\AutoTask"
$TempDir = "$BaseDir\Snapshot_Temp"

# --- [生成帶時間戳的檔名] ---
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ZipName = "AutoTask_Snapshot_$TimeStamp.zip"
$ZipPath = Join-Path $BaseDir $ZipName
$DateLimit = (Get-Date).AddHours(-24)

# 閾值設定
$SizeLimitBytes = 10MB  # 超過 10MB 視為大檔案
$TailLines      = 20000 # 保留最後 20000 行
$HeadLines      = 100   # 保留前 100 行 (保留啟動參數)

# --- [0. 權限檢查] ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "請求管理員權限..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host "正在執行系統快照 (V5 - 智慧截斷)..." -ForegroundColor Cyan
Write-Host "目標檔案: $ZipName" -ForegroundColor Gray

# --- [1. 準備目錄] ---
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -Path $TempDir -ItemType Directory | Out-Null

# --- [輔助函式: 合併檔案] ---
function Merge-Files {
    param($SourcePath, $Filter, $OutputFile, $HeaderMsg)
    Write-Host "   處理: $HeaderMsg"
    
    $OutPath = $OutputFile.FullName
    "=== [ $HeaderMsg ] ===" | Set-Content $OutPath -Encoding UTF8
    
    if (Test-Path $SourcePath) {
        $Files = Get-ChildItem $SourcePath -Include $Filter -Recurse | Sort-Object Name
        
        # 若是日誌，只取最近 24 小時
        if ($HeaderMsg -match "Log") {
            $Files = $Files | Where-Object { $_.LastWriteTime -ge $DateLimit }
        }

        foreach ($file in $Files) {
            # 略過壓縮檔本身與暫存檔
            if ($file.FullName -like "*.zip" -or $file.FullName -like "*Snapshot_Temp*") { continue }

            "`n`n" | Add-Content $OutPath -Encoding UTF8
            "=============================================================================" | Add-Content $OutPath -Encoding UTF8
            "FILE: $($file.Name)  |  PATH: $($file.FullName)" | Add-Content $OutPath -Encoding UTF8
            "TIME: $($file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))  |  SIZE: $([math]::Round($file.Length/1MB, 2)) MB" | Add-Content $OutPath -Encoding UTF8
            "=============================================================================" | Add-Content $OutPath -Encoding UTF8
            
            try {
                if ($file.Length -gt $SizeLimitBytes) {
                    Write-Host "      [截斷] $($file.Name) 大於 10MB，執行智慧縮減..." -ForegroundColor DarkGray
                    
                    "[SYSTEM NOTICE: File too large ($([math]::Round($file.Length/1MB, 2)) MB). Content Truncated.]" | Add-Content $OutPath -Encoding UTF8
                    "[--- HEAD (First $HeadLines lines) ---]" | Add-Content $OutPath -Encoding UTF8
                    Get-Content $file.FullName -Head $HeadLines | Out-String | Add-Content $OutPath -Encoding UTF8
                    
                    "`n... (Middle content omitted) ...`n" | Add-Content $OutPath -Encoding UTF8
                    
                    "[--- TAIL (Last $TailLines lines) ---]" | Add-Content $OutPath -Encoding UTF8
                    Get-Content $file.FullName -Tail $TailLines | Out-String | Add-Content $OutPath -Encoding UTF8
                } else {
                    Get-Content $file.FullName -Raw | Add-Content $OutPath -Encoding UTF8
                }
            } catch {
                "[ERROR READING FILE: $_]" | Add-Content $OutPath -Encoding UTF8
            }
        }
    } else {
        "`n[Path Not Found: $SourcePath]" | Add-Content $OutPath -Encoding UTF8
    }
}

# --- [2. 合併 AutoTask 核心] ---
Write-Host "1. 合併核心腳本與設定..."
Merge-Files -SourcePath "$BaseDir\Scripts" -Filter "*.ps1" -OutputFile (New-Item "$TempDir\1_AT_Scripts.txt" -Force) -HeaderMsg "AutoTask PowerShell Scripts"
Merge-Files -SourcePath "$BaseDir\Configs" -Filter "*.json","*.map","*.log" -OutputFile (New-Item "$TempDir\2_AT_Configs.txt" -Force) -HeaderMsg "AutoTask Configs"

# 附加 Flags 狀態
$FlagFile = "$TempDir\2_AT_Configs.txt"
"`n`n=== [ Current Flags ] ===" | Add-Content $FlagFile -Encoding UTF8
Get-ChildItem "$BaseDir\Flags" | Select Name, LastWriteTime | Out-String | Add-Content $FlagFile -Encoding UTF8

# --- [3. 匯出與合併工作排程] ---
Write-Host "2. 匯出工作排程..."
$TaskFile = "$TempDir\3_Task_Schedules.xml.txt"
"=== Task Schedules XML Export ===" | Set-Content $TaskFile -Encoding UTF8
"`n--- [ Auto_BetterGI_Payload ] ---" | Add-Content $TaskFile -Encoding UTF8
schtasks /query /tn "Auto_BetterGI_Payload" /xml | Out-String | Add-Content $TaskFile -Encoding UTF8
"`n`n--- [ Auto_1Remote_Master ] ---" | Add-Content $TaskFile -Encoding UTF8
schtasks /query /tn "Auto_1Remote_Master" /xml | Out-String | Add-Content $TaskFile -Encoding UTF8

# --- [4. 合併 AutoTask 日誌] ---
Write-Host "3. 合併 AutoTask 日誌 (24h)..."
Merge-Files -SourcePath "$BaseDir\Logs" -Filter "*.log" -OutputFile (New-Item "$TempDir\4_Logs_AutoTask.txt" -Force) -HeaderMsg "AutoTask System Logs (Last 24h)"

# --- [5. 合併外部程式日誌] ---
Write-Host "4. 偵測並合併 BetterGI 與 1Remote 日誌..."

$BetterGI_Path = "C:\Program Files\BetterGI"
$OneRemote_Path = $null
if (Test-Path "$BaseDir\Configs\EnvConfig.json") {
    try {
        $env = Get-Content "$BaseDir\Configs\EnvConfig.json" -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env.Path1Remote) { $OneRemote_Path = Split-Path $env.Path1Remote -Parent }
    } catch {}
}

# BetterGI
$BGFile = New-Item "$TempDir\5_Logs_BetterGI.txt" -Force
"=== BetterGI User Config ===" | Set-Content $BGFile -Encoding UTF8
if (Test-Path "$BetterGI_Path\User\config.json") { Get-Content "$BetterGI_Path\User\config.json" -Raw | Add-Content $BGFile -Encoding UTF8 }
Merge-Files -SourcePath "$BetterGI_Path\log" -Filter "*.log" -OutputFile $BGFile -HeaderMsg "BetterGI Logs (Last 24h)"

# 1Remote
if ($OneRemote_Path -and (Test-Path $OneRemote_Path)) {
    Merge-Files -SourcePath "$OneRemote_Path\.logs" -Filter "*.log","*.md" -OutputFile (New-Item "$TempDir\6_Logs_1Remote.txt" -Force) -HeaderMsg "1Remote Logs (Last 24h)"
}

# --- [6. 生成目錄結構樹] ---
Write-Host "5. 生成檔案結構樹..."
$TreeFile = "$TempDir\0_File_Structure.txt"
"=== AutoTask Directory Structure ===" | Set-Content $TreeFile -Encoding UTF8
Get-ChildItem $BaseDir -Recurse | Select-Object FullName, LastWriteTime, Length | Out-String | Add-Content $TreeFile -Encoding UTF8

# --- [7. 壓縮打包] ---
Write-Host "6. 正在壓縮..." -ForegroundColor Yellow
$CompressScript = {
    param($src, $dest)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($src, $dest)
}
& $CompressScript -src $TempDir -dest $ZipPath

# --- [8. 清理] ---
Remove-Item $TempDir -Recurse -Force

Write-Host "`n✅ [完成]" -ForegroundColor Green
Write-Host "   檔案: $ZipName" -ForegroundColor Cyan
Write-Host "   路徑: $ZipPath" -ForegroundColor Gray
Pause
