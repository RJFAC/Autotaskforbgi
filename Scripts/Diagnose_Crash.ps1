# Diagnose_Crash.ps1 - Payload 啟動崩潰診斷工具
$ErrorActionPreference = "Continue"

Write-Host "=== 開始診斷 Payload 啟動流程 ===" -ForegroundColor Cyan

# 1. 測試日誌寫入
Write-Host "[1] 測試日誌寫入權限..." -NoNewline
try {
    $LogFile = "C:\AutoTask\Logs\Diag_Test.log"
    "Test" | Out-File $LogFile -Force
    Write-Host "通過" -ForegroundColor Green
} catch {
    Write-Host "失敗! $_" -ForegroundColor Red; pause; exit
}

# 2. 測試 WMI (自我清理模組)
Write-Host "[2] 測試 WMI (Get-CimInstance)..." -NoNewline
try {
    $proc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop
    Write-Host "通過 (找到 $($proc.Count) 個進程)" -ForegroundColor Green
} catch {
    Write-Host "失敗! WMI 可能損毀或權限不足: $_" -ForegroundColor Red
    Write-Host "建議: 請嘗試重啟電腦修復 WMI。" -ForegroundColor Yellow
}

# 3. 測試環境變數路徑 (模擬 Payload)
Write-Host "[3] 測試環境路徑變數..." 
$ConfigsDir = "C:\AutoTask\Configs"

$FilesToCheck = @(
    "$ConfigsDir\TaskStatus.json",
    "$ConfigsDir\WeeklyConfig.json",
    "$ConfigsDir\DateConfig.map",
    "$ConfigsDir\EnvConfig.json"
)

foreach ($f in $FilesToCheck) {
    Write-Host "    檢查 $f ... " -NoNewline
    if (Test-Path $f) {
        Write-Host "存在" -ForegroundColor Green
    } else {
        Write-Host "不存在 (非致命)" -ForegroundColor Gray
    }
}

# 4. 測試 1Remote 路徑權限 (高機率崩潰點)
Write-Host "[4] 測試 1Remote 日誌路徑存取..." 
# 這裡模擬 Payload 寫死的路徑 (如果是用預設值)
$PotentialPath = "%USERPROFILE%\Downloads\1Remote-1.2.1-net9-x64\.logs" 
Write-Host "    嘗試存取: $PotentialPath"
try {
    # 只是測試 Test-Path，照理說不該崩潰，但如果權限極度嚴格可能會報錯
    $exists = Test-Path $PotentialPath -ErrorAction Stop
    Write-Host "    存取結果: $exists" -ForegroundColor Green
} catch {
    Write-Host "    ⚠️ 存取報錯 (預期中，因為跨帳戶): $_" -ForegroundColor Yellow
    Write-Host "    (Payload 腳本中若直接使用此路徑進行操作可能會崩潰)" -ForegroundColor Gray
}

# 5. 測試 BetterGI 路徑
Write-Host "[5] 測試 BetterGI 路徑..."
$BetterGI = "C:\Program Files\BetterGI\BetterGI.exe"
if (Test-Path $BetterGI) {
    Write-Host "    BetterGI 存在。" -ForegroundColor Green
} else {
    Write-Host "    ⚠️ 找不到 BetterGI！" -ForegroundColor Red
}

Write-Host "`n=== 診斷結束 ===" -ForegroundColor Cyan
pause
