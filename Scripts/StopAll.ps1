# ============================================================
# AutoTask 強制停止腳本 V3.1 (同步嚴格登出邏輯)
# ============================================================
$ErrorActionPreference = 'SilentlyContinue'

# --- [定義路徑] ---
$BaseDir  = "C:\AutoTask"
$FlagDir  = "$BaseDir\Flags"
$ConfigDir = "$BaseDir\Configs"

Write-Host "正在初始化強制清理程序..." -ForegroundColor Cyan

# 1. 檢查管理員權限
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "⚠️ 警告：未以系統管理員身分執行，無法清理遠端帳戶 Session！" -ForegroundColor Yellow
    Write-Host "建議右鍵點選 .bat 檔 -> 以系統管理員身分執行。" -ForegroundColor Yellow
    Start-Sleep 2
}

# 2. [最優先] 停止控制腳本
Write-Host "`n[1/5] 正在凍結自動化大腦 (停止腳本)..." -ForegroundColor Cyan
$MyPID = $PID
try {
    $Targets = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
               Where-Object { ($_.CommandLine -like '*AutoTask*' -or $_.CommandLine -like '*Payload.ps1*') -and $_.ProcessId -ne $MyPID }
    if ($Targets) {
        foreach ($p in $Targets) {
            Stop-Process -Id $p.ProcessId -Force
            Write-Host "    - 已終止腳本 PID: $($p.ProcessId)" -ForegroundColor Green
        }
    } else {
        Write-Host "    - 未發現其他運行中的腳本。" -ForegroundColor Gray
    }
} catch {
    Write-Host "    - 掃描腳本時發生錯誤 (WMI 可能忙碌)。" -ForegroundColor Red
}

# 3. [次優先] 強制清理業務應用程式
Write-Host "`n[2/5] 正在強制清理應用程式..." -ForegroundColor Cyan
$Apps = @("BetterGI", "GenshinImpact", "YuanShen")
foreach ($app in $Apps) {
    $pApp = Get-Process -Name $app -ErrorAction SilentlyContinue
    if ($pApp) {
        Stop-Process -Name $app -Force
        Write-Host "    - $app 已強制終止。" -ForegroundColor Green
    } else {
        Write-Host "    - 未發現 $app。" -ForegroundColor Gray
    }
}

# 4. [核心] 強制登出 Remote 帳戶 (嚴格等待版)
Write-Host "`n[3/5] 正在清理遠端工作階段 (Remote Session)..." -ForegroundColor Cyan
$SessionOutput = qwinsta 2>$null | Select-String "\bRemote\b"
if ($SessionOutput) {
    $Line = $SessionOutput.ToString().Trim() -replace "\s+", " "
    $Parts = $Line.Split(" ")
    $SessionID = $null
    foreach ($part in $Parts) { if ($part -match "^\d+$") { $SessionID = $part; break } }

    if ($SessionID) {
        Write-Host "    - 發現 Remote Session (ID: $SessionID)，執行強制登出..."
        logoff $SessionID
        
        Write-Host "    - 正在等待 Session 銷毀..." -NoNewline
        $Timeout = 0
        while ($true) {
            if (-not (qwinsta 2>$null | Select-String "\bRemote\b")) {
                Write-Host " [已確認登出]" -ForegroundColor Green
                break
            }
            if ($Timeout -ge 10) {
                Write-Host " [警告: 逾時]" -ForegroundColor Yellow
                break
            }
            Write-Host "." -NoNewline
            Start-Sleep 1
            $Timeout++
        }
    }
} else {
    Write-Host "    - Remote 帳戶目前未登入。" -ForegroundColor Gray
}

# 5. 停止 1Remote
Write-Host "`n[4/5] 正在停止本機 1Remote..." -ForegroundColor Cyan
$p1 = Get-Process -Name "1Remote" -ErrorAction SilentlyContinue
if ($p1) {
    Stop-Process -Name "1Remote" -Force
    Write-Host "    - 1Remote 已終止。" -ForegroundColor Green
} else {
    Write-Host "    - 未發現 1Remote 進程。" -ForegroundColor Gray
}

# 6. 清理所有狀態檔與旗標
Write-Host "`n[5/5] 正在清理狀態檔案..." -ForegroundColor Cyan

$Flags = Get-ChildItem -Path $FlagDir -Filter "*.flag"
if ($Flags) {
    Remove-Item "$FlagDir\*.flag" -Force
    Write-Host "    - 旗標已清理: $($Flags.Name -join ', ')" -ForegroundColor Green
}

if (Test-Path "$ConfigDir\TaskStatus.json") {
    Remove-Item "$ConfigDir\TaskStatus.json" -Force
    Write-Host "    - TaskStatus.json 已重置。" -ForegroundColor Green
}

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "   系統已完全重置 (Clean & Stop)"
Write-Host "========================================" -ForegroundColor Yellow