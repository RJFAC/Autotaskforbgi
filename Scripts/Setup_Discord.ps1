<#
    .SYNOPSIS
    AutoTask Discord Manager (Dashboard Module)
    
    .DESCRIPTION
    提供選單式介面，允許使用者：
    1. 檢視或修改 Webhook URL
    2. 發送測試訊息 (驗證功能是否正常)
    3. 查看最近的 Log (除錯用)
#>

# ==========================================
# 初始化環境
# ==========================================
$ConfigPath = "C:\AutoTask\Configs\EnvConfig.json"
$LibPath = "$PSScriptRoot\Lib_Discord.ps1"

# 載入設定檔函式
function Load-Config {
    if (Test-Path $ConfigPath) {
        try {
            $Global:Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            $Global:Config = @{ DiscordWebhook = "" }
        }
    } else {
        $Global:Config = @{ DiscordWebhook = "" }
    }
}

# 載入 Discord 函式庫
if (Test-Path $LibPath) {
    . $LibPath
} else {
    Write-Host "錯誤: 找不到 $LibPath，功能將受限。" -ForegroundColor Red
    Pause
}

# ==========================================
# 功能函式
# ==========================================

function Show-Header {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      AutoTask Discord 管理面板" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # 顯示目前狀態
    Write-Host "目前狀態: " -NoNewline
    if ($Global:Config.DiscordWebhook) {
        Write-Host "✅ 已設定 Webhook" -ForegroundColor Green
        Write-Host "URL: $($Global:Config.DiscordWebhook.Substring(0, 30))..." -ForegroundColor DarkGray
    } else {
        Write-Host "❌ 尚未設定" -ForegroundColor Red
    }
    Write-Host ""
}

function Update-Webhook {
    Write-Host "--- 修改 Webhook ---" -ForegroundColor Yellow
    $InputUrl = Read-Host "請輸入新的 Webhook URL (留空取消)"
    
    if (-not [string]::IsNullOrWhiteSpace($InputUrl)) {
        if ($InputUrl -match "^https://discord") {
            $Global:Config.DiscordWebhook = $InputUrl.Trim()
            
            # 確保目錄存在
            $ConfigDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
            if (-not (Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }
            
            # 儲存
            $Global:Config | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
            Write-Host "✅ 設定已儲存！" -ForegroundColor Green
        } else {
            Write-Host "⚠️ URL 格式錯誤 (需以 https://discord 開頭)" -ForegroundColor Red
        }
    } else {
        Write-Host "已取消。" -ForegroundColor Gray
    }
    Start-Sleep -Seconds 1
}

function Run-Test {
    Write-Host "--- 發送測試訊息 ---" -ForegroundColor Yellow
    
    if (-not $Global:Config.DiscordWebhook) {
        Write-Host "錯誤: 請先設定 Webhook URL。" -ForegroundColor Red
        Pause
        return
    }

    if (Get-Command "Send-DiscordNotification" -ErrorAction SilentlyContinue) {
        Write-Host "正在發送..." -ForegroundColor Gray
        Send-DiscordNotification -Title "📡 通訊測試" -Message "這是來自 AutoTask Dashboard 的手動測試訊息。`n如果您看到這則訊息，代表 Webhook 設定完全正常！" -Color "Green"
        Write-Host "指令已發送，請檢查 Discord 頻道。" -ForegroundColor Cyan
    } else {
        Write-Host "錯誤: 找不到 Send-DiscordNotification 函式。請檢查 Lib_Discord.ps1。" -ForegroundColor Red
    }
    
    Write-Host "`n按任意鍵返回..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ==========================================
# 主迴圈
# ==========================================

Load-Config

while ($true) {
    Show-Header
    
    Write-Host "請選擇功能:" -ForegroundColor Yellow
    Write-Host " [1] 修改/設定 Webhook URL"
    Write-Host " [2] 發送測試訊息 (Test Message)"
    Write-Host " [Q] 退出"
    Write-Host ""
    
    $Choice = Read-Host "輸入選項"
    
    switch ($Choice) {
        "1" { Update-Webhook }
        "2" { Run-Test }
        "Q" { exit }
        "q" { exit }
        default { Write-Host "無效選項" -ForegroundColor Red; Start-Sleep -Milliseconds 500 }
    }
}