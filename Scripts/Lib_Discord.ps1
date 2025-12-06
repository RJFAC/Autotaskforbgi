<#
    .SYNOPSIS
    AutoTask Discord 設定精靈 (Fix Version)
    
    .DESCRIPTION
    引導使用者設定 Webhook URL 並發送測試訊息。
    已修正函式呼叫名稱不匹配的問題。
#>

# 定義路徑
$ConfigPath = "C:\AutoTask\Configs\EnvConfig.json"
$LibPath = "$PSScriptRoot\Lib_Discord.ps1"

# 1. 嘗試載入現有設定
if (Test-Path $ConfigPath) {
    try {
        $Global:Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Host "設定檔格式錯誤，將建立新設定。" -ForegroundColor Yellow
        $Global:Config = @{ DiscordWebhook = "" }
    }
} else {
    Write-Host "找不到設定檔，將建立新設定。" -ForegroundColor Yellow
    $Global:Config = @{ DiscordWebhook = "" }
}

# 2. 介面顯示
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      AutoTask Discord 設定精靈" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "目前的 Webhook URL:" -ForegroundColor Gray
if ($Global:Config.DiscordWebhook) {
    Write-Host $Global:Config.DiscordWebhook -ForegroundColor Green
} else {
    Write-Host "(尚未設定)" -ForegroundColor Red
}
Write-Host ""

# 3. 接收輸入
$InputUrl = Read-Host "請輸入新的 Webhook URL (若不修改請直接按 Enter)"

if (-not [string]::IsNullOrWhiteSpace($InputUrl)) {
    # 簡單驗證
    if ($InputUrl -match "^https://discord") {
        $Global:Config.DiscordWebhook = $InputUrl.Trim()
        
        # 儲存設定
        $ConfigDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
        if (-not (Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }
        
        $Global:Config | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
        Write-Host "✅ 設定已儲存至 EnvConfig.json" -ForegroundColor Green
    } else {
        Write-Host "⚠️ 網址格式似乎不正確 (應以 https://discord 開頭)，本次未儲存。" -ForegroundColor Yellow
    }
} else {
    Write-Host "維持原設定。" -ForegroundColor Gray
}

# 4. 執行測試
if ($Global:Config.DiscordWebhook) {
    Write-Host "`n正在準備發送測試訊息..." -ForegroundColor Yellow

    # 檢查並載入 Discord 函式庫
    if (Test-Path $LibPath) {
        . $LibPath
        
        # [Fix] 使用正確的函式名稱 Send-DiscordNotification
        if (Get-Command "Send-DiscordNotification" -ErrorAction SilentlyContinue) {
            Send-DiscordNotification -Title "🔔 設定測試" -Message "恭喜！您的 AutoTask Discord 通知設定已成功生效。" -Color "Green"
            Write-Host "測試指令已發送，請檢查您的 Discord 頻道。" -ForegroundColor Cyan
        } else {
            Write-Host "錯誤: 載入了 $LibPath 但找不到 Send-DiscordNotification 函式。" -ForegroundColor Red
        }
    } else {
        Write-Host "錯誤: 找不到 $LibPath，無法發送測試。" -ForegroundColor Red
    }
}

Write-Host "`n按任意鍵退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")