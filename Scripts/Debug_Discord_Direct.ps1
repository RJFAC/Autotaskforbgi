<#
    .SYNOPSIS
    AutoTask Discord 獨立診斷工具
    
    .DESCRIPTION
    完全獨立運作，不依賴 Lib_Discord.ps1。
    用於排除檔案版本舊、路徑錯誤等干擾因素，直接驗證 Webhook URL 是否有效。
#>

# 1. 自動抓取目前的 Webhook 設定
$ConfigPath = "C:\AutoTask\Configs\EnvConfig.json"
$WebhookUrl = ""

Write-Host "=== AutoTask Discord 連線診斷工具 ===" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $ConfigPath) {
    try {
        $Json = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $WebhookUrl = $Json.DiscordWebhook
        Write-Host "已讀取設定檔: $ConfigPath" -ForegroundColor Gray
    } catch {
        Write-Host "設定檔讀取失敗: $_" -ForegroundColor Red
    }
}

# 2. 確認 URL
if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
    Write-Host "目前的 Webhook URL: $($WebhookUrl.Substring(0, 30))..." -ForegroundColor Yellow
} else {
    Write-Host "⚠️ 設定檔中沒有 Webhook URL。" -ForegroundColor Red
    $WebhookUrl = Read-Host "請手動貼上 Webhook URL 進行測試"
}

if (-not $WebhookUrl -match "^https://discord") {
    Write-Host "❌ URL 格式錯誤或為空，停止測試。" -ForegroundColor Red
    Pause
    exit
}

# 3. 測試 A: 最簡測試 (不帶 Timestamp)
# 用途: 排除所有時間格式錯誤的可能性，確認網址本身是否能通。
Write-Host "`n[測試 1/2] 發送純文字訊息 (無時間戳)..." -ForegroundColor Yellow
$Payload_Simple = @{
    content = "🚨 **診斷測試 A**: 這是一則不帶時間戳記的純文字訊息。如果你看得到這則，代表 Webhook URL 是正確的。"
}

try {
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body ($Payload_Simple | ConvertTo-Json) -ErrorAction Stop
    Write-Host "✅ 測試 1 發送成功！(HTTP 200 OK)" -ForegroundColor Green
} catch {
    Write-Host "❌ 測試 1 失敗: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   可能原因: Webhook URL 無效、被 Discord 封鎖、或網路問題。" -ForegroundColor Gray
}

# 4. 測試 B: UTC 時間戳測試
# 用途: 驗證是否因為時區問題導致訊息消失。
Write-Host "`n[測試 2/2] 發送 Embed 訊息 (含 UTC 時間戳)..." -ForegroundColor Yellow

# 正確的 UTC 時間 (解決時光機問題)
$TimeUTC = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

$Payload_Embed = @{
    username = "AutoTask Debugger"
    embeds = @(
        @{
            title = "🛠️ 診斷測試 B"
            description = "這則訊息帶有修正後的 UTC 時間戳記。`n原本的時間: $(Get-Date -Format 'HH:mm:ss')`n轉換後 UTC: $TimeUTC"
            color = 15105570 # Orange
            timestamp = $TimeUTC
        }
    )
}

try {
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body ($Payload_Embed | ConvertTo-Json -Depth 10) -ErrorAction Stop
    Write-Host "✅ 測試 2 發送成功！(HTTP 200 OK)" -ForegroundColor Green
} catch {
    Write-Host "❌ 測試 2 失敗: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=========================================="
Write-Host "診斷完成。請檢查 Discord 頻道。"
Write-Host "如果 [測試 1] 成功但原本腳本失敗 -> 請務必更新 Lib_Discord.ps1"
Write-Host "如果 [測試 1] 失敗 -> 你的 Webhook URL 是壞的，請重新產生一個。"
Write-Host "=========================================="
Pause