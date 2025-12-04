# =======================================================
# 檔案名稱: Setup_Discord.ps1
# 功能: 設定 Discord Webhook URL 並測試
# =======================================================

$ConfigPath = "C:\AutoTask\Configs\EnvConfig.json"
$LibPath = "C:\AutoTask\Scripts\Lib_Discord.ps1"

# 1. 檢查 Config 檔案
if (-not (Test-Path $ConfigPath)) {
    Write-Host "找不到 EnvConfig.json，正在建立..." -ForegroundColor Yellow
    $Config = @{ GenshinPath = ""; Path1Remote = "" }
} else {
    $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
}

# 2. 顯示目前設定
Write-Host "=== AutoTask Discord 設定精靈 ===" -ForegroundColor Cyan
if ($Config.PSObject.Properties.Match('DiscordWebhook').Count -gt 0 -and $Config.DiscordWebhook) {
    Write-Host "目前 Webhook: $($Config.DiscordWebhook.Substring(0, 30))..." -ForegroundColor Gray
} else {
    Write-Host "目前尚未設定 Webhook。" -ForegroundColor Red
}

# 3. 輸入網址
$NewUrl = Read-Host "`n請輸入你的 Discord Webhook URL (留空則不修改)"

if (-not [string]::IsNullOrWhiteSpace($NewUrl)) {
    if ($Config.PSObject.Properties.Match('DiscordWebhook').Count -eq 0) {
        $Config | Add-Member -Type NoteProperty -Name "DiscordWebhook" -Value $NewUrl
    } else {
        $Config.DiscordWebhook = $NewUrl
    }
    
    $Config | ConvertTo-Json -Depth 4 | Set-Content $ConfigPath -Encoding UTF8
    Write-Host "設定已儲存！" -ForegroundColor Green
}

# 4. 測試發送
if (Test-Path $LibPath) {
    . $LibPath
    Write-Host "`n正在發送測試訊息..." -ForegroundColor Cyan
    Send-DiscordWebhook `
        -WebhookUrl $Config.DiscordWebhook `
        -Title "🔔 AutoTask 通知測試" `
        -Description "如果您看到這則訊息，代表 Discord 通知設定已成功！" `
        -Color "5814783" `
        -Fields @{ "測試結果" = "成功"; "時間" = (Get-Date).ToString() }
} else {
    Write-Error "找不到 Lib_Discord.ps1，請確認檔案位置。"
}

Write-Host "`n按任意鍵退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
