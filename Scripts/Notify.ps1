# ============================================================
# AutoTask Discord 通知模組 (Notify.ps1)
# ============================================================
param (
    [string]$Title = "通知",
    [string]$Message = "",
    [string]$Color = "Green" # 支援 Green, Red, Yellow
)

$ConfigDir = "C:\AutoTask\Configs"
$UrlFile = "$ConfigDir\Webhook.url"

# 1. 檢查設定檔是否存在
if (-not (Test-Path $UrlFile)) {
    Write-Host "錯誤：找不到 Webhook 設定檔 ($UrlFile)" -ForegroundColor Red
    return
}

$WebhookUrl = Get-Content $UrlFile -Raw
$WebhookUrl = $WebhookUrl.Trim()

if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
    Write-Host "錯誤：Webhook 網址為空" -ForegroundColor Red
    return
}

# 2. 設定顏色代碼 (Decimal)
$ColorCode = switch ($Color) {
    "Green"  { 5763719 }  # 綠色
    "Red"    { 15548997 } # 紅色
    "Yellow" { 16776960 } # 黃色
    Default  { 3447003 }  # 藍色
}

# 3. 組合 JSON Payload
$Payload = @{
    username = "AutoTask Bot"
    embeds = @(
        @{
            title = $Title
            description = $Message
            color = $ColorCode
            footer = @{
                text = "來自: $env:COMPUTERNAME | 時間: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            }
        }
    )
}

# 強制轉換為 UTF-8 字串以避免亂碼
$JsonPayload = $Payload | ConvertTo-Json -Depth 4

# 4. 發送請求
try {
    # [修正] 加入 charset=utf-8
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $JsonPayload -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
    Write-Host "通知已發送至 Discord。" -ForegroundColor Green
} catch {
    Write-Host "發送通知失敗: $_" -ForegroundColor Red
}