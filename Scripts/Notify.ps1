# ============================================================
# AutoTask Discord 通知模組 V2.0 (支援 Embed Fields)
# ============================================================
param (
    [string]$Title = "通知",
    [string]$Message = "",
    [string]$Color = "Green", # Green, Red, Yellow, Blue
    [hashtable]$Fields = @{}  # [新功能] 支援 key=value 的詳細欄位
)

$ConfigDir = "C:\AutoTask\Configs"
$UrlFile = "$ConfigDir\Webhook.url"

if (-not (Test-Path $UrlFile)) { return }
$WebhookUrl = (Get-Content $UrlFile -Raw).Trim()
if ([string]::IsNullOrWhiteSpace($WebhookUrl)) { return }

$ColorCode = switch ($Color) {
    "Green"  { 5763719 }  # 綠色 (成功)
    "Red"    { 15548997 } # 紅色 (失敗)
    "Yellow" { 16776960 } # 黃色 (警告)
    "Blue"   { 3447003 }  # 藍色 (資訊)
    Default  { 3447003 }
}

# 構建 Embed Fields
$EmbedFields = @()
foreach ($key in $Fields.Keys) {
    $EmbedFields += @{
        name = $key
        value = $Fields[$key]
        inline = $false
    }
}

$Payload = @{
    username = "AutoTask Bot"
    embeds = @(
        @{
            title = $Title
            description = $Message
            color = $ColorCode
            fields = $EmbedFields
            footer = @{
                text = "來自: $env:COMPUTERNAME | 時間: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            }
        }
    )
}

# 強制 UTF-8 轉換
$JsonPayload = $Payload | ConvertTo-Json -Depth 4 -Compress
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $JsonPayload -ContentType 'application/json; charset=utf-8'