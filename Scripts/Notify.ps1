# ============================================================
# AutoTask Discord 通知模組 V2.2 (Config Fix)
# ============================================================
param (
    [string]$Title = "通知",
    [string]$Message = "",
    [string]$Color = "Green", 
    [hashtable]$Fields = @{}, 
    [string]$LogPath = "",    
    [bool]$Mention = $false   
)

# --- [Webhook 讀取邏輯修正] ---
# 優先順序: EnvConfig.json > Webhook.url
$WebhookUrl = ""
$ConfigDir = "C:\AutoTask\Configs"
$EnvFile = "$ConfigDir\EnvConfig.json"
$UrlFile = "$ConfigDir\Webhook.url"

# 1. 嘗試讀取 EnvConfig.json (新版)
if (Test-Path $EnvFile) {
    try {
        $Json = Get-Content $EnvFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($Json.DiscordWebhook) { $WebhookUrl = $Json.DiscordWebhook }
    } catch {}
}

# 2. 如果沒找到，嘗試讀取 Webhook.url (舊版)
if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
    if (Test-Path $UrlFile) {
        $WebhookUrl = (Get-Content $UrlFile -Raw).Trim()
    }
}

# 如果還是空的，就結束
if ([string]::IsNullOrWhiteSpace($WebhookUrl)) { 
    # 可選：寫入錯誤日誌方便除錯
    # Add-Content "C:\AutoTask\Logs\Notify_Debug.log" "No Webhook URL found."
    return 
}

# --- [顏色代碼] ---
$ColorCode = switch ($Color) {
    "Green"  { 5763719 }  # 綠色 (成功)
    "Red"    { 15548997 } # 紅色 (失敗)
    "Yellow" { 16776960 } # 黃色 (警告)
    "Blue"   { 3447003 }  # 藍色 (資訊)
    Default  { 3447003 }
}

# --- [內容處理] ---
$Description = $Message

if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath)) {
    try {
        $LogContent = Get-Content $LogPath -Tail 20 -Encoding UTF8 -ErrorAction Stop
        if ($LogContent) {
            $LogBlock = $LogContent -join "`n"
            if ($LogBlock.Length -gt 1000) { $LogBlock = $LogBlock.Substring($LogBlock.Length - 1000) }
            $Description += "`n`n**📋 Log 摘要:**`n```text`n$LogBlock`n```"
        }
    } catch {
        $Description += "`n`n⚠️ *無法讀取 Log: $_*"
    }
}

$ContentText = ""
if ($Mention -and ($Color -eq "Red" -or $Color -eq "Yellow")) {
    $ContentText = "@everyone"
}

$EmbedFields = @()
foreach ($key in $Fields.Keys) {
    $Val = $Fields[$key]
    if ([string]::IsNullOrWhiteSpace($Val)) { $Val = "N/A" }
    $EmbedFields += @{ name = $key; value = $Val; inline = $false }
}

$Payload = @{
    username = "AutoTask Bot"
    content = $ContentText
    embeds = @(
        @{
            title = $Title
            description = $Description
            color = $ColorCode
            fields = $EmbedFields
            footer = @{ text = "來自: $env:COMPUTERNAME | $(Get-Date -Format 'HH:mm:ss')" }
            timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        }
    )
}

try {
    $JsonPayload = $Payload | ConvertTo-Json -Depth 5 -Compress
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $JsonPayload -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
} catch {
    # 失敗不做任何事，避免卡死腳本
}