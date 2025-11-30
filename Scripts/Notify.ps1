# ============================================================
# AutoTask Discord 通知模組 V2.1 (診斷增強版)
# ============================================================
param (
    [string]$Title = "通知",
    [string]$Message = "",
    [string]$Color = "Green", # Green, Red, Yellow, Blue
    [hashtable]$Fields = @{}, # key=value 的詳細欄位
    [string]$LogPath = "",    # [新] 指定 Log 檔案路徑，自動擷取末尾內容
    [bool]$Mention = $false   # [新] 是否 Tag @everyone
)

$ConfigDir = "C:\AutoTask\Configs"
$UrlFile = "$ConfigDir\Webhook.url"

if (-not (Test-Path $UrlFile)) { return }
$WebhookUrl = (Get-Content $UrlFile -Raw).Trim()
if ([string]::IsNullOrWhiteSpace($WebhookUrl)) { return }

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

# [診斷增強] 若有指定 Log 檔案，讀取最後 20 行並附加
if (-not [string]::IsNullOrWhiteSpace($LogPath) -and (Test-Path $LogPath)) {
    try {
        $LogContent = Get-Content $LogPath -Tail 20 -Encoding UTF8 -ErrorAction Stop
        if ($LogContent) {
            $LogBlock = $LogContent -join "`n"
            # 截斷以防超過 Discord 限制 (4096 chars)
            if ($LogBlock.Length -gt 1500) { $LogBlock = $LogBlock.Substring($LogBlock.Length - 1500) }
            $Description += "`n`n**📋 Log 摘要 (Last 20 lines):**`n```text`n$LogBlock`n```"
        }
    } catch {
        $Description += "`n`n⚠️ *無法讀取 Log 檔案: $_*"
    }
}

# --- [提及設定] ---
$ContentText = ""
if ($Mention -and ($Color -eq "Red" -or $Color -eq "Yellow")) {
    $ContentText = "@everyone ⚠️ 偵測到異常！"
}

# --- [構建 Embed] ---
$EmbedFields = @()
foreach ($key in $Fields.Keys) {
    $Val = $Fields[$key]
    if ([string]::IsNullOrWhiteSpace($Val)) { $Val = "N/A" }
    $EmbedFields += @{
        name = $key
        value = $Val
        inline = $false
    }
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
            footer = @{
                text = "來自: $env:COMPUTERNAME | 時間: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            }
        }
    )
}

# --- [發送請求] ---
try {
    # 強制 UTF-8 轉換
    $JsonPayload = $Payload | ConvertTo-Json -Depth 4 -Compress
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $JsonPayload -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
} catch {
    # 本地記錄發送失敗，避免遞迴錯誤
    $ErrLog = "C:\AutoTask\Logs\Notify_Error.log"
    Add-Content -Path $ErrLog -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Notify Send Failed: $_" -Force
}