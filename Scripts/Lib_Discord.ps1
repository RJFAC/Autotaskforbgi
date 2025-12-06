<#
    .SYNOPSIS
    Discord Webhook 通知模組 (防閃退與日誌增強版)
    
    .DESCRIPTION
    提供 Send-DiscordNotification 函式。
    v2 更新: 
    1. 增加實體檔案日誌 (Discord_Debug.log) 以便事後查閱。
    2. 發生錯誤時強制暫停視窗 (Read-Host)，防止訊息閃退。
#>

# [Fix] 防閃退與日誌機制
# 檢測 Write-Log 是否存在，若不存在(如單獨測試時)，定義一個會寫入檔案的備用函式。
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message)
        $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogContent = "[$Time] $Message"
        
        # 1. 螢幕輸出 (Cyan 色)
        Write-Host $LogContent -ForegroundColor Cyan
        
        # 2. 檔案輸出 (新增: 解決閃退後無紀錄問題)
        # 嘗試寫入到當前目錄下的 Discord_Debug.log，方便事後查看
        try {
            $LogFile = "Discord_Debug.log"
            Add-Content -Path $LogFile -Value $LogContent -ErrorAction SilentlyContinue
        } catch {
            # 忽略檔案寫入錯誤，避免干擾主流程
        }
    }
}

function Send-DiscordNotification {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [string]$Color = "Blue",
        
        [string]$Title = "AutoTask Notification"
    )

    # 檢查設定是否存在
    if (-not $Global:Config.DiscordWebhook) {
        $ErrorMsg = "錯誤: 未在設定檔 (EnvConfig.json) 中找到 DiscordWebhook。"
        Write-Log "[Notify] $ErrorMsg"
        Write-Host $ErrorMsg -ForegroundColor Red
        
        # [Fix] 暫停視窗，讓使用者有機會閱讀錯誤
        Write-Host "`n[AutoTask Debug] 腳本已暫停，請按 Enter 鍵關閉視窗..." -ForegroundColor Yellow
        if ($MyInvocation.InvocationName -ne '.') { Read-Host }
        return
    }

    $WebhookUrl = $Global:Config.DiscordWebhook

    # 顏色對照表
    $ColorMap = @{
        "Blue"   = 3447003
        "Green"  = 5763719
        "Red"    = 15548997
        "Orange" = 15105570
    }

    $ColorCode = if ($ColorMap.ContainsKey($Color)) { $ColorMap[$Color] } else { 3447003 }

    # JSON Payload (含時區修正)
    $Payload = @{
        username = "AutoTask Bot"
        embeds = @(
            @{
                title = $Title
                description = $Message
                color = $ColorCode
                timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            }
        )
    }

    try {
        # 發送請求
        $Response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body ($Payload | ConvertTo-Json -Depth 10) -ErrorAction Stop
        Write-Log "[Notify] Success: Notification sent to Discord."
    }
    Catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log "[Notify] Error: Failed to send notification. Details: $ErrorMsg"
        
        # 詳細錯誤顯示
        Write-Host "--------------------------------------------------" -ForegroundColor Red
        Write-Host "發送失敗，詳細錯誤訊息：" -ForegroundColor Red
        Write-Host $ErrorMsg -ForegroundColor Yellow
        Write-Host "--------------------------------------------------" -ForegroundColor Red
        
        if ($ErrorMsg -match "404") { Write-Host "提示: (404) 找不到網址。請檢查 Webhook URL 是否完整。" -ForegroundColor Gray }
        if ($ErrorMsg -match "401") { Write-Host "提示: (401) 權限不足。Webhook 可能已失效或被刪除。" -ForegroundColor Gray }
        if ($ErrorMsg -match "400") { Write-Host "提示: (400) 格式錯誤。可能是 JSON 結構或內容過長。" -ForegroundColor Gray }

        # [Fix] 錯誤時強制暫停，防止視窗閃退
        # 這樣你就可以看清楚上面的錯誤訊息，或者去檢查 Discord_Debug.log
        Write-Host "`n[AutoTask Debug] 為了讓您閱讀錯誤訊息，腳本已暫停。" -ForegroundColor Green
        Read-Host "請按 Enter 鍵繼續/關閉..."
    }
}

# 測試用區塊
if ($MyInvocation.InvocationName -ne '.') {
    # 這裡可以放置手動測試代碼
}