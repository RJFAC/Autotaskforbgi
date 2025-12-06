<#
    .SYNOPSIS
    Discord Webhook 通知模組 (Final Fix)
    
    .DESCRIPTION
    已驗證：
    1. 包含 ToUniversalTime() 修復，解決 Discord 顯示未來時間的問題。
    2. 包含 Dashboard 測試時自動載入 Config 的功能。
    3. 包含防閃退與日誌功能。
#>

# [Log] 定義日誌輸出函式 (防止無主程式時報錯)
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message)
        $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogContent = "[$Time] $Message"
        Write-Host $LogContent -ForegroundColor Cyan
        try {
            Add-Content -Path "Discord_Debug.log" -Value $LogContent -ErrorAction SilentlyContinue
        } catch {}
    }
}

function Send-DiscordNotification {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "Blue",
        [string]$Title = "AutoTask Notification"
    )

    # [Fix] 1. 自動載入設定檔 (針對 Dashboard 測試環境)
    if (-not $Global:Config -or -not $Global:Config.DiscordWebhook) {
        $ConfigPaths = @(
            "C:\AutoTask\Configs\EnvConfig.json",
            "$PSScriptRoot\..\Configs\EnvConfig.json"
        )

        foreach ($Path in $ConfigPaths) {
            if (Test-Path $Path) {
                try {
                    $JsonContent = Get-Content -Path $Path -Raw -Encoding UTF8
                    $Global:Config = $JsonContent | ConvertFrom-Json
                    break
                } catch {}
            }
        }
    }

    # [Fix] 2. 檢查 Webhook
    if (-not $Global:Config.DiscordWebhook) {
        $ErrorMsg = "錯誤: 找不到 Webhook 設定，無法發送通知。"
        Write-Log $ErrorMsg
        Write-Host $ErrorMsg -ForegroundColor Red
        if ($MyInvocation.InvocationName -ne '.') { Read-Host "請按 Enter 鍵繼續..." }
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

    # [Fix] 3. 建構 Payload (關鍵修復：轉為 UTC 時間)
    # 這就是剛剛測試 B 成功的關鍵邏輯
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
        $Response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body ($Payload | ConvertTo-Json -Depth 10) -ErrorAction Stop
        Write-Log "[Notify] Success: Notification sent."
    }
    Catch {
        $ErrorMsg = $_.Exception.Message
        Write-Log "[Notify] Error: $ErrorMsg"
        Write-Host "發送失敗: $ErrorMsg" -ForegroundColor Red
        
        # 錯誤時暫停，方便除錯
        if ($MyInvocation.InvocationName -ne '.') { Read-Host "請按 Enter 鍵繼續..." }
    }
}