<#
    .SYNOPSIS
    Discord Webhook 通知模組 (V2.2 Stable)
    .DESCRIPTION
    修復字串轉義問題，採用拼接方式處理 Markdown 語法。
#>

# [Log] 定義日誌輸出函式
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message)
        $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$Time] $Message" -ForegroundColor Cyan
    }
}

function Send-DiscordNotification {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Color = "Blue",
        [string]$Title = "AutoTask Notification"
    )

    # 1. 自動載入設定檔
    if (-not $Global:Config -or -not $Global:Config.DiscordWebhook) {
        $ConfigPaths = @("C:\AutoTask\Configs\EnvConfig.json", "$PSScriptRoot\..\Configs\EnvConfig.json")
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

    # 2. 檢查 Webhook
    if (-not $Global:Config.DiscordWebhook) {
        Write-Host "錯誤: 找不到 Webhook 設定 (EnvConfig.json)，無法發送通知。" -ForegroundColor Red
        return
    }

    $WebhookUrl = $Global:Config.DiscordWebhook
    $ColorMap = @{ "Blue"=3447003; "Green"=5763719; "Red"=15548997; "Orange"=15105570 }
    $ColorCode = if ($ColorMap.ContainsKey($Color)) { $ColorMap[$Color] } else { 3447003 }

    # 3. 建構 Payload
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
        $JsonBody = $Payload | ConvertTo-Json -Depth 10
        $null = Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json; charset=utf-8' -Body $JsonBody -ErrorAction Stop
        Write-Host "[Notify] 通知已發送: $Title" -ForegroundColor Green
    }
    Catch {
        Write-Host "[Notify] 發送失敗: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# [新增] 兼容 Master.ps1 的報告函數
function Send-AutoTaskReport {
    param(
        [string]$Status,
        [string]$LogFile
    )
    
    $Title = "任務報告: $Status"
    $Color = "Blue"
    $Msg = "系統狀態更新。"

    if ($Status -eq "Success") {
        $Title = "✅ 任務執行成功"
        $Color = "Green"
        $Msg = "所有排程任務已完成，系統準備進入休眠/關機流程。"
    } elseif ($Status -eq "Error") {
        $Title = "❌ 任務執行失敗"
        $Color = "Red"
        $Msg = "偵測到嚴重錯誤，請檢查主機日誌。"
    }

    # 嘗試讀取日誌最後幾行
    if (Test-Path $LogFile) {
        try {
            $LogContent = Get-Content $LogFile -Tail 5 -Encoding UTF8
            $LogText = $LogContent -join [Environment]::NewLine
            
            # [修正重點] 使用單引號來處理 Markdown 的 ``` 符號，避免 PowerShell 轉義錯誤
            # 這樣寫 PowerShell 絕對不會誤判
            $Msg += "`n`n**📋 Master Log:**`n" + '```text' + "`n$LogText`n" + '```'
            
        } catch {}
    }

    Send-DiscordNotification -Title $Title -Message $Msg -Color $Color
}