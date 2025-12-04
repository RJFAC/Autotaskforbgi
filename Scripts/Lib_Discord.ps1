# =======================================================
# 檔案名稱: Lib_Discord.ps1
# 功能: Discord 通知模組 (v2.5 編碼修復版)
# =======================================================

# 強制設定控制台輸出編碼為 UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-EnvConfig {
    param($Key)
    $ConfigPath = "$PSScriptRoot\..\Configs\EnvConfig.json"
    if (Test-Path $ConfigPath) {
        try {
            $Json = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            return $Json.$Key
        } catch {}
    }
    return $null
}

function Send-DiscordWebhook {
    param(
        [string]$WebhookUrl,
        [string]$Title,
        [string]$Description,
        [string]$Color,
        [hashtable]$Fields
    )

    if ([string]::IsNullOrWhiteSpace($WebhookUrl)) { return }

    $EmbedFields = @()
    if ($Fields) {
        foreach ($key in $Fields.Keys) {
            $EmbedFields += @{ name = $key; value = $Fields[$key]; inline = $true }
        }
    }

    $Payload = @{
        username = "AutoTask Bot"
        embeds = @(
            @{
                title = $Title
                description = $Description
                color = $Color
                fields = $EmbedFields
                footer = @{ text = "AutoTask | Host: $env:COMPUTERNAME" }
                timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            }
        )
    }

    try {
        # [關鍵修正] 明確指定 ContentType 為 charset=utf-8
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json; charset=utf-8' -Body ($Payload | ConvertTo-Json -Depth 10 -Compress)
    } catch {
        Write-Warning "Discord 發送失敗: $_"
    }
}

function Send-AutoTaskReport {
    param([string]$Status, [string]$LogFile)

    $WebhookUrl = Get-EnvConfig -Key "DiscordWebhook"
    if (-not $WebhookUrl) { Write-Host "Discord Webhook 未設定，跳過通知。" -ForegroundColor Yellow; return }

    if ($Status -eq "Success") {
        $Color = "5763719" # Green
        $Title = "✅ AutoTask 任務執行成功"
    } else {
        $Color = "15548997" # Red
        $Title = "❌ AutoTask 任務執行失敗"
    }

    $LogSummary = "無日誌"
    $DurationText = "未知"

    if ($LogFile -and (Test-Path $LogFile)) {
        # 讀取摘要
        $Logs = Get-Content $LogFile -Tail 50 -Encoding UTF8
        $LogSummary = ($Logs | Where-Object { $_ -match "\S" } | Select-Object -Last 5) -join "`n"
        
        # 計算耗時
        try {
            $FullLog = Get-Content $LogFile -Encoding UTF8
            if ($FullLog.Count -ge 2) {
                $Start = $null; $End = $null
                
                if ($FullLog[0] -match "\[(.*?)\]") { $Start = [DateTime]::Parse($matches[1]) }
                if ($FullLog[-1] -match "\[(.*?)\]") { $End = [DateTime]::Parse($matches[1]) }
                
                if ($Start -and $End) {
                    $Duration = $End - $Start
                    $DurationText = "{0:00}:{1:00}:{2:00}" -f $Duration.Hours, $Duration.Minutes, $Duration.Seconds
                }
            }
        } catch {
            $DurationText = "計算錯誤"
        }
    }

    $Fields = [ordered]@{
        "⏱️ 耗時" = $DurationText
        # 使用 24 小時制 (HH) 避免出現上午/下午的中文字元
        "📅 時間" = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    }

    $SafeDescription = '```text' + "`n" + $LogSummary + "`n" + '```'

    Send-DiscordWebhook -WebhookUrl $WebhookUrl -Title $Title -Description $SafeDescription -Color $Color -Fields $Fields
}
