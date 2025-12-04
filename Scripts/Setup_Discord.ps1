$ConfigPath = "C:\AutoTask\Configs\EnvConfig.json"
$LibPath = "$PSScriptRoot\Lib_Discord.ps1"

if (-not (Test-Path $ConfigPath)) { Write-Error "找不到設定檔！"; exit }
if (-not (Test-Path $LibPath)) { Write-Error "找不到 Lib_Discord.ps1"; exit }

. $LibPath
$Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host "=== Discord 設定 ===" -ForegroundColor Cyan
$Url = Read-Host "請輸入 Webhook URL (留空不修改)"
if ($Url) {
    if ($Config.PSObject.Properties.Match('DiscordWebhook').Count -eq 0) {
        $Config | Add-Member -Name "DiscordWebhook" -Value $Url -MemberType NoteProperty
    } else {
        $Config.DiscordWebhook = $Url
    }
    $Config | ConvertTo-Json -Depth 4 | Set-Content $ConfigPath -Encoding UTF8
    Write-Host "設定已儲存。" -ForegroundColor Green
    
    # 測試發送
    Send-DiscordWebhook -WebhookUrl $Url -Title "🔔 測試通知" -Description "設定成功！" -Color "5814783"
}