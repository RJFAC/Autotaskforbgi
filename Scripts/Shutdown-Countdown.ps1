<#
.SYNOPSIS
    自動關機倒數計時器 V4 (Lib_Discord 整合版)
    - 改用 Lib_Discord.ps1 發送通知，移除對舊版 Notify.ps1 的依賴。
    - 支援 GUI 視窗倒數 (預設) 與 Headless 模式。
#>

$ErrorActionPreference = "Stop"

# --- [設定] ---
$BaseDir      = "C:\AutoTask"
$LogDir       = "$BaseDir\Logs"
$LogFile      = Join-Path $LogDir "Shutdown.log"
$LibDiscord   = "$BaseDir\Scripts\Lib_Discord.ps1"  # 指向新版 Lib
$CountdownSec = 300 # 5 分鐘
$SoundInterval= 30  # 背景模式下每 30 秒嗶一聲

# --- [載入 Lib] ---
if (Test-Path $LibDiscord) {
    . $LibDiscord
} else {
    Write-Warning "找不到 Lib_Discord.ps1，Discord 通知將失效。"
    function Send-DiscordNotification { param($Message, $Title, $Color) Write-Host "Mock Notify: $Title - $Message" }
}

# --- [輔助函數] ---
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Type="INFO")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$TimeStamp][$Type] $Message"
    try { Add-Content -Path $LogFile -Value $Line -Encoding UTF8 -Force } catch {}
    if ($Type -eq "ERROR") { Write-Host $Line -ForegroundColor Red } else { Write-Host $Line -ForegroundColor Cyan }
}

function Send-Toast {
    param([string]$Title, [string]$Message)
    $code = @"
    Windows.Data.Xml.Dom.XmlDocument toastXml = Windows.UI.Notifications.ToastNotificationManager.GetTemplateContent(Windows.UI.Notifications.ToastTemplateType.ToastImageAndText02);
    Windows.Data.Xml.Dom.XmlNodeList stringElements = toastXml.GetElementsByTagName("text");
    stringElements.Item(0).AppendChild(toastXml.CreateTextNode("$Title"));
    stringElements.Item(1).AppendChild(toastXml.CreateTextNode("$Message"));
    Windows.UI.Notifications.ToastNotification toast = new Windows.UI.Notifications.ToastNotification(toastXml);
    Windows.UI.Notifications.ToastNotificationManager.CreateToastNotifier("AutoTask").Show(toast);
"@
    try {
        if (-not ([System.Management.Automation.PSTypeName]'WinRT.Toast').Type) {
            Add-Type -TypeDefinition "using System; using Windows.UI.Notifications; using Windows.Data.Xml.Dom; public class WinRT { public static void Toast() {} }" -ErrorAction SilentlyContinue
        }
        # PowerShell 7+ Toast 支援較複雜，此處為簡易相容嘗試，若失敗則忽略
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
        $xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText02)
        $text = $xml.GetElementsByTagName("text")
        $text[0].AppendChild($xml.CreateTextNode($Title)) | Out-Null
        $text[1].AppendChild($xml.CreateTextNode($Message)) | Out-Null
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("AutoTask").Show([Windows.UI.Notifications.ToastNotification]::new($xml))
    } catch {
        Write-Log "Toast 通知發送失敗 (可能不支援): $($_.Exception.Message)" "WARN"
    }
}

function Play-AlertSound {
    [System.Console]::Beep(1000, 500)
    [System.Console]::Beep(1500, 500)
}

# --- [主邏輯] ---
Write-Log "=== 啟動關機倒數程序 ($CountdownSec 秒) ==="

try {
    # 1. --- [GUI 模式] ---
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "AutoTask - 任務完成"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.Color]::Black
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "任務已完成`n系統將在 $CountdownSec 秒後關機"
    $label.Font = New-Object System.Drawing.Font("Consolas", 14, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::Cyan
    $label.TextAlign = "MiddleCenter"
    $label.Dock = "Top"
    $label.Height = 100

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消關機"
    $btnCancel.Font = New-Object System.Drawing.Font("Microsoft JhengHei", 12)
    $btnCancel.Size = New-Object System.Drawing.Size(150, 50)
    $btnCancel.Location = New-Object System.Drawing.Point(115, 130)
    $btnCancel.BackColor = [System.Drawing.Color]::DarkRed
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.Add_Click({
        $script:timer.Stop()
        Write-Log "使用者取消關機。"
        Send-DiscordNotification -Title "🛑 關機已取消" -Message "使用者在倒數期間手動取消了關機程序。" -Color "Orange"
        $form.Close()
    })

    $form.Controls.Add($label)
    $form.Controls.Add($btnCancel)

    # 計時器
    $script:remaining = $CountdownSec
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $script:remaining--
        $script:label.Text = "任務已完成`n系統將在 $script:remaining 秒後關機"
        
        if ($script:remaining % 30 -eq 0) { Play-AlertSound }

        if ($script:remaining -le 0) {
            $script:timer.Stop()
            $script:label.Text = "正在關機..."
            $form.Refresh()
            
            Write-Log "倒數結束，執行關機。"
            Send-DiscordNotification -Title "🔌 系統關機" -Message "AutoTask 任務完成，系統自動關機。" -Color "Green"
            
            Stop-Computer -Force
            $form.Close()
        }
    })

    $timer.Start()
    Write-Log "GUI 介面啟動成功。"
    Send-Toast "AutoTask" "任務完成，5 分鐘後自動關機。"
    
    $form.ShowDialog() | Out-Null

} catch {
    # 2. --- [Headless 模式 (Fallback)] ---
    Write-Log "GUI 初始化失敗，切換至背景模式: $($_.Exception.Message)" "WARN"
    
    Send-DiscordNotification -Title "⚠️ 自動關機倒數 (背景)" -Message "GUI 啟動失敗。系統將在 5 分鐘後關機。請檢查遠端連線！" -Color "Yellow"
    
    for ($i = $CountdownSec; $i -gt 0; $i--) {
        if ($i % $SoundInterval -eq 0) { Play-AlertSound }
        Start-Sleep 1
    }
    
    Write-Log "倒數結束 (Headless)，執行關機。"
    Send-DiscordNotification -Title "🔌 系統關機" -Message "AutoTask (背景模式) 執行關機。" -Color "Green"
    Stop-Computer -Force
}