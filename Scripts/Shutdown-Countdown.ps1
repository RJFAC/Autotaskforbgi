<#
.SYNOPSIS
    自動關機倒數計時器 V3 (通知增強版)
    - 支援 GUI 視窗倒數 (預設)
    - 支援背景無頭模式 (Headless Mode)
    - 整合 Discord 通知、Windows Toast 通知、聲音警報
#>

$ErrorActionPreference = "Stop"

# --- [設定] ---
$LogDir       = "C:\AutoTask\Logs"
$LogFile      = Join-Path $LogDir "Shutdown.log"
$NotifyScript = "C:\AutoTask\Scripts\Notify.ps1"
$CountdownSec = 300 # 5 分鐘
$SoundInterval= 30  # 背景模式下每 30 秒嗶一聲

# --- [輔助函數] ---
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Type="INFO")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$TimeStamp][$Type] $Message"
    try { Add-Content -Path $LogFile -Value $Line -Encoding UTF8 -Force } catch {}
    if ($Type -eq "ERROR") { Write-Host $Line -ForegroundColor Red } else { Write-Host $Line -ForegroundColor Cyan }
}

function Send-Discord {
    param($Title, $Msg, $IsEmergency=$false)
    if (Test-Path $NotifyScript) {
        $Color = if ($IsEmergency) { "Red" } else { "Yellow" }
        # 呼叫 Notify.ps1
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$NotifyScript`" -Title `"$Title`" -Message `"$Msg`" -Color `"$Color`" -Mention `$true" -WindowStyle Hidden
    }
}

function Send-Toast {
    param($Title, $Message)
    $code = @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace ToastNotify
{
    public class Toaster
    {
        public static void Show(string title, string message)
        {
            string xml = "<toast><visual><binding template=\"ToastGeneric\"><text>" + title + "</text><text>" + message + "</text></binding></visual><audio src=\"ms-winsoundevent:Notification.Looping.Alarm\" loop=\"false\"/></toast>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            ToastNotification toast = new ToastNotification(doc);
            toast.Tag = "AutoTaskShutdown";
            toast.Group = "AutoTask";
            ToastNotificationManager.CreateToastNotifier("AutoTask System").Show(toast);
        }
    }
}
"@
    try {
        # 嘗試載入 Windows Runtime API 發送通知
        Add-Type -TypeDefinition $code -Language CSharp -ReferencedAssemblies "Windows.Data.Xml.Dom.dll","Windows.UI.Notifications.dll" -ErrorAction SilentlyContinue
        [ToastNotify.Toaster]::Show($Title, $Message)
    } catch {
        Write-Log "Toast 通知發送失敗 (可能不支援此環境): $_" "WARN"
    }
}

function Play-AlertSound {
    try { [System.Media.SystemSounds]::Hand.Play() } catch { [Console]::Beep(1000, 500) }
}

# --- [主邏輯] ---
Write-Log "=== 關機倒數程序啟動 (PID: $PID) ==="

try {
    # 1. 嘗試初始化 GUI
    Write-Log "正在初始化 GUI..."
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "⚠️ 任務完成 - 自動關機"
    $form.Size = New-Object System.Drawing.Size(450, 220)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.ControlBox = $false 
    $form.BackColor = [System.Drawing.Color]::White

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "初始化中..."
    $label.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 14, [System.Drawing.FontStyle]::Bold)
    $label.TextAlign = "MiddleCenter"
    $label.Location = New-Object System.Drawing.Point(20, 30)
    $label.Size = New-Object System.Drawing.Size(400, 60)
    $label.ForeColor = [System.Drawing.Color]::DarkRed
    $form.Controls.Add($label)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "取消關機"
    $button.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
    $button.Location = New-Object System.Drawing.Point(140, 110)
    $button.Size = New-Object System.Drawing.Size(160, 45)
    $button.BackColor = [System.Drawing.Color]::WhiteSmoke
    $form.Controls.Add($button)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000

    # GUI 事件綁定
    $button.Add_Click({
        $script:timer.Stop()
        Write-Log "使用者手動取消關機。" "WARN"
        $script:label.Text = "關機已取消"
        $script:label.ForeColor = "Green"
        $script:button.Enabled = $false
        $form.Refresh()
        Start-Sleep 2
        $script:form.Close()
    })

    $timer.Add_Tick({
        $script:CountdownSec--
        $min = [math]::Floor($script:CountdownSec / 60)
        $sec = $script:CountdownSec % 60
        $timeStr = "{0:00}:{1:00}" -f $min, $sec
        
        $script:label.Text = "任務完成`n將在 $timeStr 後自動關機..."
        
        if ($script:CountdownSec % 30 -eq 0) { Play-AlertSound } # GUI 模式下每 30 秒提醒一次

        if ($script:CountdownSec -le 0) {
            $script:timer.Stop()
            $script:label.Text = "正在關機..."
            $form.Refresh()
            Write-Log "倒數結束 (GUI)，執行關機。"
            Stop-Computer -Force
            $script:form.Close()
        }
    })

    $timer.Start()
    Write-Log "GUI 建立成功，顯示視窗。"
    
    # 即使 GUI 成功，也發送一個 Toast 提醒 (防呆)
    Send-Toast "AutoTask" "任務完成，5 分鐘後自動關機。"
    
    $form.ShowDialog() | Out-Null

} catch {
    # 2. --- [背景模式 (Headless Mode)] ---
    Write-Log "GUI 初始化失敗 (可能無桌面環境)，切換至背景模式。錯誤: $($_.Exception.Message)" "ERROR"
    
    # 發送 Discord 強力警報
    Send-Discord "⚠️ 自動關機警報 (背景模式)" "GUI 介面啟動失敗，系統將在 5 分鐘後強制關機。請檢查遠端連線！" $true
    
    # 發送 Windows 通知
    Send-Toast "⚠️ AutoTask 警報" "GUI 失敗！系統將在 5 分鐘後強制關機！"

    Write-Log "背景倒數開始 ($CountdownSec 秒)..."
    
    # 手動倒數迴圈
    $Remaining = $CountdownSec
    while ($Remaining -gt 0) {
        if ($Remaining % 60 -eq 0) { Write-Log "背景倒數: 剩餘 $($Remaining/60) 分鐘" }
        
        # 聲音警報 (確保開啟聲音)
        Play-AlertSound
        
        Start-Sleep 1
        $Remaining--
        
        # 每 30 秒再發一次 Toast 刷存在感
        if ($Remaining % 30 -eq 0) {
            Send-Toast "AutoTask 關機倒數" "剩餘 $Remaining 秒"
        }
    }

    Write-Log "背景倒數結束，執行強制關機..."
    Send-Discord "系統關機" "AutoTask 已執行強制關機。" $false
    Stop-Computer -Force
}
