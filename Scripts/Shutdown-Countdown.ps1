<#
.SYNOPSIS
    自動關機倒數計時器 (5 分鐘)
    - 彈出一個置頂 (TopMost) 視窗
    - 顯示倒數秒數
    - 提供「取消」按鈕
    - 倒數結束時，強制關機
#>

# --- [定義路徑] ---
$LogDir = "C:\AutoTask\Logs"
$ErrorLog = Join-Path $LogDir "Shutdown_Error.log"

try {
    # 1. --- [載入 Windows Forms 元件] ---
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # 2. --- [全域變數] ---
    $CountdownSeconds = 300 # 5 分鐘 * 60 秒

    # 3. --- [建立表單 (Form)] ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "自動化任務完成"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true # [關鍵] 置頂顯示
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false # [優化] 移除右上角 X 按鈕，強制使用者點擊取消

    # 4. --- [建立標籤 (Label)] ---
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "將在 $CountdownSeconds 秒後自動關機..."
    $label.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 14, [System.Drawing.FontStyle]::Bold)
    $label.TextAlign = "MiddleCenter" # [優化] 文字置中
    $label.Location = New-Object System.Drawing.Point(10, 30)
    $label.Size = New-Object System.Drawing.Size(360, 50)
    $form.Controls.Add($label)

    # 5. --- [建立按鈕 (Button)] ---
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "取消關機"
    $button.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
    $button.Location = New-Object System.Drawing.Point(120, 100)
    $button.Size = New-Object System.Drawing.Size(150, 40)
    $button.BackColor = [System.Drawing.Color]::WhiteSmoke

    # 6. --- [建立計時器 (Timer)] ---
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # 1000 毫秒 = 1 秒

    # 7. --- [定義事件動作] ---

    # (A) 按鈕點擊事件
    $button.Add_Click({
        $script:timer.Stop() # 停止計時
        $script:label.Text = "關機已取消。"
        $script:label.ForeColor = "Green"
        $script:button.Enabled = $false
        $script:button.Text = "已取消"
        
        # [關鍵修正] 強制刷新 UI，讓使用者看到文字變化
        $form.Refresh()
        
        # 短暫停留 2 秒讓使用者看到訊息，然後關閉
        Start-Sleep -Seconds 2
        $script:form.Close()
    })

    # (B) 計時器 Tick 事件 (每秒觸發一次)
    $timer.Add_Tick({
        $script:CountdownSeconds--
        
        # 倒數顯示邏輯
        $min = [math]::Floor($script:CountdownSeconds / 60)
        $sec = $script:CountdownSeconds % 60
        $timeStr = "{0:00}:{1:00}" -f $min, $sec
        
        $script:label.Text = "任務完成，將在 $timeStr 後關機..."
        
        # 最後 10 秒變紅色警示
        if ($script:CountdownSeconds -le 10) {
            $script:label.ForeColor = "Red"
        }

        if ($script:CountdownSeconds -le 0) {
            $script:timer.Stop()
            $script:label.Text = "正在關機..."
            $form.Refresh()
            
            # [關鍵] 執行強制關機
            Stop-Computer -Force
            $script:form.Close()
        }
    })

    # 8. --- [啟動] ---
    $form.Controls.Add($button)
    $timer.Start()
    $form.ShowDialog()

} catch {
    # 發生錯誤時，記錄下來
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Msg = "[$TimeStamp] GUI Error: $($_.Exception.Message)"
    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }
    Out-File -FilePath $ErrorLog -InputObject $Msg -Append -Encoding UTF8
}