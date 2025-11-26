# =============================================================================
# AutoTask Master V4.8 - 腳本監督版 (Payload 復活機制)
# =============================================================================

# ... (前段權限檢查、路徑定義、輔助函數保持不變) ...
# (請複製 V4.7 的前段代碼)

# --- [Master 監督迴圈 (增強版)] ---
$SupervisorStart = Get-Date
Write-Log ">>> Master 監督模式已啟動" "Green"

$PayloadLaunched = $false

while ($true) {
    Start-Sleep 5

    # 1. 檢查成功
    if (Test-Path $DoneFlag) {
        Write-Log "偵測到任務成功 (Done.flag)！" "Green"
        break
    }

    # 2. 檢查失敗
    if (Test-Path $FailFlag) {
        # ... (失敗處理保持不變) ...
        exit
    }

    # 3. [核心新增] 監控 Payload 存活狀態
    $PayloadProc = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*Payload.ps1*" }
    
    if (-not $PayloadProc) {
        # Payload 不見了，檢查是否應該重啟
        # 如果 $PayloadLaunched 為真 (代表我們啟動過它)，且沒有 Done/Fail 旗標
        if ($PayloadLaunched) {
            Write-Log "⚠️ 警報：Payload 腳本意外消失！正在嘗試重啟..." "Red"
            # 再次觸發 Remote 任務
            schtasks /run /tn "Auto_BetterGI_Payload"
            Start-Sleep 10
        } else {
            # 尚未啟動過，或者是剛開始
            # 檢查是否超時未啟動
            if ((Get-Date) -gt $SupervisorStart.AddMinutes(15)) {
                 Write-Log "警告：Payload 啟動超時 (15分)，嘗試重連 RDP..." "Yellow"
                 Start-Process -FilePath $1RemoteExe -ArgumentList "-r Remote" -WorkingDirectory $1RemoteDir
                 $SupervisorStart = Get-Date
            }
        }
    } else {
        # Payload 活著
        $PayloadLaunched = $true
        # 更新計時器，防止因長時間執行被誤判為啟動超時
        $SupervisorStart = Get-Date 
    }
}

# ... (後續清理邏輯保持不變) ...