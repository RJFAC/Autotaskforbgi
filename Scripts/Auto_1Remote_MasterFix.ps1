# ========================================================
# AutoTask 排程器修正腳本 V2 (XML 直接修補版)
# 功能：修正 Master 任務每 4 小時被強制殺死的問題
# ========================================================

$TaskName = "Auto_1Remote_Master"
$ErrorActionPreference = "Stop"

try {
    Write-Host "正在讀取工作排程: $TaskName ..." -ForegroundColor Cyan
    
    # 1. 獲取並導出 XML (避開物件屬性設定的格式地雷)
    $Task = Get-ScheduledTask -TaskName $TaskName
    $TaskXML = $Task | Export-ScheduledTask
    
    # 2. 修正 ExecutionTimeLimit (執行時間限制)
    # 使用正規表達式將 PT4H (4小時) 或其他值替換為 PT0S (無限制)
    if ($TaskXML -match "<ExecutionTimeLimit>.*?</ExecutionTimeLimit>") {
        $TaskXML = $TaskXML -replace "<ExecutionTimeLimit>.*?</ExecutionTimeLimit>", "<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>"
        Write-Host " -> XML 修正: ExecutionTimeLimit => PT0S (已改為無限制)" -ForegroundColor Green
    } else {
        Write-Host " -> ExecutionTimeLimit 未設定 (預設為無限制)，無需修改。" -ForegroundColor Gray
    }

    # 3. 修正 MultipleInstancesPolicy (重複實例策略)
    # 確保為 IgnoreNew，這樣即使 4 小時重複觸發也不會殺死正在跑的腳本
    if ($TaskXML -match "<MultipleInstancesPolicy>.*?</MultipleInstancesPolicy>") {
         $TaskXML = $TaskXML -replace "<MultipleInstancesPolicy>.*?</MultipleInstancesPolicy>", "<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
         Write-Host " -> XML 修正: MultipleInstancesPolicy => IgnoreNew" -ForegroundColor Green
    }

    # 4. 重新註冊任務
    Write-Host "正在套用設定..." -ForegroundColor Cyan
    # 使用原始 Principal 設定覆蓋
    Register-ScheduledTask -TaskName $TaskName -Xml $TaskXML -Force | Out-Null
    
    Write-Host "✅ 修正成功！" -ForegroundColor Green
    Write-Host "   現在 Master 任務將會持續運行，不再受 4 小時限制。"

} catch {
    Write-Host "❌ 修正失敗: $_" -ForegroundColor Red
    Write-Host "詳細錯誤: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "按 Enter 鍵退出..."
Read-Host
