# ========================================================
# AutoTask 排程器修正腳本 - Payload 專用版
# 功能：移除 Auto_BetterGI_Payload 的 4 小時強制執行限制
# ========================================================

$TaskName = "Auto_BetterGI_Payload"
$ErrorActionPreference = "Stop"

try {
    Write-Host "正在讀取工作排程: $TaskName ..." -ForegroundColor Cyan
    
    # 1. 檢查任務是否存在
    $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    
    # 2. 獲取 XML
    $TaskXML = $Task | Export-ScheduledTask
    
    # 3. 修正 ExecutionTimeLimit (執行時間限制)
    # 將 PT4H (4小時) 替換為 PT0S (無限制)
    if ($TaskXML -match "<ExecutionTimeLimit>PT4H</ExecutionTimeLimit>") {
        $TaskXML = $TaskXML -replace "<ExecutionTimeLimit>PT4H</ExecutionTimeLimit>", "<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>"
        Write-Host " -> XML 修正: 發現 4 小時限制，已移除 (ExecutionTimeLimit => PT0S)" -ForegroundColor Green
    } elseif ($TaskXML -match "<ExecutionTimeLimit>.*?</ExecutionTimeLimit>") {
        $TaskXML = $TaskXML -replace "<ExecutionTimeLimit>.*?</ExecutionTimeLimit>", "<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>"
        Write-Host " -> XML 修正: 強制將時間限制改為無限制 (PT0S)" -ForegroundColor Green
    } else {
        Write-Host " -> 未發現時間限制設定，無需修改。" -ForegroundColor Gray
    }

    # 4. 重新註冊任務
    Write-Host "正在套用設定..." -ForegroundColor Cyan
    Register-ScheduledTask -TaskName $TaskName -Xml $TaskXML -Force | Out-Null
    
    Write-Host "✅ Payload 修正成功！" -ForegroundColor Green
    Write-Host "   現在 Payload 任務將不會在 07:55 被強制結束。"

} catch {
    Write-Host "❌ 修正失敗: $_" -ForegroundColor Red
    Write-Host "詳細錯誤: $($_.Exception.Message)" -ForegroundColor Yellow
    if ($_.Exception.Message -match "存取被拒") {
        Write-Host "⚠️ 請嘗試以「系統管理員身分」執行此腳本！" -ForegroundColor Magenta
    }
}

Write-Host "按 Enter 鍵退出..."
Read-Host
