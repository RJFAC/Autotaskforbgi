<#
.SYNOPSIS
    AutoTask Scheduler Update Tool (V3.34)
    新增或更新每日 04:05 觸發器，配合新的時間視窗。
#>

$TaskName = "Auto_1Remote_Master"
$TargetTime = "04:05"

Write-Host "=== AutoTask 排程器更新工具 ===" -ForegroundColor Cyan

$Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if (-not $Task) {
    Write-Host "❌ 找不到排程任務: $TaskName" -ForegroundColor Red
    exit
}

Write-Host "-> 找到任務: $TaskName" -ForegroundColor Green

# 檢查是否已有 04:05 觸發器
$HasTargetTrigger = $false
foreach ($Trig in $Task.Triggers) {
    # 檢查是否為每日觸發 (TimeTrigger 且 Repetition 間隔為 1天)
    if ($Trig.ToString() -match "MSFT_TaskDailyTrigger") {
        if ($Trig.StartBoundary -match "T$($TargetTime.Replace(':',''))" -or $Trig.StartBoundary -match "T$TargetTime") {
            $HasTargetTrigger = $true
            Write-Host "-> 已存在每日 $TargetTime 的觸發器。" -ForegroundColor Yellow
        }
    }
}

if (-not $HasTargetTrigger) {
    Write-Host "-> 正在新增每日 $TargetTime 觸發器..." -ForegroundColor Cyan
    
    # 建立新的觸發器
    $Trigger = New-ScheduledTaskTrigger -Daily -At $TargetTime
    
    # 將新觸發器加入現有任務
    # 注意: Set-ScheduledTask 接受 Trigger 陣列，我們需要保留舊的觸發器
    $CurrentTriggers = $Task.Triggers
    $NewTriggers = $CurrentTriggers + $Trigger
    
    try {
        Set-ScheduledTask -TaskName $TaskName -Trigger $NewTriggers -ErrorAction Stop
        Write-Host "✅ 更新成功！任務將在每日 $TargetTime 自動執行。" -ForegroundColor Green
    } catch {
        Write-Host "❌ 更新失敗: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "請嘗試以系統管理員身分執行此腳本。" -ForegroundColor Yellow
    }
} else {
    Write-Host "✅ 無需變更。" -ForegroundColor Green
}

Read-Host "按 Enter 鍵結束..."