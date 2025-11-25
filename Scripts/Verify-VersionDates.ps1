# ============================================================
# 原神版本更新與紊亂爆發期 (幽境危戰) 預測工具
# ============================================================

# 基準日: 5.0 版本更新日 (2024-08-28)
$BaseDate = [datetime]"2024-08-28"
$CurrentVerMajor = 5
$CurrentVerMinor = 0

Write-Host "=== 原神版本日程預測 (5.0 ~ 7.0) ===" -ForegroundColor Cyan
Write-Host "基準: 5.0 (2024/08/28), 週期: 42天"
Write-Host "規則: 紊亂爆發期為版本第8天(週三) 10:00 ~ 第18天(週六) 03:59"
Write-Host ""
Write-Host "版本  | 更新日 (週三) | 紊亂開始 (週三) | 紊亂結束 (週六) | 備註"
Write-Host "------+---------------+-----------------+-----------------+----------------"

# 預測未來 18 個版本
for ($i = 0; $i -le 18; $i++) {
    $VerDate = $BaseDate.AddDays($i * 42)
    $VerStr = "$CurrentVerMajor.$CurrentVerMinor"
    
    # 計算紊亂爆發期
    # 第8天 (Update + 7 days)
    $TurbStart = $VerDate.AddDays(7)
    # 第18天 (Update + 17 days)
    $TurbEnd   = $VerDate.AddDays(17)

    # 格式化輸出
    $D1 = $VerDate.ToString("yyyy/MM/dd")
    $D2 = $TurbStart.ToString("MM/dd")
    $D3 = $TurbEnd.ToString("MM/dd")
    
    $Note = ""
    if ($VerStr -eq "5.2") { $Note = "<-- 目前版本?" }
    if ($VerStr -eq "6.0") { $Note = "納塔完結/至冬?" }

    Write-Host "$VerStr   | $D1    | $D2 (10:00)     | $D3 (03:59)     | $Note"

    # 版本號遞增 (假設 x.0 ~ x.8)
    $CurrentVerMinor++
    if ($CurrentVerMinor -gt 8) { 
        $CurrentVerMajor++
        $CurrentVerMinor = 0 
    }
}
Write-Host ""
Write-Host "請確認上述日期是否與官方公告或您的經驗一致。" -ForegroundColor Yellow
pause