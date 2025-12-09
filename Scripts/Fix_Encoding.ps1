$TargetDir = "C:\AutoTask\Scripts"

if (Test-Path $TargetDir) {
    Write-Host "🔍 正在掃描目錄: $TargetDir" -ForegroundColor Cyan
    
    # 取得目錄下所有 .ps1 檔案
    $Files = Get-ChildItem -Path $TargetDir -Filter "*.ps1"
    
    # 建立 UTF-8 (No BOM) 編碼物件
    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False

    foreach ($File in $Files) {
        $FilePath = $File.FullName
        Write-Host "檢查檔案: $($File.Name)" -NoNewline
        
        try {
            # 讀取內容 (Get-Content 自動處理 BOM)
            $Content = Get-Content -Path $FilePath -Raw
            
            # 使用 .NET 強制寫入為 No BOM 格式
            # 注意: 必須使用 [Class]::Method 語法
            [System.IO.File]::WriteAllText($FilePath, $Content, $Utf8NoBomEncoding)
            
            Write-Host " -> [OK] 已修正 (UTF-8 NoBOM)" -ForegroundColor Green
        }
        catch {
            Write-Host " -> [ERROR] 失敗: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`n✅ 所有腳本編碼修正作業完成。" -ForegroundColor Yellow
    Write-Host "請重新啟動 Master 與 Payload 測試是否仍有紅字錯誤。"
} else {
    Write-Warning "❌ 找不到 Scripts 目錄: $TargetDir"
}