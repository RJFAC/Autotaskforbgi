# =======================================================
# 檔案名稱: Fix_Encoding.ps1
# 功能: 將所有腳本轉換為 UTF-8 with BOM (解決中文亂碼 ?? 問題)
# =======================================================

$TargetDir = "C:\AutoTask\Scripts"
Write-Host "正在修復腳本編碼 (UTF-8 BOM)..." -ForegroundColor Cyan

if (Test-Path $TargetDir) {
    $Files = Get-ChildItem -Path $TargetDir -Filter "*.ps1" -Recurse
    
    foreach ($File in $Files) {
        Write-Host "處理: $($File.Name)" -NoNewline
        try {
            # 讀取內容 (嘗試自動偵測編碼)
            $Content = Get-Content $File.FullName -Raw
            
            # 強制以 UTF-8 (BOM) 寫回
            # 在 PowerShell 5.1 中，"UTF8" 預設就是帶 BOM 的，這是我們需要的
            Set-Content -Path $File.FullName -Value $Content -Encoding UTF8 -Force
            Write-Host " [成功]" -ForegroundColor Green
        } catch {
            Write-Host " [失敗: $_]" -ForegroundColor Red
        }
    }
}

Write-Host "`n全部完成！請重新測試 Discord 通知。" -ForegroundColor Yellow
Write-Host "按任意鍵退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
