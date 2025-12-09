$FilePath = "C:\AutoTask\Scripts\Payload.ps1"

if (Test-Path $FilePath) {
    Write-Host "正在檢查檔案編碼: $FilePath"
    
    # 讀取內容 (Get-Content 預設會自動處理 BOM)
    $Content = Get-Content -Path $FilePath -Raw
    
    # 重新寫入為 UTF8NoBOM
    # 注意: PowerShell Core (pwsh) 的預設 UTF8 即為 NoBOM
    # 但為了相容 Windows PowerShell 5.1，我們明確指定編碼物件
    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
    [System.IO.File]::WriteAllText($FilePath, $Content, $Utf8NoBomEncoding)
    
    Write-Host "✅ 已成功將檔案轉換為 UTF-8 (No BOM) 格式。" -ForegroundColor Green
    Write-Host "請重新執行 Payload.ps1 測試是否正常。"
} else {
    Write-Warning "❌ 找不到檔案: $FilePath"
}