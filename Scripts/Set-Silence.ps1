<#
.SYNOPSIS
    AutoTask Silence & Focus Mode Controller
    用於在系統啟動時強制靜音與開啟勿擾模式。
.DESCRIPTION
    1. 音量控制: 模擬發送 50 次 VolumeDown 按鍵，確保主音量歸零。
    2. 勿擾模式: 修改登錄檔以禁用 Toast 通知 (Focus Assist)。
#>

# --- 1. 強制靜音 (Volume Zero) ---
# 使用 WScript.Shell 模擬按鍵，這是最不依賴外部 DLL 的原生方法
try {
    $wsh = New-Object -ComObject WScript.Shell
    # 傳送 50 次音量降低鍵 (Code 174)，確保音量降至 0
    # 我們不使用靜音鍵 (Code 173)，因為那是切換(Toggle)開關，若原先是靜音反而會變成有聲
    for ($i = 0; $i -lt 50; $i++) {
        $wsh.SendKeys([char]174)
        Start-Sleep -Milliseconds 10
    }
    Write-Host "[Silence] 系統音量已歸零。" -ForegroundColor Green
} catch {
    Write-Warning "[Silence] 音量控制失敗: $($_.Exception.Message)"
}

# --- 2. 開啟勿擾/禁用通知 (Disable Toasts) ---
try {
    $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
    if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
    
    # ToastEnabled: 0 = 禁用通知 (勿擾), 1 = 啟用
    Set-ItemProperty -Path $RegPath -Name "ToastEnabled" -Value 0 -Type DWord -Force
    Write-Host "[Silence] Windows 通知已禁用 (勿擾模式)。" -ForegroundColor Green
} catch {
    Write-Warning "[Silence] 勿擾模式設定失敗: $($_.Exception.Message)"
}