# 檢查是否以管理員身分執行
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "請以系統管理員身分執行此腳本！" -ForegroundColor Red
    Start-Sleep -Seconds 3
    Exit
}

Write-Host "正在執行 USB 穩定性修復..." -ForegroundColor Cyan

# 1. 關閉電源計畫中的「USB 選擇性暫停」(USB Selective Suspend)
# 修改當前與所有電源計畫的設定 (AC 與 DC 模式)
Write-Host "1. 正在關閉 USB 選擇性暫停..." -ForegroundColor Yellow
powercfg /SETACVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-728eb46c4996 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
powercfg /SETDCVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-728eb46c4996 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
powercfg /SetActive SCHEME_CURRENT
Write-Host "   - 已完成。" -ForegroundColor Green

# 2. 關閉 Windows 快速啟動 (Fast Startup)
# 這能確保每次關機都是完全斷電，避免錯誤狀態殘留
Write-Host "2. 正在關閉 Windows 快速啟動 (Hiberboot)..." -ForegroundColor Yellow
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
try {
    Set-ItemProperty -Path $regPath -Name "HiberbootEnabled" -Value 0 -ErrorAction Stop
    Write-Host "   - 已完成 (HiberbootEnabled = 0)。" -ForegroundColor Green
} catch {
    Write-Host "   - 修改註冊表失敗，請確認權限。" -ForegroundColor Red
}

# 3. 掃描所有 USB 控制器並關閉「允許電腦關閉這個裝置以節省電源」
# 這部分透過 WMI 修改 PnP 電源管理設定
Write-Host "3. 正在修改裝置管理員 USB 電源設定 (這可能需要一點時間)..." -ForegroundColor Yellow
$usbDevices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.Name -match 'USB' -or $_.Service -match 'USB' }

foreach ($device in $usbDevices) {
    try {
        # 嘗試尋找該裝置的電源管理設定 (MSP_PowerManagementSettings)
        # 注意：並非所有驅動都暴露此 WMI 介面，此操作為盡力而為
        $pmSettings = Get-CimInstance -Namespace "root\wmi" -ClassName "MSP_PowerManagementSettings" -Filter "InstanceName like '$($device.PNPDeviceID)%'" -ErrorAction SilentlyContinue
        
        if ($pmSettings) {
            $pmSettings.AllowComputerToTurnOffDevice = $false
            Set-CimInstance -InputObject $pmSettings -ErrorAction SilentlyContinue
            Write-Host "   - 已修正: $($device.Name)" -ForegroundColor Gray
        }
    } catch {
        # 忽略無法修改的裝置
    }
}
Write-Host "   - 掃描與修正完成。" -ForegroundColor Green

Write-Host "`n==============================================="
Write-Host "所有修復已套用。"
Write-Host "請務必【重新開機】一次以讓設定生效。" -ForegroundColor Cyan
Write-Host "==============================================="
Read-Host "按 Enter 鍵退出..."