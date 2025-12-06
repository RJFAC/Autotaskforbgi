# ==========================================
# ASUS ROG STRIX B350-F USB Freeze Fixer & Diagnostic
# 功能:
# 1. [診斷] 檢查 Filter Drivers 與硬體狀態
# 2. [修復] 強制關閉 Windows USB 選擇性暫停 (Selective Suspend)
# 3. [修復] 掃描並關閉所有 USB 裝置的硬體省電模式 (PnP Power Management)
# ==========================================

# 確保以管理員身分執行
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "請以系統管理員身分執行此腳本！(右鍵 -> 使用 PowerShell 執行)"
    Start-Sleep -Seconds 5
    Exit
}

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ReportFile = "$DesktopPath\USB_Fix_Report.txt"
$OutputEncoding = [System.Text.Encoding]::UTF8

Function Write-Log {
    Param ([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
    Add-Content -Path $ReportFile -Value $Message -Encoding UTF8
}

Function Write-Section {
    Param ([string]$Title)
    $Line = "=" * 50
    Add-Content -Path $ReportFile -Value "`n$Line" -Encoding UTF8
    Add-Content -Path $ReportFile -Value " $Title" -Encoding UTF8
    Add-Content -Path $ReportFile -Value "$Line" -Encoding UTF8
    Write-Host "`n[$Title]" -ForegroundColor Yellow
}

# 初始化報告
New-Item -Path $ReportFile -ItemType File -Force | Out-Null
Write-Log "執行時間: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
try {
    Write-Log "主機板: $(Get-CimInstance Win32_BaseBoard | Select-Object -ExpandProperty Product)"
} catch {}

# ==========================================
# 1. 執行修復：強制關閉 USB 選擇性暫停
# ==========================================
Write-Section "1. 正在套用電源修復..."

# A. 修改登錄檔 (全域設定)
$RegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\USB"
try {
    if (!(Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
    Set-ItemProperty -Path $RegPath -Name "DisableSelectiveSuspend" -Value 1 -Type DWord -Force
    Write-Log "[修復成功] 已強制登錄檔 DisableSelectiveSuspend = 1"
} catch {
    Write-Log "[修復失敗] 無法寫入登錄檔: $_"
}

# B. 修改個別裝置電源管理 (PnP 層級)
Write-Log "正在掃描 USB 裝置並關閉個別省電設定..."
try {
    # 使用 MSP_PowerManagementSettings 修改 "允許電腦關閉這個裝置以節省電源"
    # 這通常是導致裝置睡死的元兇
    $PowerSettings = Get-CimInstance -Namespace "root\wmi" -ClassName "MSP_PowerManagementSettings" -ErrorAction SilentlyContinue
    
    $Count = 0
    if ($PowerSettings) {
        foreach ($setting in $PowerSettings) {
            if ($setting.AllowComputerToTurnOffDevice -eq $true) {
                $setting.AllowComputerToTurnOffDevice = $false
                Set-CimInstance -InputObject $setting -ErrorAction SilentlyContinue
                $Count++
            }
        }
        Write-Log "[修復成功] 已關閉 $Count 個 USB 裝置的硬體省電選項。"
    } else {
        Write-Log "[資訊] 未偵測到可透過 WMI 修改電源的 USB 裝置 (或已全部關閉)。"
    }
} catch {
    Write-Log "[錯誤] 修改裝置電源設定時發生異常: $_"
}

# ==========================================
# 2. 診斷 Filter Drivers (再次確認)
# ==========================================
Write-Section "2. Filter Drivers 狀態確認"
$USBClassGUID = "{36FC9E60-C465-11CF-8056-444553540000}"
$RegPathClass = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\$USBClassGUID"
try {
    $Upper = (Get-ItemProperty -Path $RegPathClass -Name UpperFilters -ErrorAction SilentlyContinue).UpperFilters
    $Lower = (Get-ItemProperty -Path $RegPathClass -Name LowerFilters -ErrorAction SilentlyContinue).LowerFilters
    
    if ($Upper) { Write-Log "UpperFilters: $Upper (注意)" } else { Write-Log "UpperFilters: 無 (安全)" }
    if ($Lower) { Write-Log "LowerFilters: $Lower (注意)" } else { Write-Log "LowerFilters: 無 (安全)" }
} catch {}

# ==========================================
# 3. 硬體狀態檢查 (使用 Get-PnpDevice 修正版)
# ==========================================
Write-Section "3. USB 硬體異常掃描"
try {
    # 這裡過濾所有 USB 相關且狀態不為 OK 的裝置
    $BadDevices = Get-PnpDevice -Class "USB" -Status "Error","Degraded","Unknown" -ErrorAction SilentlyContinue
    
    if ($BadDevices) {
        foreach ($dev in $BadDevices) {
            Write-Log "發現異常: $($dev.FriendlyName)"
            Write-Log "  - 狀態: $($dev.Status)"
            Write-Log "  - 錯誤碼: $($dev.Problem) (CM_PROB)"
        }
    } else {
        Write-Log "目前系統未偵測到帶有錯誤代碼的 USB 裝置。"
        Write-Log "(若目前滑鼠鍵盤可用，此結果為正常)"
    }
} catch {
    Write-Log "掃描失敗: $_"
}

# ==========================================
# 4. 結論
# ==========================================
Write-Section "完成"
Write-Log "修復操作已完成。"
Write-Log "請務必【重新開機】以確保電源設定生效。"
Write-Log "報告已儲存: $ReportFile"

Start-Sleep -Seconds 2
Invoke-Item $ReportFile