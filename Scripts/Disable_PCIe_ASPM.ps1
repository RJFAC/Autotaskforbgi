# ==========================================
# ASUS B350-F USB 進階修復工具 (註冊表強制版)
# 功能:
# 1. 繞過 powercfg 指令，直接修改註冊表電源設定
# 2. 強制寫入 PCIe ASPM = Off (0)
# 3. 強制寫入 USB 3.0 Link Power Management = Off (0)
# ==========================================

# 確保以管理員身分執行
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "請以系統管理員身分執行此腳本！(右鍵 -> 使用 PowerShell 執行)"
    Start-Sleep -Seconds 5
    Exit
}

$LogFile = "$([Environment]::GetFolderPath('Desktop'))\PCIe_Fix_Log.txt"
Function Write-Log { Param ([string]$Msg) Write-Host $Msg -ForegroundColor Cyan; Add-Content $LogFile "$Msg" }

Write-Log "開始執行進階電源修復 (註冊表強制寫入模式)..."

# 1. 獲取當前使用中的電源計畫 GUID
Write-Log "正在獲取當前電源計畫 GUID..."
$ActiveSchemeRaw = powercfg /getactivescheme
if ($ActiveSchemeRaw -match "([a-f0-9]{8}-([a-f0-9]{4}-){3}[a-f0-9]{12})") {
    $TargetScheme = $matches[0]
    Write-Log "鎖定目標電源計畫: $TargetScheme"
} else {
    Write-Log "無法自動獲取 GUID，使用平衡(Balanced)預設值..."
    $TargetScheme = "381b4222-f694-41f0-9685-ff5bb260df2e"
}

# 定義要修改的 GUID
$PCIe_Subgroup = "501a4d13-42af-4429-9fd1-a8218c268e29"
$PCIe_Setting  = "ee12f906-d277-404b-b6da-e5fa1a576df5" # Link State Power Management

$USB_Subgroup  = "2a737441-1930-4402-8d77-728eb46c4996"
$USB_Setting   = "d4e98f31-5ffe-4ce1-be31-1b38b384c009" # USB 3.0 LPM

# 註冊表根路徑 (所有電源計畫的儲存位置)
$RegRoot = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerSchemes"

Function Force-Reg-Power-Setting {
    Param (
        [string]$Scheme,
        [string]$Subgroup,
        [string]$Setting,
        [string]$Name
    )
    
    $Path = "$RegRoot\$Scheme\$Subgroup\$Setting"
    
    Write-Log "正在設定: $Name"
    
    # 檢查路徑是否存在，若不存在則建立 (防止無效路徑錯誤)
    if (!(Test-Path $Path)) {
        Write-Log "  - 路徑不存在，正在建立..."
        New-Item -Path $Path -Force | Out-Null
    }

    try {
        # 設定 AC (插電) 與 DC (電池) 值為 0 (Off)
        Set-ItemProperty -Path $Path -Name "ACSettingIndex" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $Path -Name "DCSettingIndex" -Value 0 -Type DWord -Force
        Write-Log "  - [成功] 已強制寫入數值 0 (Off)"
    } catch {
        Write-Log "  - [失敗] 寫入註冊表時發生錯誤: $_"
    }
}

# 執行修改
Force-Reg-Power-Setting -Scheme $TargetScheme -Subgroup $PCIe_Subgroup -Setting $PCIe_Setting -Name "PCIe ASPM (Link State)"
Force-Reg-Power-Setting -Scheme $TargetScheme -Subgroup $USB_Subgroup  -Setting $USB_Setting  -Name "USB 3.0 Link Power Management"

# 最後重新載入電源設定以通知核心
Write-Log "正在重新載入電源設定..."
powercfg /setactive $TargetScheme

Write-Log "修復完成。請【重新開機】以確保硬體設定生效。"
Start-Sleep -Seconds 3