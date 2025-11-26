# =============================================================================
# AutoTask 嚴格排程註冊工具 V3.0 (XML 強制定義版)
# 功能：使用 XML 定義檔繞過 CIM 物件錯誤，精確綁定 Remote 帳戶
# =============================================================================

$TaskName = "Auto_BetterGI_Payload"
$TargetUser = "Remote" 
$ComputerName = $env:COMPUTERNAME
$FullUserName = "$ComputerName\$TargetUser" # 格式: DESKTOP-XXX\Remote
$ScriptPath = "C:\AutoTask\Scripts\Payload.ps1"

Write-Host "正在設定 [$FullUserName] 專用的嚴格觸發任務..." -ForegroundColor Cyan

# 定義 XML 內容 (包含登入、連線、解鎖三種觸發器，且綁定 UserID)
$TaskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>AutoTask</Author>
    <Description>Auto-start Payload for BetterGI automation on Remote account login/connection.</Description>
    <URI>\$TaskName</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$FullUserName</UserId>
    </LogonTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>RemoteConnect</StateChange>
      <UserId>$FullUserName</UserId>
    </SessionStateChangeTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>ConsoleConnect</StateChange>
      <UserId>$FullUserName</UserId>
    </SessionStateChangeTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>SessionUnlock</StateChange>
      <UserId>$FullUserName</UserId>
    </SessionStateChangeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$FullUserName</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
    <Priority>1</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "$ScriptPath"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

# 執行註冊
try {
    # 先嘗試刪除舊任務 (避免衝突)
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    
    # 使用 XML 註冊新任務
    Register-ScheduledTask -Xml $TaskXml -TaskName $TaskName -Force
    
    Write-Host "✅ 任務 [$TaskName] 已成功註冊！" -ForegroundColor Green
    Write-Host "   - 執行身分: $FullUserName"
    Write-Host "   - 觸發機制: XML 定義 (登入/連線/解鎖)"
    Write-Host "   - 腳本路徑: $ScriptPath"
} catch {
    Write-Host "❌ 註冊失敗: $_" -ForegroundColor Red
    Write-Host "請確保您是以 [系統管理員身分] 執行此腳本。" -ForegroundColor Yellow
    Write-Host "如果錯誤提示『使用者帳戶未知』，請確認 [$FullUserName] 是否正確。" -ForegroundColor Yellow
}

Write-Host "`n請按 Enter 結束..."
Read-Host