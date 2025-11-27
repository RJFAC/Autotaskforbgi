# =============================================================================
# AutoTask Dashboard V8.3 - 多重配置與暫停旗艦版
# =============================================================================

# --- [隱藏 Console 黑窗] ---
$code = @"
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
"@
$win = Add-Type -MemberDefinition $code -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
$hwnd = $win::GetConsoleWindow()
if ($hwnd -ne [IntPtr]::Zero) { $win::ShowWindow($hwnd, 0) } 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic 

# --- [定義檔案路徑] ---
$Dir = "C:\AutoTask"
$ScriptDir = "$Dir\Scripts"
$ConfigsDir = "$Dir\Configs"
$LogsDir = "$Dir\Logs"
$WeeklyConf = "$ConfigsDir\WeeklyConfig.json"
$DateMap = "$ConfigsDir\DateConfig.map"
$TaskStatus = "$ConfigsDir\TaskStatus.json"
$PauseLog = "$ConfigsDir\PauseDates.log"
$NoShutdownLog = "$ConfigsDir\NoShutdown.log"
$ResinConf = "$ConfigsDir\ResinConfig.json"
$ManualFlag = "$Dir\Flags\ManualTrigger.flag"
$BetterGI_UserDir = "C:\Program Files\BetterGI\User\OneDragon"
$MasterScript = "$ScriptDir\Master.ps1"
$StopScript = "$ScriptDir\StopAll.ps1"
$PublishScript = "$ScriptDir\PublishRelease.ps1"
$HashFile = "$ConfigsDir\ScriptHash.txt"

# --- [全域變數] ---
$Global:ConfigList = @() 
$Global:WeeklyRules = @{}
$Global:TurbulenceRules = @{}
$Global:WeeklyNoShut = @{} 
$Global:TurbulenceNoShut = @{}
$Global:GenshinPath = "" 
$Global:ResinData = @{} 
$Global:InitialHash = ""
$Script:IsDirty = $false
$Script:IsLoading = $false
$WindowTitle = "AutoTask 控制台 V8.3"

# 字型
$MainFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 10)
$BoldFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 10, [System.Drawing.FontStyle]::Bold)
$TitleFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12, [System.Drawing.FontStyle]::Bold)
$MonoFont = New-Object System.Drawing.Font("Consolas", 10) 

function Get-CurrentScriptsHash {
    $str = ""
    Get-ChildItem $ScriptDir -Include "*.ps1", "*.bat" -Recurse | Sort-Object Name | ForEach-Object { 
        $str += (Get-FileHash $_.FullName).Hash 
    }
    return $str
}
$Global:InitialHash = Get-CurrentScriptsHash

# --- [輔助函數] ---
function Get-JsonConf ($path) {
    if (Test-Path $path) { 
        try { return Get-Content $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } 
        catch { return $null }
    }
    return $null
}

function Load-BetterGIConfigs {
    $Global:ConfigList = @() 
    if (Test-Path $BetterGI_UserDir) {
        $Files = Get-ChildItem "$BetterGI_UserDir\*.json"
        foreach ($f in $Files) {
            try {
                $json = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                $Global:ConfigList += $json.Name
            } catch {}
        }
    }
}

function Load-EnvConfig {
    $env = Get-JsonConf "$ConfigsDir\EnvConfig.json"
    if ($env -and $env.GenshinPath) {
        $Global:GenshinPath = $env.GenshinPath
    } else {
        $wk = Get-JsonConf $WeeklyConf
        if ($wk -and $wk.GenshinPath) { 
            $Global:GenshinPath = $wk.GenshinPath
            $newEnv = @{ GenshinPath = $wk.GenshinPath }
            $newEnv | ConvertTo-Json | Set-Content "$ConfigsDir\EnvConfig.json" -Encoding UTF8
        } else {
            $Global:GenshinPath = "尚未設定"
        }
    }
}

function Load-WeeklyRules {
    $wk = Get-JsonConf $WeeklyConf
    
    $Global:WeeklyRules = @{ "Monday"="monday"; "Tuesday"="day"; "Wednesday"="day"; "Thursday"="day"; "Friday"="day"; "Saturday"="day"; "Sunday"="day" }
    $Global:TurbulenceRules = @{ "Monday"="day"; "Tuesday"="day"; "Wednesday"="day"; "Thursday"="day"; "Friday"="day"; "Saturday"="day"; "Sunday"="day" }
    $Global:WeeklyNoShut = @{ "Monday"=$false; "Tuesday"=$false; "Wednesday"=$false; "Thursday"=$false; "Friday"=$false; "Saturday"=$false; "Sunday"=$false }
    $Global:TurbulenceNoShut = @{ "Monday"=$false; "Tuesday"=$false; "Wednesday"=$false; "Thursday"=$false; "Friday"=$false; "Saturday"=$false; "Sunday"=$false }

    if ($wk) {
        foreach ($k in $Global:WeeklyRules.Keys) { if ($wk.$k) { $Global:WeeklyRules[$k] = $wk.$k } }
        if ($wk.NoShutdown) {
            foreach ($k in $Global:WeeklyNoShut.Keys) { 
                if ($wk.NoShutdown.$k -ne $null) { $Global:WeeklyNoShut[$k] = [bool]$wk.NoShutdown.$k } 
            }
        }
        if ($wk.Turbulence) {
            foreach ($k in $Global:TurbulenceRules.Keys) { if ($wk.Turbulence.$k) { $Global:TurbulenceRules[$k] = $wk.Turbulence.$k } }
            if ($wk.Turbulence.NoShutdown) {
                 foreach ($k in $Global:TurbulenceNoShut.Keys) {
                    if ($wk.Turbulence.NoShutdown.$k -ne $null) { $Global:TurbulenceNoShut[$k] = [bool]$wk.Turbulence.NoShutdown.$k }
                 }
            }
        }
    }
}

function Load-ResinConfig {
    $json = Get-JsonConf $ResinConf
    $Global:ResinData = if ($json) { $json } else { @{} }
}

function Test-GenshinUpdateDay ($CheckDate) {
    $RefDate = [datetime]"2024-08-28"
    $DiffDays = ($CheckDate.Date - $RefDate).Days
    if ($DiffDays -ge 0 -and ($DiffDays % 42) -eq 0) { return $true }
    return $false
}

function Test-TurbulencePeriod ($CheckDate) {
    $RefDate = [datetime]"2024-08-28"
    $DiffDays = ($CheckDate.Date - $RefDate).Days
    if ($DiffDays -ge 0) {
        $CycleDay = $DiffDays % 42
        if ($CycleDay -ge 8 -and $CycleDay -le 17) { return $CycleDay }
    }
    return 0
}

function Get-DisplayConfigName ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    $dWeek = $dateObj.DayOfWeek.ToString()
    if (Test-Path $DateMap) {
        $map = Get-Content $DateMap
        foreach ($line in $map) { if ($line -match "^$dStr=(.+)$") { return "$($matches[1]) (指定)" } }
    }
    if (Test-TurbulencePeriod $dateObj) {
        $tConf = $Global:TurbulenceRules.$dWeek
        if ($tConf) { return "$tConf (紊亂期)" }
    }
    return "$($Global:WeeklyRules.$dWeek) (每週)"
}

function Get-StatusText {
    $dStr = (Get-Date).AddHours(-3).ToString("yyyyMMdd")
    $st = Get-JsonConf $TaskStatus
    $txt = "尚未執行"
    $color = [System.Drawing.Color]::Gray
    if ($st -and $st.Date -eq $dStr) {
        $txt = $st.Status
        if ($st.RetryCount -gt 0) { $txt += " (重試: $($st.RetryCount))" }
        if ($txt -match "Failed") { $color = [System.Drawing.Color]::Red }
        elseif ($txt -match "Success") { $color = [System.Drawing.Color]::Green }
        elseif ($txt -match "Running") { $color = [System.Drawing.Color]::Blue }
        elseif ($txt -match "Maintenance") { $color = [System.Drawing.Color]::Orange }
    }
    if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains $dStr) { $txt = "已排程暫停"; $color = [System.Drawing.Color]::Orange } }
    return @{Text=$txt; Color=$color}
}

function Get-ShutdownPolicy ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    $dWeek = $dateObj.DayOfWeek.ToString()
    if (Test-Path $NoShutdownLog) { if ((Get-Content $NoShutdownLog) -contains $dStr) { return "不關機 (指定)" } }
    if (Test-TurbulencePeriod $dateObj) {
        if ($Global:TurbulenceNoShut.$dWeek) { return "不關機 (紊亂)" }
    }
    if ($Global:WeeklyNoShut.$dWeek) { return "不關機 (每週)" }
    return "自動關機"
}

function Get-WeekName ($dateObj) { return (@{ "Monday"="週一"; "Tuesday"="週二"; "Wednesday"="週三"; "Thursday"="週四"; "Friday"="週五"; "Saturday"="週六"; "Sunday"="週日" })[$dateObj.DayOfWeek.ToString()] }

function Mark-Dirty { if (-not $Script:IsLoading) { $Script:IsDirty = $true; $Form.Text = "$WindowTitle * (未儲存)" } }
function Mark-Clean { $Script:IsDirty = $false; $Form.Text = $WindowTitle }

function Auto-Detect-GenshinPath {
    $GameExes = @("YuanShen.exe", "GenshinImpact.exe")
    try {
        $WmicOutput = wmic process where "name='YuanShen.exe' or name='GenshinImpact.exe'" get ExecutablePath 2>$null | Out-String
        if ($WmicOutput -match "(.:\\.*\.exe)") { return (Split-Path $matches[1] -Parent) }
    } catch {}
    $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Genshin Impact", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\原神", "HKCU:\Software\miHoYo\Genshin Impact")
    foreach ($reg in $RegPaths) { if (Test-Path $reg) { $p=(Get-ItemProperty $reg).InstallLocation; $p2=(Get-ItemProperty $reg).InstallPath; foreach($b in @($p,$p2)){if($b-and(Test-Path $b)){$s=Get-ChildItem -Path $b -Include $GameExes -Recurse -Depth 3 -File -EA SilentlyContinue|Select -First 1; if($s){return $s.DirectoryName}}} } }
    $CommonPaths = @("C:\Program Files\Genshin Impact","C:\Program Files\HoYoPlay\games\Genshin Impact Game","D:\Genshin Impact Game","E:\Genshin Impact Game")
    foreach ($cp in $CommonPaths) { if (Test-Path $cp) { $s=Get-ChildItem -Path $cp -Include $GameExes -Recurse -Depth 3 -File -EA SilentlyContinue|Select -First 1; if($s){return $s.DirectoryName} } }
    return $null
}

# --- [核心功能] 配置選擇器 (V8.3) ---
function Show-ConfigSelectorGUI {
    param([string]$CurrentSelection) 
    $SelForm = New-Object System.Windows.Forms.Form; $SelForm.Text="配置選擇 (拖曳排序)"; $SelForm.Size="700,500"; $SelForm.StartPosition="CenterParent"; $SelForm.Font=$MainFont
    
    $lblSrc = New-Object System.Windows.Forms.Label; $lblSrc.Text="可用配置 (可多選)"; $lblSrc.Location="20,10"; $lblSrc.AutoSize=$true
    $listSrc = New-Object System.Windows.Forms.ListBox; $listSrc.Location="20,30"; $listSrc.Size="250,350"; $listSrc.SelectionMode="MultiExtended"
    $RealConfigs = $Global:ConfigList | Where-Object { $_ -ne "PAUSE" }; $listSrc.Items.AddRange($RealConfigs)
    
    $lblDst = New-Object System.Windows.Forms.Label; $lblDst.Text="執行佇列 (由上而下執行)"; $lblDst.Location="380,10"; $lblDst.AutoSize=$true
    $listDst = New-Object System.Windows.Forms.ListBox; $listDst.Location="380,30"; $listDst.Size="250,350"; $listDst.SelectionMode="One"; $listDst.AllowDrop=$true 
    
    if (-not [string]::IsNullOrWhiteSpace($CurrentSelection)) { 
        if ($CurrentSelection -eq "PAUSE") {
            $listDst.Items.Add("PAUSE") | Out-Null
        } else {
            $parts = $CurrentSelection -split ","
            foreach ($p in $parts) { 
                if (-not [string]::IsNullOrWhiteSpace($p)) { $listDst.Items.Add($p) | Out-Null } 
            } 
        }
    }
    
    # 按鈕區
    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text="加入 ->"; $btnAdd.Location="280,100"; $btnAdd.Size="90,30"
    $btnAdd.Add_Click({ 
        if ($listDst.Items.Contains("PAUSE")) { $listDst.Items.Clear() }
        foreach ($item in $listSrc.SelectedItems) { $listDst.Items.Add($item) | Out-Null } 
    })
    
    $btnRem = New-Object System.Windows.Forms.Button; $btnRem.Text="<- 移除"; $btnRem.Location="280,150"; $btnRem.Size="90,30"
    $btnRem.Add_Click({ if ($listDst.SelectedIndex -ge 0) { $listDst.Items.RemoveAt($listDst.SelectedIndex) } })

    # [新] 暫停與清空
    $btnPause = New-Object System.Windows.Forms.Button; $btnPause.Text="⛔ 設為暫停"; $btnPause.Location="280,250"; $btnPause.Size="90,30"; $btnPause.BackColor="LightCoral"
    $btnPause.Add_Click({ $listDst.Items.Clear(); $listDst.Items.Add("PAUSE") | Out-Null })

    $btnClear = New-Object System.Windows.Forms.Button; $btnClear.Text="❌ 清空"; $btnClear.Location="280,300"; $btnClear.Size="90,30"
    $btnClear.Add_Click({ $listDst.Items.Clear() })

    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text="確定"; $btnOk.Location="250,400"; $btnOk.DialogResult="OK"; $btnOk.BackColor="LightGreen"; $btnOk.Size="100,40"
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text="取消"; $btnCancel.Location="360,400"; $btnCancel.DialogResult="Cancel"; $btnCancel.Size="100,40"
    
    # 拖曳排序
    $listDst.Add_MouseDown({ param($s,$e); if($listDst.SelectedItem) { $listDst.DoDragDrop($listDst.SelectedItem, [System.Windows.Forms.DragDropEffects]::Move) } })
    $listDst.Add_DragOver({ param($s,$e); $e.Effect=[System.Windows.Forms.DragDropEffects]::Move })
    $listDst.Add_DragDrop({ param($s,$e); $idx=$listDst.IndexFromPoint($listDst.PointToClient([System.Drawing.Point]::new($e.X,$e.Y))); if($idx -lt 0){$idx=$listDst.Items.Count-1}; $item=$e.Data.GetData([string]); if($item){$listDst.Items.Remove($item); $listDst.Items.Insert($idx,$item); $listDst.SelectedIndex=$idx} })
    
    $SelForm.Controls.AddRange(@($lblSrc, $listSrc, $lblDst, $listDst, $btnAdd, $btnRem, $btnPause, $btnClear, $btnOk, $btnCancel))
    $SelForm.AcceptButton = $btnOk

    if ($SelForm.ShowDialog() -eq "OK") { 
        $finalList = @()
        foreach($i in $listDst.Items){ $finalList += $i }
        return ($finalList -join ",") 
    } else { return $null }
}

# --- GUI 初始化 ---
Load-BetterGIConfigs
Load-WeeklyRules
Load-ResinConfig

$Form = New-Object System.Windows.Forms.Form
$Form.Text = $WindowTitle
$Form.Size = New-Object System.Drawing.Size(1000, 780)
$Form.StartPosition = "CenterScreen"
$Form.Font = $MainFont

$Form.Add_FormClosing({
    param($sender, $e)
    if ($Script:IsDirty) {
        if ([System.Windows.Forms.MessageBox]::Show("設定未儲存，確定要離開？", "警告", "YesNo") -eq "No") { $e.Cancel = $true; return }
    }
    $LastHash = ""; if (Test-Path $HashFile) { $LastHash = Get-Content $HashFile -Raw }
    $CurrentHash = Get-CurrentScriptsHash
    if ($CurrentHash -ne $LastHash) {
        if ([System.Windows.Forms.MessageBox]::Show("偵測到腳本核心已變更 (與上次發布不同)，是否同步至 GitHub？", "版本控制", "YesNo") -eq "Yes") {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PublishScript`""
        }
    }
})

$TabControl = New-Object System.Windows.Forms.TabControl; $TabControl.Dock = "Fill"; $TabControl.Font = $MainFont

# === 分頁 1: 即時狀態 ===
$TabStatus = New-Object System.Windows.Forms.TabPage; $TabStatus.Text = "[HOME] 即時狀態"; $TabStatus.Padding = "10"
$lblInfo = New-Object System.Windows.Forms.Label; $lblInfo.AutoSize=$true; $lblInfo.Font=$TitleFont; $lblInfo.Location="20,20"
$btnMan = New-Object System.Windows.Forms.Button; $btnMan.Text="[!] 強制啟動"; $btnMan.Location="20,150"; $btnMan.Size="300,50"; $btnMan.BackColor="LightCoral"; $btnMan.Font=$TitleFont
$btnMan.Add_Click({ if([System.Windows.Forms.MessageBox]::Show("確定強制啟動？","確認","YesNo") -eq "Yes"){ New-Item -Path $ManualFlag -Force|Out-Null; Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`"" -Verb RunAs } })
$btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="重新整理"; $btnRef.Location="20,210"; $btnRef.Width=300
$btnRef.Font = $MainFont
$btnRef.Add_Click({ Update-StatusUI })
$TabStatus.Controls.AddRange(@($lblInfo, $btnMan, $btnRef))

function Update-StatusUI {
    $today = (Get-Date).AddHours(-3)
    $st = Get-StatusText
    $finalConf = Get-DisplayConfigName $today
    if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains $today.ToString("yyyyMMdd")) { $finalConf = "PAUSED" } }
    $Note = ""; if (Test-GenshinUpdateDay $today) { $Note = " (⚠️ 版本更新日)" }; $ITDay = Test-TurbulencePeriod $today; if ($ITDay -gt 0) { $Note = " (🔥 紊亂期 Day $ITDay)" }
    $lblInfo.Text = "今日: $($today.ToString('yyyy/MM/dd')) ($($today.DayOfWeek))$Note`n配置: $finalConf`n狀態: $($st.Text)"; $lblInfo.ForeColor = $st.Color
}

# === 分頁 2: 排程網格 ===
$TabGrid = New-Object System.Windows.Forms.TabPage; $TabGrid.Text = "[GRID] 排程編輯器"
$pTool = New-Object System.Windows.Forms.Panel; $pTool.Dock="Top"; $pTool.Height=40
$btnGSave = New-Object System.Windows.Forms.Button; $btnGSave.Text="[SAVE]"; $btnGSave.Dock="Left"; $btnGSave.Width=100; $btnGSave.BackColor="LightGreen"; $btnGSave.Font=$BoldFont
$btnGSave.Add_Click({ Save-GridData })
$lblHint = New-Object System.Windows.Forms.Label; $lblHint.Text="操作提示: 支援批量勾選 [不關機] (Ctrl/Shift) | 雙擊配置欄排序 | Ctrl+C/V | Del"; $lblHint.Dock="Fill"; $lblHint.TextAlign="MiddleLeft"; $lblHint.Padding="10,0,0,0"; $lblHint.Font=$MainFont
$pTool.Controls.Add($lblHint); $pTool.Controls.Add($btnGSave)
$grid = New-Object System.Windows.Forms.DataGridView; $grid.Dock="Fill"; $grid.EditMode="EditProgrammatically"; $grid.Font=$MonoFont; $grid.MultiSelect=$true
$grid.Columns.Add("Date","日期"); $grid.Columns[0].ReadOnly=$true; $grid.Columns[0].Width=120
$grid.Columns.Add("Week","星期"); $grid.Columns[1].ReadOnly=$true; $grid.Columns[1].Width=60
$grid.Columns.Add("Def","每週預設"); $grid.Columns[2].ReadOnly=$true; $grid.Columns[2].Width=100
$grid.Columns.Add("Conf","執行配置 (雙擊)"); $grid.Columns[3].Width=250
$grid.Columns.Add("Shut","不關機"); $grid.Columns[4].Width=60; $grid.Columns[4].CellTemplate=New-Object System.Windows.Forms.DataGridViewCheckBoxCell
$grid.Columns.Add("Note","備註"); $grid.Columns[5].ReadOnly=$true; $grid.Columns[5].Width=150
$grid.Add_CellClick({ param($s,$e); if($e.RowIndex-lt 0){return}; if($e.ColumnIndex-eq 4){ $c=$grid.Rows[$e.RowIndex].Cells[4]; $v=-not [bool]$c.Value; $sel=$grid.SelectedCells|Where{$_.ColumnIndex-eq 4}; if($sel.Count-gt 0 -and ($sel|Where{$_.RowIndex-eq $e.RowIndex})){foreach($x in $sel){$x.Value=$v}}else{$c.Value=$v}; Mark-Dirty } })
$grid.Add_CellDoubleClick({ param($s,$e); if($e.RowIndex-lt 0-or $e.ColumnIndex-ne 3){return}; $c=$grid.Rows[$e.RowIndex].Cells[3]; $cv=$c.Value; if($cv-eq $grid.Rows[$e.RowIndex].Cells[2].Value-or $cv-eq "PAUSE"){$cv=""}; $n=Show-ConfigSelectorGUI $cv; if($n-ne $null){if($n-eq""){$c.Value=$grid.Rows[$e.RowIndex].Cells[2].Value;$c.Style=$grid.DefaultCellStyle}else{$c.Value=$n;$c.Style.ForeColor="Blue";$c.Style.Font=$BoldFont};Mark-Dirty} })
$grid.Add_KeyDown({ param($s,$e); if($e.KeyCode-eq "Delete"){foreach($c in $grid.SelectedCells){if($c.ColumnIndex-eq 3){$def=$grid.Rows[$c.RowIndex].Cells[2].Value;$c.Value=$def;$c.Style=$grid.DefaultCellStyle;Mark-Dirty}}}; if($e.Control-and $e.KeyCode-eq "V"){$t=[Windows.Forms.Clipboard]::GetText().Trim();if($t){foreach($c in $grid.SelectedCells){if($c.ColumnIndex-eq 3){$c.Value=$t;if($t-eq"PAUSE"){$c.Style.BackColor="LightCoral";$c.Style.ForeColor="White"}else{$c.Style.ForeColor="Blue";$c.Style.Font=$BoldFont;$c.Style.BackColor="White"};Mark-Dirty}}}} })
function Load-GridData { $Script:IsLoading=$true; $grid.Rows.Clear(); $MapData=@{}; if(Test-Path $DateMap){Get-Content $DateMap|ForEach{if($_-match"^(\d{8})=(.+)$"){$MapData[$matches[1]]=$matches[2]}}}; $PauseData=@(); if(Test-Path $PauseLog){$PauseData=Get-Content $PauseLog}; $NoShutData=@(); if(Test-Path $NoShutdownLog){$NoShutData=Get-Content $NoShutdownLog}; $Start=(Get-Date).AddHours(-3).Date; for($i=0;$i-lt 90;$i++){ $d=$Start.AddDays($i); $dS=$d.ToString("yyyyMMdd"); $wS=$d.DayOfWeek.ToString(); $def=$Global:WeeklyRules[$wS]; $ITDay=Test-TurbulencePeriod $d; if($ITDay-gt 0){$tConf=$Global:TurbulenceRules[$wS];if($tConf){$def="$tConf"}}; $cur=$def; $isO=$false; $isP=$false; if($PauseData-contains $dS){$cur="PAUSE";$isP=$true}elseif($MapData.ContainsKey($dS)){$cur=$MapData[$dS];$isO=$true}; $isS=$NoShutData-contains $dS; if(Test-TurbulencePeriod $d){if($Global:TurbulenceNoShut[$wS]){$isS=$true}}else{if($Global:WeeklyNoShut[$wS]){$isS=$true}}; $note=""; if(Test-GenshinUpdateDay $d){$note="⚠️ 版本更新"}; if($ITDay-gt 0){$note+=" 🔥 紊亂(Day$ITDay)"}; $idx=$grid.Rows.Add($d.ToString("yyyy/MM/dd"),$wS,$def,$cur,$isS,$note); $row=$grid.Rows[$idx]; $row.Tag=$dS; if($isP){$row.Cells[3].Style.BackColor="LightCoral";$row.Cells[3].Style.ForeColor="White"}elseif($isO){$row.Cells[3].Style.ForeColor="Blue";$row.Cells[3].Style.Font=$BoldFont}; if($note){$row.Cells[5].Style.ForeColor="Magenta";$row.Cells[5].Style.Font=$BoldFont} }; $Script:IsLoading=$false; Mark-Clean }
function Save-GridData { $newMap=@(); $newP=@(); $newS=@(); foreach($r in $grid.Rows){ $k=$r.Tag; $def=$r.Cells[2].Value; $cur=$r.Cells[3].Value; $shut=$r.Cells[4].Value; if($cur-eq"PAUSE"){$newP+=$k}elseif($cur-ne$def){$newMap+="$k=$cur"}; $dObj=[DateTime]::ParseExact($k,"yyyyMMdd",$null); $wS=$dObj.DayOfWeek.ToString(); $defShut=$false; if(Test-TurbulencePeriod $dObj){if($Global:TurbulenceNoShut[$wS]){$defShut=$true}}else{if($Global:WeeklyNoShut[$wS]){$defShut=$true}}; if($shut-and-not$defShut){$newS+=$k} }; $newMap|Sort|Set-Content $DateMap -Enc UTF8; $newP|Sort|Set-Content $PauseLog -Enc UTF8; $newS|Sort|Set-Content $NoShutdownLog -Enc UTF8; Mark-Clean; [System.Windows.Forms.MessageBox]::Show("設定已儲存！"); Load-GridData; Init-WeeklyTab }
$TabGrid.Controls.Add($grid); $TabGrid.Controls.Add($pTool)

# =============================================================================
# 分頁 3: 每週配置 GUI
# =============================================================================
$TabWeekly = New-Object System.Windows.Forms.TabPage; $TabWeekly.Text = "⚙️ 每週預設設定"
$pnlW = New-Object System.Windows.Forms.Panel; $pnlW.Dock="Fill"; $pnlW.AutoScroll=$true
$DaysKey = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
$DaysTxt = @("週一","週二","週三","週四","週五","週六","週日")
$WInputs = @{}; $TInputs = @{}; $WShutChecks = @{}; $TShutChecks = @{}
function Build-WRow ($parent, $y, $txt, $key, $store, $storeCheck) {
    $l=New-Object System.Windows.Forms.Label; $l.Text=$txt; $l.Location="30,$y"; $l.AutoSize=$true; $l.Font=$MainFont
    $t=New-Object System.Windows.Forms.TextBox; $t.Location="80,$y"; $t.Width=220; $t.ReadOnly=$false; $t.Font=$MainFont
    $t.Add_TextChanged({ Mark-Dirty }) 
    $t.Add_DoubleClick({ param($s,$e); $n=Show-ConfigSelectorGUI $s.Text; if($n-ne$null){$s.Text=$n; Mark-Dirty} })
    $b=New-Object System.Windows.Forms.Button; $b.Text="選擇"; $b.Location="310,$($y-2)"; $b.Width=60; $b.Font=$MainFont; $b.Tag=$t
    $b.Add_Click({ param($s,$e); $n=Show-ConfigSelectorGUI $this.Tag.Text; if($n-ne$null){$this.Tag.Text=$n; Mark-Dirty} }.GetNewClosure())
    $parent.Controls.AddRange(@($l,$t,$b)); $store[$key]=$t
    if($storeCheck-ne$null){ $chk=New-Object System.Windows.Forms.CheckBox; $chk.Text="不關機"; $chk.Location="380,$y"; $chk.AutoSize=$true; $chk.Font=$MainFont; $parent.Controls.Add($chk); $storeCheck[$key]=$chk; $chk.Add_CheckedChanged({ Mark-Dirty }) }
}
$lblW1 = New-Object System.Windows.Forms.Label; $lblW1.Text="=== 一般每週排程 ==="; $lblW1.Location="20,20"; $lblW1.AutoSize=$true; $lblW1.Font=$BoldFont; $lblW1.ForeColor="DarkBlue"; $pnlW.Controls.Add($lblW1); $y=50; for($i=0;$i-lt 7;$i++){ Build-WRow $pnlW $y $DaysTxt[$i] $DaysKey[$i] $WInputs $WShutChecks; $y+=40 }
$y+=10; $lblW2 = New-Object System.Windows.Forms.Label; $lblW2.Text="=== 紊亂爆發期 (幽境危戰) 專用 ==="; $lblW2.Location="20,$y"; $lblW2.AutoSize=$true; $lblW2.Font=$BoldFont; $lblW2.ForeColor="DarkRed"; $lblW3 = New-Object System.Windows.Forms.Label; $lblW3.Text="(版本更新後第8~17天，優先級高於一般排程)"; $lblW3.Location="20,$($y+25)"; $lblW3.AutoSize=$true; $lblW3.Font=$MainFont; $lblW3.ForeColor="Gray"; $pnlW.Controls.AddRange(@($lblW2, $lblW3)); $y+=60; for($i=0;$i-lt 7;$i++){ Build-WRow $pnlW $y $DaysTxt[$i] $DaysKey[$i] $TInputs $TShutChecks; $y+=40 }
$y+=30; $btnWSave = New-Object System.Windows.Forms.Button; $btnWSave.Text="儲存所有設定"; $btnWSave.Location="120,$y"; $btnWSave.Size="250,50"; $btnWSave.BackColor="LightGreen"; $btnWSave.Font=$BoldFont; $btnWSave.Add_Click({ $conf=Get-JsonConf $WeeklyConf; if(-not $conf.Turbulence){$conf|Add-Member -Name "Turbulence" -Value @{} -MemberType NoteProperty}; if(-not $conf.NoShutdown){$conf|Add-Member -Name "NoShutdown" -Value @{} -MemberType NoteProperty}; if(-not $conf.Turbulence.NoShutdown){$conf.Turbulence|Add-Member -Name "NoShutdown" -Value @{} -MemberType NoteProperty}; foreach($d in $DaysKey){$conf.$d=$WInputs[$d].Text; $conf.Turbulence.$d=$TInputs[$d].Text; $conf.NoShutdown.$d=$WShutChecks[$d].Checked; $conf.Turbulence.NoShutdown.$d=$TShutChecks[$d].Checked}; if($conf.GenshinPath-eq$null){$conf|Add-Member -Name "GenshinPath" -Value $Global:GenshinPath -MemberType NoteProperty -Force}else{$conf.GenshinPath=$Global:GenshinPath}; $conf|ConvertTo-Json -Depth 4|Set-Content $WeeklyConf; Load-WeeklyRules; [System.Windows.Forms.MessageBox]::Show("設定已儲存！"); Load-GridData }); $pnlW.Controls.Add($btnWSave); $TabWeekly.Controls.Add($pnlW)
function Init-WeeklyTab { $wk=Get-JsonConf $WeeklyConf; if($wk){ foreach($d in $DaysKey){ if($WInputs.ContainsKey($d)){$WInputs[$d].Text=$wk.$d}; if($wk.Turbulence-and $TInputs.ContainsKey($d)){$TInputs[$d].Text=$wk.Turbulence.$d}; if($wk.NoShutdown-and $WShutChecks.ContainsKey($d)){$WShutChecks[$d].Checked=[bool]$wk.NoShutdown.$d}; if($wk.Turbulence.NoShutdown-and $TShutChecks.ContainsKey($d)){$TShutChecks[$d].Checked=[bool]$wk.Turbulence.NoShutdown.$d} } } }

# ... (TabResin, TabTools, TabLogs 保持 V8.1 不變) ...
# (請複製 V8.1 的剩餘代碼)
# --------------------------------------------------------
$TabResin=New-Object System.Windows.Forms.TabPage;$TabResin.Text="🧪 樹脂策略";$pnlResin=New-Object System.Windows.Forms.Panel;$pnlResin.Dock="Fill";$pnlResin.Padding="20";$lblR1=New-Object System.Windows.Forms.Label;$lblR1.Text="選擇一條龍配置組:";$lblR1.Location="20,20";$lblR1.AutoSize=$true;$lblR1.Font=$BoldFont;$cbRConfig=New-Object System.Windows.Forms.ComboBox;$cbRConfig.Location="180,18";$cbRConfig.Width=250;$cbRConfig.DropDownStyle="DropDownList";$cbRConfig.Font=$MainFont;$grpType=New-Object System.Windows.Forms.GroupBox;$grpType.Text="任務類型";$grpType.Location="20,60";$grpType.Size="200,80";$rbDomain=New-Object System.Windows.Forms.RadioButton;$rbDomain.Text="自動秘境 (Domain)";$rbDomain.Location="20,25";$rbDomain.Width=150;$rbDomain.Checked=$true;$rbStygian=New-Object System.Windows.Forms.RadioButton;$rbStygian.Text="幽境危戰 (Stygian)";$rbStygian.Location="20,50";$rbStygian.Width=150;$grpType.Controls.AddRange(@($rbDomain,$rbStygian));$grpMode=New-Object System.Windows.Forms.GroupBox;$grpMode.Text="消耗模式";$grpMode.Location="240,60";$grpMode.Size="200,80";$rbAll=New-Object System.Windows.Forms.RadioButton;$rbAll.Text="完全消耗 (All)";$rbAll.Location="20,25";$rbAll.Width=150;$rbAll.Checked=$true;$rbCount=New-Object System.Windows.Forms.RadioButton;$rbCount.Text="指定次數 (Count)";$rbCount.Location="20,50";$rbCount.Width=150;$grpMode.Controls.AddRange(@($rbAll,$rbCount));$grpCounts=New-Object System.Windows.Forms.GroupBox;$grpCounts.Text="指定次數";$grpCounts.Location="20,150";$grpCounts.Size="420,80";$lC1=New-Object System.Windows.Forms.Label;$lC1.Text="原粹:";$lC1.Location="20,30";$lC1.AutoSize=$true;$numOrig=New-Object System.Windows.Forms.NumericUpDown;$numOrig.Location="60,28";$numOrig.Width=50;$numOrig.Minimum=0;$lC2=New-Object System.Windows.Forms.Label;$lC2.Text="濃縮:";$lC2.Location="120,30";$lC2.AutoSize=$true;$numCond=New-Object System.Windows.Forms.NumericUpDown;$numCond.Location="160,28";$numCond.Width=50;$numCond.Minimum=0;$lC3=New-Object System.Windows.Forms.Label;$lC3.Text="須臾:";$lC3.Location="220,30";$lC3.AutoSize=$true;$numTran=New-Object System.Windows.Forms.NumericUpDown;$numTran.Location="260,28";$numTran.Width=50;$numTran.Minimum=0;$lC4=New-Object System.Windows.Forms.Label;$lC4.Text="脆弱:";$lC4.Location="320,30";$lC4.AutoSize=$true;$numFrag=New-Object System.Windows.Forms.NumericUpDown;$numFrag.Location="360,28";$numFrag.Width=50;$numFrag.Minimum=0;$grpCounts.Controls.AddRange(@($lC1,$numOrig,$lC2,$numCond,$lC3,$numTran,$lC4,$numFrag));$grpPrio=New-Object System.Windows.Forms.GroupBox;$grpPrio.Text="優先級";$grpPrio.Location="20,250";$grpPrio.Size="200,200";$lstPrio=New-Object System.Windows.Forms.ListBox;$lstPrio.Location="20,30";$lstPrio.Size="120,150";$bU=New-Object System.Windows.Forms.Button;$bU.Text="▲";$bU.Location="150,50";$bU.Size="30,30";$bD=New-Object System.Windows.Forms.Button;$bD.Text="▼";$bD.Location="150,100";$bD.Size="30,30";$grpPrio.Controls.AddRange(@($lstPrio,$bU,$bD));$bRS=New-Object System.Windows.Forms.Button;$bRS.Text="儲存";$bRS.Location="250,300";$bRS.Size="180,50";$bRS.BackColor="LightGreen";$bRD=New-Object System.Windows.Forms.Button;$bRD.Text="刪除";$bRD.Location="250,360";$bRD.Size="180,40";$bRD.BackColor="LightCoral";$cbRConfig.Add_DropDown({$cbRConfig.Items.Clear();$r=$Global:ConfigList|Where{$_-ne"PAUSE"};$cbRConfig.Items.AddRange($r)});$cbRConfig.Add_SelectedIndexChanged({$s=$cbRConfig.Text;if($Global:ResinData.ContainsKey($s)){$d=$Global:ResinData.$s;if($d.TaskType-eq"Stygian"){$rbStygian.Checked=$true}else{$rbDomain.Checked=$true};if($d.ResinMode-eq"Count"){$rbCount.Checked=$true}else{$rbAll.Checked=$true};$numOrig.Value=if($d.Counts.Original){$d.Counts.Original}else{0};$numCond.Value=if($d.Counts.Condensed){$d.Counts.Condensed}else{0};$numTran.Value=if($d.Counts.Transient){$d.Counts.Transient}else{0};$numFrag.Value=if($d.Counts.Fragile){$d.Counts.Fragile}else{0};$lstPrio.Items.Clear();if($d.Priority){$lstPrio.Items.AddRange($d.Priority)}else{$lstPrio.Items.AddRange(@("浓缩树脂","原粹树脂","须臾树脂","脆弱树脂"))}}else{$rbDomain.Checked=$true;$rbAll.Checked=$true;$numOrig.Value=0;$numCond.Value=0;$numTran.Value=0;$numFrag.Value=0;$lstPrio.Items.Clear();$lstPrio.Items.AddRange(@("浓缩树脂","原粹树脂","须臾树脂","脆弱树脂"))}});$bU.Add_Click({$i=$lstPrio.SelectedIndex;if($i-gt 0){$t=$lstPrio.SelectedItem;$lstPrio.Items.RemoveAt($i);$lstPrio.Items.Insert($i-1,$t);$lstPrio.SelectedIndex=$i-1}});$bD.Add_Click({$i=$lstPrio.SelectedIndex;if($i-ge 0-and $i-lt$lstPrio.Items.Count-1){$t=$lstPrio.SelectedItem;$lstPrio.Items.RemoveAt($i);$lstPrio.Items.Insert($i+1,$t);$lstPrio.SelectedIndex=$i+1}});$bRS.Add_Click({$s=$cbRConfig.Text;if(-not$s){return};$p=@();foreach($i in $lstPrio.Items){$p+=$i};$nd=@{TaskType=if($rbStygian.Checked){"Stygian"}else{"Domain"};ResinMode=if($rbCount.Checked){"Count"}else{"All"};Priority=$p;Counts=@{Original=$numOrig.Value;Condensed=$numCond.Value;Transient=$numTran.Value;Fragile=$numFrag.Value}};$Global:ResinData.$s=$nd;$Global:ResinData|ConvertTo-Json -Depth 5|Set-Content $ResinConf -Enc UTF8;[System.Windows.Forms.MessageBox]::Show("Saved")});$bRD.Add_Click({$s=$cbRConfig.Text;if($Global:ResinData.ContainsKey($s)){$Global:ResinData.Remove($s);$Global:ResinData|ConvertTo-Json -Depth 5|Set-Content $ResinConf -Enc UTF8;[System.Windows.Forms.MessageBox]::Show("Deleted");$cbRConfig.SelectedIndex=-1}});$pnlResin.Controls.AddRange(@($lblR1,$cbRConfig,$grpType,$grpMode,$grpCounts,$grpPrio,$bRS,$bRD));$TabResin.Controls.Add($pnlResin)
$TabTools=New-Object System.Windows.Forms.TabPage;$TabTools.Text="[TOOL] 工具與維護";$flpTools=New-Object System.Windows.Forms.FlowLayoutPanel;$flpTools.Dock="Fill";$flpTools.FlowDirection="TopDown";$flpTools.Padding="20";$flpTools.AutoSize=$true;function Add-ToolBtn($t,$c,$a){$b=New-Object System.Windows.Forms.Button;$b.Text=$t;$b.Width=400;$b.Height=50;$b.BackColor=$c;$b.Font=$BoldFont;$b.Margin="0,0,0,15";$b.Add_Click($a);$flpTools.Controls.Add($b)};$lblPath=New-Object System.Windows.Forms.Label;$lblPath.AutoSize=$true;$lblPath.Font=$MainFont;$lblPath.ForeColor="Gray";$lblPath.Text="目前遊戲路徑: 載入中...";$flpTools.Controls.Add($lblPath);function Update-PathLabel{$p="尚未設定";if($Global:GenshinPath){$p=$Global:GenshinPath};$lblPath.Text="目前遊戲路徑: $p"};Add-ToolBtn "📂 設定原神遊戲路徑" "LightYellow" {$f=Auto-Detect-GenshinPath;$u=$false;if($f){if([System.Windows.Forms.MessageBox]::Show("找到路徑:\n$f\n使用?","偵測","YesNo")-eq"Yes"){$Global:GenshinPath=$f;$u=$true}};if(-not$u){$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$Global:GenshinPath=$d.SelectedPath;$u=$true}};if($u){$e=@{GenshinPath=$Global:GenshinPath};$e|ConvertTo-Json|Set-Content "$ConfigsDir\EnvConfig.json";Update-PathLabel}};Add-ToolBtn "[COPY] 複製配置" "LightBlue" {$s=Show-ConfigSelectorGUI "";if($s){$s=($s-split",")[0];$n=[Microsoft.VisualBasic.Interaction]::InputBox("新名稱:","複製","$s-Copy");if($n){$src=Join-Path $BetterGI_UserDir "$s.json";$dst=Join-Path $BetterGI_UserDir "$n.json";if(Test-Path $src){Copy-Item $src $dst -Force;$j=Get-Content $dst -Raw|ConvertFrom-Json;$j.Name=$n;$j|ConvertTo-Json -Depth 20|Set-Content $dst;Load-BetterGIConfigs}}}};Add-ToolBtn "[SYNC] 同步配置名稱" "LightBlue" {$r=[System.Windows.Forms.MessageBox]::Show("修正內部 Name?","確認","YesNo");if($r-eq"Yes"){if(Test-Path $BetterGI_UserDir){Get-ChildItem "$BetterGI_UserDir\*.json"|ForEach{try{$j=Get-Content $_.FullName -Raw|ConvertFrom-Json;if($j.Name-ne$_.BaseName){$j.Name=$_.BaseName;$j|ConvertTo-Json|Set-Content $_.FullName}}catch{}}};Load-BetterGIConfigs}};Add-ToolBtn "[STOP] 強制停止" "LightCoral" {if([System.Windows.Forms.MessageBox]::Show("停止?","警","YesNo")-eq"Yes"){Start-Process powershell -Arg "-File `"$StopScript`"" -Verb RunAs}};Add-ToolBtn "[FIX] 修復權限" "LightBlue" {Start-Process powershell -Arg "-Command `"takeown /F '$Dir' /R /D Y; icacls '$Dir' /grant Everyone:(OI)(CI)F /T /C`"" -Verb RunAs};Add-ToolBtn "[GIT] 發布至 GitHub" "LightGray" {Start-Process powershell -Arg "-File `"$PublishScript`""};$TabTools.Controls.Add($flpTools)
$TabLogs=New-Object System.Windows.Forms.TabPage;$TabLogs.Text="[LOG] 日誌檢視";$pLog=New-Object System.Windows.Forms.Panel;$pLog.Dock="Top";$pLog.Height=40;$cbL=New-Object System.Windows.Forms.ComboBox;$cbL.Width=300;$cbL.Location="10,10";$cbL.DropDownStyle="DropDownList";$btnL=New-Object System.Windows.Forms.Button;$btnL.Text="讀取";$btnL.Location="320,8";$txtL=New-Object System.Windows.Forms.TextBox;$txtL.Dock="Fill";$txtL.Multiline=$true;$txtL.ScrollBars="Vertical";$txtL.ReadOnly=$true;function Refresh-LogList{$cbL.Items.Clear();if(Test-Path "$Dir\Logs"){Get-ChildItem "$Dir\Logs\*.log"|Sort LastWriteTime -Des|ForEach{$cbL.Items.Add($_.Name)}};if($cbL.Items.Count-gt 0){$cbL.SelectedIndex=0}};$btnL.Add_Click({if($cbL.SelectedItem){$p=Join-Path "$Dir\Logs" $cbL.SelectedItem;$txtL.Text=Get-Content $p -Enc UTF8|Out-String;$txtL.SelectionStart=$txtL.Text.Length;$txtL.ScrollToCaret()}});$cbL.Add_SelectedIndexChanged({$btnL.PerformClick()});$pLog.Controls.Add($cbL);$pLog.Controls.Add($btnL);$TabLogs.Controls.Add($txtL);$TabLogs.Controls.Add($pLog);$TabLogs.Add_Enter({Refresh-LogList})

# --- 組合 ---
$TabControl.Controls.AddRange(@($TabStatus, $TabGrid, $TabWeekly, $TabResin, $TabTools, $TabLogs))
$Form.Controls.Add($TabControl)
$Form.Add_Load({ Update-StatusUI; Load-GridData; Init-WeeklyTab; Load-EnvConfig; Update-PathLabel; Load-ResinConfig })
$Form.ShowDialog()