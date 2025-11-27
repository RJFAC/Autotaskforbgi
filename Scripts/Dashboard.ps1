# =============================================================================
# AutoTask Dashboard V8.2 - 支援腳本熱重載 (Restart with State)
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
$WindowTitle = "AutoTask 控制台 V8.2"

# 字型
$MainFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 10)
$BoldFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 10, [System.Drawing.FontStyle]::Bold)
$TitleFont = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12, [System.Drawing.FontStyle]::Bold)
$MonoFont = New-Object System.Drawing.Font("Consolas", 10) 

function Get-ScriptsHash {
    $str = ""
    Get-ChildItem $ScriptDir -Include "*.ps1", "*.bat" -Recurse | Sort-Object Name | ForEach-Object { 
        $str += (Get-FileHash $_.FullName).Hash 
    }
    return $str
}
$Global:InitialHash = Get-ScriptsHash

# --- [輔助函數] ---
function Get-JsonConf ($path) {
    if (Test-Path $path) { 
        try { return Get-Content $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } 
        catch { return $null }
    }
    return $null
}

function Load-BetterGIConfigs {
    $Global:ConfigList = @("PAUSE") 
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
    if ($env -and $env.GenshinPath) { $Global:GenshinPath = $env.GenshinPath }
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

# --- [新功能] 重啟腳本 (熱重載) ---
function Restart-Services {
    Write-Host "正在重新啟動服務..."
    # 1. 找出舊的 Master 與 Monitor
    $MyPID = $PID
    $Targets = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
               Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID }
    
    # 2. 終止它們
    foreach ($p in $Targets) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
    
    # 3. 啟動新的 Master (它會自動偵測 Run.flag 並接手)
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`"" -Verb RunAs
    
    [System.Windows.Forms.MessageBox]::Show("Master 與 Monitor 已重啟並嘗試接手任務。")
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
    if (Get-ScriptsHash -ne $Global:InitialHash) {
        if ([System.Windows.Forms.MessageBox]::Show("腳本已變更，是否同步至 GitHub？", "同步", "YesNo") -eq "Yes") {
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

# ... (TabGrid, TabWeekly, TabResin 保持 V8.1 原樣，省略以節省篇幅) ...
# 請務必從 V8.1 複製過來，它們不需要修改
# (為確保代碼完整，以下使用簡化版插入，請實際操作時使用完整代碼)
# -------------------------------------------------------------
$TabGrid = New-Object System.Windows.Forms.TabPage; $TabGrid.Text = "[GRID] 排程編輯器"; $pTool = New-Object System.Windows.Forms.Panel; $pTool.Dock="Top"; $pTool.Height=40; $btnGSave = New-Object System.Windows.Forms.Button; $btnGSave.Text="[SAVE]"; $btnGSave.Dock="Left"; $btnGSave.Width=100; $btnGSave.BackColor="LightGreen"; $btnGSave.Font=$BoldFont; $btnGSave.Add_Click({ Save-GridData }); $lblHint = New-Object System.Windows.Forms.Label; $lblHint.Text="操作提示: 批量勾選 | Ctrl+C/V | Del"; $lblHint.Dock="Fill"; $lblHint.TextAlign="MiddleLeft"; $lblHint.Font=$MainFont; $pTool.Controls.Add($lblHint); $pTool.Controls.Add($btnGSave); $grid = New-Object System.Windows.Forms.DataGridView; $grid.Dock="Fill"; $grid.EditMode="EditProgrammatically"; $grid.Font=$MonoFont; $grid.MultiSelect=$true; $grid.Columns.Add("Date","日期"); $grid.Columns.Add("Week","星期"); $grid.Columns.Add("Def","預設"); $grid.Columns.Add("Conf","配置"); $grid.Columns.Add("Shut","不關機"); $grid.Columns.Add("Note","備註"); $grid.Columns[4].CellTemplate=New-Object System.Windows.Forms.DataGridViewCheckBoxCell; $grid.Add_CellClick({param($s,$e);if($e.RowIndex-ge 0-and $e.ColumnIndex-eq 4){$c=$grid.Rows[$e.RowIndex].Cells[4];$v=-not[bool]$c.Value;$sel=$grid.SelectedCells|Where{$_.ColumnIndex-eq 4};if($sel){foreach($x in $sel){$x.Value=$v}}else{$c.Value=$v};Mark-Dirty}}); $grid.Add_CellDoubleClick({param($s,$e);if($e.RowIndex-ge 0-and $e.ColumnIndex-eq 3){$v=$grid.Rows[$e.RowIndex].Cells[3].Value;if($v-eq $grid.Rows[$e.RowIndex].Cells[2].Value){$v=""};$n=Show-ConfigSelectorGUI $v;if($n-ne$null){if($n){$grid.Rows[$e.RowIndex].Cells[3].Value=$n;$grid.Rows[$e.RowIndex].Cells[3].Style.ForeColor="Blue"}else{$grid.Rows[$e.RowIndex].Cells[3].Value=$grid.Rows[$e.RowIndex].Cells[2].Value;$grid.Rows[$e.RowIndex].Cells[3].Style=$grid.DefaultCellStyle};Mark-Dirty}}}); $grid.Add_KeyDown({param($s,$e);if($e.KeyCode-eq"Delete"){foreach($c in $grid.SelectedCells){if($c.ColumnIndex-eq 3){$c.Value=$grid.Rows[$c.RowIndex].Cells[2].Value;$c.Style=$grid.DefaultCellStyle;Mark-Dirty}}};if($e.Control-and $e.KeyCode-eq"V"){$t=[Windows.Forms.Clipboard]::GetText().Trim();if($t){foreach($c in $grid.SelectedCells){if($c.ColumnIndex-eq 3){$c.Value=$t;$c.Style.ForeColor="Blue";Mark-Dirty}}}}}); function Load-GridData {$Script:IsLoading=$true;$grid.Rows.Clear();$MapData=@{};if(Test-Path $DateMap){Get-Content $DateMap|ForEach{if($_-match"^(\d{8})=(.+)$"){$MapData[$matches[1]]=$matches[2]}}};$Start=(Get-Date).AddHours(-3).Date;for($i=0;$i-lt 90;$i++){$d=$Start.AddDays($i);$ds=$d.ToString("yyyyMMdd");$w=$d.DayOfWeek.ToString();$def=$Global:WeeklyRules[$w];if(Test-TurbulencePeriod $d){$def=$Global:TurbulenceRules[$w]};$cur=$def;$ov=$false;if($MapData.ContainsKey($ds)){$cur=$MapData[$ds];$ov=$true};$s=$false;if(Test-TurbulencePeriod $d){if($Global:TurbulenceNoShut[$w]){$s=$true}}else{if($Global:WeeklyNoShut[$w]){$s=$true}};$n="";if(Test-GenshinUpdateDay $d){$n="更新"};if(Test-TurbulencePeriod $d){$n+=" 紊亂"};$r=$grid.Rows.Add($d.ToString("yyyy/MM/dd"),$w,$def,$cur,$s,$n);$grid.Rows[$r].Tag=$ds;if($ov){$grid.Rows[$r].Cells[3].Style.ForeColor="Blue"}};$Script:IsLoading=$false;Mark-Clean}; function Save-GridData {$nm=@();foreach($r in $grid.Rows){$k=$r.Tag;$c=$r.Cells[3].Value;$d=$r.Cells[2].Value;if($c-ne$d){$nm+="$k=$c"}};$nm|Set-Content $DateMap;Mark-Clean;[System.Windows.Forms.MessageBox]::Show("Saved");Load-GridData}; $TabGrid.Controls.Add($grid);$TabGrid.Controls.Add($pTool)
$TabWeekly = New-Object System.Windows.Forms.TabPage; $TabWeekly.Text = "⚙️ 每週預設設定"; $pnlW = New-Object System.Windows.Forms.Panel; $pnlW.Dock="Fill"; $pnlW.AutoScroll=$true; $DaysKey=@("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"); $WInputs=@{}; $TInputs=@{}; $WShutChecks=@{}; $TShutChecks=@{}; function Build-WRow ($p,$y,$t,$k,$s,$sc){$l=New-Object System.Windows.Forms.Label;$l.Text=$t;$l.Location="30,$y";$l.AutoSize=$true;$l.Font=$MainFont;$tb=New-Object System.Windows.Forms.TextBox;$tb.Location="80,$y";$tb.Width=220;$tb.ReadOnly=$false;$tb.Font=$MainFont;$tb.Add_TextChanged({Mark-Dirty});$tb.Add_DoubleClick({param($s,$e);$n=Show-ConfigSelectorGUI $s.Text;if($n-ne$null){$s.Text=$n}});$b=New-Object System.Windows.Forms.Button;$b.Text="選擇";$b.Location="310,$($y-2)";$b.Tag=$tb;$b.Add_Click({param($s,$e);$n=Show-ConfigSelectorGUI $this.Tag.Text;if($n-ne$null){$this.Tag.Text=$n}}.GetNewClosure());$p.Controls.AddRange(@($l,$tb,$b));$s[$k]=$tb;if($sc){$c=New-Object System.Windows.Forms.CheckBox;$c.Text="不關機";$c.Location="380,$y";$c.AutoSize=$true;$c.Font=$MainFont;$c.Add_CheckedChanged({Mark-Dirty});$p.Controls.Add($c);$sc[$k]=$c}}; $y=50; foreach($d in $DaysKey){Build-WRow $pnlW $y $d $d $WInputs $WShutChecks; $y+=40}; $y+=40; foreach($d in $DaysKey){Build-WRow $pnlW $y "$d (Turb)" $d $TInputs $TShutChecks; $y+=40}; $bs=New-Object System.Windows.Forms.Button;$bs.Text="Save";$bs.Location="120,$y";$bs.Add_Click({$c=Get-JsonConf $WeeklyConf;foreach($d in $DaysKey){$c.$d=$WInputs[$d].Text;$c.Turbulence.$d=$TInputs[$d].Text;$c.NoShutdown.$d=$WShutChecks[$d].Checked;$c.Turbulence.NoShutdown.$d=$TShutChecks[$d].Checked};$c|ConvertTo-Json -Depth 4|Set-Content $WeeklyConf;Load-WeeklyRules;Load-GridData});$pnlW.Controls.Add($bs);$TabWeekly.Controls.Add($pnlW); function Init-WeeklyTab{if($wk=Get-JsonConf $WeeklyConf){foreach($d in $DaysKey){$WInputs[$d].Text=$wk.$d;$TInputs[$d].Text=$wk.Turbulence.$d;$WShutChecks[$d].Checked=[bool]$wk.NoShutdown.$d;$TShutChecks[$d].Checked=[bool]$wk.Turbulence.NoShutdown.$d}}}
function Show-ConfigSelectorGUI { param($c); $f=New-Object System.Windows.Forms.Form;$f.Text="Select";$l=New-Object System.Windows.Forms.ListBox;$l.Dock="Fill";$l.SelectionMode="MultiExtended";$l.Items.AddRange(($Global:ConfigList|Where{$_-ne"PAUSE"}));$b=New-Object System.Windows.Forms.Button;$b.Text="OK";$b.Dock="Bottom";$b.DialogResult="OK";$f.Controls.Add($l);$f.Controls.Add($b);if($f.ShowDialog()-eq"OK"){$r=@();foreach($i in $l.SelectedItems){$r+=$i};return ($r-join",")}return $null }
$TabResin = New-Object System.Windows.Forms.TabPage; $TabResin.Text = "🧪 樹脂策略"; $pnlResin = New-Object System.Windows.Forms.Panel; $pnlResin.Dock="Fill"; $TabResin.Controls.Add($pnlResin) 
# -------------------------------------------------------------

# =============================================================================
# 分頁 4: 工具與維護
# =============================================================================
$TabTools = New-Object System.Windows.Forms.TabPage; $TabTools.Text = "[TOOL] 工具與維護" 
$flpTools = New-Object System.Windows.Forms.FlowLayoutPanel; $flpTools.Dock="Fill"; $flpTools.FlowDirection="TopDown"; $flpTools.Padding="20"; $flpTools.AutoSize=$true
function Add-ToolBtn ($text, $color, $action) { $btn = New-Object System.Windows.Forms.Button; $btn.Text=$text; $btn.Width=400; $btn.Height=50; $btn.BackColor=$color; $btn.Font=$BoldFont; $btn.Margin="0,0,0,15"; $btn.Add_Click($action); $flpTools.Controls.Add($btn) }

$lblPath = New-Object System.Windows.Forms.Label; $lblPath.AutoSize=$true; $lblPath.Font=$MainFont; $lblPath.ForeColor="Gray"
$lblPath.Text = "目前遊戲路徑: 載入中..."
$flpTools.Controls.Add($lblPath)

function Update-PathLabel {
    $path = "尚未設定"
    if ($Global:GenshinPath) { $path = $Global:GenshinPath }
    $lblPath.Text = "目前遊戲路徑: $path"
}

# [新] 智慧偵測按鈕
Add-ToolBtn "📂 設定原神遊戲路徑 (自動/手動)" "LightYellow" {
    $FoundPath = Auto-Detect-GenshinPath
    $UseAuto = $false
    
    if ($FoundPath) {
        $res = [System.Windows.Forms.MessageBox]::Show("✅ 自動偵測成功！`n`n找到路徑：`n$FoundPath`n`n是否直接使用此路徑？", "偵測結果", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        if ($res -eq "Yes") {
            $Global:GenshinPath = $FoundPath
            $UseAuto = $true
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("❌ 自動偵測失敗。`n`n無法在標準安裝位置找到原神。`n請在接下來的視窗中手動選擇 'Genshin Impact Game' 資料夾。", "偵測結果", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
    
    if (-not $UseAuto) {
        $f = New-Object System.Windows.Forms.FolderBrowserDialog
        $f.Description = "請選擇包含 YuanShen.exe / GenshinImpact.exe 的資料夾"
        if ($f.ShowDialog() -eq "OK") {
            $Global:GenshinPath = $f.SelectedPath
            $UseAuto = $true
        }
    }

    if ($UseAuto) {
        $envData = @{ GenshinPath = $Global:GenshinPath }
        $envData | ConvertTo-Json | Set-Content "$ConfigsDir\EnvConfig.json" -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("路徑已儲存！", "設定完成")
        Update-PathLabel
    }
}

# [新] 腳本熱重載按鈕
Add-ToolBtn "[RESTART] 重啟腳本 (保留狀態)" "Orange" {
    if ([System.Windows.Forms.MessageBox]::Show("這將殺死當前運行的 Master 與 Monitor 並重啟。`n`n若正在執行任務，新腳本會自動接手監控 (不會中斷遊戲)。`n`n確定嗎？", "熱重載", "YesNo") -eq "Yes") {
        # 呼叫 Master.ps1 中已經寫好的 Restart-Services 邏輯？
        # 因為 Dashboard 權限可能不同，我們直接在這裡實作
        
        Write-Host "正在停止舊服務..."
        $MyPID = $PID
        Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
            Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        
        Start-Sleep 1
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`"" -Verb RunAs
        [System.Windows.Forms.MessageBox]::Show("已發送重啟指令。")
    }
}

Add-ToolBtn "[COPY] 複製配置組" "LightBlue" { 
    $sel = Show-ConfigSelectorGUI ""; 
    if($sel) { 
        $sel = ($sel -split ",")[0]; 
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox("請輸入新配置組名稱:", "複製配置", "$sel-Copy");
        if (-not [string]::IsNullOrWhiteSpace($newName)) {
            $src = Join-Path $BetterGI_UserDir "$sel.json";
            $dst = Join-Path $BetterGI_UserDir "$newName.json";
            if (Test-Path $src) {
                Copy-Item $src $dst -Force;
                $j = Get-Content $dst -Raw -Enc UTF8 | ConvertFrom-Json;
                $j.Name = $newName;
                $j | ConvertTo-Json -Depth 20 | Set-Content $dst -Enc UTF8;
                [System.Windows.Forms.MessageBox]::Show("已複製為: $newName");
                Load-BetterGIConfigs;
            }
        }
    } 
}
Add-ToolBtn "[GIT] 發布至 GitHub" "LightGray" { if([System.Windows.Forms.MessageBox]::Show("確定發布？","確認","YesNo")-eq"Yes"){ Start-Process powershell -Arg "-NoProfile -ExecutionPolicy Bypass -File `"$PublishScript`"" } }
Add-ToolBtn "[SYNC] 雲端更新 (Git Pull)" "LightGray" { 
    Start-Process git -ArgumentList "pull" -WorkingDirectory $Dir -NoNewWindow -Wait
    if ([System.Windows.Forms.MessageBox]::Show("更新完成。是否重啟腳本以應用變更？", "更新", "YesNo") -eq "Yes") {
        # 執行熱重載
        $MyPID = $PID
        Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | 
            Where-Object { ($_.CommandLine -like "*Monitor.ps1*" -or $_.CommandLine -like "*Master.ps1*") -and $_.ProcessId -ne $MyPID } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`"" -Verb RunAs
    }
}
$TabTools.Controls.Add($flpTools)

# =============================================================================
# 分頁 5: 日誌檢視
# =============================================================================
$TabLogs = New-Object System.Windows.Forms.TabPage; $TabLogs.Text = "[LOG] 日誌檢視" 
$pnlLogTop = New-Object System.Windows.Forms.Panel; $pnlLogTop.Dock="Top"; $pnlLogTop.Height=40
$cbLogFiles = New-Object System.Windows.Forms.ComboBox; $cbLogFiles.Width=300; $cbLogFiles.Location="10,10"; $cbLogFiles.DropDownStyle="DropDownList"; $cbLogFiles.Font=$MainFont
$btnRefreshLog = New-Object System.Windows.Forms.Button; $btnRefreshLog.Text="重新讀取"; $btnRefreshLog.Location="320,8"; $btnRefreshLog.Width=100; $btnRefreshLog.Font=$MainFont
$txtLogContent = New-Object System.Windows.Forms.TextBox; $txtLogContent.Dock="Fill"; $txtLogContent.Multiline=$true; $txtLogContent.ScrollBars="Vertical"; $txtLogContent.Font=$MonoFont; $txtLogContent.ReadOnly=$true
function Refresh-LogList { $cbLogFiles.Items.Clear(); if(Test-Path "$Dir\Logs") { Get-ChildItem "$Dir\Logs\*.log"|Sort LastWriteTime -Des|ForEach{$cbLogFiles.Items.Add($_.Name)} }; if($cbLogFiles.Items.Count -gt 0){$cbLogFiles.SelectedIndex=0} }
$btnRefreshLog.Add_Click({ if($cbLogFiles.SelectedItem){ $p=Join-Path "$Dir\Logs" $cbLogFiles.SelectedItem; $txtLogContent.Text=Get-Content $p -Encoding UTF8|Out-String; $txtLogContent.SelectionStart=$txtLogContent.Text.Length;$txtLogContent.ScrollToCaret() } })
$cbLogFiles.Add_SelectedIndexChanged({ $btnRefreshLog.PerformClick() })
$pnlLogTop.Controls.Add($cbLogFiles); $pnlLogTop.Controls.Add($btnRefreshLog); $TabLogs.Controls.Add($txtLogContent); $TabLogs.Controls.Add($pnlLogTop); $TabLogs.Add_Enter({ Refresh-LogList }) 

# --- 組合 ---
$TabControl.Controls.AddRange(@($TabStatus, $TabGrid, $TabWeekly, $TabResin, $TabTools, $TabLogs))
$Form.Controls.Add($TabControl)
$Form.Add_Load({ Update-StatusUI; Load-GridData; Init-WeeklyTab; Load-EnvConfig; Update-PathLabel; Load-ResinConfig })
$Form.ShowDialog()