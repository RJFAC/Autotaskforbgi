# =============================================================================
# AutoTask Dashboard V7.4 - 介面排版修復 & 紊亂期不關機功能
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
$Global:TurbulenceNoShut = @{} # [新] 紊亂期不關機設定
$Global:InitialHash = ""
$Script:IsDirty = $false
$Script:IsLoading = $false
$WindowTitle = "AutoTask 控制台 V7.4"

# 字型設定
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

function Load-WeeklyRules {
    $wk = Get-JsonConf $WeeklyConf
    
    # 初始化預設結構
    $Global:WeeklyRules = @{ "Monday"="monday"; "Tuesday"="day"; "Wednesday"="day"; "Thursday"="day"; "Friday"="day"; "Saturday"="day"; "Sunday"="day" }
    $Global:TurbulenceRules = @{ "Monday"="day"; "Tuesday"="day"; "Wednesday"="day"; "Thursday"="day"; "Friday"="day"; "Saturday"="day"; "Sunday"="day" }
    $Global:WeeklyNoShut = @{ "Monday"=$false; "Tuesday"=$false; "Wednesday"=$false; "Thursday"=$false; "Friday"=$false; "Saturday"=$false; "Sunday"=$false }
    $Global:TurbulenceNoShut = @{ "Monday"=$false; "Tuesday"=$false; "Wednesday"=$false; "Thursday"=$false; "Friday"=$false; "Saturday"=$false; "Sunday"=$false }

    if ($wk) {
        # 載入一般週排程
        foreach ($k in $Global:WeeklyRules.Keys) { if ($wk.$k) { $Global:WeeklyRules[$k] = $wk.$k } }
        if ($wk.NoShutdown) {
            foreach ($k in $Global:WeeklyNoShut.Keys) { 
                if ($wk.NoShutdown.$k -ne $null) { $Global:WeeklyNoShut[$k] = [bool]$wk.NoShutdown.$k } 
            }
        }
        
        # 載入紊亂期設定
        if ($wk.Turbulence) {
            foreach ($k in $Global:TurbulenceRules.Keys) { if ($wk.Turbulence.$k) { $Global:TurbulenceRules[$k] = $wk.Turbulence.$k } }
            # [新] 載入紊亂期不關機
            if ($wk.Turbulence.NoShutdown) {
                 foreach ($k in $Global:TurbulenceNoShut.Keys) {
                    if ($wk.Turbulence.NoShutdown.$k -ne $null) { $Global:TurbulenceNoShut[$k] = [bool]$wk.Turbulence.NoShutdown.$k }
                 }
            }
        }
    }
}

# 判斷是否為版本更新日
function Test-GenshinUpdateDay ($CheckDate) {
    $RefDate = [datetime]"2024-08-28"
    $DiffDays = ($CheckDate.Date - $RefDate).Days
    if ($DiffDays -ge 0 -and ($DiffDays % 42) -eq 0) { return $true }
    return $false
}

# 判斷是否為紊亂爆發期
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
    
    # 1. 指定日期 (最高優先)
    if (Test-Path $NoShutdownLog) { if ((Get-Content $NoShutdownLog) -contains $dStr) { return "不關機 (指定)" } }
    
    # 2. [新] 紊亂期設定 (次高優先)
    if (Test-TurbulencePeriod $dateObj) {
        if ($Global:TurbulenceNoShut.$dWeek) { return "不關機 (紊亂)" }
    }
    
    # 3. 一般每週設定
    if ($Global:WeeklyNoShut.$dWeek) { return "不關機 (每週)" }
    
    return "自動關機"
}

function Get-WeekName ($dateObj) { return (@{ "Monday"="週一"; "Tuesday"="週二"; "Wednesday"="週三"; "Thursday"="週四"; "Friday"="週五"; "Saturday"="週六"; "Sunday"="週日" })[$dateObj.DayOfWeek.ToString()] }

# --- 變更追蹤 ---
function Mark-Dirty { if (-not $Script:IsLoading) { $Script:IsDirty = $true; $Form.Text = "$WindowTitle * (未儲存)" } }
function Mark-Clean { $Script:IsDirty = $false; $Form.Text = $WindowTitle }

# --- GUI 初始化 ---
Load-BetterGIConfigs
Load-WeeklyRules

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
$btnMan.Add_Click({ if([System.Windows.Forms.MessageBox]::Show("確定強制啟動？","確認","YesNo") -eq "Yes"){ New-Item -Path $ManualFlag -Force|Out-Null; Start-Process powershell -Arg "-File `"$MasterScript`"" } })
$btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="重新整理"; $btnRef.Location="20,210"; $btnRef.Width=300
$btnRef.Add_Click({ Update-StatusUI })
$TabStatus.Controls.AddRange(@($lblInfo, $btnMan, $btnRef))

function Update-StatusUI {
    $today = (Get-Date).AddHours(-3)
    $st = Get-StatusText
    $finalConf = Get-DisplayConfigName $today
    if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains $today.ToString("yyyyMMdd")) { $finalConf = "PAUSED" } }
    
    $Note = ""
    if (Test-GenshinUpdateDay $today) { $Note = " (⚠️ 版本更新日)" }
    $ITDay = Test-TurbulencePeriod $today
    if ($ITDay -gt 0) { $Note = " (🔥 紊亂期 Day $ITDay)" }

    $lblInfo.Text = "今日: $($today.ToString('yyyy/MM/dd')) ($($today.DayOfWeek))$Note`n配置: $finalConf`n狀態: $($st.Text)"
    $lblInfo.ForeColor = $st.Color
}

# === 分頁 2: 排程網格 (省略重複代碼，邏輯與前版相同) ===
$TabGrid = New-Object System.Windows.Forms.TabPage; $TabGrid.Text = "[GRID] 排程編輯器"
$pTool = New-Object System.Windows.Forms.Panel; $pTool.Dock="Top"; $pTool.Height=40
$btnGSave = New-Object System.Windows.Forms.Button; $btnGSave.Text="[SAVE]"; $btnGSave.Dock="Left"; $btnGSave.Width=100; $btnGSave.BackColor="LightGreen"; $btnGSave.Font=$BoldFont
$btnGSave.Add_Click({ Save-GridData })
$lblHint = New-Object System.Windows.Forms.Label; $lblHint.Text="操作提示: 批量勾選不關機 (Ctrl/Shift) | 雙擊配置欄 | Ctrl+C/V | Del"; $lblHint.Dock="Fill"; $lblHint.TextAlign="MiddleLeft"; $lblHint.Padding="10,0,0,0"; $lblHint.Font=$MainFont
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

function Load-GridData {
    $Script:IsLoading = $true; $grid.Rows.Clear()
    $MapData = @{}; if (Test-Path $DateMap) { Get-Content $DateMap | ForEach { if ($_ -match "^(\d{8})=(.+)$") { $MapData[$matches[1]] = $matches[2] } } }
    $PauseData = @(); if (Test-Path $PauseLog) { $PauseData = Get-Content $PauseLog }
    $NoShutData = @(); if (Test-Path $NoShutdownLog) { $NoShutData = Get-Content $NoShutdownLog }
    $Start = (Get-Date).AddHours(-3).Date
    for ($i=0; $i -lt 90; $i++) {
        $d=$Start.AddDays($i); $dS=$d.ToString("yyyyMMdd"); $wS=$d.DayOfWeek.ToString()
        
        # [更新] 顯示邏輯 (紊亂期 > 一般)
        $def=$Global:WeeklyRules[$wS]
        $ITDay = Test-TurbulencePeriod $d
        if ($ITDay -gt 0) { 
             $tConf = $Global:TurbulenceRules[$wS]
             if ($tConf) { $def = "$tConf" }
        }

        $cur=$def; $isO=$false; $isP=$false
        if ($PauseData -contains $dS) { $cur="PAUSE"; $isP=$true } elseif ($MapData.ContainsKey($dS)) { $cur=$MapData[$dS]; $isO=$true }
        
        # [更新] 不關機顯示邏輯
        $isS = $NoShutData -contains $dS
        if (Test-TurbulencePeriod $d) {
            if ($Global:TurbulenceNoShut[$wS]) { $isS = $true }
        } else {
            if ($Global:WeeklyNoShut[$wS]) { $isS = $true }
        }

        $note=""; if(Test-GenshinUpdateDay $d){$note="⚠️ 版本更新"}
        if ($ITDay -gt 0) { $note += " 🔥 紊亂(Day$ITDay)" }

        $idx=$grid.Rows.Add($d.ToString("yyyy/MM/dd"), $wS, $def, $cur, $isS, $note)
        $row=$grid.Rows[$idx]; $row.Tag=$dS
        if($isP){$row.Cells[3].Style.BackColor="LightCoral";$row.Cells[3].Style.ForeColor="White"}elseif($isO){$row.Cells[3].Style.ForeColor="Blue";$row.Cells[3].Style.Font=$BoldFont}
        if($note){$row.Cells[5].Style.ForeColor="Magenta";$row.Cells[5].Style.Font=$BoldFont}
    }
    $Script:IsLoading=$false; Mark-Clean
}

function Save-GridData {
    $newMap=@(); $newP=@(); $newS=@()
    foreach ($r in $grid.Rows) {
        $k=$r.Tag; $def=$r.Cells[2].Value; $cur=$r.Cells[3].Value; $shut=$r.Cells[4].Value
        if ($cur-eq"PAUSE") { $newP+=$k } elseif ($cur-ne$def) { $newMap+="$k=$cur" }
        
        # [更新] 存檔時判斷是否與預設值不同
        $dObj = [DateTime]::ParseExact($k, "yyyyMMdd", $null)
        $wDay = $dObj.DayOfWeek.ToString()
        $defShut = $false
        
        if (Test-TurbulencePeriod $dObj) {
             if ($Global:TurbulenceNoShut[$wDay]) { $defShut = $true }
        } else {
             if ($Global:WeeklyNoShut[$wDay]) { $defShut = $true }
        }
        
        # 只儲存 "額外指定的不關機" (如果預設關機但這裡勾選了)
        # 目前邏輯簡化：只要有勾且非預設，就寫入 NoShutdown.log
        if ($shut -and -not $defShut) { $newS += $k }
    }
    $newMap|Sort|Set-Content $DateMap -Enc UTF8; $newP|Sort|Set-Content $PauseLog -Enc UTF8; $newS|Sort|Set-Content $NoShutdownLog -Enc UTF8
    Mark-Clean; [System.Windows.Forms.MessageBox]::Show("設定已儲存！"); Load-GridData; Init-WeeklyTab
}
$TabGrid.Controls.Add($grid); $TabGrid.Controls.Add($pTool)

# =============================================================================
# 分頁 3: 每週配置 GUI (V7.2 - 雙區塊 + 不關機設定)
# =============================================================================
$TabWeekly = New-Object System.Windows.Forms.TabPage; $TabWeekly.Text = "⚙️ 每週預設設定"
$pnlW = New-Object System.Windows.Forms.Panel; $pnlW.Dock="Fill"; $pnlW.AutoScroll=$true
$DaysKey = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
$DaysTxt = @("週一","週二","週三","週四","週五","週六","週日")
$WInputs = @{}; $TInputs = @{}; $WShutChecks = @{}; $TShutChecks = @{}

function Build-WRow ($parent, $y, $txt, $key, $store, $storeCheck) {
    $l=New-Object System.Windows.Forms.Label; $l.Text=$txt; $l.Location="30,$y"; $l.AutoSize=$true; $l.Font=$MainFont
    $t=New-Object System.Windows.Forms.TextBox; $t.Location="80,$y"; $t.Width=220; $t.ReadOnly=$true; $t.Font=$MainFont
    $b=New-Object System.Windows.Forms.Button; $b.Text="選擇"; $b.Location="310,$($y-2)"; $b.Width=60; $b.Font=$MainFont; $b.Tag=$t
    $b.Add_Click({ param($s,$e); $n=Show-ConfigSelectorGUI $this.Tag.Text; if($n-ne$null){$this.Tag.Text=$n} }.GetNewClosure())
    $parent.Controls.AddRange(@($l,$t,$b))
    $store[$key] = $t
    
    # [修正] 不關機 Checkbox 位置調整
    if ($storeCheck -ne $null) {
        $chk = New-Object System.Windows.Forms.CheckBox; $chk.Text="不關機"; $chk.Location="380,$y"; $chk.AutoSize=$true; $chk.Font=$MainFont
        $parent.Controls.Add($chk)
        $storeCheck[$key] = $chk
    }
}

$lblW1 = New-Object System.Windows.Forms.Label; $lblW1.Text="=== 一般每週排程 ==="; $lblW1.Location="20,20"; $lblW1.AutoSize=$true; $lblW1.Font=$BoldFont; $lblW1.ForeColor="DarkBlue"
$pnlW.Controls.Add($lblW1)
$y=50
for($i=0;$i-lt 7;$i++) { Build-WRow $pnlW $y $DaysTxt[$i] $DaysKey[$i] $WInputs $WShutChecks; $y+=40 }

$y+=10
$lblW2 = New-Object System.Windows.Forms.Label; $lblW2.Text="=== 紊亂爆發期 (幽境危戰) 專用 ==="; $lblW2.Location="20,$y"; $lblW2.AutoSize=$true; $lblW2.Font=$BoldFont; $lblW2.ForeColor="DarkRed"
$lblW3 = New-Object System.Windows.Forms.Label; $lblW3.Text="(版本更新後第8~17天，優先級高於一般排程)"; $lblW3.Location="20,$($y+25)"; $lblW3.AutoSize=$true; $lblW3.Font=$MainFont; $lblW3.ForeColor="Gray"
$pnlW.Controls.AddRange(@($lblW2, $lblW3))
$y+=60
# [修正] 這裡也傳入 TShutChecks
for($i=0;$i-lt 7;$i++) { Build-WRow $pnlW $y $DaysTxt[$i] $DaysKey[$i] $TInputs $TShutChecks; $y+=40 }

# [修正] 儲存按鈕移到底部，避免重疊
$y+=30
$btnWSave = New-Object System.Windows.Forms.Button; $btnWSave.Text="儲存所有設定"; $btnWSave.Location="120,$y"; $btnWSave.Size="250,50"; $btnWSave.BackColor="LightGreen"; $btnWSave.Font=$BoldFont
$btnWSave.Add_Click({
    $conf = Get-JsonConf $WeeklyConf
    if (-not $conf.Turbulence) { $conf | Add-Member -Name "Turbulence" -Value @{} -MemberType NoteProperty }
    if (-not $conf.NoShutdown) { $conf | Add-Member -Name "NoShutdown" -Value @{} -MemberType NoteProperty }
    
    # 確保 Turbulence.NoShutdown 存在
    if (-not $conf.Turbulence.NoShutdown) { $conf.Turbulence | Add-Member -Name "NoShutdown" -Value @{} -MemberType NoteProperty }

    foreach ($d in $DaysKey) { 
        $conf.$d = $WInputs[$d].Text 
        $conf.Turbulence.$d = $TInputs[$d].Text
        $conf.NoShutdown.$d = $WShutChecks[$d].Checked
        $conf.Turbulence.NoShutdown.$d = $TShutChecks[$d].Checked
    }
    $conf | ConvertTo-Json -Depth 4 | Set-Content $WeeklyConf # Depth 4 確保巢狀物件正確儲存
    Load-WeeklyRules; [System.Windows.Forms.MessageBox]::Show("設定已儲存！"); Load-GridData
})
$pnlW.Controls.Add($btnWSave); $TabWeekly.Controls.Add($pnlW)

function Init-WeeklyTab {
    $wk = Get-JsonConf $WeeklyConf
    if ($wk) {
        foreach ($d in $DaysKey) {
            if ($WInputs.ContainsKey($d)) { $WInputs[$d].Text = $wk.$d }
            if ($wk.Turbulence -and $TInputs.ContainsKey($d)) { $TInputs[$d].Text = $wk.Turbulence.$d }
            if ($wk.NoShutdown -and $WShutChecks.ContainsKey($d)) { $WShutChecks[$d].Checked = [bool]$wk.NoShutdown.$d }
            
            # [修正] 讀取紊亂期不關機
            if ($wk.Turbulence.NoShutdown -and $TShutChecks.ContainsKey($d)) {
                $TShutChecks[$d].Checked = [bool]$wk.Turbulence.NoShutdown.$d
            }
        }
    }
}

# --- Config Selector ---
function Show-ConfigSelectorGUI {
    param([string]$CurrentSelection) 
    $SelForm = New-Object System.Windows.Forms.Form; $SelForm.Text="配置選擇 (拖曳排序)"; $SelForm.Size="700,500"; $SelForm.StartPosition="CenterParent"; $SelForm.Font=$MainFont
    $lblSrc = New-Object System.Windows.Forms.Label; $lblSrc.Text="可用配置"; $lblSrc.Location="20,10"; $lblSrc.AutoSize=$true
    $listSrc = New-Object System.Windows.Forms.ListBox; $listSrc.Location="20,30"; $listSrc.Size="250,350"; $listSrc.SelectionMode="MultiExtended"
    $RealConfigs = $Global:ConfigList | Where-Object { $_ -ne "PAUSE" }; $listSrc.Items.AddRange($RealConfigs)
    $lblDst = New-Object System.Windows.Forms.Label; $lblDst.Text="執行佇列"; $lblDst.Location="380,10"; $lblDst.AutoSize=$true
    $listDst = New-Object System.Windows.Forms.ListBox; $listDst.Location="380,30"; $listDst.Size="250,350"; $listDst.SelectionMode="One"; $listDst.AllowDrop=$true 
    if (-not [string]::IsNullOrWhiteSpace($CurrentSelection) -and $CurrentSelection -ne "PAUSE") { $parts = $CurrentSelection -split ","; foreach ($p in $parts) { if($p){$listDst.Items.Add($p)} } }
    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text="加入 ->"; $btnAdd.Location="280,150"; $btnAdd.Size="90,30"; $btnAdd.Add_Click({ foreach ($item in $listSrc.SelectedItems) { $listDst.Items.Add($item) } })
    $btnRem = New-Object System.Windows.Forms.Button; $btnRem.Text="<- 移除"; $btnRem.Location="280,200"; $btnRem.Size="90,30"; $btnRem.Add_Click({ if ($listDst.SelectedIndex -ge 0) { $listDst.Items.RemoveAt($listDst.SelectedIndex) } })
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text="確定"; $btnOk.Location="250,400"; $btnOk.DialogResult="OK"; $btnOk.BackColor="LightGreen"
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text="取消"; $btnCancel.Location="360,400"; $btnCancel.DialogResult="Cancel"
    $listDst.Add_MouseDown({ param($s,$e); if($listDst.SelectedItem) { $listDst.DoDragDrop($listDst.SelectedItem, [System.Windows.Forms.DragDropEffects]::Move) } })
    $listDst.Add_DragOver({ param($s,$e); $e.Effect=[System.Windows.Forms.DragDropEffects]::Move })
    $listDst.Add_DragDrop({ param($s,$e); $idx=$listDst.IndexFromPoint($listDst.PointToClient([System.Drawing.Point]::new($e.X,$e.Y))); if($idx -lt 0){$idx=$listDst.Items.Count-1}; $item=$e.Data.GetData([string]); if($item){$listDst.Items.Remove($item); $listDst.Items.Insert($idx,$item); $listDst.SelectedIndex=$idx} })
    $SelForm.Controls.AddRange(@($lblSrc, $listSrc, $lblDst, $listDst, $btnAdd, $btnRem, $btnOk, $btnCancel)); $SelForm.AcceptButton = $btnOk; if ($SelForm.ShowDialog() -eq "OK") { $f=@(); foreach($i in $listDst.Items){$f+=$i}; return ($f -join ",") } else { return $null }
}

# ... (TabTools 與 TabLogs 保持不變) ...
$TabTools = New-Object System.Windows.Forms.TabPage; $TabTools.Text = "[TOOL] 工具與維護" 
$flpTools = New-Object System.Windows.Forms.FlowLayoutPanel; $flpTools.Dock="Fill"; $flpTools.FlowDirection="TopDown"; $flpTools.Padding="20"; $flpTools.AutoSize=$true
function Add-ToolBtn ($text, $color, $action) { $btn = New-Object System.Windows.Forms.Button; $btn.Text=$text; $btn.Width=400; $btn.Height=50; $btn.BackColor=$color; $btn.Font=$BoldFont; $btn.Margin="0,0,0,15"; $btn.Add_Click($action); $flpTools.Controls.Add($btn) }
Add-ToolBtn "[STOP] 強制停止所有任務" "LightCoral" { if([System.Windows.Forms.MessageBox]::Show("確定停止？","警告","YesNo")-eq"Yes"){ Start-Process powershell -Arg "-NoProfile -ExecutionPolicy Bypass -File `"$StopScript`"" -Verb RunAs } }
Add-ToolBtn "[FIX] 修復檔案權限" "LightBlue" { Start-Process powershell -Arg "-Command `"takeown /F '$Dir' /R /D Y; icacls '$Dir' /grant Everyone:(OI)(CI)F /T /C`"" -Verb RunAs; [System.Windows.Forms.MessageBox]::Show("完成") }
Add-ToolBtn "[GIT] 發布至 GitHub" "LightGray" { if([System.Windows.Forms.MessageBox]::Show("確定發布？","確認","YesNo")-eq"Yes"){ Start-Process powershell -Arg "-NoProfile -ExecutionPolicy Bypass -File `"$PublishScript`"" } }
Add-ToolBtn "[RDP] 修復 RDP 最小化" "LightGray" { Start-Process powershell -Arg "-Command `"reg add 'HKLM\Software\Microsoft\Terminal Server Client' /v 'RemoteDesktop_SuppressWhenMinimized' /t REG_DWORD /d 2 /f`"" -Verb RunAs; [System.Windows.Forms.MessageBox]::Show("完成") }
$TabTools.Controls.Add($flpTools)

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
$TabControl.Controls.AddRange(@($TabStatus, $TabGrid, $TabWeekly, $TabTools, $TabLogs))
$Form.Controls.Add($TabControl)
$Form.Add_Load({ Update-StatusUI; Load-GridData; Init-WeeklyTab })
$Form.ShowDialog()