# =============================================================================
# AutoTask Dashboard V7.2 - 終極排程管理版 (每週 Grid 化 + 顯示修復)
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
$Global:WeeklyNoShut = @{} # [新] 每週不關機設定
$Global:InitialHash = ""
$Script:IsDirty = $false
$Script:IsLoading = $false
$WindowTitle = "AutoTask 控制台 V7.2"
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
    if (Test-Path $path) { return Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json }
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

    if ($wk) {
        # 載入一般週排程
        foreach ($k in $Global:WeeklyRules.Keys) { if ($wk.$k) { $Global:WeeklyRules[$k] = $wk.$k } }
        
        # 載入紊亂期
        if ($wk.Turbulence) {
            foreach ($k in $Global:TurbulenceRules.Keys) { if ($wk.Turbulence.$k) { $Global:TurbulenceRules[$k] = $wk.Turbulence.$k } }
        }
        
        # [新] 載入每週不關機
        if ($wk.NoShutdown) {
            foreach ($k in $Global:WeeklyNoShut.Keys) { 
                if ($wk.NoShutdown.$k -ne $null) { $Global:WeeklyNoShut[$k] = [bool]$wk.NoShutdown.$k } 
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

# 判斷是否為紊亂爆發期 (回傳天數 8~18，若無則回傳 0)
function Test-TurbulencePeriod ($CheckDate) {
    $RefDate = [datetime]"2024-08-28"
    $DiffDays = ($CheckDate.Date - $RefDate).Days
    if ($DiffDays -ge 0) {
        $CycleDay = $DiffDays % 42
        if ($CycleDay -ge 8 -and $CycleDay -le 18) { return $CycleDay }
    }
    return 0
}

function Get-DisplayConfigName ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    $dWeek = $dateObj.DayOfWeek.ToString()
    
    # 1. 指定日期
    if (Test-Path $DateMap) {
        $map = Get-Content $DateMap
        foreach ($line in $map) { if ($line -match "^$dStr=(.+)$") { return "$($matches[1]) (指定)" } }
    }
    
    # 2. 紊亂爆發期
    if (Test-TurbulencePeriod $dateObj) {
        $tConf = $Global:TurbulenceRules.$dWeek
        if ($tConf) { return "$tConf (紊亂期)" }
    }

    # 3. 每週配置
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
    }
    if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains $dStr) { $txt = "已排程暫停"; $color = [System.Drawing.Color]::Orange } }
    return @{Text=$txt; Color=$color}
}

function Get-ShutdownPolicy ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    $dWeek = $dateObj.DayOfWeek.ToString()
    
    # 1. 指定日期不關機
    if (Test-Path $NoShutdownLog) { if ((Get-Content $NoShutdownLog) -contains $dStr) { return "不關機 (指定)" } }
    
    # 2. [新] 每週預設不關機
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
$Form.Size = New-Object System.Drawing.Size(1000, 750)
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
    if (Test-GenshinUpdateDay $today) { $Note = " ⚠️ 版本更新日)" }
    $ITDay = Test-TurbulencePeriod $today
    if ($ITDay -gt 0) { $Note = " (🔥 紊亂期 Day $ITDay)" }

    $lblInfo.Text = "今日: $($today.ToString('yyyy/MM/dd')) ($($today.DayOfWeek))$Note`n配置: $finalConf`n狀態: $($st.Text)"
    $lblInfo.ForeColor = $st.Color
}

# --- 通用 Grid 建構函數 ---
function Build-Grid ($parent, $isWeekly) {
    $panelTool = New-Object System.Windows.Forms.Panel; $panelTool.Dock="Top"; $panelTool.Height=40
    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text="[SAVE] 儲存"; $btnSave.Dock="Left"; $btnSave.Width=100; $btnSave.BackColor="LightGreen"; $btnSave.Font=$BoldFont
    $lblHint = New-Object System.Windows.Forms.Label; $lblHint.Dock="Fill"; $lblHint.TextAlign="MiddleLeft"; $lblHint.Padding="10,0,0,0"
    $lblHint.Text = if($isWeekly){"每週設定: 雙擊配置欄選擇 | 勾選不關機 | 支援 Ctrl+C/V"}else{"排程網格: 支援 Ctrl/Shift 批量勾選 | 雙擊配置 | Ctrl+C/V | Del"}
    
    if ($isWeekly) { 
        $btnSave.Add_Click({ Save-WeeklyGrid }) 
    } else { 
        $btnSave.Add_Click({ Save-DailyGrid }) 
    }

    $panelTool.Controls.Add($lblHint); $panelTool.Controls.Add($btnSave)

    $grid = New-Object System.Windows.Forms.DataGridView; $grid.Dock="Fill"; $grid.EditMode="EditProgrammatically"; $grid.Font=$MonoFont; $grid.MultiSelect=$true

    # 定義欄位
    if ($isWeekly) {
        $grid.Columns.Add("Day", "星期"); $grid.Columns[0].ReadOnly=$true; $grid.Columns[0].Width=100
        $grid.Columns.Add("Norm", "一般週配置 (雙擊)"); $grid.Columns[1].Width=250
        $grid.Columns.Add("Turb", "紊亂期配置 (雙擊)"); $grid.Columns[2].Width=250
        $grid.Columns.Add("Shut", "預設不關機"); $grid.Columns[3].Width=100; $grid.Columns[3].CellTemplate = New-Object System.Windows.Forms.DataGridViewCheckBoxCell
    } else {
        $grid.Columns.Add("Date", "日期"); $grid.Columns[0].ReadOnly=$true; $grid.Columns[0].Width=100
        $grid.Columns.Add("Week", "星期"); $grid.Columns[1].ReadOnly=$true; $grid.Columns[1].Width=60
        $grid.Columns.Add("Def", "每週預設"); $grid.Columns[2].ReadOnly=$true; $grid.Columns[2].Width=100
        $grid.Columns.Add("Conf", "執行配置 (雙擊)"); $grid.Columns[3].Width=250
        $grid.Columns.Add("Shut", "不關機"); $grid.Columns[4].Width=60; $grid.Columns[4].CellTemplate = New-Object System.Windows.Forms.DataGridViewCheckBoxCell
        $grid.Columns.Add("Note", "備註"); $grid.Columns[5].ReadOnly=$true; $grid.Columns[5].Width=120
    }

    # 通用事件綁定
    $grid.Add_CellClick({ param($s,$e); Handle-CellClick $s $e })
    $grid.Add_CellDoubleClick({ param($s,$e); Handle-CellDoubleClick $s $e })
    $grid.Add_KeyDown({ param($s,$e); Handle-KeyDown $s $e })
    
    $parent.Controls.Add($grid)
    $parent.Controls.Add($panelTool)
    return $grid
}

# --- 事件處理邏輯 ---
function Handle-CellClick ($grid, $e) {
    if ($e.RowIndex -lt 0) { return }
    # 判斷是否為 Checkbox 欄位 (Daily:4, Weekly:3)
    $isCheckCol = ($grid.Columns.Count -eq 6 -and $e.ColumnIndex -eq 4) -or ($grid.Columns.Count -eq 4 -and $e.ColumnIndex -eq 3)
    
    if ($isCheckCol) {
        $clickedCell = $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
        $val = -not [bool]$clickedCell.Value
        
        # 批量勾選邏輯
        $targetCells = $grid.SelectedCells | Where-Object { $_.ColumnIndex -eq $e.ColumnIndex }
        if ($targetCells.Count -gt 0 -and ($targetCells | Where-Object { $_.RowIndex -eq $e.RowIndex })) {
            foreach ($cell in $targetCells) { $cell.Value = $val }
        } else {
            $clickedCell.Value = $val
        }
        Mark-Dirty
    }
}

function Handle-CellDoubleClick ($grid, $e) {
    if ($e.RowIndex -lt 0) { return }
    # 判斷配置欄 (Daily:3, Weekly:1,2)
    $isConfCol = ($grid.Columns.Count -eq 6 -and $e.ColumnIndex -eq 3) -or ($grid.Columns.Count -eq 4 -and ($e.ColumnIndex -eq 1 -or $e.ColumnIndex -eq 2))
    
    if ($isConfCol) {
        $cur = $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
        if ($cur -eq "PAUSE" -or ($grid.Columns.Count -eq 6 -and $cur -eq $grid.Rows[$e.RowIndex].Cells[2].Value)) { $cur="" }
        
        $new = Show-ConfigSelectorGUI $cur
        if ($new -ne $null) {
            if ($new -eq "") { 
                if ($grid.Columns.Count -eq 6) { # Daily 還原預設
                    $grid.Rows[$e.RowIndex].Cells[3].Value = $grid.Rows[$e.RowIndex].Cells[2].Value
                    $grid.Rows[$e.RowIndex].Cells[3].Style = $grid.DefaultCellStyle
                } else { # Weekly 清空
                    $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value = ""
                }
            } else {
                $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value = $new
                $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.ForeColor = "Blue"
                $grid.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.Font = $BoldFont
            }
            Mark-Dirty
        }
    }
}

function Handle-KeyDown ($grid, $e) {
    # Del 鍵
    if ($e.KeyCode -eq "Delete") {
        foreach ($cell in $grid.SelectedCells) {
            # Daily 配置欄 (3)
            if ($grid.Columns.Count -eq 6 -and $cell.ColumnIndex -eq 3) {
                $def = $grid.Rows[$cell.RowIndex].Cells[2].Value
                $cell.Value = $def; $cell.Style = $grid.DefaultCellStyle; Mark-Dirty
            }
            # Weekly 配置欄 (1,2)
            if ($grid.Columns.Count -eq 4 -and ($cell.ColumnIndex -eq 1 -or $cell.ColumnIndex -eq 2)) {
                 # 不允許刪除一般配置，只能重置為 day
                 if ($cell.ColumnIndex -eq 1) { $cell.Value = "day" } else { $cell.Value = "" } # 紊亂期可為空(繼承一般)
                 Mark-Dirty
            }
        }
    }
    # Ctrl+V
    if ($e.Control -and $e.KeyCode -eq "V") {
        $txt = [System.Windows.Forms.Clipboard]::GetText().Trim()
        if (-not [string]::IsNullOrWhiteSpace($txt)) {
            foreach ($cell in $grid.SelectedCells) {
                # 判斷是否為配置欄
                $isConf = ($grid.Columns.Count -eq 6 -and $cell.ColumnIndex -eq 3) -or ($grid.Columns.Count -eq 4 -and ($cell.ColumnIndex -eq 1 -or $cell.ColumnIndex -eq 2))
                if ($isConf) {
                    $cell.Value = $txt
                    $cell.Style.ForeColor = "Blue"; $cell.Style.Font = $BoldFont
                    Mark-Dirty
                }
            }
        }
    }
}

# --- Grid 載入與存檔 ---

# Daily Grid
$TabGrid = New-Object System.Windows.Forms.TabPage; $TabGrid.Text = "[GRID] 排程編輯器"
$GridDaily = Build-Grid $TabGrid $false

function Load-DailyGrid {
    $GridDaily.Rows.Clear()
    
    $MapData = @{}
    if (Test-Path $DateMap) { Get-Content $DateMap | ForEach-Object { if ($_ -match "^(\d{8})=(.+)$") { $MapData[$matches[1]] = $matches[2] } } }
    $PauseData = @(); if (Test-Path $PauseLog) { $PauseData = Get-Content $PauseLog }
    $NoShutData = @(); if (Test-Path $NoShutdownLog) { $NoShutData = Get-Content $NoShutdownLog }

    $StartDate = (Get-Date).AddHours(-3).Date
    for ($i = 0; $i -lt 90; $i++) {
        $d = $StartDate.AddDays($i); $dStr = $d.ToString("yyyyMMdd")
        $wStr = $d.DayOfWeek.ToString()
        
        # 計算每週預設 (需考慮紊亂期)
        $defConf = $Global:WeeklyRules[$wStr]
        $ITDay = Test-TurbulencePeriod $d
        if ($ITDay -gt 0) {
            $tConf = $Global:TurbulenceRules[$wStr]
            if ($tConf) { $defConf = "$tConf" }
        }

        $currConf = $defConf; $isOverride = $false; $isPaused = $false
        if ($PauseData -contains $dStr) { $currConf = "PAUSE"; $isPaused = $true; $isOverride = $true }
        elseif ($MapData.ContainsKey($dStr)) { $currConf = $MapData[$dStr]; $isOverride = $true }

        $isNoShut = $NoShutData -contains $dStr
        # 檢查每週不關機預設
        if ($Global:WeeklyNoShut[$wStr]) { $isNoShut = $true } # 顯示勾選，但在存檔時要區分是預設還是指定 (這裡簡化為顯示最終結果)

        $Note = ""; if (Test-GenshinUpdateDay $d) { $Note = "⚠️ 版本更新" }
        if ($ITDay -gt 0) { $Note += " 🔥 紊亂(Day$ITDay)" }

        $idx = $GridDaily.Rows.Add($d.ToString("yyyy/MM/dd"), $wStr, $defConf, $currConf, $isNoShut, $Note)
        $row = $GridDaily.Rows[$idx]; $row.Tag = $dStr

        if ($isPaused) { $row.Cells[3].Style.BackColor = "LightCoral"; $row.Cells[3].Style.ForeColor = "White" }
        elseif ($isOverride) { $row.Cells[3].Style.ForeColor = "Blue"; $row.Cells[3].Style.Font = $BoldFont }
        if ($Note) { $row.Cells[5].Style.ForeColor = "Magenta"; $row.Cells[5].Style.Font = $BoldFont }
    }
}

function Save-DailyGrid {
    $newMap = @(); $newPause = @(); $newNoShut = @()
    foreach ($row in $GridDaily.Rows) {
        $dStr = $row.Tag; $def = $row.Cells[2].Value; $cur = $row.Cells[3].Value; $shut = $row.Cells[4].Value
        if ($cur -eq "PAUSE") { $newPause += $dStr } elseif ($cur -ne $def) { $newMap += "$dStr=$cur" }
        
        # 不關機存檔邏輯：
        # 如果該日被勾選，且 該日不是「每週預設不關機」，則寫入 NoShutdown.log
        # 如果該日沒被勾選，且 該日是「每週預設不關機」，(目前無機制處理「強制關機」例外，假設使用者只會在特殊日設定不關機)
        # 為了簡化：只要有勾，且不等於每週預設，就寫入。
        # 取得該日原本是否應該不關機
        $wDay = [DateTime]::ParseExact($dStr, "yyyyMMdd", $null).DayOfWeek.ToString()
        $isDefShut = $Global:WeeklyNoShut[$wDay]
        
        if ($shut -and -not $isDefShut) { $newNoShut += $dStr }
        # 若原本是不關機，但使用者取消勾選 -> 目前 NoShutdown.log 邏輯只存「不關機日期」，無法存「強制關機」。
        # 暫時維持：只存「額外指定的不關機」。
    }
    $newMap | Sort-Object | Set-Content $DateMap -Encoding UTF8
    $newPause | Sort-Object | Set-Content $PauseLog -Encoding UTF8
    $newNoShut | Sort-Object | Set-Content $NoShutdownLog -Encoding UTF8
    Mark-Clean; [System.Windows.Forms.MessageBox]::Show("排程已儲存！"); Load-GridData
}

# Weekly Grid
$TabWeekly = New-Object System.Windows.Forms.TabPage; $TabWeekly.Text = "⚙️ 每週預設設定"
$GridWeekly = Build-Grid $TabWeekly $true

function Load-WeeklyGrid {
    $GridWeekly.Rows.Clear()
    $DaysKey = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
    $DaysTxt = @("週一","週二","週三","週四","週五","週六","週日")
    
    for ($i=0; $i -lt 7; $i++) {
        $k = $DaysKey[$i]
        $n = $Global:WeeklyRules[$k]
        $t = $Global:TurbulenceRules[$k]
        $s = $Global:WeeklyNoShut[$k]
        $GridWeekly.Rows.Add($DaysTxt[$i], $n, $t, $s)
        $GridWeekly.Rows[$i].Tag = $k
    }
}

function Save-WeeklyGrid {
    $conf = Get-JsonConf $WeeklyConf
    if (-not $conf.Turbulence) { $conf | Add-Member -Name "Turbulence" -Value @{} -MemberType NoteProperty }
    if (-not $conf.NoShutdown) { $conf | Add-Member -Name "NoShutdown" -Value @{} -MemberType NoteProperty }
    
    foreach ($row in $GridWeekly.Rows) {
        $k = $row.Tag
        $conf.$k = $row.Cells[1].Value
        $conf.Turbulence.$k = $row.Cells[2].Value
        $conf.NoShutdown.$k = $row.Cells[3].Value
    }
    $conf | ConvertTo-Json -Depth 3 | Set-Content $WeeklyConf
    Load-WeeklyRules; Mark-Clean
    [System.Windows.Forms.MessageBox]::Show("每週設定已儲存！"); Load-GridData; Load-WeeklyGrid
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
    
    $listDst.Add_MouseDown({ param($s,$e); if($listDst.SelectedItem){$listDst.DoDragDrop($listDst.SelectedItem, [System.Windows.Forms.DragDropEffects]::Move)} })
    $listDst.Add_DragOver({ param($s,$e); $e.Effect=[System.Windows.Forms.DragDropEffects]::Move })
    $listDst.Add_DragDrop({ param($s,$e); $idx=$listDst.IndexFromPoint($listDst.PointToClient([System.Drawing.Point]::new($e.X,$e.Y))); if($idx -lt 0){$idx=$listDst.Items.Count-1}; $item=$e.Data.GetData([string]); if($item){$listDst.Items.Remove($item); $listDst.Items.Insert($idx,$item); $listDst.SelectedIndex=$idx} })
    
    $SelForm.Controls.AddRange(@($lblSrc, $listSrc, $lblDst, $listDst, $btnAdd, $btnRem, $btnOk, $btnCancel))
    if ($SelForm.ShowDialog() -eq "OK") { $f=@(); foreach($i in $listDst.Items){$f+=$i}; return ($f -join ",") } else { return $null }
}

# =============================================================================
# 分頁 4: 工具與維護
# =============================================================================
$TabTools = New-Object System.Windows.Forms.TabPage; $TabTools.Text = "[TOOL] 工具與維護" 
$flpTools = New-Object System.Windows.Forms.FlowLayoutPanel; $flpTools.Dock="Fill"; $flpTools.FlowDirection="TopDown"; $flpTools.Padding="20"; $flpTools.AutoSize=$true
function Add-ToolBtn ($text, $color, $action) {
    $btn = New-Object System.Windows.Forms.Button; $btn.Text=$text; $btn.Width=400; $btn.Height=50; $btn.BackColor=$color; $btn.Font=$BoldFont; $btn.Margin="0,0,0,15"
    $btn.Add_Click($action); $flpTools.Controls.Add($btn)
}
Add-ToolBtn "[STOP] 強制停止所有任務" "LightCoral" { if([System.Windows.Forms.MessageBox]::Show("確定停止？","警告","YesNo")-eq"Yes"){ Start-Process powershell -Arg "-File `"$StopScript`"" -Verb RunAs } }
Add-ToolBtn "[FIX] 修復檔案權限" "LightBlue" { Start-Process powershell -Arg "-Command `"takeown /F '$Dir' /R /D Y; icacls '$Dir' /grant Everyone:(OI)(CI)F /T /C`"" -Verb RunAs; [System.Windows.Forms.MessageBox]::Show("完成") }
Add-ToolBtn "[GIT] 發布至 GitHub" "LightGray" { if([System.Windows.Forms.MessageBox]::Show("確定發布？","確認","YesNo")-eq"Yes"){ Start-Process powershell -Arg "-File `"$PublishScript`"" } }
Add-ToolBtn "[RDP] 修復 RDP 最小化" "LightGray" { Start-Process powershell -Arg "-Command `"reg add 'HKLM\Software\Microsoft\Terminal Server Client' /v 'RemoteDesktop_SuppressWhenMinimized' /t REG_DWORD /d 2 /f`"" -Verb RunAs; [System.Windows.Forms.MessageBox]::Show("完成") }
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
$pnlLogTop.Controls.Add($cbLogFiles); $pnlLogTop.Controls.Add($btnRefreshLog)
$TabLogs.Controls.Add($txtLogContent); $TabLogs.Controls.Add($pnlLogTop); $TabLogs.Add_Enter({ Refresh-LogList }) 

# --- 組合 ---
$TabControl.Controls.AddRange(@($TabStatus, $TabGrid, $TabWeekly, $TabTools, $TabLogs))
$Form.Controls.Add($TabControl)
$Form.Add_Load({ Update-StatusUI; Load-GridData; Load-WeeklyGrid })
$Form.ShowDialog()