# =============================================================================
# AutoTask Dashboard V6.0 - 旗艦整合版 (含工具箱與發布功能)
# =============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- [定義檔案路徑] ---
$Dir = "C:\AutoTask"
$WeeklyConf = "$Dir\Configs\WeeklyConfig.json"
$DateMap = "$Dir\Configs\DateConfig.map"
$TaskStatus = "$Dir\Configs\TaskStatus.json"
$PauseLog = "$Dir\Configs\PauseDates.log"
$NoShutdownLog = "$Dir\Configs\NoShutdown.log"
$ManualFlag = "$Dir\Flags\ManualTrigger.flag"
$BetterGI_UserDir = "C:\Program Files\BetterGI\User\OneDragon"
$MasterScript = "$Dir\Scripts\Master.ps1"
$StopScript = "$Dir\Scripts\StopAll.ps1"
$PublishScript = "$Dir\Scripts\PublishRelease.ps1"

# --- [全域變數] ---
$Global:ConfigList = @() 
$Global:WeeklyRules = @{}
$Script:IsDirty = $false
$Script:IsLoading = $false
$WindowTitle = "AutoTask 控制台 V6.0"

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
    if ($wk) {
        $Global:WeeklyRules = $wk
    } else {
        $Global:WeeklyRules = @{
            "Monday"="monday"; "Tuesday"="day"; "Wednesday"="day"; "Thursday"="day";
            "Friday"="day"; "Saturday"="day"; "Sunday"="day"
        }
    }
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
    
    if (Test-Path $PauseLog) {
        if ((Get-Content $PauseLog) -contains $dStr) { 
            $txt = "已排程暫停 (PAUSED)" 
            $color = [System.Drawing.Color]::Orange
        }
    }
    return @{Text=$txt; Color=$color}
}

function Get-WeekName ($dateObj) {
    $Map = @{ "Monday"="週一"; "Tuesday"="週二"; "Wednesday"="週三"; "Thursday"="週四"; "Friday"="週五"; "Saturday"="週六"; "Sunday"="週日" }
    return $Map[$dateObj.DayOfWeek.ToString()]
}

function Get-DisplayConfigName ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    $dWeek = $dateObj.DayOfWeek.ToString()
    if (Test-Path $DateMap) {
        $map = Get-Content $DateMap
        foreach ($line in $map) {
            if ($line -match "^$dStr=(.+)$") {
                return "$($matches[1]) (指定)"
            }
        }
    }
    $wk = Get-JsonConf $WeeklyConf
    if ($wk) {
        return "$($wk.$dWeek) (每週)"
    }
    return "day (預設)"
}

function Get-ShutdownPolicy ($dateObj) {
    $dStr = $dateObj.ToString("yyyyMMdd")
    if (Test-Path $NoShutdownLog) {
        if ((Get-Content $NoShutdownLog) -contains $dStr) {
            return "不關機"
        }
    }
    return "自動關機"
}

function Parse-SmartDate ($inputStr, $anchorDate) {
    if ([string]::IsNullOrWhiteSpace($inputStr)) { return $null }
    $inputStr = $inputStr.Trim()
    $y = $anchorDate.Year
    $m = $anchorDate.Month

    if ($inputStr -match "^\d{8}$") {
        return [DateTime]::ParseExact($inputStr, "yyyyMMdd", $null)
    }
    if ($inputStr -match "^(\d{2})(\d{2})$") {
        $mM = [int]$matches[1]
        $dD = [int]$matches[2]
        try {
            $testDate = Get-Date -Year $y -Month $mM -Day $dD -Hour 0 -Minute 0 -Second 0
        } catch { throw "日期無效" }
        if ($testDate -lt $anchorDate) { return $testDate.AddYears(1) } else { return $testDate }
    }
    if ($inputStr -match "^(\d{2})$") {
        $dD = [int]$matches[1]
        try {
            $testDate = Get-Date -Year $y -Month $m -Day $dD -Hour 0 -Minute 0 -Second 0
        } catch { throw "日期無效" }
        if ($testDate -lt $anchorDate) { return $testDate.AddMonths(1) } else { return $testDate }
    }
    throw "格式錯誤 (支援 yyyyMMdd, MMDD, DD)"
}

# --- [變更狀態追蹤] ---
function Mark-Dirty {
    if ($Script:IsLoading) { return }
    $Script:IsDirty = $true
    $Form.Text = "$WindowTitle * (未儲存)"
}

function Mark-Clean {
    $Script:IsDirty = $false
    $Form.Text = $WindowTitle
}

# --- [GUI 初始化] ---
Load-BetterGIConfigs
Load-WeeklyRules

$Form = New-Object System.Windows.Forms.Form
$Form.Text = $WindowTitle
$Form.Size = New-Object System.Drawing.Size(1000, 720)
$Form.StartPosition = "CenterScreen"

$Form.Add_FormClosing({
    param($sender, $e)
    if ($Script:IsDirty) {
        $result = [System.Windows.Forms.MessageBox]::Show("您有尚未儲存的變更，確定要離開嗎？", "未儲存警告", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq "No") { $e.Cancel = $true }
    }
})

$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Dock = "Fill"

# =============================================================================
# 分頁 1: 即時狀態
# =============================================================================
$TabStatus = New-Object System.Windows.Forms.TabPage
$TabStatus.Text = "🏠 即時狀態"
$TabStatus.Padding = New-Object System.Windows.Forms.Padding(10)

$lblTodayInfo = New-Object System.Windows.Forms.Label
$lblTodayInfo.AutoSize = $true
$lblTodayInfo.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)
$lblTodayInfo.Location = "20, 20"

$btnManual = New-Object System.Windows.Forms.Button
$btnManual.Text = "🚨 手動救援 / 強制啟動"
$btnManual.Location = "20, 150"; $btnManual.Size = "300, 50"
$btnManual.BackColor = [System.Drawing.Color]::LightCoral
$btnManual.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12, [System.Drawing.FontStyle]::Bold)
$btnManual.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("確定要強制啟動任務嗎？", "確認", "YesNo") -eq "Yes") {
        New-Item -Path $ManualFlag -ItemType File -Force | Out-Null
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$MasterScript`""
    }
})

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "重新整理"
$btnRefresh.Location = "20, 210"; $btnRefresh.Width = 300
$btnRefresh.Add_Click({ Update-StatusUI })

$TabStatus.Controls.AddRange(@($lblTodayInfo, $btnManual, $btnRefresh))

function Update-StatusUI {
    $today = (Get-Date).AddHours(-3)
    $st = Get-StatusText
    $wkName = $Global:WeeklyRules[$today.DayOfWeek.ToString()]
    
    $finalConf = $wkName
    $dStr = $today.ToString("yyyyMMdd")
    if (Test-Path $DateMap) {
        Get-Content $DateMap | ForEach-Object { if ($_ -match "^$dStr=(.+)$") { $finalConf = "$($matches[1]) (指定)" } }
    }
    if (Test-Path $PauseLog) { if ((Get-Content $PauseLog) -contains $dStr) { $finalConf = "PAUSED (暫停)" } }

    $lblTodayInfo.Text = "今日: $($today.ToString('yyyy/MM/dd')) ($($today.DayOfWeek))`n" +
                         "配置: $finalConf`n" +
                         "狀態: $($st.Text)"
    $lblTodayInfo.ForeColor = $st.Color
}

# =============================================================================
# 分頁 2: 排程網格編輯器
# =============================================================================
$TabGrid = New-Object System.Windows.Forms.TabPage
$TabGrid.Text = "📅 排程編輯器"

$panelTool = New-Object System.Windows.Forms.Panel
$panelTool.Dock = "Top"; $panelTool.Height = 40

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "💾 儲存變更"
$btnSave.Dock = "Left"; $btnSave.Width = 100; $btnSave.BackColor = [System.Drawing.Color]::LightGreen
$btnSave.Add_Click({ Save-GridData })

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "操作提示: 支援批量勾選 [不關機] (Ctrl/Shift) | 雙擊配置欄排序 | Ctrl+C/V | Del"
$lblHint.Dock = "Fill"; $lblHint.TextAlign = "MiddleLeft"; $lblHint.Padding = "10,0,0,0"

$panelTool.Controls.Add($lblHint)
$panelTool.Controls.Add($btnSave)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "CellSelect"
$grid.MultiSelect = $true
$grid.ClipboardCopyMode = "EnableWithoutHeaderText"
$grid.Font = New-Object System.Drawing.Font("Consolas", 10)
$grid.EditMode = "EditProgrammatically" 

$colDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDate.HeaderText = "日期"; $colDate.ReadOnly = $true; $colDate.Width = 100
$colWeek = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colWeek.HeaderText = "星期"; $colWeek.ReadOnly = $true; $colWeek.Width = 80
$colDef  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDef.HeaderText = "每週預設"; $colDef.ReadOnly = $true; $colDef.Width = 100
$colConf = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colConf.HeaderText = "執行配置 (雙擊)"; $colConf.Width = 250
$colShut = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colShut.HeaderText = "不關機 (批量勾選)"; $colShut.Width = 120

$grid.Columns.Add($colDate) | Out-Null
$grid.Columns.Add($colWeek) | Out-Null
$grid.Columns.Add($colDef) | Out-Null
$grid.Columns.Add($colConf) | Out-Null
$grid.Columns.Add($colShut) | Out-Null

function Load-GridData {
    $Script:IsLoading = $true
    $grid.Rows.Clear()
    
    $MapData = @{}
    if (Test-Path $DateMap) {
        Get-Content $DateMap | ForEach-Object { if ($_ -match "^(\d{8})=(.+)$") { $MapData[$matches[1]] = $matches[2] } }
    }
    $PauseData = @()
    if (Test-Path $PauseLog) { $PauseData = Get-Content $PauseLog }
    $NoShutData = @()
    if (Test-Path $NoShutdownLog) { $NoShutData = Get-Content $NoShutdownLog }

    $StartDate = (Get-Date).AddHours(-3).Date
    for ($i = 0; $i -lt 90; $i++) {
        $d = $StartDate.AddDays($i)
        $dStr = $d.ToString("yyyyMMdd")
        $wStr = $d.DayOfWeek.ToString()
        $defConf = $Global:WeeklyRules[$wStr]
        
        $currConf = $defConf
        $isOverride = $false
        $isPaused = $false

        if ($PauseData -contains $dStr) {
            $currConf = "PAUSE"
            $isPaused = $true
            $isOverride = $true
        } elseif ($MapData.ContainsKey($dStr)) {
            $currConf = $MapData[$dStr]
            $isOverride = $true
        }

        $isNoShut = $NoShutData -contains $dStr

        $idx = $grid.Rows.Add($d.ToString("yyyy/MM/dd"), $wStr, $defConf, $currConf, $isNoShut)
        $row = $grid.Rows[$idx]
        $row.Tag = $dStr

        if ($isPaused) {
            $row.Cells[3].Style.BackColor = [System.Drawing.Color]::LightCoral
            $row.Cells[3].Style.ForeColor = [System.Drawing.Color]::White
        } elseif ($isOverride) {
            $row.Cells[3].Style.ForeColor = [System.Drawing.Color]::Blue
            $row.Cells[3].Style.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
        }
    }
    $Script:IsLoading = $false
    Mark-Clean
}

$grid.Add_CellClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }

    if ($e.ColumnIndex -eq 4) {
        $clickedCell = $grid.Rows[$e.RowIndex].Cells[4]
        $currentVal = if ($clickedCell.Value) { $true } else { $false }
        $newValue = -not $currentVal
        
        $targetCells = $grid.SelectedCells | Where-Object { $_.ColumnIndex -eq 4 }
        
        if ($targetCells.Count -gt 0 -and ($targetCells | Where-Object { $_.RowIndex -eq $e.RowIndex })) {
            foreach ($cell in $targetCells) {
                $cell.Value = $newValue
            }
        } else {
            $clickedCell.Value = $newValue
        }
        Mark-Dirty
    }
})

$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 3) { return }
    
    $currentVal = $grid.Rows[$e.RowIndex].Cells[3].Value
    if ($currentVal -eq $grid.Rows[$e.RowIndex].Cells[2].Value -or $currentVal -eq "PAUSE") { $currentVal = "" }
    
    $newVal = Show-ConfigSelectorGUI $currentVal
    
    if ($newVal -ne $null) {
        if ($newVal -eq "") { 
            $grid.Rows[$e.RowIndex].Cells[3].Value = $grid.Rows[$e.RowIndex].Cells[2].Value
            $grid.Rows[$e.RowIndex].Cells[3].Style = $grid.DefaultCellStyle
        } else {
            $grid.Rows[$e.RowIndex].Cells[3].Value = $newVal
            $grid.Rows[$e.RowIndex].Cells[3].Style.ForeColor = [System.Drawing.Color]::Blue
            $grid.Rows[$e.RowIndex].Cells[3].Style.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
        }
        Mark-Dirty
    }
})

$grid.Add_KeyDown({
    param($sender, $e)
    
    if ($e.KeyCode -eq "Delete") {
        foreach ($cell in $grid.SelectedCells) {
            if ($cell.ColumnIndex -eq 3) {
                $defVal = $grid.Rows[$cell.RowIndex].Cells[2].Value
                $cell.Value = $defVal
                $cell.Style = $grid.DefaultCellStyle
                Mark-Dirty
            }
        }
    }

    if ($e.Control -and $e.KeyCode -eq "V") {
        $clipText = [System.Windows.Forms.Clipboard]::GetText().Trim()
        if (-not [string]::IsNullOrWhiteSpace($clipText)) {
            foreach ($cell in $grid.SelectedCells) {
                if ($cell.ColumnIndex -eq 3) {
                    if ($cell.Value -ne $clipText) {
                        $cell.Value = $clipText
                        if ($clipText -eq "PAUSE") {
                            $cell.Style.BackColor = [System.Drawing.Color]::LightCoral
                            $cell.Style.ForeColor = [System.Drawing.Color]::White
                        } else {
                            $cell.Style.ForeColor = [System.Drawing.Color]::Blue
                            $cell.Style.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                            $cell.Style.BackColor = [System.Drawing.Color]::White
                        }
                        Mark-Dirty
                    }
                }
            }
        }
    }
})

function Save-GridData {
    $newMap = @()
    $newPause = @()
    $newNoShut = @()

    foreach ($row in $grid.Rows) {
        $dStr = $row.Tag
        $defVal = $row.Cells[2].Value
        $curVal = $row.Cells[3].Value
        $isNoShut = $row.Cells[4].Value

        if ($curVal -eq "PAUSE") {
            $newPause += $dStr
        } elseif ($curVal -ne $defVal) {
            $newMap += "$dStr=$curVal"
        }

        if ($isNoShut) { $newNoShut += $dStr }
    }

    $newMap | Sort-Object | Set-Content $DateMap -Encoding UTF8
    $newPause | Sort-Object | Set-Content $PauseLog -Encoding UTF8
    $newNoShut | Sort-Object | Set-Content $NoShutdownLog -Encoding UTF8

    Mark-Clean
    [System.Windows.Forms.MessageBox]::Show("設定已儲存！`n(每週配置請至分頁 3 儲存)")
    Load-GridData
    Init-WeeklyTab
}

# =============================================================================
# 分頁 3: 每週配置 GUI
# =============================================================================
$TabWeekly = New-Object System.Windows.Forms.TabPage
$TabWeekly.Text = "⚙️ 每週預設設定"

$pnlWeekly = New-Object System.Windows.Forms.Panel
$pnlWeekly.Dock = "Fill"
$pnlWeekly.AutoScroll = $true

$DaysKey = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
$DaysTxt = @("週一", "週二", "週三", "週四", "週五", "週六", "週日")
$WeeklyInputs = @{} 

$y = 30
for ($i=0; $i -lt 7; $i++) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $DaysTxt[$i]
    $l.Location = "50, $y"
    $l.AutoSize = $true
    $l.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 11)
    
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = "120, $y"
    $txt.Width = 250
    $txt.ReadOnly = $true
    
    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = "選擇/排序"
    $btnEdit.Location = "380, $($y-2)"
    $btnEdit.Width = 80
    $btnEdit.Tag = $txt
    $btnEdit.Add_Click({ param($sender, $e); $current = $this.Tag.Text; $new = Show-ConfigSelectorGUI $current; if ($new -ne $null) { $this.Tag.Text = $new } }.GetNewClosure())
    
    $pnlWeekly.Controls.Add($l)
    $pnlWeekly.Controls.Add($txt)
    $pnlWeekly.Controls.Add($btnEdit)
    $WeeklyInputs[$DaysKey[$i]] = $txt
    
    $y += 50
}

$btnSaveWeekly = New-Object System.Windows.Forms.Button
$btnSaveWeekly.Text = "儲存每週預設值"
$btnSaveWeekly.Location = "120, $y"
$btnSaveWeekly.Size = "250, 40"
$btnSaveWeekly.BackColor = [System.Drawing.Color]::LightGreen
$btnSaveWeekly.Add_Click({
    $newConf = @{}
    foreach ($d in $DaysKey) { $newConf[$d] = $WeeklyInputs[$d].Text }
    $newConf | ConvertTo-Json | Set-Content $WeeklyConf
    Load-WeeklyRules
    [System.Windows.Forms.MessageBox]::Show("每週預設配置已更新！`n(排程網格將同步更新)")
    Load-GridData
})

$pnlWeekly.Controls.Add($btnSaveWeekly)
$TabWeekly.Controls.Add($pnlWeekly)

function Init-WeeklyTab {
    $wk = Get-JsonConf $WeeklyConf
    if ($wk) {
        foreach ($d in $DaysKey) {
            if ($WeeklyInputs.ContainsKey($d)) { $WeeklyInputs[$d].Text = $wk.$d }
        }
    }
}

# --- 配置選擇器 (拖曳排序) ---
function Show-ConfigSelectorGUI {
    param([string]$CurrentSelection) 

    $SelForm = New-Object System.Windows.Forms.Form
    $SelForm.Text = "配置組排程器 (拖曳排序)"
    $SelForm.Size = New-Object System.Drawing.Size(700, 500)
    $SelForm.StartPosition = "CenterParent"
    $SelForm.FormBorderStyle = "FixedDialog"
    $SelForm.MaximizeBox = $false
    $SelForm.MinimizeBox = $false

    $lblSrc = New-Object System.Windows.Forms.Label; $lblSrc.Text = "可用配置 (可多選)"; $lblSrc.Location = "20,10"; $lblSrc.AutoSize = $true
    $listSrc = New-Object System.Windows.Forms.ListBox; $listSrc.Location = "20,30"; $listSrc.Size = "250,350"; $listSrc.SelectionMode = "MultiExtended"
    $RealConfigs = $Global:ConfigList | Where-Object { $_ -ne "PAUSE" }
    $listSrc.Items.AddRange($RealConfigs)

    $lblDst = New-Object System.Windows.Forms.Label; $lblDst.Text = "執行佇列 (拖曳可排序)"; $lblDst.Location = "380,10"; $lblDst.AutoSize = $true
    $listDst = New-Object System.Windows.Forms.ListBox; $listDst.Location = "380,30"; $listDst.Size = "250,350"; $listDst.SelectionMode = "One" 
    $listDst.AllowDrop = $true 

    if (-not [string]::IsNullOrWhiteSpace($CurrentSelection) -and $CurrentSelection -ne "PAUSE") {
        $parts = $CurrentSelection -split ","
        foreach ($p in $parts) { if (-not [string]::IsNullOrWhiteSpace($p)) { $listDst.Items.Add($p) } }
    }

    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text = "加入 ->"; $btnAdd.Location = "280,150"; $btnAdd.Size = "90,30"
    $btnRem = New-Object System.Windows.Forms.Button; $btnRem.Text = "<- 移除"; $btnRem.Location = "280,200"; $btnRem.Size = "90,30"
    
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "確定儲存"; $btnOk.Location = "250,400"; $btnOk.Size = "100,40"; $btnOk.DialogResult = "OK"; $btnOk.BackColor = [System.Drawing.Color]::LightGreen
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "取消"; $btnCancel.Location = "360,400"; $btnCancel.Size = "100,40"; $btnCancel.DialogResult = "Cancel"

    $btnAdd.Add_Click({ foreach ($item in $listSrc.SelectedItems) { $listDst.Items.Add($item) } })
    $RemoveAction = { if ($listDst.SelectedIndex -ge 0) { $listDst.Items.RemoveAt($listDst.SelectedIndex) } }
    $btnRem.Add_Click($RemoveAction)
    $listDst.Add_KeyDown({ if ($_.KeyCode -eq "Delete") { & $RemoveAction } })

    $listDst.Add_MouseDown({ 
        param($sender, $e)
        if ($listDst.SelectedItem -eq $null) { return }
        $listDst.DoDragDrop($listDst.SelectedItem, [System.Windows.Forms.DragDropEffects]::Move)
    })
    $listDst.Add_DragOver({ 
        param($sender, $e)
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    })
    $listDst.Add_DragDrop({ 
        param($sender, $e)
        $point = $listDst.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
        $index = $listDst.IndexFromPoint($point)
        if ($index -lt 0) { $index = $listDst.Items.Count - 1 }
        $data = $e.Data.GetData([string])
        $item = $data
        if ($item) {
             if ($listDst.SelectedIndex -ge 0) {
                $moveItem = $listDst.SelectedItem
                $listDst.Items.Remove($moveItem)
                $listDst.Items.Insert($index, $moveItem)
                $listDst.SelectedIndex = $index
             }
        }
    })

    $SelForm.Controls.AddRange(@($lblSrc, $listSrc, $lblDst, $listDst, $btnAdd, $btnRem, $btnOk, $btnCancel))
    $SelForm.AcceptButton = $btnOk
    $SelForm.CancelButton = $btnCancel

    $result = $SelForm.ShowDialog()
    if ($result -eq "OK") {
        $finalList = @()
        foreach ($i in $listDst.Items) { $finalList += $i }
        return ($finalList -join ",")
    } else {
        return $null
    }
}

# =============================================================================
# 分頁 4: 工具與維護 (新增)
# =============================================================================
$TabTools = New-Object System.Windows.Forms.TabPage
$TabTools.Text = "🛠️ 工具與維護"

$flpTools = New-Object System.Windows.Forms.FlowLayoutPanel
$flpTools.Dock = "Fill"; $flpTools.FlowDirection = "TopDown"; $flpTools.Padding = New-Object System.Windows.Forms.Padding(20); $flpTools.AutoSize = $true

function Add-ToolBtn ($text, $color, $action) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text; $btn.Width = 400; $btn.Height = 50; $btn.BackColor = $color
    $btn.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 11, [System.Drawing.FontStyle]::Bold)
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 15)
    $btn.Add_Click($action)
    $flpTools.Controls.Add($btn)
}

Add-ToolBtn "🛑 強制停止所有任務 (StopAll)" "LightCoral" {
    if ([System.Windows.Forms.MessageBox]::Show("這將強制關閉所有相關程序 (1Remote, BetterGI, 腳本)，確定嗎？", "警告", "YesNo", "Warning") -eq "Yes") {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$StopScript`"" -Verb RunAs
    }
}

Add-ToolBtn "🔧 修復檔案權限 (Fix Permissions)" "LightBlue" {
    $cmd = "takeown /F `"$Dir`" /R /D Y; icacls `"$Dir`" /grant Everyone:(OI)(CI)F /T /C"
    Start-Process powershell.exe -ArgumentList "-Command `"$cmd`"" -Verb RunAs
    [System.Windows.Forms.MessageBox]::Show("權限修復指令已發送。")
}

Add-ToolBtn "☁️ 發布至 GitHub (淨化+上傳)" "LightGray" {
    if ([System.Windows.Forms.MessageBox]::Show("這將建立/更新公開發布資料夾並推送到 GitHub，確定嗎？", "發布確認", "YesNo") -eq "Yes") {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PublishScript`""
    }
}

Add-ToolBtn "🖥️ 修復 RDP 最小化停止渲染" "LightGray" {
    $reg = "reg add `"HKEY_LOCAL_MACHINE\Software\Microsoft\Terminal Server Client`" /v `"RemoteDesktop_SuppressWhenMinimized`" /t REG_DWORD /d 2 /f"
    Start-Process powershell.exe -ArgumentList "-Command `"$reg`"" -Verb RunAs
    [System.Windows.Forms.MessageBox]::Show("註冊表已修正。")
}

$TabTools.Controls.Add($flpTools)

# =============================================================================
# 分頁 5: 日誌檢視 (新增)
# =============================================================================
$TabLogs = New-Object System.Windows.Forms.TabPage
$TabLogs.Text = "📝 日誌檢視"

$pnlLogTop = New-Object System.Windows.Forms.Panel; $pnlLogTop.Dock = "Top"; $pnlLogTop.Height = 40
$cbLogFiles = New-Object System.Windows.Forms.ComboBox; $cbLogFiles.Width = 300; $cbLogFiles.Location = "10, 10"; $cbLogFiles.DropDownStyle = "DropDownList"
$btnRefreshLog = New-Object System.Windows.Forms.Button; $btnRefreshLog.Text = "重新讀取"; $btnRefreshLog.Location = "320, 8"; $btnRefreshLog.Width = 100

$txtLogContent = New-Object System.Windows.Forms.TextBox; $txtLogContent.Dock = "Fill"; $txtLogContent.Multiline = $true; $txtLogContent.ScrollBars = "Vertical"; $txtLogContent.Font = New-Object System.Drawing.Font("Consolas", 10); $txtLogContent.ReadOnly = $true

function Refresh-LogList {
    $cbLogFiles.Items.Clear()
    if (Test-Path "$Dir\Logs") {
        $files = Get-ChildItem "$Dir\Logs\*.log" | Sort-Object LastWriteTime -Descending
        foreach ($f in $files) { $cbLogFiles.Items.Add($f.Name) }
    }
    if ($cbLogFiles.Items.Count -gt 0) { $cbLogFiles.SelectedIndex = 0 }
}

$btnRefreshLog.Add_Click({
    if ($cbLogFiles.SelectedItem) {
        $path = Join-Path "$Dir\Logs" $cbLogFiles.SelectedItem
        $txtLogContent.Text = Get-Content $path -Encoding UTF8 | Out-String
        $txtLogContent.SelectionStart = $txtLogContent.Text.Length; $txtLogContent.ScrollToCaret()
    }
})

$cbLogFiles.Add_SelectedIndexChanged({ $btnRefreshLog.PerformClick() })

$pnlLogTop.Controls.Add($cbLogFiles); $pnlLogTop.Controls.Add($btnRefreshLog)
$TabLogs.Controls.Add($txtLogContent); $TabLogs.Controls.Add($pnlLogTop)

$TabLogs.Add_Enter({ Refresh-LogList }) 

# --- [啟動邏輯] ---
$TabControl.Controls.Add($TabStatus)
$TabControl.Controls.Add($TabGrid)
$TabControl.Controls.Add($TabWeekly)
$TabControl.Controls.Add($TabTools) # 新增工具頁
$TabControl.Controls.Add($TabLogs)  # 新增日誌頁
$Form.Controls.Add($TabControl)

$Form.Add_Load({ 
    Update-StatusUI
    Load-GridData
    Init-WeeklyTab
})

$Form.ShowDialog()