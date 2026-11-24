function New-OpsToolbar($Parent) {
    $bar = New-Object Windows.Forms.FlowLayoutPanel
    $bar.Dock = 'Top'
    $bar.AutoSize = $true
    $bar.Padding = '8,5,8,5'
    $bar.WrapContents = $true
    $Parent.Controls.Add($bar)
    $bar.BringToFront()
    return $bar
}
function Add-OpsControl($Parent, $Control) {
    $Control.Margin = '3,4,8,3'
    $Parent.Controls.Add($Control)
    return $Control
}
function New-OpsLabel($Parent, [string]$Text) {
    $label = New-Object Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Margin = '3,9,4,0'
    $Parent.Controls.Add($label)
    return $label
}
function New-OpsButton($Parent, [string]$Text) {
    $button = New-Object Windows.Forms.Button
    $button.Text = $Text
    $button.AutoSize = $true
    Add-OpsControl $Parent $button
}
function New-OpsCheck($Parent, [string]$Text, [bool]$Checked = $false) {
    $check = New-Object Windows.Forms.CheckBox
    $check.Text = $Text
    $check.Checked = $Checked
    $check.AutoSize = $true
    $check.Margin = '3,8,8,0'
    $Parent.Controls.Add($check)
    return $check
}
function New-OpsText($Parent, [int]$Width = 180) {
    $box = New-Object Windows.Forms.TextBox
    $box.Width = $Width
    Add-OpsControl $Parent $box
}
function Set-OpsChoices($Control, [object[]]$Items, [string]$Value) {
    $Control.BeginUpdate()
    try {
        $Control.Items.Clear()
        foreach ($item in $Items) { if ($null -ne $item) { [void]$Control.Items.Add([string]$item) } }
        if ($Value -and $Control.Items.Contains($Value)) { $Control.SelectedItem = $Value }
        elseif ($Value -and $Control.DropDownStyle -ne 'DropDownList') { $Control.Text = $Value }
        elseif ($Control.Items.Count) { $Control.SelectedIndex = 0 }
        elseif ($Control.DropDownStyle -ne 'DropDownList') { $Control.Text = '' }
    } finally { $Control.EndUpdate() }
}
function New-OpsCombo($Parent, [object[]]$Items, [int]$Width = 150, [switch]$Searchable) {
    $combo = New-Object Windows.Forms.ComboBox
    $combo.Width = $Width
    $combo.DropDownStyle = 'DropDownList'
    if ($Searchable) {
        $combo.DropDownStyle = 'DropDown'
        $combo.AutoCompleteMode = 'SuggestAppend'
        $combo.AutoCompleteSource = 'ListItems'
    }
    Set-OpsChoices $combo $Items ''
    Add-OpsControl $Parent $combo
}
function New-OpsDetails($Parent) {
    $box = New-Object Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.WordWrap = $false
    $box.ScrollBars = 'Both'
    $box.Dock = 'Fill'
    $box.Font = New-Object Drawing.Font('Consolas', 10)
    $Parent.Controls.Add($box)
    $box.BringToFront()
    return $box
}
function New-OpsSplit($Parent, [int]$TopHeight = 460) {
    $split = New-Object Windows.Forms.SplitContainer
    $split.Orientation = 'Horizontal'
    $split.Dock = 'Fill'
    $split.Height = $TopHeight + 135
    $split.SplitterDistance = $TopHeight
    $split.FixedPanel = 'Panel2'
    $split.Panel2MinSize = 70
    $Parent.Controls.Add($split)
    $split.BringToFront()
    return $split
}
function New-OpsGrid($Parent, [System.Collections.IDictionary]$Columns = @{}) {
    $grid = New-Object Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToOrderColumns = $true
    $grid.RowHeadersVisible = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.BackgroundColor = [Drawing.Color]::White
    $grid.BorderStyle = 'None'
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(241,245,249)
    Set-OpsColumns $grid $Columns
    $Parent.Controls.Add($grid)
    $grid.BringToFront()
    return ,$grid
}
function Set-OpsColumns($Grid, [System.Collections.IDictionary]$Columns) {
    $Grid.Columns.Clear()
    foreach ($name in $Columns.Keys) {
        $index = $Grid.Columns.Add([string]$name, [string]$Columns[$name])
        $Grid.Columns[$index].Width = if ($name -match 'message|folder|path|sourceName|detail') { 245 } else { 130 }
        $Grid.Columns[$index].SortMode = 'Programmatic'
    }
    $Grid.Tag = $null
}
function Get-OpsValue($Row, [string]$Name) {
    if ($null -eq $Row) { return }
    if ($Row -is [Data.DataRow]) {
        if ($Row.Table.Columns.Contains($Name) -and $Row[$Name] -ne [DBNull]::Value) { return ,$Row[$Name] }
    } elseif ($Row -is [System.Collections.IDictionary]) {
        return ,$Row[$Name]
    } elseif ($Row.PSObject.Properties[$Name]) {
        return ,$Row.$Name
    }
}
function Get-OpsDisplay($Value) {
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return (Format-QCDisplayTime -Value $Value) }
    if ($Value -is [byte[]]) { return ('{0} bytes' -f $Value.Length) }
    return [string]$Value
}
function Get-OpsSelected($Grid) {
    if ($Grid.CurrentRow) { return ,$Grid.CurrentRow.Tag }
}
function Set-OpsRows($Grid, [object[]]$Rows, [string]$Key, [switch]$ColorSeverity) {
    $prior = @($Grid.Tag)
    $selected = Get-OpsValue (Get-OpsSelected $Grid) $Key
    $scroll = $null
    $first = $Grid.FirstDisplayedScrollingRowIndex
    if ($Key -and $first -ge 0) { $scroll = Get-OpsValue $Grid.Rows[$first].Tag $Key }
    $keys = @()
    if ($Key) { $keys = @($Rows | ForEach-Object { [string](Get-OpsValue $_ $Key) }) }
    # Add new log rows ahead of the existing rows without recreating the surviving records.
    $reuse = $ColorSeverity -and $prior.Count -gt 0 -and $prior[0] -and $keys.Count -gt 0
    $start = 0
    if ($reuse) {
        $start = [array]::IndexOf($keys, $prior[0])
        $overlap = [math]::Min($prior.Count, $keys.Count - $start)
        if ($start -lt 0 -or $overlap -le 0) { $reuse = $false }
        for ($i = 0; $reuse -and $i -lt $overlap; $i++) {
            if ($prior[$i] -cne $keys[$start + $i]) { $reuse = $false }
        }
    }
    $script:UpdatingGrid = $true
    $Grid.SuspendLayout()
    try {
        if (-not $reuse) { $Grid.Rows.Clear() }
        elseif ($overlap -eq $prior.Count -and $overlap -eq $keys.Count) { return }
        else {
            while ($Grid.Rows.Count -gt $overlap) { $Grid.Rows.RemoveAt($Grid.Rows.Count - 1) }
        }
        for ($i = 0; $i -lt $Rows.Count; $i++) {
            if ($reuse -and $i -ge $start -and $i -lt ($start + $overlap)) { continue }
            $record = $Rows[$i]
            $cells = [object[]]@($Grid.Columns | ForEach-Object { Get-OpsDisplay (Get-OpsValue $record $_.Name) })
            if ($i -lt $Grid.Rows.Count) { $Grid.Rows.Insert($i, $cells) }
            else { [void]$Grid.Rows.Add($cells) }
            $row = $Grid.Rows[$i]
            $row.Tag = $record
            if ($ColorSeverity) {
                switch -Regex ($record.severity) {
                    'error|fatal|critical' { $row.DefaultCellStyle.ForeColor = [Drawing.Color]::Firebrick }
                    'warn' { $row.DefaultCellStyle.ForeColor = [Drawing.Color]::SaddleBrown }
                }
            }
        }
        $Grid.Tag = $keys
        if ($Key -and ($selected -or $scroll)) {
            $column = $Grid.Columns.GetFirstColumn([Windows.Forms.DataGridViewElementStates]::Visible)
            if ($column) {
                foreach ($row in $Grid.Rows) {
                    if ($selected -and (Get-OpsValue $row.Tag $Key) -ceq $selected) { $Grid.CurrentCell = $row.Cells[$column.Index] }
                }
            }
            if ($scroll) {
                for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
                    if ((Get-OpsValue $Grid.Rows[$i].Tag $Key) -ceq $scroll) { $Grid.FirstDisplayedScrollingRowIndex = $i; break }
                }
            }
        }
    } finally {
        $Grid.ResumeLayout()
        $script:UpdatingGrid = $false
    }
}
function Format-OpsRecord($Record) {
    if ($null -eq $Record) { return '' }
    if ($Record -is [Data.DataRow]) {
        return (@(foreach ($c in $Record.Table.Columns) {
            $value = $Record[$c]
            $text = if ($value -eq [DBNull]::Value) { '(NULL)' }
                elseif ($value -is [byte[]]) { [Convert]::ToBase64String($value) }
                else { Get-OpsDisplay $value }
            "$($c.ColumnName): $text"
        }) -join [Environment]::NewLine)
    }
    return ($Record | ConvertTo-Json -Depth 14)
}
function Set-OpsNotice($Label, [string]$Text, [switch]$ErrorState) {
    $Label.ForeColor = if ($ErrorState) { [Drawing.Color]::Firebrick } else { [Drawing.Color]::FromArgb(51,65,85) }
    $Label.Text = $Text
}
function Confirm-Ops([string]$Text) {
    [Windows.Forms.MessageBox]::Show($form, $Text, 'QC Pipeline Ops', 'YesNo', 'Warning', 'Button2') -eq 'Yes'
}
function New-OpsFolderPicker($Parent, [string]$Project = '', [string]$Folder = '') {
    [void](New-OpsLabel $Parent 'Project')
    $projects = @('All projects') + @($script:Presets.project | Sort-Object -Unique)
    $projectBox = New-OpsCombo $Parent $projects 160 -Searchable
    [void](New-OpsLabel $Parent 'Folder')
    $folderBox = New-OpsCombo $Parent @('') 300 -Searchable
    Set-OpsChoices $projectBox $projects $Project
    $projectBox.Tag = $folderBox
    $projectBox.Add_TextChanged({
        $current = $this.Tag.Text
        $options = @($script:Presets | Where-Object { $this.Text -eq 'All projects' -or $_.project -eq $this.Text })
        Set-OpsChoices $this.Tag (@('') + @($options.folder)) $current
    })
    $options = @($script:Presets | Where-Object { $projectBox.Text -eq 'All projects' -or $_.project -eq $projectBox.Text })
    Set-OpsChoices $folderBox (@('') + @($options.folder)) $Folder
    @{ project = $projectBox; folder = $folderBox }
}
function Get-OpsFolder($Picker, [switch]$ExplicitOnly) {
    if ($Picker.folder.Text.Trim()) { return $Picker.folder.Text.Trim() }
    if (-not $ExplicitOnly) {
        $preset = $script:Presets | Where-Object project -eq $Picker.project.Text | Select-Object -First 1
        if ($preset) { return $preset.root }
    }
    return ''
}
