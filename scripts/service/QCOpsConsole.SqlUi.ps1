$tabSql = New-OpsTab 'SQL Explorer'
$sqlMode = New-Object Windows.Forms.TabControl
$sqlMode.Dock = 'Fill'
$tabSql.Controls.Add($sqlMode)
$sqlBrowse = New-Object Windows.Forms.TabPage('Browse')
$sqlCustom = New-Object Windows.Forms.TabPage('Custom SELECT')
[void]$sqlMode.TabPages.Add($sqlBrowse)
[void]$sqlMode.TabPages.Add($sqlCustom)

$sqlDatasetBar = New-OpsToolbar $sqlBrowse
[void](New-OpsLabel $sqlDatasetBar 'Find dataset')
$sqlDatasetFind = New-OpsText $sqlDatasetBar 125
[void](New-OpsLabel $sqlDatasetBar 'Dataset')
$sqlDatasetChoice = New-OpsCombo $sqlDatasetBar @() 315 -Searchable
$sqlLoad = New-OpsButton $sqlDatasetBar 'Load results'
$sqlReload = New-OpsButton $sqlDatasetBar 'Reload datasets'
$sqlColumns = New-OpsButton $sqlDatasetBar 'Columns...'
$sqlExport = New-OpsButton $sqlDatasetBar 'Export filtered CSV...'
$sqlScopeBar = New-OpsToolbar $sqlBrowse
$sqlPicker = New-OpsFolderPicker $sqlScopeBar $script:Prefs.SqlProject $script:Prefs.SqlFolder
[void](New-OpsLabel $sqlScopeBar 'Range (Arizona)')
$sqlRange = New-OpsCombo $sqlScopeBar @(Get-QCOpsSqlTimeRangeChoices) 115
Set-OpsChoices $sqlRange @(Get-QCOpsSqlTimeRangeChoices) $(if ($script:Prefs.SqlRange) { $script:Prefs.SqlRange } else { 'Last 24 hours' })
function New-OpsSqlDate($Parent, [datetime]$Value) {
    $picker = New-Object Windows.Forms.DateTimePicker
    $picker.Format = 'Custom'
    $picker.CustomFormat = 'yyyy-MM-dd HH:mm'
    $picker.Width = 145
    $picker.Value = $Value
    Add-OpsControl $Parent $picker
}
$azNow = Get-QCWallClockNow
$sqlFrom = New-OpsSqlDate $sqlScopeBar $azNow.AddDays(-1)
$sqlTo = New-OpsSqlDate $sqlScopeBar $azNow
foreach ($pair in @(@{ c = $sqlFrom; value = $script:Prefs.SqlFrom }, @{ c = $sqlTo; value = $script:Prefs.SqlTo })) {
    if ($pair.value) { try { $pair.c.Value = [datetime]::ParseExact($pair.value,'yyyy-MM-dd HH:mm',[Globalization.CultureInfo]::InvariantCulture) } catch { } }
}
$sqlFilterBar = New-OpsToolbar $sqlBrowse
[void](New-OpsLabel $sqlFilterBar 'Filter column')
$sqlFilterColumn = New-OpsCombo $sqlFilterBar @('None') 162 -Searchable
$sqlFilterText = New-OpsText $sqlFilterBar 185
$sqlExact = New-OpsCheck $sqlFilterBar 'Exact'
$sqlSpecial = New-OpsCheck $sqlFilterBar 'Failures only'
$sqlSpecial.Visible = $false
[void](New-OpsLabel $sqlFilterBar 'Sort by')
$sqlSort = New-OpsCombo $sqlFilterBar @('Default') 153 -Searchable
$sqlOrder = New-OpsCombo $sqlFilterBar @('Newest / descending','Oldest / ascending') 155
Set-OpsChoices $sqlOrder @('Newest / descending','Oldest / ascending') $script:Prefs.SqlOrder
$sqlPager = New-OpsToolbar $sqlBrowse
$sqlPrevious = New-OpsButton $sqlPager 'Previous'
$sqlNext = New-OpsButton $sqlPager 'Next'
$sqlPrevious.Enabled = $false
$sqlNext.Enabled = $false
$sqlSize = New-OpsCombo $sqlPager @('50','100','250') 70
Set-OpsChoices $sqlSize @('50','100','250') $(if ($script:Prefs.SqlSize) { $script:Prefs.SqlSize } else { '100' })
[void](New-OpsLabel $sqlPager 'rows/page')
$sqlRelated = New-OpsCombo $sqlPager @('Related records') 190
$sqlRelatedGo = New-OpsButton $sqlPager 'Open'
$sqlNotice = New-OpsLabel $sqlPager 'Load a dataset.'
$sqlSplit = New-OpsSplit $sqlBrowse
$sqlGrid = New-OpsGrid $sqlSplit.Panel1
$sqlDetails = New-OpsDetails $sqlSplit.Panel2
$sqlDatasetBar.BringToFront()
$sqlScopeBar.BringToFront()
$sqlFilterBar.BringToFront()
$sqlPager.BringToFront()
# Fill is laid out after the top-docked bars.
$sqlSplit.BringToFront()
$sqlDatasetBar.SendToBack()
$sqlScopeBar.SendToBack()
$sqlFilterBar.SendToBack()
$sqlPager.SendToBack()
# Order top-docked bars from dataset selector to pager.
$sqlBrowse.Controls.SetChildIndex($sqlDatasetBar, $sqlBrowse.Controls.Count - 1)
$sqlBrowse.Controls.SetChildIndex($sqlScopeBar, $sqlBrowse.Controls.Count - 2)
$sqlBrowse.Controls.SetChildIndex($sqlFilterBar, $sqlBrowse.Controls.Count - 3)

$customBar = New-OpsToolbar $sqlCustom
$customRun = New-OpsButton $customBar 'Run SELECT'
$customNotice = New-OpsLabel $customBar 'Separate SELECT-only SQL connection | 250 rows | 10-second query timeout'
$customText = New-Object Windows.Forms.TextBox
$customText.Dock = 'Top'
$customText.Height = 110
$customText.Multiline = $true
$customText.ScrollBars = 'Both'
$customText.Font = New-Object Drawing.Font('Consolas', 10)
$customText.MaxLength = 8000
$sqlCustom.Controls.Add($customText)
$customBar.SendToBack()
$customSplit = New-OpsSplit $sqlCustom 280
$customGrid = New-OpsGrid $customSplit.Panel1
$customDetails = New-OpsDetails $customSplit.Panel2

$script:SqlDatasets = @()
$script:SqlLoaded = $null
$script:SqlTarget = $null
$script:SqlSelected = $null
$script:SqlChanging = $false
$script:SqlRelations = @{}

function Get-OpsFriendlyColumns([string[]]$Names) {
    $result = [ordered]@{}
    foreach ($name in $Names) {
        if (-not $name -or $result.Contains($name)) { continue }
        $label = [Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase(($name -replace '_',' '))
        $label = [regex]::Replace($label, '\b(Id|Guid|Pw|Qc|Utc|Pdf|Dgn|Sql)\b', { param($match) $match.Value.ToUpperInvariant() })
        if ($name -eq 'pw_acttime') { $label = 'ProjectWise event time' }
        if ($name -eq 'captured_at') { $label = 'Captured at' }
        $result[$name] = $label
    }
    return $result
}
function Get-OpsDataset {
    $value = $sqlDatasetChoice.Text
    $script:SqlDatasets | Where-Object { $_.label -eq $value -or $_.id -eq $value } | Select-Object -First 1
}
function Show-OpsDatasetChoices([string]$SelectedLabel) {
    $needle = $sqlDatasetFind.Text.Trim()
    $choices = @($script:SqlDatasets | Where-Object {
        -not $needle -or $_.label.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
    })
    if ($SelectedLabel -notin $choices.label) { $SelectedLabel = '' }
    $script:SqlChanging = $true
    try { Set-OpsChoices $sqlDatasetChoice @($choices.label) $SelectedLabel }
    finally { $script:SqlChanging = $false }
    Update-OpsSqlDataset
}
function Request-OpsCatalog {
    if ($SmokeTest -or $script:Workers.ContainsKey('catalog')) { return }
    Set-OpsNotice $sqlNotice 'Loading datasets...'
    Start-OpsWork catalog {
        param($config)
        $ErrorActionPreference = 'Stop'
        Get-QCOpsExplorerCatalog -Config $config
    } $script:Cfg {
        param($result,$errorText)
        if (-not $errorText -and -not $result.IsSuccess) { $errorText = $result.Message }
        if ($errorText) { Set-OpsNotice $sqlNotice "Dataset list unavailable | $errorText" -ErrorState; return }
        $selected = Get-OpsDataset
        $script:SqlDatasets = @($result.Data.items)
        $id = if ($selected) { $selected.id } else { $script:Prefs.SqlDataset }
        $previous = $script:SqlDatasets | Where-Object id -eq $id | Select-Object -First 1
        Show-OpsDatasetChoices $previous.label
        if ($script:SqlTarget) { Complete-OpsSqlTarget }
        elseif ($previous -and $previous.id -eq $script:Prefs.SqlDataset) {
            if ($script:Prefs.SqlFilterColumn -in $previous.columns.name) { $sqlFilterColumn.Text = $script:Prefs.SqlFilterColumn; $sqlFilterText.Text = [string]$script:Prefs.SqlFilterText }
            $sqlExact.Checked = [bool]$script:Prefs.SqlExact
        }
        if (-not $script:Workers.ContainsKey('sql')) { Set-OpsNotice $sqlNotice $result.Message }
    }
}
function Update-OpsSqlDataset {
    if ($script:SqlChanging) { return }
    $dataset = Get-OpsDataset
    $sqlLoad.Enabled = [bool]$dataset
    $sqlColumns.Enabled = [bool]$dataset
    if ($script:SqlSelected -and $dataset.id -eq $script:SqlSelected.id) { return }
    $script:SqlSelected = $dataset
    $sqlPrevious.Enabled = $false
    $sqlNext.Enabled = $false
    $sqlExport.Enabled = $false
    $sqlSpecial.Visible = $dataset.name -in @('poll_runs','notification_log')
    $sqlSpecial.Text = if ($dataset.name -eq 'poll_runs') { 'Reconciliations only' } else { 'Failures only' }
    $sqlSpecial.Checked = $false
    $sqlPicker.project.Enabled = [bool]$dataset.folder
    $sqlPicker.folder.Enabled = [bool]$dataset.folder
    $sqlRange.Enabled = [bool]$dataset.time
    $sqlFrom.Visible = $sqlTo.Visible = $sqlRange.Enabled -and $sqlRange.Text -eq 'Custom'
    Set-OpsChoices $sqlFilterColumn (@('None') + @($dataset.columns.name)) 'None'
    $sqlFilterText.Clear()
    $sqlExact.Checked = $false
    $sortables = @($dataset.columns | Where-Object sortable | ForEach-Object { $_.name })
    $sort = if ($dataset) { [string]$script:SortPrefs[$dataset.id] } else { '' }
    Set-OpsChoices $sqlSort (@('Default') + $sortables) $sort
    if ($dataset) {
        Set-OpsNotice $sqlNotice 'Load results to apply this dataset and its filters.'
    }
}
function Get-OpsSqlColumns($Dataset) {
    $all = @($Dataset.columns.name)
    $names = @($script:ColumnPrefs[$Dataset.id] | Where-Object { $_ -in $all })
    if (-not $names.Count) { $names = @($Dataset.defaults) }
    return $names
}
function Save-OpsSqlColumnOrder {
    if (-not $script:SqlLoaded -or -not $sqlGrid.Columns.Count) { return }
    $script:ColumnPrefs[$script:SqlLoaded.dataset.id] = @($sqlGrid.Columns | Sort-Object DisplayIndex | ForEach-Object { $_.Name })
}
function Show-OpsColumnDialog {
    $dataset = Get-OpsDataset
    if (-not $dataset) { return }
    $selected = @(Get-OpsSqlColumns $dataset)
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = 'Visible SQL columns'
    $dialog.Size = New-Object Drawing.Size(390, 535)
    $dialog.StartPosition = 'CenterParent'
    $list = New-Object Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true
    $list.Dock = 'Fill'
    $dialog.Controls.Add($list)
    foreach ($name in ($selected + @($dataset.columns.name | Where-Object { $_ -notin $selected }))) {
        [void]$list.Items.Add($name, ($name -in $selected))
    }
    $ok = New-Object Windows.Forms.Button
    $ok.Text = 'Apply columns'
    $ok.Dock = 'Bottom'
    $ok.DialogResult = 'OK'
    $dialog.AcceptButton = $ok
    $dialog.Controls.Add($ok)
    $list.BringToFront()
    try {
        if ($dialog.ShowDialog($form) -eq 'OK' -and $list.CheckedItems.Count) {
            $script:ColumnPrefs[$dataset.id] = @($list.CheckedItems)
            if ($script:SqlLoaded -and $script:SqlLoaded.dataset.id -eq $dataset.id) { Show-OpsSqlPage }
        }
    } finally { $dialog.Dispose() }
}
function Get-OpsSqlRequest {
    $dataset = Get-OpsDataset
    if (-not $dataset) { throw 'Choose an available dataset.' }
    $range = Resolve-QCOpsSqlTimeRange -Range $sqlRange.Text -CustomFrom $sqlFrom.Value -CustomTo $sqlTo.Value
    if ($dataset.time -and -not $range.IsSuccess) { throw $range.Message }
    $sort = if ($sqlSort.Text -ne 'Default') { $sqlSort.Text } else { '' }
    $script:SortPrefs[$dataset.id] = $sort
    return @{
        dataset = $dataset; sort = $sort; ascending = ($sqlOrder.SelectedIndex -eq 1); size = [int]$sqlSize.Text
        filter = @{
            Folder = (Get-OpsFolder $sqlPicker); FromUtc = $range.Data.fromUtc; ToUtc = $range.Data.toUtc
            ProjectFolders = @($script:Presets | Where-Object project -eq $sqlPicker.project.Text | ForEach-Object { $_.folder })
            Column = if ($sqlFilterColumn.Text -ne 'None') { $sqlFilterColumn.Text } else { '' }
            Text = $sqlFilterText.Text; Exact = $sqlExact.Checked
            ReconciliationOnly = ($dataset.name -eq 'poll_runs' -and $sqlSpecial.Checked)
            FailuresOnly = ($dataset.name -eq 'notification_log' -and $sqlSpecial.Checked)
        }
        config = $script:Cfg
    }
}
function Request-OpsSqlPage([int]$Page = 0, [switch]$ReuseFilters) {
    if ($script:Workers.ContainsKey('sql')) { return }
    try {
        $request = if ($ReuseFilters -and $script:SqlLoaded) { $script:SqlLoaded.request.Clone() } else { Get-OpsSqlRequest }
        $request.page = [math]::Max(0,$Page)
        if ($script:SqlLoaded -and $script:SqlLoaded.dataset.id -eq $request.dataset.id) { Save-OpsSqlColumnOrder }
    } catch { Set-OpsNotice $sqlNotice $_.Exception.Message -ErrorState; return }
    $sqlPrevious.Enabled = $false
    $sqlNext.Enabled = $false
    $sqlExport.Enabled = $false
    Set-OpsNotice $sqlNotice 'Loading results...'
    Start-OpsWork sql {
        param($p)
        $ErrorActionPreference = 'Stop'
        $result = Invoke-QCOpsExplorerPage -Config $p.config -Dataset $p.dataset -Filter $p.filter -Sort $p.sort -Ascending:$p.ascending -Page $p.page -PageSize $p.size
        if (-not $result.IsSuccess) { throw $result.Message }
        @{ dataset = $p.dataset; request = $p; table = $result.Data.table; more = $result.Data.hasMore }
    } $request {
        param($value,$errorText)
        if ($script:SqlTarget) { Complete-OpsSqlTarget; return }
        if ($errorText) { Set-OpsNotice $sqlNotice "Stale results | $errorText" -ErrorState; return }
        if ((Get-OpsDataset).id -ne $value.dataset.id) { return }
        $script:SqlLoaded = $value
        Show-OpsSqlPage
    }
}
function Show-OpsSqlPage {
    $page = $script:SqlLoaded
    if (-not $page) { return }
    Set-OpsColumns $sqlGrid (Get-OpsFriendlyColumns @(Get-OpsSqlColumns $page.dataset))
    Set-OpsRows $sqlGrid @($page.table.Rows)
    $sortColumn = if ($page.request.sort) { $page.request.sort } else { $page.dataset.time }
    if ($sortColumn -and $sqlGrid.Columns.Contains($sortColumn)) {
        $sqlGrid.Columns[$sortColumn].HeaderCell.SortGlyphDirection = if ($page.request.ascending) { 'Ascending' } else { 'Descending' }
    }
    $sqlPrevious.Enabled = $page.request.page -gt 0
    $sqlNext.Enabled = $page.more
    $sqlExport.Enabled = $true
    Show-OpsSqlDetails
    Set-OpsNotice $sqlNotice "Page $($page.request.page + 1) | $($page.table.Rows.Count) rows | $(if ($page.more) { 'More results' } else { 'End of results' })"
}
function Show-OpsSqlDetails {
    $row = Get-OpsSelected $sqlGrid
    $sqlDetails.Text = Format-OpsRecord $row
    $script:SqlRelations = @{}
    $choices = [System.Collections.Generic.List[string]]::new()
    $choices.Add('Related records')
    $targets = @(
        @{ label = 'Job timeline'; id = 'dbo.v_mcp_job_timeline'; column = 'job_id'; fields = @('job_id','enqueued_job_id') }
        @{ label = 'Job history'; id = 'dbo.processing_jobs'; column = 'job_id'; fields = @('job_id','enqueued_job_id') }
        @{ label = 'Trigger audit event'; id = 'dbo.audit_events'; column = 'id'; fields = @('trigger_audit_id','audit_event_id') }
        @{ label = 'Jobs from audit event'; id = 'dbo.processing_jobs'; column = 'trigger_audit_id'; fields = @('id'); only = @('audit_events','v_audit_events_with_user') }
        @{ label = 'Sheet package'; id = 'dbo.v_sheet_package_status'; column = 'sheet_package_id'; fields = @('sheet_package_id') }
        @{ label = 'Lane PDFs'; id = 'dbo.v_sheet_package_qc_pdf_matrix'; column = 'sheet_package_id'; fields = @('sheet_package_id') }
        @{ label = 'Document notifications'; id = 'dbo.notification_log'; column = 'document_guid'; fields = @('document_guid','pw_objguid') }
        @{ label = 'Document audit history'; id = 'dbo.v_audit_events_with_user'; column = 'pw_objguid'; fields = @('document_guid','pw_objguid') }
    )
    if ($row -and $script:SqlLoaded) {
        foreach ($target in $targets) {
            if ($target.id -eq $script:SqlLoaded.dataset.id -or $target.id -notin $script:SqlDatasets.id) { continue }
            if ($target.only -and $script:SqlLoaded.dataset.name -notin $target.only) { continue }
            $value = $target.fields | ForEach-Object { Get-OpsValue $row $_ } | Where-Object { $_ } | Select-Object -First 1
            if ($value) {
                $script:SqlRelations[$target.label] = @{ id = $target.id; column = $target.column; value = [string]$value }
                $choices.Add($target.label)
            }
        }
    }
    Set-OpsChoices $sqlRelated $choices.ToArray() ''
    $sqlRelatedGo.Enabled = $choices.Count -gt 1
}
function Open-OpsSqlTarget([string]$Id, [string]$Column, [string]$Value) {
    $script:SqlTarget = @{ id = $Id; column = $Column; value = $Value }
    $main.SelectedTab = $tabSql
    $sqlMode.SelectedTab = $sqlBrowse
    if ($script:SqlDatasets.Count) { Complete-OpsSqlTarget }
    else { Request-OpsCatalog }
}
function Complete-OpsSqlTarget {
    if (-not $script:SqlTarget -or $script:Workers.ContainsKey('sql')) { return }
    $target = $script:SqlTarget
    $script:SqlTarget = $null
    $dataset = $script:SqlDatasets | Where-Object id -eq $target.id | Select-Object -First 1
    if (-not $dataset) { Set-OpsNotice $sqlNotice "$($target.id) is unavailable." -ErrorState; return }
    $sqlDatasetFind.Clear()
    $sqlDatasetChoice.Text = $dataset.label
    $sqlPicker.project.Text = 'All projects'
    $sqlPicker.folder.Text = ''
    $sqlRange.Text = 'All'
    $sqlSpecial.Checked = $false
    $sqlFilterColumn.Text = if ($target.column -in $dataset.columns.name -and $target.value) { $target.column } else { 'None' }
    $sqlFilterText.Text = [string]$target.value
    $sqlExact.Checked = [bool]$target.value
    Request-OpsSqlPage
}
function Start-OpsSqlExport {
    if ($script:Workers.ContainsKey('export') -or -not $script:SqlLoaded) { return }
    Save-OpsSqlColumnOrder
    $dialog = New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = $script:SqlLoaded.dataset.name + '.csv'
    try {
        if ($dialog.ShowDialog($form) -ne 'OK') { return }
        $path = $dialog.FileName
    } finally { $dialog.Dispose() }
    $request = $script:SqlLoaded.request.Clone()
    $request.path = $path
    $request.columns = @(Get-OpsSqlColumns $request.dataset)
    Set-OpsNotice $sqlNotice 'Exporting up to 5,000 matching rows...'
    Start-OpsWork export {
        param($p)
        $ErrorActionPreference = 'Stop'
        $result = Invoke-QCOpsExplorerPage -Config $p.config -Dataset $p.dataset -Filter $p.filter -Sort $p.sort -Ascending:$p.ascending -Export -Projection $p.columns
        if (-not $result.IsSuccess) { throw $result.Message }
        Export-QCOpsExplorerCsv -Table $result.Data.table -Columns $p.columns -Path $p.path
        @{ count = $result.Data.rowCount; capped = $result.Data.hasMore; path = $p.path }
    } $request {
        param($value,$errorText)
        if ($errorText) { Set-OpsNotice $sqlNotice "Export failed | $errorText" -ErrorState; return }
        $limit = if ($value.capped) { ' (limit reached)' } else { '' }
        Set-OpsNotice $sqlNotice "$($value.count) rows$limit | $($value.path)"
    }
}
function Save-OpsPreferences {
    Save-OpsSqlColumnOrder
    $settings = [ordered]@{
        Width = $form.RestoreBounds.Width; Height = $form.RestoreBounds.Height
        QueueState = $queueState.Text; QueueOrder = $queueOrder.Text; QueueText = $queueFind.Text; QueueType = $queueType.Text
        HistoryKind = $historyKind.Text; HistoryOrder = $historyOrder.Text
        LogOrder = $logOrder.Text; LogRange = $logRange.Text; LogSource = $logSource.Text; LogLevel = $logLevel.Text; LogText = $logFind.Text
        SheetProject = $sheetPicker.project.Text; SheetFolder = $sheetPicker.folder.Text; Action = $actionChoice.Text
        SqlDataset = if ((Get-OpsDataset)) { (Get-OpsDataset).id } else { $script:Prefs.SqlDataset }
        SqlColumns = $script:ColumnPrefs; SqlSort = $script:SortPrefs; SqlOrder = $sqlOrder.Text; SqlSize = $sqlSize.Text
        SqlProject = $sqlPicker.project.Text; SqlFolder = $sqlPicker.folder.Text; SqlRange = $sqlRange.Text
        SqlFrom = $sqlFrom.Value.ToString('yyyy-MM-dd HH:mm'); SqlTo = $sqlTo.Value.ToString('yyyy-MM-dd HH:mm')
        SqlFilterColumn = $sqlFilterColumn.Text; SqlFilterText = $sqlFilterText.Text; SqlExact = $sqlExact.Checked
    }
    try {
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $script:PrefPath))
        $settings | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:PrefPath -Encoding utf8
    } catch { Write-Warning "Could not save console preferences: $($_.Exception.Message)" }
}
$sqlDatasetChoice.Add_TextChanged({ Update-OpsSqlDataset })
$sqlDatasetFind.Add_TextChanged({ Show-OpsDatasetChoices $sqlDatasetChoice.Text })
$sqlRange.Add_SelectedIndexChanged({ $sqlFrom.Visible = $sqlTo.Visible = ($sqlRange.Enabled -and $sqlRange.Text -eq 'Custom') })
$sqlReload.Add_Click({ Request-OpsCatalog })
$sqlLoad.Add_Click({ Request-OpsSqlPage })
$sqlColumns.Add_Click({ Show-OpsColumnDialog })
$sqlExport.Add_Click({ Start-OpsSqlExport })
$sqlPrevious.Add_Click({ Request-OpsSqlPage ($script:SqlLoaded.request.page - 1) -ReuseFilters })
$sqlNext.Add_Click({ Request-OpsSqlPage ($script:SqlLoaded.request.page + 1) -ReuseFilters })
$sqlRelatedGo.Add_Click({
    $target = $script:SqlRelations[$sqlRelated.Text]
    if ($target) { Open-OpsSqlTarget $target.id $target.column $target.value }
})
$sqlGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { Show-OpsSqlDetails } })
$sqlGrid.Add_ColumnHeaderMouseClick({
    $name = $sqlGrid.Columns[$args[1].ColumnIndex].Name
    if ($sqlSort.Items.Contains($name)) {
        if ($sqlSort.Text -eq $name) { $sqlOrder.SelectedIndex = 1 - $sqlOrder.SelectedIndex }
        else { $sqlSort.Text = $name; $sqlOrder.SelectedIndex = 0 }
        Request-OpsSqlPage
    }
})
$customRun.Add_Click({
    if ($script:Workers.ContainsKey('custom')) { return }
    Set-OpsNotice $customNotice 'Running SELECT...'
    Start-OpsWork custom {
        param($query)
        $ErrorActionPreference = 'Stop'
        Invoke-QCOpsCustomSelect -Query $query
    } $customText.Text {
        param($result,$errorText)
        if (-not $errorText -and -not $result.IsSuccess) { $errorText = $result.Message }
        if ($errorText) { Set-OpsNotice $customNotice "Stale results | $errorText" -ErrorState; return }
        $table = $result.Data.table
        Set-OpsColumns $customGrid (Get-OpsFriendlyColumns @($table.Columns.ColumnName))
        Set-OpsRows $customGrid @($table.Rows)
        $customDetails.Text = Format-OpsRecord (Get-OpsSelected $customGrid)
        Set-OpsNotice $customNotice "$($table.Rows.Count) rows$(if ($result.Data.hasMore) { ' | 250-row limit reached' })"
    }
})
$customGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { $customDetails.Text = Format-OpsRecord (Get-OpsSelected $customGrid) } })
Update-OpsSqlDataset
