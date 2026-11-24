$tabOverview = New-OpsTab 'Overview'
$overviewBar = New-OpsToolbar $tabOverview
$startPipeline = New-OpsButton $overviewBar 'Start pipeline'
$stopPipeline = New-OpsButton $overviewBar 'Stop pipeline'
$scanButton = New-OpsButton $overviewBar 'Request full scan'
$tickButton = New-OpsButton $overviewBar 'Run one watcher pass'
$overviewRefresh = New-OpsButton $overviewBar 'Refresh'
$overviewFresh = New-OpsLabel $overviewBar 'Loading...'
$overview = New-OpsGrid $tabOverview ([ordered]@{ item = 'Health'; value = 'Current state' })
$overview.Columns['item'].Width = 225
$overview.Columns['value'].AutoSizeMode = 'Fill'
$script:LatestStatus = $null
foreach ($button in @($startPipeline,$stopPipeline,$scanButton,$tickButton)) { $button.Enabled = $false }

$tabQueue = New-OpsTab 'Queue'
$queueBar = New-OpsToolbar $tabQueue
$queueState = New-OpsCombo $queueBar @('Failed','Held','Pending','Running','Succeeded','All') 115
$queueType = New-OpsCombo $queueBar @('All types') 160 -Searchable
[void](New-OpsLabel $queueBar 'Find job, document or folder')
$queueFind = New-OpsText $queueBar 220
$queueOrder = New-OpsCombo $queueBar @('Newest first','Oldest first') 112
$queueRefresh = New-OpsButton $queueBar 'Refresh'
$queueTools = New-OpsToolbar $tabQueue
$queueRequeue = New-OpsButton $queueTools 'Requeue selected'
$queueWriteback = New-OpsCheck $queueTools 'Writeback only'
$queueRecover = New-OpsButton $queueTools 'Recover stale jobs'
$queueSql = New-OpsButton $queueTools 'SQL job history'
$queueNotice = New-OpsLabel $queueTools 'Waiting for queue snapshot.'
$queueSplit = New-OpsSplit $tabQueue
$queueGrid = New-OpsGrid $queueSplit.Panel1 ([ordered]@{
    lastWriteTime = 'Updated (Arizona)'; state = 'State'; type = 'Job type'; sourceName = 'Document / work'
    sourceFolder = 'Folder'; machineName = 'Worker'; disposition = 'Disposition'; checkpoint = 'Checkpoint'
    errorDetail = 'Last error'; priority = 'Priority'; jobId = 'Job ID'
})
$queueDetails = New-OpsDetails $queueSplit.Panel2
$script:QueueRows = @()
Set-OpsChoices $queueState @('Failed','Held','Pending','Running','Succeeded','All') $script:Prefs.QueueState
Set-OpsChoices $queueOrder @('Newest first','Oldest first') $script:Prefs.QueueOrder
if ($script:Prefs.QueueType) { $queueType.Text = $script:Prefs.QueueType }
$queueFind.Text = [string]$script:Prefs.QueueText

$tabLogs = New-OpsTab 'Logs'
$logBar = New-OpsToolbar $tabLogs
$logPaused = New-OpsCheck $logBar 'Pause live updates'
$logOrder = New-OpsCombo $logBar @('Newest first','Oldest first') 112
$logRange = New-OpsCombo $logBar @('All loaded','Last hour','Last 6 hours','Last 24 hours') 105
$logSource = New-OpsCombo $logBar @('All components','watcher','processor','dashboard') 145
$logLevel = New-OpsCombo $logBar @('All severities','Errors','Warnings','Information','Debug') 125
[void](New-OpsLabel $logBar 'Search')
$logFind = New-OpsText $logBar 225
$logRefresh = New-OpsButton $logBar 'Refresh'
$logRelated = New-OpsButton $logBar 'Related SQL'
$logRelated.Enabled = $false
$logNotice = New-OpsLabel $logBar 'Waiting for log snapshot.'
$logSplit = New-OpsSplit $tabLogs
$logGrid = New-OpsGrid $logSplit.Panel1 ([ordered]@{
    time = 'Timestamp (Arizona)'; severity = 'Severity'; source = 'Component'; code = 'Event'
    job = 'Job ID'; document = 'Document'; message = 'Message'
})
$logGrid.Columns['message'].AutoSizeMode = 'Fill'
$logDetails = New-OpsDetails $logSplit.Panel2
$logSearchDelay = New-Object Windows.Forms.Timer
$logSearchDelay.Interval = 225
$script:LogRows = @()
Set-OpsChoices $logOrder @('Newest first','Oldest first') $script:Prefs.LogOrder
Set-OpsChoices $logRange @('All loaded','Last hour','Last 6 hours','Last 24 hours') $script:Prefs.LogRange
Set-OpsChoices $logSource @('All components','watcher','processor','dashboard') $script:Prefs.LogSource
Set-OpsChoices $logLevel @('All severities','Errors','Warnings','Information','Debug') $script:Prefs.LogLevel
$logFind.Text = [string]$script:Prefs.LogText

$tabHistory = New-OpsTab 'History'
$historyBar = New-OpsToolbar $tabHistory
$historyMap = [ordered]@{
    'Audit polls' = 'dbo.poll_runs'; 'Reconciliations' = 'dbo.poll_runs'
    'Audit events' = 'dbo.v_audit_events_with_user'; 'Job outcomes' = 'dbo.processing_jobs'
    'Notification failures' = 'dbo.notification_log'; 'All notifications' = 'dbo.notification_log'
    'Automation errors' = 'dbo.v_mcp_recent_errors'
}
$historyKind = New-OpsCombo $historyBar @($historyMap.Keys) 176
$historyOrder = New-OpsCombo $historyBar @('Newest first','Oldest first') 112
$historyRefresh = New-OpsButton $historyBar 'Refresh'
$historySql = New-OpsButton $historyBar 'Explore in SQL'
$historyNotice = New-OpsLabel $historyBar 'Open history to load recent records.'
$historySplit = New-OpsSplit $tabHistory
$historyGrid = New-OpsGrid $historySplit.Panel1
$historyDetails = New-OpsDetails $historySplit.Panel2
$script:HistoryDataset = $null
Set-OpsChoices $historyKind @($historyMap.Keys) $script:Prefs.HistoryKind
Set-OpsChoices $historyOrder @('Newest first','Oldest first') $script:Prefs.HistoryOrder

$tabSheets = New-OpsTab 'Sheets'
$sheetBar = New-OpsToolbar $tabSheets
[void](New-OpsLabel $sheetBar 'Sheet stem, document GUID or PW path')
$sheetQuery = New-OpsText $sheetBar 270
$sheetSearch = New-OpsButton $sheetBar 'Search SQL'
$sheetCompare = New-OpsButton $sheetBar 'Compare selected with ProjectWise'
$sheetMatrix = New-OpsButton $sheetBar 'Lane PDF matrix'
$sheetNotice = New-OpsLabel $sheetBar ''
$sheetFolders = New-OpsToolbar $tabSheets
$sheetPicker = New-OpsFolderPicker $sheetFolders $script:Prefs.SheetProject $script:Prefs.SheetFolder
$sheetSplit = New-OpsSplit $tabSheets
$sheetGrid = New-OpsGrid $sheetSplit.Panel1
$sheetDetails = New-OpsDetails $sheetSplit.Panel2

function Show-OpsSnapshot($Snapshot) {
    $status = $Snapshot.status
    $script:LatestStatus = $status
    $startPipeline.Enabled = $status.task.registered -and -not $status.pipelineOn
    $stopPipeline.Enabled = $status.task.registered -and ($status.task.enabled -or $status.lock.alive)
    $tickButton.Enabled = $status.hostAllowed -and $status.task.registered -and -not $status.task.enabled -and -not $status.lock.alive -and $status.task.state -ne 'Running'
    $scanButton.Enabled = $status.pipelineOn
    $stamp = Format-QCDisplayTime -Value $Snapshot.takenAt
    $age = if ($null -ne $status.watermarkAgeSeconds) { [timespan]::FromSeconds([math]::Max(0,$status.watermarkAgeSeconds)).ToString('g') } else { 'Unavailable' }
    $oldest = if ($Snapshot.oldest) {
        $wait = ($Snapshot.takenAt - $Snapshot.oldest.CreationTimeUtc).ToString('g')
        "$wait | $($Snapshot.oldest.BaseName)"
    } else { 'No pending files' }
    $scan = if ($Snapshot.scan.slotKey) { "$($Snapshot.scan.completedFolders.Count) folders complete, $($Snapshot.scan.folderQueue.Count) remaining | $($Snapshot.scan.slotKey)" } else { 'No checkpoint in progress' }
    if ($status.opsRequest) { $scan = 'Full scan requested | ' + $scan }
    $items = [ordered]@{
        'Watcher' = "$($status.stateLabel) | ProjectWise $($status.pwText)"
        'Last watcher tick' = if ($status.lastTick) { Format-QCDisplayTime -Value $status.lastTick.ts } else { 'Unavailable' }
        'Audit watermark age' = $age
        'Queue' = if ($Snapshot.queueAvailable) { "$($status.queue.pending) pending | $($status.queue.running) running | $($status.queue.failed) failed | $($Snapshot.holds) lane holds" } else { 'Queue folder unavailable' }
        'Oldest pending' = $oldest
        'Reconciliation' = $scan
        'Scheduled scans (Arizona)' = $Snapshot.schedule -join ', '
        'Worker hosts' = ($status.hostRows.text -join ' | ')
        'Status Set batch' = $status.statusSetBatchText
        'Execution profile' = "Dry run: $($status.dryRunText) | SQL telemetry: $($status.sqlEnabled)"
    }
    if ($status.lastStall) { $items['Last watcher stall'] = "$($status.lastStall.code) | $(Format-QCDisplayTime -Value $status.lastStall.ts)" }
    if ($status.lastNotificationFail) { $items['Last notification failure'] = "$($status.lastNotificationFail.code) | $(Format-QCDisplayTime -Value $status.lastNotificationFail.ts)" }
    Set-OpsRows $overview @(foreach ($name in $items.Keys) { @{ item = $name; value = $items[$name] } }) item
    foreach ($row in $overview.Rows) {
        $item = $row.Tag.item
        if ($item -in @('Last watcher stall','Last notification failure') -or
            ($item -eq 'Queue' -and (-not $Snapshot.queueAvailable -or $Snapshot.holds -gt 0 -or $status.queue.failed -gt 0))) {
            $row.DefaultCellStyle.ForeColor = [Drawing.Color]::Firebrick
        }
    }
    Set-OpsNotice $overviewFresh "Updated $stamp"
    if ($Snapshot.includeQueue) {
        if (-not $Snapshot.queueAvailable) { Set-OpsNotice $queueNotice 'Stale data | Queue folder unavailable' -ErrorState; return }
        $script:QueueRows = @($Snapshot.queue)
        $types = @('All types') + @($script:QueueRows.type | Where-Object { $_ } | Sort-Object -Unique)
        if ($queueType.Text -and $queueType.Text -notin $types) { $types += $queueType.Text }
        if (($types -join '|') -ne (@($queueType.Items) -join '|')) { Set-OpsChoices $queueType $types $queueType.Text }
        Show-OpsQueue
        Set-OpsNotice $queueNotice "$($queueGrid.Rows.Count) shown | up to 250 per queue state, plus deferred | $stamp"
    }
    if ($Snapshot.includeLogs -and -not $logPaused.Checked) {
        if (-not $Snapshot.logsAvailable) { Set-OpsNotice $logNotice 'Log folder unavailable or not created yet.' -ErrorState; return }
        $script:LogRows = @($Snapshot.logs)
        Show-OpsLogs
        Set-OpsNotice $logNotice "$($logGrid.Rows.Count) shown | $stamp"
    }
}
function Show-OpsQueue {
    $query = $queueFind.Text.Trim()
    $rows = @(foreach ($row in $script:QueueRows) {
        if ($queueState.Text -eq 'Held') {
            if ($row.disposition -notin @('needs_operator','lane_held')) { continue }
        } elseif ($queueState.Text -ne 'All' -and $row.state -ne $queueState.Text) { continue }
        if ($queueType.Text -and $queueType.Text -ne 'All types' -and $row.type -ne $queueType.Text) { continue }
        if ($query -and ($row.jobId,$row.sourceName,$row.sourceFolder,$row.errorDetail -join ' ').IndexOf($query,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $row
    })
    if ($queueOrder.Text -eq 'Oldest first') { [array]::Reverse($rows) }
    Set-OpsRows $queueGrid $rows key
    Show-OpsQueueDetails
}
function Show-OpsQueueDetails {
    $record = Get-OpsSelected $queueGrid
    $queueRequeue.Enabled = $false
    $queueSql.Enabled = [bool]$record.jobId -and -not $record.deferred
    $queueDetails.Clear()
    if (-not $record) { return }
    $reason = Get-QCOpsRequeueBlockReason -JobRow $record -WritebackOnly:$queueWriteback.Checked
    $queueRequeue.Enabled = -not $reason -and [bool]$record.path
    $detail = if ($record.detail) { $record.detail } else { $record }
    $queueDetails.Text = $(if ($reason) { $reason } else { 'Eligible for requeue.' }) + [Environment]::NewLine + (Format-OpsRecord $detail)
    if (-not $record.path) { return }
    Start-OpsWork queueRecord {
        param($p)
        @{ id = $p.id; path = $p.path; text = (Get-Content -LiteralPath $p.path -Raw -ErrorAction Stop) }
    } @{ id = $record.jobId; path = $record.path } {
        param($result, $errorText)
        $current = Get-OpsSelected $queueGrid
        if ($result -and $current.jobId -eq $result.id -and $current.path -eq $result.path) {
            $reason = Get-QCOpsRequeueBlockReason -JobRow $current -WritebackOnly:$queueWriteback.Checked
            $queueDetails.Text = $(if ($reason) { $reason } else { 'Eligible for requeue.' }) + [Environment]::NewLine + $result.text
        } elseif ($current -and $result -and $current.path -ne $result.path) {
            Show-OpsQueueDetails
        }
    }
}
function Show-OpsLogs {
    $query = $logFind.Text.Trim()
    $cutoff = switch ($logRange.Text) {
        'Last hour' { [datetime]::UtcNow.AddHours(-1) }
        'Last 6 hours' { [datetime]::UtcNow.AddHours(-6) }
        'Last 24 hours' { [datetime]::UtcNow.AddHours(-24) }
        default { [datetime]::MinValue }
    }
    $rows = @(foreach ($row in $script:LogRows) {
        if ($row.when -lt $cutoff) { continue }
        if ($logSource.Text -ne 'All components' -and $row.source -ne $logSource.Text) { continue }
        if ($logLevel.Text -eq 'Errors' -and $row.severity -notmatch 'error|fatal|critical') { continue }
        if ($logLevel.Text -eq 'Warnings' -and $row.severity -notmatch 'warn') { continue }
        if ($logLevel.Text -in @('Information','Debug') -and $row.severity -ne $logLevel.Text) { continue }
        if ($query -and ($row.code,$row.job,$row.document,$row.message,$row.source -join ' ').IndexOf($query,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $row
    })
    $rows = @($rows | Select-Object -First 1500)
    if ($logOrder.Text -eq 'Oldest first') { [array]::Reverse($rows) }
    Set-OpsRows $logGrid $rows key -ColorSeverity
    Show-OpsLogDetails
}
function Show-OpsLogDetails {
    $record = Get-OpsSelected $logGrid
    $logRelated.Enabled = [bool]($record.job -or $record.document)
    if (-not $record) { $logDetails.Clear(); return }
    $detail = [string]$record.detail.message + [Environment]::NewLine + [Environment]::NewLine + (Format-OpsRecord $record.detail)
    if ($detail -cne $logDetails.Text) { $logDetails.Text = $detail }
}
function Request-OpsHistory {
    if ($SmokeTest -or $script:Workers.ContainsKey('history')) { return }
    $kind = $historyKind.Text
    $dataset = $script:SqlDatasets | Where-Object id -eq $historyMap[$kind] | Select-Object -First 1
    Set-OpsNotice $historyNotice 'Loading history...'
    Start-OpsWork history {
        param($p)
        $ErrorActionPreference = 'Stop'
        $dataset = $p.dataset
        if (-not $dataset) {
            $catalog = Get-QCOpsExplorerCatalog -Config $p.config
            if (-not $catalog.IsSuccess) { throw $catalog.Message }
            $dataset = $catalog.Data.items | Where-Object id -eq $p.id | Select-Object -First 1
        }
        if (-not $dataset) { throw "Dataset $($p.id) is unavailable in this database." }
        $filter = @{ ReconciliationOnly = ($p.kind -eq 'Reconciliations'); FailuresOnly = ($p.kind -eq 'Notification failures') }
        if ($p.kind -eq 'Audit polls') { $filter.AuditOnly = $true }
        $page = Invoke-QCOpsExplorerPage -Config $p.config -Dataset $dataset -Filter $filter -Ascending:$p.ascending -PageSize 150
        if (-not $page.IsSuccess) { throw $page.Message }
        @{ dataset = $dataset; page = $page.Data; kind = $p.kind; ascending = $p.ascending }
    } @{ config = $script:Cfg; dataset = $dataset; id = $historyMap[$kind]; kind = $kind; ascending = ($historyOrder.Text -eq 'Oldest first') } {
        param($value, $errorText)
        if ($errorText) { Set-OpsNotice $historyNotice "Stale or unavailable | $errorText" -ErrorState; return }
        if ($value.kind -ne $historyKind.Text -or $value.ascending -ne ($historyOrder.Text -eq 'Oldest first')) { Request-OpsHistory; return }
        $script:HistoryDataset = $value.dataset
        Set-OpsColumns $historyGrid (Get-OpsFriendlyColumns $value.dataset.defaults)
        Set-OpsRows $historyGrid @($value.page.table.Rows)
        $historyDetails.Text = Format-OpsRecord (Get-OpsSelected $historyGrid)
        Set-OpsNotice $historyNotice "$($value.page.rowCount) rows | $(if ($value.page.hasMore) { 'More in SQL Explorer' } else { 'Up to date' })"
    }
}
function Request-OpsSheets([bool]$Compare) {
    if ($script:Workers.ContainsKey('sheets')) { return }
    $selected = Get-OpsSelected $sheetGrid
    $query = $sheetQuery.Text.Trim()
    $folder = Get-OpsFolder $sheetPicker
    $guid = ''
    $package = ''
    $documentPath = ''
    if ($Compare) {
        $guid = [string](Get-OpsValue $selected 'document_guid')
        $package = [string](Get-OpsValue $selected 'sheet_package_id')
        if (-not $guid -and ($query -as [guid])) { $guid = $query }
        if ($query -match '[\\/]') { $documentPath = $query }
    }
    if (-not ($query -or $folder -or $guid -or $package)) { Set-OpsNotice $sheetNotice 'Enter a stem, GUID, folder or document path.' -ErrorState; return }
    if (-not $query) { $query = $folder }
    Set-OpsNotice $sheetNotice 'Loading sheet diagnostics...'
    Start-OpsWork sheets {
        param($p)
        $ErrorActionPreference = 'Stop'
        if ($p.compare) {
            $result = Invoke-QCOpsPwCompareChild -AppSettingsPath $p.settings -SheetNumber $p.query -DocumentGuid $p.guid -PackageId $p.package -DocumentPath $p.documentPath
        } else {
            $result = Search-QCOpsSheets -Config $p.config -Query $p.query -Folder $p.folder
        }
        if (-not $result.IsSuccess) { throw $result.Message }
        @{ result = $result.Data; compare = $p.compare; package = $p.package }
    } @{ compare = $Compare; settings = $AppSettingsPath; query = $query; guid = $guid; package = $package; documentPath = $documentPath; folder = $folder; config = $script:Cfg } {
        param($value, $errorText)
        if ($errorText) { Set-OpsNotice $sheetNotice "Stale or unavailable | $errorText" -ErrorState; return }
        $result = $value.result
        if ($value.compare) {
            Set-OpsColumns $sheetGrid (Get-OpsFriendlyColumns @('role','document_name','database_state','projectwise_state','state_match','database_qc_process_type','projectwise_qc_process_type','qc_process_type_match','document_guid'))
            $rows = @($result.data.comparisons | Where-Object { $_ })
            if ($value.package) { foreach ($row in $rows) { $row.sheet_package_id = $value.package } }
        } else {
            Set-OpsColumns $sheetGrid (Get-OpsFriendlyColumns @('dataset','sheet_stem','folder_path','document_name','document_guid','pw_state_name','qc_process_type','sheet_package_id'))
            $rows = @(foreach ($key in $result.data.matches_by_table.Keys) {
                foreach ($row in $result.data.matches_by_table[$key]) {
                    $row.dataset = $key
                    $row
                }
            })
        }
        Set-OpsRows $sheetGrid $rows
        $sheetDetails.Text = Format-OpsRecord (Get-OpsSelected $sheetGrid)
        $warnings = @($result.warnings | Where-Object { $_ }).Count
        $message = "$($rows.Count) records | $warnings warnings"
        if ($value.compare -and -not $result.data.projectwise_available) { $message += ' | ProjectWise unavailable' }
        Set-OpsNotice $sheetNotice $message -ErrorState:($warnings -gt 0)
        $script:SheetDiagnostic = $result
        if ($warnings -or -not $rows.Count) { $sheetDetails.Text = Format-OpsRecord $result }
    }
}

$overviewRefresh.Add_Click({ Request-OpsSnapshot })
$startPipeline.Add_Click({ Invoke-OpsControl Start-QCOpsPipeline @{} 'Starting the scheduled pipeline...' })
$stopPipeline.Add_Click({
    if (Confirm-Ops 'Disable the scheduled task and stop its watcher and workers?') {
        Invoke-OpsControl Stop-QCOpsPipeline @{ Config = $script:Cfg } 'Stopping the pipeline...'
    }
})
$scanButton.Add_Click({ Invoke-OpsControl Request-QCOpsFullScan @{ Config = $script:Cfg } 'Requesting full reconciliation...' })
$tickButton.Add_Click({ Invoke-OpsManualPass })
$queueRefresh.Add_Click({ Request-OpsSnapshot })
$queueState.Add_SelectedIndexChanged({ Show-OpsQueue })
$queueType.Add_TextChanged({ Show-OpsQueue })
$queueFind.Add_TextChanged({ Show-OpsQueue })
$queueOrder.Add_SelectedIndexChanged({ Show-OpsQueue })
$queueGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { Show-OpsQueueDetails } })
$queueWriteback.Add_CheckedChanged({ Show-OpsQueueDetails })
$queueRequeue.Add_Click({
    $record = Get-OpsSelected $queueGrid
    if ($record -and (Confirm-Ops "Requeue $($record.jobId)?")) {
        Invoke-OpsControl Invoke-QCOpsRequeueJobs @{ Config = $script:Cfg; JobPaths = @($record.path); WritebackOnly = $queueWriteback.Checked } 'Requeueing job...'
    }
})
$queueRecover.Add_Click({
    if (Confirm-Ops 'Run the existing stale-job recovery for this queue?') {
        Invoke-OpsControl Invoke-QCOpsRecoverStaleJobs @{ Config = $script:Cfg } 'Checking stale jobs...'
    }
})
$queueSql.Add_Click({ $row = Get-OpsSelected $queueGrid; if ($row) { Open-OpsSqlTarget 'dbo.processing_jobs' 'job_id' $row.jobId } })
$logSource.Add_SelectedIndexChanged({ Show-OpsLogs })
$logLevel.Add_SelectedIndexChanged({ Show-OpsLogs })
$logOrder.Add_SelectedIndexChanged({ Show-OpsLogs })
$logRange.Add_SelectedIndexChanged({ Show-OpsLogs })
$logFind.Add_TextChanged({ $logSearchDelay.Stop(); $logSearchDelay.Start() })
$logSearchDelay.Add_Tick({ $this.Stop(); Show-OpsLogs })
$logPaused.Add_CheckedChanged({
    $logRefresh.Enabled = -not $this.Checked
    if ($this.Checked) { Set-OpsNotice $logNotice "Paused | $($logGrid.Rows.Count) shown" }
    else { Request-OpsSnapshot }
})
$logRefresh.Add_Click({ Request-OpsSnapshot })
$logRelated.Add_Click({
    $record = Get-OpsSelected $logGrid
    if ($record.job) { Open-OpsSqlTarget 'dbo.v_mcp_job_timeline' 'job_id' $record.job }
    elseif ($record.document -as [guid]) { Open-OpsSqlTarget 'dbo.v_audit_events_with_user' 'pw_objguid' $record.document }
    elseif ($record.document) { Open-OpsSqlTarget 'dbo.notification_log' 'document_name' $record.document }
})
$logGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { Show-OpsLogDetails } })
$historyKind.Add_SelectedIndexChanged({ Request-OpsHistory })
$historyOrder.Add_SelectedIndexChanged({ Request-OpsHistory })
$historyRefresh.Add_Click({ Request-OpsHistory })
$historySql.Add_Click({
    $kind = $historyKind.Text
    $column = ''
    $value = ''
    if ($kind -in @('Audit polls','Reconciliations')) {
        $column = 'is_reconciliation'
        $value = if ($kind -eq 'Reconciliations') { '1' } else { '0' }
    } elseif ($kind -eq 'Notification failures') { $column = 'success'; $value = '0' }
    Open-OpsSqlTarget $historyMap[$kind] $column $value
})
$historyGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { $historyDetails.Text = Format-OpsRecord (Get-OpsSelected $historyGrid) } })
$sheetSearch.Add_Click({ Request-OpsSheets $false })
$sheetCompare.Add_Click({ Request-OpsSheets $true })
$sheetMatrix.Add_Click({
    $package = [string](Get-OpsValue (Get-OpsSelected $sheetGrid) 'sheet_package_id')
    if ($package) { Open-OpsSqlTarget 'dbo.v_sheet_package_qc_pdf_matrix' 'sheet_package_id' $package }
    else { Open-OpsSqlTarget 'dbo.v_sheet_package_qc_pdf_matrix' 'sheet_number' $sheetQuery.Text }
})
$sheetGrid.Add_SelectionChanged({ if (-not $script:UpdatingGrid) { $sheetDetails.Text = Format-OpsRecord (Get-OpsSelected $sheetGrid) } })
