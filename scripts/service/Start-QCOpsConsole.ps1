<#
.SYNOPSIS
Logon-session WinForms console for the QC pipeline (PXBENTLEY01).

.DESCRIPTION
Observes and controls the Session 0 QC-PipelineDashboard task. Does not start
Start-QCPipelineDashboard.ps1. Close this window anytime; the pipeline keeps running.

.EXAMPLE
powershell.exe -STA -NoProfile -File .\scripts\service\Start-QCOpsConsole.ps1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppSettingsPath,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'

if ($Host.Name -eq 'ConsoleHost' -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $arg = @(
        '-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath
    )
    if ($AppSettingsPath) { $arg += @('-AppSettingsPath', $AppSettingsPath) }
    if ($Force.IsPresent) { $arg += '-Force' }
    Start-Process -FilePath 'powershell.exe' -ArgumentList $arg
    return
}

$scriptsRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptsRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $scriptsRoot 'Restore-QCModuleExports.ps1') -RepoRoot $repoRoot
Import-QCModuleBootstrapSet -RepoRoot $repoRoot -FeatureModules @(
    'Core\Core.Results.psm1'
    'Core\Core.Runtime.psm1'
    'Queue\QC.Queue.Json.psm1'
    'Ops\QC.OpsConsole.psm1'
) -RequiredCommands @(
    'Get-QCAppSettingsConfig'
    'Get-QCOpsPipelineStatus'
    'Start-QCOpsPipeline'
    'Stop-QCOpsPipeline'
) -Context 'ops console bootstrap'

if ($NoGui.IsPresent) {
    & (Join-Path $PSScriptRoot 'Watch-QCPipelineDashboardConsole.ps1') -AppSettingsPath $AppSettingsPath
    return
}

if (-not (Test-QCOpsHostAllowed -Force:$Force)) {
    $msg = "QC ops console is for $(Get-QCOpsExpectedHost) (this host is $env:COMPUTERNAME). Pass -Force to override."
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($msg, 'QC ops console') | Out-Null
    } catch {
        Write-Host $msg -ForegroundColor Red
    }
    exit 2
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Cfg = Get-QCAppSettingsConfig -Path $AppSettingsPath
try { $script:Cfg['_appSettingsPath'] = $AppSettingsPath } catch { }
$script:AppSettingsPath = $AppSettingsPath
$script:SuppressToggle = $false
$script:Status = $null

function _Ops-Msg([string]$Text, [string]$Title = 'QC ops') {
    [System.Windows.Forms.MessageBox]::Show($Text, $Title) | Out-Null
}

function _Ops-Confirm([string]$Text) {
    $r = [System.Windows.Forms.MessageBox]::Show($Text, 'Confirm', 'YesNo', 'Warning')
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function _Ops-FormatEvent($Obj) {
    if (-not $Obj) { return '' }
    $code = ''; $msg = ''; $ts = ''
    try { $code = [string]$Obj.code } catch { }
    try { $msg = [string]$Obj.message } catch { }
    try { if ($Obj.ts) { $ts = [string]$Obj.ts } } catch { }
    return ('{0} {1} {2}' -f $ts, $code, $msg)
}

function _Ops-RefreshStatus {
    $script:Status = Get-QCOpsPipelineStatus -Config $script:Cfg
    return $script:Status
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'QC Pipeline Ops'
$form.Size = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 88
$form.Controls.Add($header)

$chkOn = New-Object System.Windows.Forms.CheckBox
$chkOn.Text = 'Pipeline on'
$chkOn.Location = New-Object System.Drawing.Point(12, 10)
$chkOn.AutoSize = $true
$header.Controls.Add($chkOn)

$lblState = New-Object System.Windows.Forms.Label
$lblState.Location = New-Object System.Drawing.Point(140, 12)
$lblState.Size = New-Object System.Drawing.Size(400, 20)
$lblState.Text = 'State: ...'
$header.Controls.Add($lblState)

$lblLine2 = New-Object System.Windows.Forms.Label
$lblLine2.Location = New-Object System.Drawing.Point(12, 36)
$lblLine2.Size = New-Object System.Drawing.Size(1080, 20)
$header.Controls.Add($lblLine2)

$lblLine3 = New-Object System.Windows.Forms.Label
$lblLine3.Location = New-Object System.Drawing.Point(12, 58)
$lblLine3.Size = New-Object System.Drawing.Size(1080, 22)
$header.Controls.Add($lblLine3)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Padding = New-Object System.Drawing.Point(8, 8)
$form.Controls.Add($tabs)
$form.Controls.SetChildIndex($tabs, 0)
$form.Controls.SetChildIndex($header, 1)

function _New-Tab([string]$Title) {
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = $Title
    $p.UseVisualStyleBackColor = $true
    [void]$tabs.TabPages.Add($p)
    return $p
}

# --- Overview ---
$tabOverview = _New-Tab 'Overview'
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = 'Both'
$logBox.ReadOnly = $true
$logBox.Dock = 'Fill'
$logBox.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabOverview.Controls.Add($logBox)

# --- Queue ---
$tabQueue = _New-Tab 'Queue'
$queueTop = New-Object System.Windows.Forms.Panel
$queueTop.Dock = 'Top'
$queueTop.Height = 36
$tabQueue.Controls.Add($queueTop)
$cmbQueueState = New-Object System.Windows.Forms.ComboBox
$cmbQueueState.DropDownStyle = 'DropDownList'
@('pending', 'running', 'failed', 'succeeded') | ForEach-Object { [void]$cmbQueueState.Items.Add($_) }
$cmbQueueState.SelectedIndex = 1
$cmbQueueState.Location = New-Object System.Drawing.Point(8, 6)
$cmbQueueState.Width = 140
$queueTop.Controls.Add($cmbQueueState)
$btnQueueRefresh = New-Object System.Windows.Forms.Button
$btnQueueRefresh.Text = 'Refresh'
$btnQueueRefresh.Location = New-Object System.Drawing.Point(156, 4)
$queueTop.Controls.Add($btnQueueRefresh)
$btnRequeue = New-Object System.Windows.Forms.Button
$btnRequeue.Text = 'Requeue selected'
$btnRequeue.Location = New-Object System.Drawing.Point(250, 4)
$btnRequeue.Width = 140
$queueTop.Controls.Add($btnRequeue)
$btnRecover = New-Object System.Windows.Forms.Button
$btnRecover.Text = 'Recover stale'
$btnRecover.Location = New-Object System.Drawing.Point(400, 4)
$btnRecover.Width = 120
$queueTop.Controls.Add($btnRecover)
$btnOpenJson = New-Object System.Windows.Forms.Button
$btnOpenJson.Text = 'Open JSON'
$btnOpenJson.Location = New-Object System.Drawing.Point(528, 4)
$btnOpenJson.Width = 100
$queueTop.Controls.Add($btnOpenJson)
$chkWritebackOnly = New-Object System.Windows.Forms.CheckBox
$chkWritebackOnly.Text = 'Writeback-only requeue'
$chkWritebackOnly.Location = New-Object System.Drawing.Point(640, 8)
$chkWritebackOnly.AutoSize = $true
$queueTop.Controls.Add($chkWritebackOnly)
$gridQueue = New-Object System.Windows.Forms.DataGridView
$gridQueue.Dock = 'Fill'
$gridQueue.ReadOnly = $true
$gridQueue.AllowUserToAddRows = $false
$gridQueue.SelectionMode = 'FullRowSelect'
$gridQueue.AutoSizeColumnsMode = 'Fill'
$tabQueue.Controls.Add($gridQueue)
$tabQueue.Controls.SetChildIndex($gridQueue, 0)

# --- Runs ---
$tabRuns = _New-Tab 'Runs'
$runsSplit = New-Object System.Windows.Forms.SplitContainer
$runsSplit.Dock = 'Fill'
$runsSplit.Orientation = 'Horizontal'
$runsSplit.SplitterDistance = 280
$tabRuns.Controls.Add($runsSplit)
$gridRuns = New-Object System.Windows.Forms.DataGridView
$gridRuns.Dock = 'Fill'
$gridRuns.ReadOnly = $true
$gridRuns.AllowUserToAddRows = $false
$gridRuns.AutoSizeColumnsMode = 'Fill'
$runsSplit.Panel1.Controls.Add($gridRuns)
$runsBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$runsBtns.Dock = 'Fill'
$runsBtns.Padding = New-Object System.Windows.Forms.Padding(8)
$runsSplit.Panel2.Controls.Add($runsBtns)
$btnFullScan = New-Object System.Windows.Forms.Button
$btnFullScan.Text = 'Request full folder scan'
$btnFullScan.AutoSize = $true
$btnReconcileSs = New-Object System.Windows.Forms.Button
$btnReconcileSs.Text = 'Reconcile status sets'
$btnReconcileSs.AutoSize = $true
$btnEnqReport = New-Object System.Windows.Forms.Button
$btnEnqReport.Text = 'Enqueue QC_REPORTING_SCAN'
$btnEnqReport.AutoSize = $true
$btnEnqComment = New-Object System.Windows.Forms.Button
$btnEnqComment.Text = 'Enqueue QC_COMMENT_STATUS_SYNC'
$btnEnqComment.AutoSize = $true
$btnTickHint = New-Object System.Windows.Forms.Button
$btnTickHint.Text = 'Run audit tick now'
$btnTickHint.AutoSize = $true
[void]$runsBtns.Controls.AddRange(@($btnFullScan, $btnReconcileSs, $btnEnqReport, $btnEnqComment, $btnTickHint))

# --- Sheets ---
$tabSheets = _New-Tab 'Sheets'
$sheetTop = New-Object System.Windows.Forms.Panel
$sheetTop.Dock = 'Top'
$sheetTop.Height = 36
$tabSheets.Controls.Add($sheetTop)
$txtSheet = New-Object System.Windows.Forms.TextBox
$txtSheet.Location = New-Object System.Drawing.Point(8, 6)
$txtSheet.Width = 360
$sheetTop.Controls.Add($txtSheet)
$btnSheetSearch = New-Object System.Windows.Forms.Button
$btnSheetSearch.Text = 'Search'
$btnSheetSearch.Location = New-Object System.Drawing.Point(376, 4)
$sheetTop.Controls.Add($btnSheetSearch)
$btnPwCompare = New-Object System.Windows.Forms.Button
$btnPwCompare.Text = 'Compare PW (child)'
$btnPwCompare.Location = New-Object System.Drawing.Point(470, 4)
$btnPwCompare.Width = 150
$sheetTop.Controls.Add($btnPwCompare)
$txtSheetOut = New-Object System.Windows.Forms.TextBox
$txtSheetOut.Multiline = $true
$txtSheetOut.ScrollBars = 'Both'
$txtSheetOut.ReadOnly = $true
$txtSheetOut.Dock = 'Fill'
$txtSheetOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabSheets.Controls.Add($txtSheetOut)
$tabSheets.Controls.SetChildIndex($txtSheetOut, 0)

# --- SQL ---
$tabSql = _New-Tab 'SQL'
$sqlTop = New-Object System.Windows.Forms.Panel
$sqlTop.Dock = 'Top'
$sqlTop.Height = 72
$tabSql.Controls.Add($sqlTop)
$cmbSql = New-Object System.Windows.Forms.ComboBox
$cmbSql.DropDownStyle = 'DropDownList'
@('v_poller_health', 'v_job_summary', 'poll_runs', 'processing_jobs', 'audit_events', 'notification_log') | ForEach-Object { [void]$cmbSql.Items.Add($_) }
$cmbSql.SelectedIndex = 0
$cmbSql.Location = New-Object System.Drawing.Point(8, 6)
$cmbSql.Width = 220
$sqlTop.Controls.Add($cmbSql)
$btnSqlLoad = New-Object System.Windows.Forms.Button
$btnSqlLoad.Text = 'Load'
$btnSqlLoad.Location = New-Object System.Drawing.Point(236, 4)
$sqlTop.Controls.Add($btnSqlLoad)
$lblWatermark = New-Object System.Windows.Forms.Label
$lblWatermark.Location = New-Object System.Drawing.Point(330, 8)
$lblWatermark.Size = New-Object System.Drawing.Size(500, 20)
$sqlTop.Controls.Add($lblWatermark)
$txtWmUtc = New-Object System.Windows.Forms.TextBox
$txtWmUtc.Location = New-Object System.Drawing.Point(8, 40)
$txtWmUtc.Width = 280
$txtWmUtc.Text = ''
$sqlTop.Controls.Add($txtWmUtc)
$txtWmConfirm = New-Object System.Windows.Forms.TextBox
$txtWmConfirm.Location = New-Object System.Drawing.Point(296, 40)
$txtWmConfirm.Width = 180
$sqlTop.Controls.Add($txtWmConfirm)
$btnWmRewind = New-Object System.Windows.Forms.Button
$btnWmRewind.Text = 'Rewind watermark'
$btnWmRewind.Location = New-Object System.Drawing.Point(484, 38)
$btnWmRewind.Width = 140
$sqlTop.Controls.Add($btnWmRewind)
$lblWmHint = New-Object System.Windows.Forms.Label
$lblWmHint.Text = 'UTC datetime + type REWIND WATERMARK'
$lblWmHint.Location = New-Object System.Drawing.Point(632, 42)
$lblWmHint.AutoSize = $true
$sqlTop.Controls.Add($lblWmHint)
$gridSql = New-Object System.Windows.Forms.DataGridView
$gridSql.Dock = 'Fill'
$gridSql.ReadOnly = $true
$gridSql.AllowUserToAddRows = $false
$gridSql.AutoSizeColumnsMode = 'DisplayedCells'
$tabSql.Controls.Add($gridSql)
$tabSql.Controls.SetChildIndex($gridSql, 0)

# --- Notifications ---
$tabNotif = _New-Tab 'Notifications'
$txtNotif = New-Object System.Windows.Forms.TextBox
$txtNotif.Multiline = $true
$txtNotif.ScrollBars = 'Both'
$txtNotif.ReadOnly = $true
$txtNotif.Dock = 'Fill'
$txtNotif.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabNotif.Controls.Add($txtNotif)

# --- Hosts ---
$tabHosts = _New-Tab 'Hosts'
$txtHosts = New-Object System.Windows.Forms.TextBox
$txtHosts.Multiline = $true
$txtHosts.ScrollBars = 'Both'
$txtHosts.ReadOnly = $true
$txtHosts.Dock = 'Fill'
$txtHosts.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabHosts.Controls.Add($txtHosts)

# --- Maintenance ---
$tabMaint = _New-Tab 'Maintenance'
$lstMaint = New-Object System.Windows.Forms.ListBox
$lstMaint.Dock = 'Top'
$lstMaint.Height = 180
$tabMaint.Controls.Add($lstMaint)
$maintPanel = New-Object System.Windows.Forms.Panel
$maintPanel.Dock = 'Top'
$maintPanel.Height = 70
$tabMaint.Controls.Add($maintPanel)
$txtConfirm = New-Object System.Windows.Forms.TextBox
$txtConfirm.Location = New-Object System.Drawing.Point(8, 8)
$txtConfirm.Width = 280
$txtConfirm.Text = ''
$maintPanel.Controls.Add($txtConfirm)
$lblConfirm = New-Object System.Windows.Forms.Label
$lblConfirm.Text = 'High-danger: type YES (or REWIND WATERMARK)'
$lblConfirm.Location = New-Object System.Drawing.Point(300, 12)
$lblConfirm.AutoSize = $true
$maintPanel.Controls.Add($lblConfirm)
$btnMaint = New-Object System.Windows.Forms.Button
$btnMaint.Text = 'Run selected'
$btnMaint.Location = New-Object System.Drawing.Point(8, 36)
$maintPanel.Controls.Add($btnMaint)
$txtMaintOut = New-Object System.Windows.Forms.TextBox
$txtMaintOut.Multiline = $true
$txtMaintOut.ScrollBars = 'Both'
$txtMaintOut.ReadOnly = $true
$txtMaintOut.Dock = 'Fill'
$txtMaintOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabMaint.Controls.Add($txtMaintOut)
$tabMaint.Controls.SetChildIndex($txtMaintOut, 0)

$script:MaintCatalog = @(Get-QCOpsMaintenanceCatalog)
foreach ($item in $script:MaintCatalog) {
    [void]$lstMaint.Items.Add(('{0}  [{1}]  {2}' -f $item.label, $item.danger, $item.script))
}

function _Ops-ApplyHeader {
    $st = _Ops-RefreshStatus
    $d = $st.Data
    $script:SuppressToggle = $true
    $chkOn.Checked = [bool]$d.pipelineOn
    $script:SuppressToggle = $false
    $lblState.Text = ('State: {0}   task={1} ({2})   lock={3}' -f $d.stateLabel, $d.task.state, $(if ($d.task.enabled) { 'enabled' } else { 'disabled' }), $d.lock.text)
    $dryTxt = 'n/a'
    try { if ($d.dryRun -and $d.dryRun.globalDryRun) { $dryTxt = 'global dry-run ON' } elseif ($d.dryRun) { $dryTxt = 'dry-run layered' } else { $dryTxt = 'live' } } catch { }
    $wm = 'n/a'
    if ($null -ne $d.watermarkAgeSeconds) { $wm = ('{0}s ago' -f $d.watermarkAgeSeconds) }
    $lblLine2.Text = ('PW={0}   SQL={1}   DryRun={2}   watermark={3}' -f $d.pwText, $d.sqlEnabled, $dryTxt, $wm)
    $tick = ''
    if ($d.lastTick -and $d.lastTick.ts) { $tick = [string]$d.lastTick.ts }
    $stall = ''
    if ($d.lastStall) { $stall = [string]$d.lastStall.code }
    $notifFail = ''
    if ($d.lastNotificationFail) { $notifFail = [string]$d.lastNotificationFail.code }
    $lblLine3.Text = ('Queue pending={0} running={1} failed={2} locks={3}  local={4} remote={5}  lastTick={6}  stall={7}  notifFail={8}' -f `
            $d.queue.pending, $d.queue.running, $d.queue.failed, $d.queue.locks, $d.localRunning, $d.remoteRunning, $tick, $stall, $notifFail)
}

function _Ops-BindGrid($Grid, $Rows) {
    $al = New-Object System.Collections.ArrayList
    foreach ($r in @($Rows)) { [void]$al.Add($r) }
    $Grid.DataSource = $null
    $Grid.DataSource = $al
}

function _Ops-RefreshOverviewLog {
    $root = [string]$script:Status.Data.queueRoot
    $ev = Get-QCOpsRecentLogEvents -QueueRoot $root -MaxLines 50
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($o in @($ev.watcher)) { [void]$lines.Add((_Ops-FormatEvent $o)) }
    foreach ($o in @($ev.processor)) { [void]$lines.Add((_Ops-FormatEvent $o)) }
    $logBox.Text = (($lines | Select-Object -Last 80) -join [Environment]::NewLine)
}

function _Ops-RefreshQueue {
    $state = [string]$cmbQueueState.SelectedItem
    if (-not $state) { $state = 'running' }
    $rows = @(Get-QCOpsRunningJobRows -QueueRoot $script:Status.Data.queueRoot -State $state)
    _Ops-BindGrid $gridQueue $rows
}

function _Ops-RefreshRuns {
    $res = Get-QCOpsLastRuns -Config $script:Cfg
    $rows = @()
    if ($res.IsSuccess) { $rows = @($res.Data.rows) }
    _Ops-BindGrid $gridRuns $rows
}

function _Ops-RefreshNotif {
    $res = Get-QCOpsNotificationSummary -Config $script:Cfg
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(('Dedupe: {0}' -f $res.Data.dedupePath))
    [void]$sb.AppendLine('--- recent sent-keys ---')
    foreach ($k in @($res.Data.recentKeys)) { [void]$sb.AppendLine([string]$k) }
    [void]$sb.AppendLine('--- notification_log ---')
    foreach ($r in @($res.Data.sqlRows)) {
        [void]$sb.AppendLine(('{0}  {1}  success={2}  {3}' -f $r.sent_at, $r.event_type, $r.success, $r.document_name))
    }
    $txtNotif.Text = $sb.ToString()
}

function _Ops-RefreshHosts {
    $res = Get-QCOpsHostSummary -Config $script:Cfg
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(('This host: {0}   modelling PC: {1}' -f $res.Data.localHost, $res.Data.modellingHost))
    foreach ($h in @($res.Data.byHost.Keys)) {
        $jobs = @($res.Data.byHost[$h])
        [void]$sb.AppendLine(('{0}: {1} running' -f $h, $jobs.Count))
        foreach ($j in $jobs) {
            [void]$sb.AppendLine(('  {0} {1} {2}' -f $j.type, $j.sourceName, $j.checkpoint))
        }
    }
    foreach ($t in @($res.Data.throttleFiles)) {
        [void]$sb.AppendLine(('Throttle {0}: {1}' -f $t.file, ($t.doc | ConvertTo-Json -Compress)))
    }
    $txtHosts.Text = $sb.ToString()
}

$chkOn.Add_CheckedChanged({
    if ($script:SuppressToggle) { return }
    if ($chkOn.Checked) {
        $r = Start-QCOpsPipeline
        if (-not $r.IsSuccess) {
            _Ops-Msg $r.Message
        }
    } else {
        if (-not (_Ops-Confirm 'Disable the boot task and stop watcher/workers? In-flight prepend will block unless you confirm force.')) {
            $script:SuppressToggle = $true
            $chkOn.Checked = $true
            $script:SuppressToggle = $false
            return
        }
        $r = Stop-QCOpsPipeline -Config $script:Cfg
        if (-not $r.IsSuccess) {
            if ($r.Code -eq 'OPS_STOP_IN_FLIGHT_PREPEND') {
                if (_Ops-Confirm ($r.Message + ' Force stop anyway?')) {
                    $r = Stop-QCOpsPipeline -Config $script:Cfg -Force
                }
            }
            if (-not $r.IsSuccess) { _Ops-Msg $r.Message }
        }
    }
    _Ops-ApplyHeader
})

$btnQueueRefresh.Add_Click({ _Ops-RefreshQueue })
$cmbQueueState.Add_SelectedIndexChanged({ _Ops-RefreshQueue })
$btnRequeue.Add_Click({
    $paths = @()
    foreach ($row in @($gridQueue.SelectedRows)) {
        try { $paths += [string]$row.DataBoundItem.path } catch { }
    }
    if ($paths.Count -eq 0) { _Ops-Msg 'Select one or more jobs.'; return }
    $r = Invoke-QCOpsRequeueJobs -Config $script:Cfg -JobPaths $paths -WritebackOnly:$chkWritebackOnly.Checked
    _Ops-Msg $r.Message
    _Ops-RefreshQueue
    _Ops-ApplyHeader
})
$btnRecover.Add_Click({
    $r = Invoke-QCOpsRecoverStaleJobs -Config $script:Cfg
    _Ops-Msg $(if ($r.IsSuccess) { $r.Message } else { $r.Message })
    _Ops-RefreshQueue
})
$btnOpenJson.Add_Click({
    foreach ($row in @($gridQueue.SelectedRows)) {
        $p = ''
        try { $p = [string]$row.DataBoundItem.path } catch { }
        if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath 'notepad.exe' -ArgumentList $p }
    }
})
$gridQueue.Add_CellDoubleClick({
    $row = $gridQueue.CurrentRow
    if (-not $row) { return }
    $p = ''
    try { $p = [string]$row.DataBoundItem.path } catch { }
    if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath 'notepad.exe' -ArgumentList $p }
})

$btnFullScan.Add_Click({
    $r = Request-QCOpsFullScan -Config $script:Cfg
    _Ops-Msg $(if ($r.IsSuccess) { 'Full scan requested. The live watcher will pick it up on the next tick.' } else { $r.Message })
    _Ops-RefreshRuns
})
$btnTickHint.Add_Click({
    $on = [bool]$script:Status.Data.pipelineOn
    if ($on) {
        _Ops-Msg 'Watcher is already cycling. Use Request full folder scan for an off-schedule reconcile. Do not start a second Watch-QCTrigger.'
        return
    }
    $watch = Join-Path $PSScriptRoot 'Watch-QCTrigger.ps1'
    Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-MTA', '-File', $watch, '-AppSettingsPath', $script:AppSettingsPath)
    _Ops-Msg 'Started a one-shot Watch-QCTrigger because the pipeline is off.'
})
$btnReconcileSs.Add_Click({
    $cat = @($script:MaintCatalog | Where-Object { $_.id -eq 'reconcile-status-sets' })[0]
    $r = Invoke-QCOpsMaintenanceScript -ScriptPath $cat.script -AppSettingsPath $script:AppSettingsPath
    _Ops-Msg $r.Message
})
$btnEnqReport.Add_Click({
    $r = Invoke-QCOpsEnqueueJobType -Config $script:Cfg -JobType 'QC_REPORTING_SCAN'
    _Ops-Msg $r.Message
})
$btnEnqComment.Add_Click({
    $r = Invoke-QCOpsEnqueueJobType -Config $script:Cfg -JobType 'QC_COMMENT_STATUS_SYNC'
    _Ops-Msg $r.Message
})

$btnSheetSearch.Add_Click({
    $q = [string]$txtSheet.Text
    $r = Search-QCOpsSheets -Config $script:Cfg -Query $q
    if ($r.IsSuccess) {
        $txtSheetOut.Text = ($r.Data | ConvertTo-Json -Depth 8)
    } else {
        $txtSheetOut.Text = $r.Message
    }
})
$btnPwCompare.Add_Click({
    $q = [string]$txtSheet.Text
    $txtSheetOut.Text = 'Running PW compare in MTA child...'
    $r = Invoke-QCOpsPwCompareChild -AppSettingsPath $script:AppSettingsPath -SheetNumber $q
    if ($r.IsSuccess) {
        $txtSheetOut.Text = ($r.Data | ConvertTo-Json -Depth 6)
    } else {
        $txtSheetOut.Text = $r.Message
    }
})

$btnSqlLoad.Add_Click({
    $name = [string]$cmbSql.SelectedItem
    $r = Get-QCOpsSqlTablePreview -Config $script:Cfg -TableOrView $name -Top 80
    if (-not $r.IsSuccess) { _Ops-Msg $r.Message; return }
    $gridSql.DataSource = $r.Data.table
    $wm = $script:Status.Data.watermarkUtc
    $lblWatermark.Text = ('Watermark: {0}' -f $wm)
})
$btnWmRewind.Add_Click({
    $raw = [string]$txtWmUtc.Text
    $dt = $null
    try {
        $dt = [datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        if ($dt.Kind -ne [DateTimeKind]::Utc) { $dt = $dt.ToUniversalTime() }
    } catch { $dt = $null }
    if (-not $dt) { _Ops-Msg 'Enter a UTC datetime (ISO-8601).'; return }
    $r = Set-QCOpsAuditWatermark -Config $script:Cfg -WatermarkUtc $dt -ConfirmText ([string]$txtWmConfirm.Text)
    _Ops-Msg $r.Message
    _Ops-ApplyHeader
})

$btnMaint.Add_Click({
    $idx = $lstMaint.SelectedIndex
    if ($idx -lt 0) { _Ops-Msg 'Select a maintenance action.'; return }
    $item = $script:MaintCatalog[$idx]
    $confirmNeeded = 'YES'
    if ($item.confirmText) { $confirmNeeded = [string]$item.confirmText }
    if ([string]$item.danger -eq 'High' -and [string]$txtConfirm.Text -ne $confirmNeeded) {
        _Ops-Msg ('High-danger action. Type {0} in the confirm box.' -f $confirmNeeded)
        return
    }
    if ($item.needsFolder) {
        _Ops-Msg 'This action needs a folder path. Run the printed script from PowerShell with -FolderPath (or equivalent).'
        $txtMaintOut.Text = $item.script
        return
    }
    if ([string]$item.id -eq 'rewind-watermark') {
        _Ops-Msg 'Use the SQL tab Rewind watermark box (UTC datetime + REWIND WATERMARK).'
        return
    }
    if (-not (_Ops-Confirm ('Run {0}?' -f $item.script))) { return }
    $maintArgs = @()
    if ($item.args) { $maintArgs = @($item.args) }
    $r = Invoke-QCOpsMaintenanceScript -ScriptPath $item.script -ArgumentList $maintArgs -AppSettingsPath $script:AppSettingsPath
    $txtMaintOut.Text = ('{0}`r`n{1}' -f $r.Message, $r.Data.output)
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2000
$script:SlowTick = 0
$timer.Add_Tick({
    try {
        _Ops-ApplyHeader
        if ($tabs.SelectedTab -eq $tabOverview) { _Ops-RefreshOverviewLog }
        $script:SlowTick++
        if ($script:SlowTick -ge 5) {
            $script:SlowTick = 0
            if ($tabs.SelectedTab -eq $tabQueue) { _Ops-RefreshQueue }
            if ($tabs.SelectedTab -eq $tabRuns) { _Ops-RefreshRuns }
            if ($tabs.SelectedTab -eq $tabNotif) { _Ops-RefreshNotif }
            if ($tabs.SelectedTab -eq $tabHosts) { _Ops-RefreshHosts }
        }
    } catch { }
})

$form.Add_Shown({
    _Ops-ApplyHeader
    _Ops-RefreshOverviewLog
    _Ops-RefreshQueue
    _Ops-RefreshRuns
    $timer.Start()
})
$form.Add_FormClosed({ $timer.Stop() })

[void]$form.ShowDialog()
