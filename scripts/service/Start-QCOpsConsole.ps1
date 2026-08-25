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
$script:StateColor = [System.Drawing.Color]::Gray
$script:LogEvents = @()
$script:LogStamp = ''
$script:QueueStamp = ''
$script:TickN = 0

function _Ops-Msg([string]$Text, [string]$Title = 'QC ops') {
    [System.Windows.Forms.MessageBox]::Show($Text, $Title) | Out-Null
}

function _Ops-Confirm([string]$Text) {
    $r = [System.Windows.Forms.MessageBox]::Show($Text, 'Confirm', 'YesNo', 'Warning')
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function _Ops-FormatEvent($Obj) {
    if (-not $Obj) { return '' }
    $code = ''; $msg = ''; $ts = ''; $src = ''; $level = ''
    try { $code = [string]$Obj.code } catch { }
    try { $msg = [string]$Obj.message } catch { }
    try { if ($Obj.ts) { $ts = [string]$Obj.ts } } catch { }
    try { if ($Obj.logSource) { $src = [string]$Obj.logSource } } catch { }
    try { if ($Obj.level) { $level = [string]$Obj.level } } catch { }
    return ('{0}  {1}  {2}  {3}  {4}' -f $ts, $src, $level, $code, $msg)
}

function _Ops-RefreshStatus {
    param([switch]$Light)
    $st = Get-QCOpsPipelineStatus -Config $script:Cfg -Light:$Light
    if ($Light.IsPresent -and $script:Status -and $script:Status.Data) {
        $prev = $script:Status.Data
        $d = $st.Data
        $d.queue = $prev.queue
        $d.pwText = $prev.pwText
        $d.lastTick = $prev.lastTick
        $d.lastStall = $prev.lastStall
        $d.lastNotificationFail = $prev.lastNotificationFail
        $d.watermarkUtc = $prev.watermarkUtc
        $d.watermarkAgeSeconds = $prev.watermarkAgeSeconds
        if (-not $d.throttleSummary) { $d.throttleSummary = $prev.throttleSummary }
        if (-not $d.hostRows -or @($d.hostRows).Count -eq 0) { $d.hostRows = $prev.hostRows }
    }
    $script:Status = $st
    return $script:Status
}

function _Ops-StateColor([string]$Label) {
    switch ($Label) {
        'Running' { return [System.Drawing.Color]::FromArgb(40, 167, 69) }
        'Ready' { return [System.Drawing.Color]::FromArgb(240, 173, 78) }
        'Stale lock' { return [System.Drawing.Color]::FromArgb(253, 126, 20) }
        'Stopped' { return [System.Drawing.Color]::FromArgb(108, 117, 125) }
        'Disabled' { return [System.Drawing.Color]::FromArgb(108, 117, 125) }
        default { return [System.Drawing.Color]::FromArgb(220, 53, 69) }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'QC Pipeline Ops'
$form.Size = New-Object System.Drawing.Size(1280, 820)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1000, 640)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 156
$header.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$header.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
$form.Controls.Add($header)

$headerSplit = New-Object System.Windows.Forms.TableLayoutPanel
$headerSplit.Dock = 'Fill'
$headerSplit.ColumnCount = 2
$headerSplit.RowCount = 1
$headerSplit.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$headerSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 58)))
[void]$headerSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 42)))
[void]$headerSplit.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$header.Controls.Add($headerSplit)

$leftHead = New-Object System.Windows.Forms.Panel
$leftHead.Dock = 'Fill'
$headerSplit.Controls.Add($leftHead, 0, 0)

$chkOn = New-Object System.Windows.Forms.CheckBox
$chkOn.Text = 'Pipeline on'
$chkOn.Location = New-Object System.Drawing.Point(4, 6)
$chkOn.AutoSize = $true
$leftHead.Controls.Add($chkOn)

$pnlBubble = New-Object System.Windows.Forms.Panel
$pnlBubble.Location = New-Object System.Drawing.Point(132, 8)
$pnlBubble.Size = New-Object System.Drawing.Size(18, 18)
$pnlBubble.Add_Paint({
    $e = $args[1]
    if (-not $e) { return }
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $brush = New-Object System.Drawing.SolidBrush $script:StateColor
    $g.FillEllipse($brush, 1, 1, 15, 15)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60, 60, 60))
    $g.DrawEllipse($pen, 1, 1, 15, 15)
    $brush.Dispose()
    $pen.Dispose()
})
$leftHead.Controls.Add($pnlBubble)

$lblState = New-Object System.Windows.Forms.Label
$lblState.Location = New-Object System.Drawing.Point(156, 6)
$lblState.Size = New-Object System.Drawing.Size(400, 20)
$lblState.Anchor = 'Top,Left,Right'
$lblState.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblState.Text = 'State: ...'
$leftHead.Controls.Add($lblState)

$lblLine2 = New-Object System.Windows.Forms.Label
$lblLine2.Location = New-Object System.Drawing.Point(4, 30)
$lblLine2.Size = New-Object System.Drawing.Size(400, 18)
$lblLine2.Anchor = 'Top,Left,Right'
$leftHead.Controls.Add($lblLine2)

$pnlHosts = New-Object System.Windows.Forms.Panel
$pnlHosts.Location = New-Object System.Drawing.Point(4, 50)
$pnlHosts.Size = New-Object System.Drawing.Size(400, 40)
$pnlHosts.Anchor = 'Top,Left,Right'
$leftHead.Controls.Add($pnlHosts)

$lblHostLocal = New-Object System.Windows.Forms.Label
$lblHostLocal.Location = New-Object System.Drawing.Point(0, 0)
$lblHostLocal.Size = New-Object System.Drawing.Size(400, 18)
$lblHostLocal.Anchor = 'Top,Left,Right'
$lblHostLocal.Text = '...'
$pnlHosts.Controls.Add($lblHostLocal)

$lblHostRemote = New-Object System.Windows.Forms.Label
$lblHostRemote.Location = New-Object System.Drawing.Point(0, 18)
$lblHostRemote.Size = New-Object System.Drawing.Size(400, 20)
$lblHostRemote.Anchor = 'Top,Left,Right'
$lblHostRemote.Text = '...'
$pnlHosts.Controls.Add($lblHostRemote)
$script:HostStatusLabels = @($lblHostLocal, $lblHostRemote)

$lblLine3 = New-Object System.Windows.Forms.Label
$lblLine3.Location = New-Object System.Drawing.Point(4, 92)
$lblLine3.Size = New-Object System.Drawing.Size(400, 48)
$lblLine3.Anchor = 'Top,Left,Right'
$lblLine3.Text = ''
$leftHead.Controls.Add($lblLine3)

$rightHead = New-Object System.Windows.Forms.Panel
$rightHead.Dock = 'Fill'
$rightHead.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$headerSplit.Controls.Add($rightHead, 1, 0)

$queueScore = New-Object System.Windows.Forms.TableLayoutPanel
$queueScore.Dock = 'Fill'
$queueScore.ColumnCount = 5
$queueScore.RowCount = 2
$queueScore.BackColor = [System.Drawing.Color]::White
$queueScore.CellBorderStyle = 'Single'
$queueScore.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$queueScore.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 38)))
[void]$queueScore.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 62)))
$headerNames = @('pending', 'running', 'failed', 'succeeded', 'locks')
$script:QueueScoreLabels = @{}
$qi = 0
foreach ($name in $headerNames) {
    [void]$queueScore.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    $h = New-Object System.Windows.Forms.Label
    $h.Text = $name
    $h.Dock = 'Fill'
    $h.TextAlign = 'MiddleCenter'
    $h.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $h.Margin = New-Object System.Windows.Forms.Padding(0)
    $queueScore.Controls.Add($h, $qi, 0)
    $c = New-Object System.Windows.Forms.Label
    $c.Text = '0'
    $c.Dock = 'Fill'
    $c.TextAlign = 'MiddleCenter'
    $c.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
    $c.Margin = New-Object System.Windows.Forms.Padding(0)
    $queueScore.Controls.Add($c, $qi, 1)
    $script:QueueScoreLabels[$name] = $c
    $qi++
}
$rightHead.Controls.Add($queueScore)

$leftHead.Add_Resize({
    $w = [Math]::Max(180, $leftHead.ClientSize.Width - 12)
    $lblState.Width = [Math]::Max(120, $w - 156)
    $lblLine2.Width = $w
    $pnlHosts.Width = $w
    $lblHostLocal.Width = $w
    $lblHostRemote.Width = $w
    $lblLine3.Width = $w
})

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
$logTop = New-Object System.Windows.Forms.Panel
$logTop.Dock = 'Top'
$logTop.Height = 36
$tabOverview.Controls.Add($logTop)
$lblLogTime = New-Object System.Windows.Forms.Label
$lblLogTime.Text = 'Time'
$lblLogTime.Location = New-Object System.Drawing.Point(8, 10)
$lblLogTime.AutoSize = $true
$logTop.Controls.Add($lblLogTime)
$cmbLogTime = New-Object System.Windows.Forms.ComboBox
$cmbLogTime.DropDownStyle = 'DropDownList'
@('Last 15 min', 'Last hour', 'Last 6 hours', 'Last 24 hours', 'All loaded') | ForEach-Object { [void]$cmbLogTime.Items.Add($_) }
$cmbLogTime.SelectedIndex = 1
$cmbLogTime.Location = New-Object System.Drawing.Point(48, 6)
$cmbLogTime.Width = 130
$logTop.Controls.Add($cmbLogTime)
$lblLogType = New-Object System.Windows.Forms.Label
$lblLogType.Text = 'Type'
$lblLogType.Location = New-Object System.Drawing.Point(190, 10)
$lblLogType.AutoSize = $true
$logTop.Controls.Add($lblLogType)
$cmbLogType = New-Object System.Windows.Forms.ComboBox
$cmbLogType.DropDownStyle = 'DropDownList'
@('All', 'Watcher', 'Processor', 'Error', 'Warning') | ForEach-Object { [void]$cmbLogType.Items.Add($_) }
$cmbLogType.SelectedIndex = 0
$cmbLogType.Location = New-Object System.Drawing.Point(228, 6)
$cmbLogType.Width = 110
$logTop.Controls.Add($cmbLogType)
$lblLogCode = New-Object System.Windows.Forms.Label
$lblLogCode.Text = 'Code'
$lblLogCode.Location = New-Object System.Drawing.Point(350, 10)
$lblLogCode.AutoSize = $true
$logTop.Controls.Add($lblLogCode)
$txtLogCode = New-Object System.Windows.Forms.TextBox
$txtLogCode.Location = New-Object System.Drawing.Point(390, 6)
$txtLogCode.Width = 180
$logTop.Controls.Add($txtLogCode)
$btnLogApply = New-Object System.Windows.Forms.Button
$btnLogApply.Text = 'Apply'
$btnLogApply.Location = New-Object System.Drawing.Point(580, 4)
$logTop.Controls.Add($btnLogApply)
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = 'Both'
$logBox.ReadOnly = $true
$logBox.Dock = 'Fill'
$logBox.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabOverview.Controls.Add($logBox)
$tabOverview.Controls.SetChildIndex($logBox, 0)

# --- Queue ---
$tabQueue = _New-Tab 'Queue'
$queueTop = New-Object System.Windows.Forms.Panel
$queueTop.Dock = 'Top'
$queueTop.Height = 36
$tabQueue.Controls.Add($queueTop)
$btnQueueRefresh = New-Object System.Windows.Forms.Button
$btnQueueRefresh.Text = 'Refresh'
$btnQueueRefresh.Location = New-Object System.Drawing.Point(8, 4)
$queueTop.Controls.Add($btnQueueRefresh)
$btnRequeue = New-Object System.Windows.Forms.Button
$btnRequeue.Text = 'Requeue selected'
$btnRequeue.Location = New-Object System.Drawing.Point(100, 4)
$btnRequeue.Width = 140
$queueTop.Controls.Add($btnRequeue)
$btnRecover = New-Object System.Windows.Forms.Button
$btnRecover.Text = 'Recover stale'
$btnRecover.Location = New-Object System.Drawing.Point(250, 4)
$btnRecover.Width = 120
$queueTop.Controls.Add($btnRecover)
$btnOpenJson = New-Object System.Windows.Forms.Button
$btnOpenJson.Text = 'Open JSON'
$btnOpenJson.Location = New-Object System.Drawing.Point(380, 4)
$btnOpenJson.Width = 100
$queueTop.Controls.Add($btnOpenJson)
$chkWritebackOnly = New-Object System.Windows.Forms.CheckBox
$chkWritebackOnly.Text = 'Writeback-only requeue'
$chkWritebackOnly.Location = New-Object System.Drawing.Point(496, 8)
$chkWritebackOnly.AutoSize = $true
$queueTop.Controls.Add($chkWritebackOnly)

$queueLayout = New-Object System.Windows.Forms.TableLayoutPanel
$queueLayout.Dock = 'Fill'
$queueLayout.ColumnCount = 2
$queueLayout.RowCount = 2
$queueLayout.Padding = New-Object System.Windows.Forms.Padding(4)
[void]$queueLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$queueLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$queueLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$queueLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tabQueue.Controls.Add($queueLayout)
$tabQueue.Controls.SetChildIndex($queueLayout, 0)

function _Ops-OpenJobJson($Grid) {
    $row = $Grid.CurrentRow
    if (-not $row) { return }
    $p = ''
    try { $p = [string]$row.DataBoundItem.path } catch { }
    if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath 'notepad.exe' -ArgumentList $p }
}

function _New-JobGrid {
    param([switch]$IncludePriority)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.RowHeadersVisible = $false
    $grid.AutoGenerateColumns = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $cols = @()
    if ($IncludePriority) {
        $cols += @{ Name = 'priority'; Header = 'priority'; Weight = 18 }
    }
    $cols += @(
        @{ Name = 'jobId'; Header = 'jobId'; Weight = 40 }
        @{ Name = 'type'; Header = 'type'; Weight = 28 }
        @{ Name = 'sourceName'; Header = 'sourceName'; Weight = 50 }
        @{ Name = 'checkpoint'; Header = 'checkpoint'; Weight = 28 }
        @{ Name = 'lastWriteTime'; Header = 'lastWriteTime'; Weight = 32 }
    )
    foreach ($c in $cols) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.Name
        $col.DataPropertyName = $c.Name
        $col.HeaderText = $c.Header
        $col.FillWeight = [single]$c.Weight
        [void]$grid.Columns.Add($col)
    }
    $grid.Add_CellFormatting({
        $e = $args[1]
        if (-not $e -or $e.ColumnIndex -lt 0 -or $e.RowIndex -lt 0) { return }
        $item = $this.Rows[$e.RowIndex].DataBoundItem
        $deferred = $false
        $locked = $false
        try { $deferred = [bool]$item.deferred } catch { }
        try { $locked = [bool]$item.locked } catch { }
        if ($deferred) {
            $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        }
        $col = $this.Columns[$e.ColumnIndex]
        if ($col.DataPropertyName -ne 'jobId') { return }
        if ($deferred) {
            $e.Value = ('[W] {0}' -f $item.jobId)
            $e.FormattingApplied = $true
        } elseif ($locked) {
            $e.Value = ('[L] {0}' -f $item.jobId)
            $e.FormattingApplied = $true
        }
    })
    $grid.Add_CellDoubleClick({ _Ops-OpenJobJson $this })
    return $grid
}

function _New-QueuePane([string]$Title, [switch]$IncludePriority) {
    $box = New-Object System.Windows.Forms.GroupBox
    $box.Text = $Title
    $box.Dock = 'Fill'
    $box.Padding = New-Object System.Windows.Forms.Padding(6)
    $grid = _New-JobGrid -IncludePriority:$IncludePriority
    $box.Controls.Add($grid)
    return @{ box = $box; grid = $grid }
}

$panePending = _New-QueuePane 'Pending' -IncludePriority
$paneRunning = _New-QueuePane 'Running'
$paneFailed = _New-QueuePane 'Failed'
$paneSucceeded = _New-QueuePane 'Succeeded'
$queueLayout.Controls.Add($panePending.box, 0, 0)
$queueLayout.Controls.Add($paneRunning.box, 1, 0)
$queueLayout.Controls.Add($paneFailed.box, 0, 1)
$queueLayout.Controls.Add($paneSucceeded.box, 1, 1)
$gridPending = $panePending.grid
$gridRunning = $paneRunning.grid
$gridFailed = $paneFailed.grid
$gridSucceeded = $paneSucceeded.grid
$script:QueueGrids = @($gridPending, $gridRunning, $gridFailed, $gridSucceeded)
$script:QueuePanes = @{
    pending = $panePending
    running = $paneRunning
    failed = $paneFailed
    succeeded = $paneSucceeded
}

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
$btnTickHint = New-Object System.Windows.Forms.Button
$btnTickHint.Text = 'Run audit tick now'
$btnTickHint.AutoSize = $true
[void]$runsBtns.Controls.AddRange(@($btnFullScan, $btnReconcileSs, $btnEnqReport, $btnTickHint))

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
$sqlTop.Height = 112
$tabSql.Controls.Add($sqlTop)

$cmbSql = New-Object System.Windows.Forms.ComboBox
$cmbSql.DropDownStyle = 'DropDownList'
@('v_poller_health', 'v_job_summary', 'poll_runs', 'processing_jobs', 'audit_events', 'notification_log') | ForEach-Object { [void]$cmbSql.Items.Add($_) }
$cmbSql.SelectedIndex = 0
$cmbSql.Location = New-Object System.Drawing.Point(8, 8)
$cmbSql.Width = 180
$sqlTop.Controls.Add($cmbSql)

$lblSqlRange = New-Object System.Windows.Forms.Label
$lblSqlRange.Text = 'Range'
$lblSqlRange.Location = New-Object System.Drawing.Point(196, 12)
$lblSqlRange.AutoSize = $true
$sqlTop.Controls.Add($lblSqlRange)

$cmbSqlRange = New-Object System.Windows.Forms.ComboBox
$cmbSqlRange.DropDownStyle = 'DropDownList'
Get-QCOpsSqlTimeRangeChoices | ForEach-Object { [void]$cmbSqlRange.Items.Add($_) }
$cmbSqlRange.SelectedItem = 'Last 24 hours'
$cmbSqlRange.Location = New-Object System.Drawing.Point(240, 8)
$cmbSqlRange.Width = 130
$sqlTop.Controls.Add($cmbSqlRange)

$btnSqlLoad = New-Object System.Windows.Forms.Button
$btnSqlLoad.Text = 'Load'
$btnSqlLoad.Location = New-Object System.Drawing.Point(378, 6)
$sqlTop.Controls.Add($btnSqlLoad)

$lblWatermark = New-Object System.Windows.Forms.Label
$lblWatermark.Location = New-Object System.Drawing.Point(460, 12)
$lblWatermark.Size = New-Object System.Drawing.Size(520, 20)
$sqlTop.Controls.Add($lblWatermark)

$lblSqlJob = New-Object System.Windows.Forms.Label
$lblSqlJob.Text = 'Job type'
$lblSqlJob.Location = New-Object System.Drawing.Point(8, 44)
$lblSqlJob.AutoSize = $true
$sqlTop.Controls.Add($lblSqlJob)

$cmbSqlJobType = New-Object System.Windows.Forms.ComboBox
$cmbSqlJobType.DropDownStyle = 'DropDownList'
Get-QCOpsSqlJobTypeChoices | ForEach-Object { [void]$cmbSqlJobType.Items.Add($_) }
$cmbSqlJobType.SelectedIndex = 0
$cmbSqlJobType.Location = New-Object System.Drawing.Point(70, 40)
$cmbSqlJobType.Width = 190
$sqlTop.Controls.Add($cmbSqlJobType)

$lblSqlPath = New-Object System.Windows.Forms.Label
$lblSqlPath.Text = 'Source folder'
$lblSqlPath.Location = New-Object System.Drawing.Point(270, 44)
$lblSqlPath.AutoSize = $true
$sqlTop.Controls.Add($lblSqlPath)

$cmbSqlSourceFolder = New-Object System.Windows.Forms.ComboBox
$cmbSqlSourceFolder.DropDownStyle = 'DropDown'
$cmbSqlSourceFolder.Location = New-Object System.Drawing.Point(360, 40)
$cmbSqlSourceFolder.Width = 420
try {
    Get-QCOpsWatchFolderChoices -Config $script:Cfg | ForEach-Object { [void]$cmbSqlSourceFolder.Items.Add($_) }
} catch { }
$sqlTop.Controls.Add($cmbSqlSourceFolder)

function _Ops-NewSqlDatePicker([int]$X, [int]$Y) {
    $dtp = New-Object System.Windows.Forms.DateTimePicker
    $dtp.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtp.CustomFormat = 'yyyy-MM-dd HH:mm'
    $dtp.ShowUpDown = $false
    $dtp.Location = New-Object System.Drawing.Point($X, $Y)
    $dtp.Width = 168
    $dtp.Value = [datetime]::Now
    return $dtp
}

$lblSqlFrom = New-Object System.Windows.Forms.Label
$lblSqlFrom.Text = 'From'
$lblSqlFrom.AutoSize = $true
$sqlTop.Controls.Add($lblSqlFrom)
$dtpSqlFrom = _Ops-NewSqlDatePicker 44 76
$dtpSqlFrom.Value = [datetime]::Now.AddDays(-1)
$sqlTop.Controls.Add($dtpSqlFrom)

$lblSqlTo = New-Object System.Windows.Forms.Label
$lblSqlTo.Text = 'To'
$lblSqlTo.AutoSize = $true
$sqlTop.Controls.Add($lblSqlTo)
$dtpSqlTo = _Ops-NewSqlDatePicker 246 76
$sqlTop.Controls.Add($dtpSqlTo)

$lblWm = New-Object System.Windows.Forms.Label
$lblWm.Text = 'Rewind to'
$lblWm.AutoSize = $true
$sqlTop.Controls.Add($lblWm)
$dtpWm = _Ops-NewSqlDatePicker 78 76
$dtpWm.Value = [datetime]::Now.AddHours(-1)
$sqlTop.Controls.Add($dtpWm)
$lblWmLocal = New-Object System.Windows.Forms.Label
$lblWmLocal.Text = '(local)'
$lblWmLocal.AutoSize = $true
$sqlTop.Controls.Add($lblWmLocal)

$txtWmConfirm = New-Object System.Windows.Forms.TextBox
$txtWmConfirm.Width = 180
$sqlTop.Controls.Add($txtWmConfirm)
$btnWmRewind = New-Object System.Windows.Forms.Button
$btnWmRewind.Text = 'Rewind watermark'
$btnWmRewind.Width = 140
$sqlTop.Controls.Add($btnWmRewind)
$lblWmHint = New-Object System.Windows.Forms.Label
$lblWmHint.Text = 'type REWIND WATERMARK'
$lblWmHint.AutoSize = $true
$sqlTop.Controls.Add($lblWmHint)

function _Ops-LayoutSqlFilters {
    $tbl = [string]$cmbSql.SelectedItem
    $custom = ([string]$cmbSqlRange.SelectedItem -eq 'Custom')
    $jobOk = ($tbl -in @('processing_jobs', 'v_job_summary'))
    $folderOk = ($tbl -in @('processing_jobs', 'v_job_summary', 'audit_events', 'notification_log'))
    $cmbSqlJobType.Enabled = $jobOk
    $cmbSqlSourceFolder.Enabled = $folderOk
    $lblSqlFrom.Visible = $custom
    $dtpSqlFrom.Visible = $custom
    $lblSqlTo.Visible = $custom
    $dtpSqlTo.Visible = $custom
    $wmX = 8
    if ($custom) {
        $lblSqlFrom.Location = New-Object System.Drawing.Point(8, 80)
        $dtpSqlFrom.Location = New-Object System.Drawing.Point(44, 76)
        $lblSqlTo.Location = New-Object System.Drawing.Point(220, 80)
        $dtpSqlTo.Location = New-Object System.Drawing.Point(246, 76)
        $wmX = 430
    }
    $lblWm.Location = New-Object System.Drawing.Point($wmX, 80)
    $dtpWm.Location = New-Object System.Drawing.Point(($wmX + 70), 76)
    $lblWmLocal.Location = New-Object System.Drawing.Point(($wmX + 242), 80)
    $txtWmConfirm.Location = New-Object System.Drawing.Point(($wmX + 292), 78)
    $btnWmRewind.Location = New-Object System.Drawing.Point(($wmX + 480), 76)
    $lblWmHint.Location = New-Object System.Drawing.Point(($wmX + 628), 80)
}

$cmbSql.Add_SelectedIndexChanged({ _Ops-LayoutSqlFilters })
$cmbSqlRange.Add_SelectedIndexChanged({ _Ops-LayoutSqlFilters })
_Ops-LayoutSqlFilters

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

function _Ops-BindGrid($Grid, $Rows) {
    $al = New-Object System.Collections.ArrayList
    foreach ($r in @($Rows)) { [void]$al.Add($r) }
    $Grid.DataSource = $null
    $Grid.DataSource = $al
}

function _Ops-ApplyHeader {
    param([switch]$Light)
    $st = _Ops-RefreshStatus -Light:$Light
    $d = $st.Data
    $script:SuppressToggle = $true
    $chkOn.Checked = [bool]$d.pipelineOn
    $script:SuppressToggle = $false
    $script:StateColor = _Ops-StateColor ([string]$d.stateLabel)
    $pnlBubble.Invalidate()
    $lblState.Text = ('{0}    task={1} ({2})    lock={3}' -f $d.stateLabel, $d.task.state, $(if ($d.task.enabled) { 'enabled' } else { 'disabled' }), $d.lock.text)
    $dryTxt = 'n/a'
    try {
        if ($d.dryRunText) { $dryTxt = [string]$d.dryRunText }
        else { $dryTxt = Get-QCOpsDryRunHeaderText -Policy $d.dryRun }
    } catch { }
    $wm = 'n/a'
    if ($null -ne $d.watermarkAgeSeconds) { $wm = ('{0}s ago' -f $d.watermarkAgeSeconds) }
    $lblLine2.Text = ('PW={0}   SQL={1}   DryRun={2}   watermark={3}' -f $d.pwText, $d.sqlEnabled, $dryTxt, $wm)

    if (-not $Light.IsPresent) {
        $counts = @{
            pending = $d.queue.pending
            running = $d.queue.running
            failed = $d.queue.failed
            succeeded = $d.queue.succeeded
            locks = $d.queue.locks
        }
        foreach ($k in @($script:QueueScoreLabels.Keys)) {
            $n = 0
            try { $n = [int]$counts[$k] } catch { $n = 0 }
            $script:QueueScoreLabels[$k].Text = [string]$n
        }
    }

    $hostRows = @()
    try { $hostRows = @($d.hostRows) } catch { $hostRows = @() }
    $hostLabels = @($script:HostStatusLabels)
    for ($i = 0; $i -lt $hostLabels.Count; $i++) {
        $lbl = $hostLabels[$i]
        $row = $null
        if ($i -lt $hostRows.Count) { $row = $hostRows[$i] }
        if ($row) {
            $lbl.Text = [string]$row.text
            $argb = Get-QCOpsHostHealthColorArgb -Health ([string]$row.health)
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb([int]$argb.R, [int]$argb.G, [int]$argb.B)
        } else {
            $lbl.Text = ''
        }
    }

    $tick = ''
    if ($d.lastTick -and $d.lastTick.ts) { $tick = [string]$d.lastTick.ts }
    $stall = ''
    if ($d.lastStall) { $stall = [string]$d.lastStall.code }
    $notifFail = ''
    if ($d.lastNotificationFail) { $notifFail = [string]$d.lastNotificationFail.code }
    $batchTxt = 'Status-set batch: n/a'
    try { if ($d.statusSetBatchText) { $batchTxt = [string]$d.statusSetBatchText } } catch { }
    $lblLine3.Text = ($batchTxt + [Environment]::NewLine + ('lastTick={0}   stall={1}   notifFail={2}' -f $tick, $stall, $notifFail))
}

function _Ops-EventTime($Obj) {
    $raw = $null
    try { if ($Obj.ts) { $raw = [string]$Obj.ts } } catch { }
    if (-not $raw) { return $null }
    try {
        return [datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal)
    } catch {
        return $null
    }
}

function _Ops-RenderOverviewLog {
    $window = [string]$cmbLogTime.SelectedItem
    $type = [string]$cmbLogType.SelectedItem
    $codeFilter = ([string]$txtLogCode.Text).Trim()
    $cutoff = $null
    $now = [datetime]::UtcNow
    switch ($window) {
        'Last 15 min' { $cutoff = $now.AddMinutes(-15) }
        'Last hour' { $cutoff = $now.AddHours(-1) }
        'Last 6 hours' { $cutoff = $now.AddHours(-6) }
        'Last 24 hours' { $cutoff = $now.AddHours(-24) }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($o in @($script:LogEvents)) {
        $src = ''
        $level = ''
        $code = ''
        try { $src = [string]$o.logSource } catch { }
        try { $level = [string]$o.level } catch { }
        try { $code = [string]$o.code } catch { }
        if ($cutoff) {
            $ts = _Ops-EventTime $o
            if ($ts -and $ts.ToUniversalTime() -lt $cutoff) { continue }
        }
        if ($type -eq 'Watcher' -and $src -ne 'watcher') { continue }
        if ($type -eq 'Processor' -and $src -ne 'processor') { continue }
        if ($type -eq 'Error' -and $level -notmatch '(?i)error') { continue }
        if ($type -eq 'Warning' -and $level -notmatch '(?i)warning') { continue }
        if ($codeFilter -and $code -notlike ('*{0}*' -f $codeFilter) -and ("$(_Ops-FormatEvent $o)" -notlike ('*{0}*' -f $codeFilter))) { continue }
        [void]$lines.Add((_Ops-FormatEvent $o))
    }
    $logBox.Text = ($lines -join [Environment]::NewLine)
}

function _Ops-RefreshOverviewLog {
    $root = [string]$script:Status.Data.queueRoot
    $window = [string]$cmbLogTime.SelectedItem
    $files = 1
    $lines = 80
    switch ($window) {
        'Last hour' { $files = 2; $lines = 80 }
        'Last 6 hours' { $files = 6; $lines = 50 }
        'Last 24 hours' { $files = 8; $lines = 40 }
        'All loaded' { $files = 3; $lines = 80 }
    }
    $logDir = Join-Path $root '_logs'
    $stamp = $window
    if (Test-Path -LiteralPath $logDir) {
        $latest = @(Get-ChildItem -LiteralPath $logDir -Filter '*.jsonl' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 2)
        foreach ($f in $latest) { $stamp += ('|{0}:{1}' -f $f.LastWriteTimeUtc.Ticks, $f.Length) }
    }
    if ($stamp -eq $script:LogStamp -and $script:LogEvents.Count -gt 0) {
        _Ops-RenderOverviewLog
        return
    }
    $ev = Get-QCOpsRecentLogEvents -QueueRoot $root -MaxLines $lines -MaxFiles $files -Unfiltered
    $script:LogEvents = @($ev.events)
    $script:LogStamp = $stamp
    _Ops-RenderOverviewLog
}

function _Ops-RefreshQueue {
    $root = [string]$script:Status.Data.queueRoot
    $stamp = ''
    foreach ($state in @('pending', 'running', 'failed', 'succeeded', 'locks')) {
        $dir = Join-Path $root $state
        if (Test-Path -LiteralPath $dir) {
            $stamp += (Get-Item -LiteralPath $dir).LastWriteTimeUtc.Ticks.ToString() + '|'
        }
    }
    try {
        $dirtyPath = Get-QCStatusSetDirtyFolderStorePath -Config $script:Cfg
        if ($dirtyPath -and (Test-Path -LiteralPath $dirtyPath)) {
            $stamp += (Get-Item -LiteralPath $dirtyPath).LastWriteTimeUtc.Ticks.ToString() + '|'
        }
    } catch { }
    $stamp += [datetime]::UtcNow.ToString('yyyyMMddHHmm')
    if ($stamp -eq $script:QueueStamp -and $stamp) { return }
    $script:QueueStamp = $stamp
    $limits = @{ pending = 150; running = 50; failed = 100; succeeded = 100 }
    $prefer = @(Get-QCOpsPreferJobTypes -Config $script:Cfg)
    foreach ($state in @('pending', 'running', 'failed', 'succeeded')) {
        $lim = [int]$limits[$state]
        $rowLimit = $lim
        if ($state -eq 'pending') { $rowLimit = 0 }
        $rows = @(Get-QCOpsRunningJobRows -QueueRoot $root -State $state -Limit $rowLimit -PreferJobTypes $prefer -Config $script:Cfg -IncludeDeferredStatusSet:($state -eq 'pending'))
        if ($state -eq 'pending') {
            $rows = @($rows | Sort-Object @{ Expression = { [bool]$_.deferred } }, @{ Expression = { [int]$_.priority } }, @{ Expression = { $_.lastWriteTimeUtc } })
        } else {
            $rows = @($rows | Sort-Object lastWriteTimeUtc -Descending)
        }
        if ($lim -gt 0 -and $rows.Count -gt $lim) { $rows = @($rows | Select-Object -First $lim) }
        $pane = $script:QueuePanes[$state]
        $shown = $rows.Count
        $extra = ''
        if ($shown -ge $lim) { $extra = '+' }
        $pane.box.Text = ('{0} ({1}{2})' -f (Get-Culture).TextInfo.ToTitleCase($state), $shown, $extra)
        _Ops-BindGrid $pane.grid $rows
    }
}

function _Ops-SelectedJobPaths {
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($g in @($script:QueueGrids)) {
        foreach ($row in @($g.SelectedRows)) {
            try {
                $p = [string]$row.DataBoundItem.path
                if ($p) { [void]$paths.Add($p) }
            } catch { }
        }
    }
    return @($paths)
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

$cmbLogTime.Add_SelectedIndexChanged({
    $script:LogStamp = ''
    _Ops-RefreshOverviewLog
})
$cmbLogType.Add_SelectedIndexChanged({ _Ops-RenderOverviewLog })
$btnLogApply.Add_Click({
    $script:LogStamp = ''
    _Ops-RefreshOverviewLog
})
$txtLogCode.Add_KeyDown({
    $e = $args[1]
    if ($e -and $e.KeyCode -eq 'Enter') { _Ops-RenderOverviewLog }
})

$btnQueueRefresh.Add_Click({ _Ops-RefreshQueue })
$btnRequeue.Add_Click({
    $paths = @(_Ops-SelectedJobPaths)
    if ($paths.Count -eq 0) { _Ops-Msg 'Select one or more jobs.'; return }
    $r = Invoke-QCOpsRequeueJobs -Config $script:Cfg -JobPaths $paths -WritebackOnly:$chkWritebackOnly.Checked
    _Ops-Msg $r.Message
    $script:QueueStamp = ''
    _Ops-RefreshQueue
    _Ops-ApplyHeader
})
$btnRecover.Add_Click({
    $r = Invoke-QCOpsRecoverStaleJobs -Config $script:Cfg
    _Ops-Msg $(if ($r.IsSuccess) { $r.Message } else { $r.Message })
    $script:QueueStamp = ''
    _Ops-RefreshQueue
})
$btnOpenJson.Add_Click({
    foreach ($p in @(_Ops-SelectedJobPaths)) {
        if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath 'notepad.exe' -ArgumentList $p }
    }
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

function _Ops-HideSqlSourceFolderColumn {
    foreach ($name in @('source_folder', 'sourceFolder')) {
        try {
            $col = $gridSql.Columns[$name]
            if ($col) { $col.Visible = $false }
        } catch { }
    }
}

$gridSql.Add_DataBindingComplete({ _Ops-HideSqlSourceFolderColumn })

$btnSqlLoad.Add_Click({
    $name = [string]$cmbSql.SelectedItem
    $range = Resolve-QCOpsSqlTimeRange -Range ([string]$cmbSqlRange.SelectedItem) -CustomFrom $dtpSqlFrom.Value -CustomTo $dtpSqlTo.Value
    if (-not $range.IsSuccess) { _Ops-Msg $range.Message; return }
    $jobType = [string]$cmbSqlJobType.SelectedItem
    if (-not $cmbSqlJobType.Enabled) { $jobType = '' }
    $sourceFolder = [string]$cmbSqlSourceFolder.Text
    if (-not $cmbSqlSourceFolder.Enabled) { $sourceFolder = '' }
    $r = Get-QCOpsSqlTablePreview -Config $script:Cfg -TableOrView $name -Top 80 -FromUtc $range.Data.fromUtc -ToUtc $range.Data.toUtc -JobType $jobType -SourceFolder $sourceFolder
    if (-not $r.IsSuccess) { _Ops-Msg $r.Message; return }
    $gridSql.DataSource = $r.Data.table
    _Ops-HideSqlSourceFolderColumn
    $wm = $script:Status.Data.watermarkUtc
    $lblWatermark.Text = ('Watermark: {0}' -f $wm)
})
$btnWmRewind.Add_Click({
    $dt = $dtpWm.Value
    if ($dt.Kind -ne [DateTimeKind]::Utc) { $dt = $dt.ToUniversalTime() }
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
$timer.Interval = 5000
$script:TickN = 0
$timer.Add_Tick({
    try {
        if ($form.WindowState -eq 'Minimized') { return }
        $script:TickN++
        $full = (($script:TickN % 3) -eq 0)
        if ($full) { _Ops-ApplyHeader } else { _Ops-ApplyHeader -Light }
        if ($full -and $tabs.SelectedTab -eq $tabOverview) { _Ops-RefreshOverviewLog }
        if ($full -and $tabs.SelectedTab -eq $tabQueue) { _Ops-RefreshQueue }
        if ((($script:TickN % 6) -eq 0) -and $tabs.SelectedTab -eq $tabRuns) { _Ops-RefreshRuns }
        if ((($script:TickN % 6) -eq 0) -and $tabs.SelectedTab -eq $tabNotif) { _Ops-RefreshNotif }
    } catch { }
})

$tabs.Add_SelectedIndexChanged({
    try {
        if ($tabs.SelectedTab -eq $tabOverview) { _Ops-RefreshOverviewLog }
        elseif ($tabs.SelectedTab -eq $tabQueue) { _Ops-RefreshQueue }
        elseif ($tabs.SelectedTab -eq $tabRuns) { _Ops-RefreshRuns }
        elseif ($tabs.SelectedTab -eq $tabNotif) { _Ops-RefreshNotif }
    } catch { }
})

$form.Add_Shown({
    _Ops-ApplyHeader
    _Ops-RefreshOverviewLog
    $timer.Start()
})
$form.Add_FormClosed({ $timer.Stop() })

[void]$form.ShowDialog()
