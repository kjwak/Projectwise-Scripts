# Runs in the logon session. UI callbacks consume completed results from an MTA runspace pool.
. (Join-Path $PSScriptRoot 'QCOpsConsole.UiControls.ps1')

$script:Presets = @(Get-QCOpsFolderPresets -Config $script:Cfg)
$script:Workers = @{}
$script:Prefs = @{}
$script:PrefPath = Join-Path $env:LOCALAPPDATA 'ProjectWiseQC\ops-console.json'
if (-not $SmokeTest -and (Test-Path -LiteralPath $script:PrefPath)) {
    try {
        $json = Get-Content -LiteralPath $script:PrefPath -Raw | ConvertFrom-Json
        foreach ($p in $json.PSObject.Properties) { $script:Prefs[$p.Name] = $p.Value }
    } catch { }
}
$script:ColumnPrefs = @{}
if ($script:Prefs.SqlColumns) {
    foreach ($p in $script:Prefs.SqlColumns.PSObject.Properties) { $script:ColumnPrefs[$p.Name] = @($p.Value) }
}
$script:SortPrefs = @{}
if ($script:Prefs.SqlSort) {
    foreach ($p in $script:Prefs.SqlSort.PSObject.Properties) { $script:SortPrefs[$p.Name] = [string]$p.Value }
}

$iss = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$iss.ImportPSModule([string[]]@(
    (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1'),
    (Join-Path $opsModules 'QC.OpsConsole.psm1'),
    (Join-Path $opsModules 'QC.OpsDiagnostics.psm1'),
    (Join-Path $opsModules 'QC.OpsSqlExplorer.psm1')
))
$script:Pool = [RunspaceFactory]::CreateRunspacePool($iss)
$script:Pool.SetMaxRunspaces(4) | Out-Null
$script:Pool.ApartmentState = 'MTA'
$script:Pool.ThreadOptions = 'ReuseThread'
$script:Pool.Open()

function Start-OpsWork([string]$Key, [scriptblock]$Work, $InputData, [scriptblock]$Done) {
    if ($script:Workers.ContainsKey($Key)) { return }
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $script:Pool
    [void]$ps.AddScript($Work.ToString()).AddArgument($InputData)
    $script:Workers[$Key] = @{ pipeline = $ps; handle = $ps.BeginInvoke(); done = $Done }
}
function Receive-OpsWork {
    foreach ($key in @($script:Workers.Keys)) {
        $item = $script:Workers[$key]
        if (-not $item.handle.IsCompleted) { continue }
        $script:Workers.Remove($key)
        $failure = ''
        $value = $null
        try {
            $output = $item.pipeline.EndInvoke($item.handle)
            if ($item.pipeline.HadErrors) { $failure = ($item.pipeline.Streams.Error | Select-Object -First 1).ToString() }
            if ($output.Count) { $value = $output[$output.Count - 1] }
        } catch { $failure = $_.Exception.Message }
        finally { $item.pipeline.Dispose() }
        try { & $item.done $value $failure }
        catch { Set-OpsNotice $statusText "Display error: $($_.Exception.Message)" -ErrorState }
    }
}
function Invoke-OpsControl([string]$Command, [hashtable]$Arguments, [string]$Message) {
    if ($script:Workers.ContainsKey('control') -or $script:Workers.ContainsKey('action')) { return }
    Set-OpsNotice $statusText $Message
    Start-OpsWork control {
        param($p)
        $ErrorActionPreference = 'Stop'
        $parameters = $p.arguments
        & $p.command @parameters
    } @{ command = $Command; arguments = $Arguments } {
        param($result, $errorText)
        if ($result.Code -eq 'OPS_STOP_IN_FLIGHT_PREPEND' -and (Confirm-Ops ($result.Message + ' Force stop anyway?'))) {
            Invoke-OpsControl Stop-QCOpsPipeline @{ Config = $script:Cfg; Force = $true } 'Forcing pipeline stop...'
            return
        }
        if (-not $errorText -and -not $result.IsSuccess) { $errorText = $result.Message }
        if (-not $errorText -and $result.Data -is [hashtable] -and $result.Data.errors.Count) {
            $errorText = $result.Message + ' ' + ($result.Data.errors -join ' ')
        }
        if ($errorText) { Set-OpsNotice $statusText $errorText -ErrorState }
        else { Set-OpsNotice $statusText $result.Message }
        Request-OpsSnapshot
    }
}

$form = New-Object Windows.Forms.Form
$form.Text = 'QC Pipeline Ops'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object Drawing.Size(1370, 860)
$form.MinimumSize = New-Object Drawing.Size(820, 570)
$form.AutoScaleMode = 'Dpi'
$form.Font = New-Object Drawing.Font('Segoe UI', 10)
$form.BackColor = [Drawing.Color]::White
if ($script:Prefs.Width -ge 820 -and $script:Prefs.Height -ge 570) {
    $screen = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $form.Width = [math]::Min([int]$script:Prefs.Width, $screen.Width)
    $form.Height = [math]::Min([int]$script:Prefs.Height, $screen.Height)
}
$main = New-Object Windows.Forms.TabControl
$main.Dock = 'Fill'
$main.Padding = '12,7'
$form.Controls.Add($main)
$statusBar = New-Object Windows.Forms.StatusStrip
$statusText = New-Object Windows.Forms.ToolStripStatusLabel
$statusText.Text = 'Starting console...'
$statusText.Spring = $true
$statusText.TextAlign = 'MiddleLeft'
[void]$statusBar.Items.Add($statusText)
$form.Controls.Add($statusBar)
$main.BringToFront()

function New-OpsTab([string]$Title) {
    $tab = New-Object Windows.Forms.TabPage($Title)
    $tab.BackColor = [Drawing.Color]::White
    $main.TabPages.Add($tab)
    return $tab
}

. (Join-Path $PSScriptRoot 'QCOpsConsole.DiagnosticsUi.ps1')
. (Join-Path $PSScriptRoot 'QCOpsConsole.ActionsUi.ps1')
. (Join-Path $PSScriptRoot 'QCOpsConsole.SqlUi.ps1')

function Request-OpsSnapshot {
    if ($SmokeTest -or $script:Workers.ContainsKey('snapshot')) { return }
    $includeQueue = $main.SelectedTab -eq $tabQueue
    $includeLogs = $main.SelectedTab -eq $tabLogs -and -not $logPaused.Checked
    Start-OpsWork snapshot {
        param($p)
        $ErrorActionPreference = 'Stop'
        Get-QCOpsSnapshot -Config $p.config -IncludeQueue:$p.queue -IncludeLogs:$p.logs
    } @{ config = $script:Cfg; queue = $includeQueue; logs = $includeLogs } {
        param($result, $errorText)
        if ($errorText -or -not $result) {
            Set-OpsNotice $overviewFresh "Stale data | $errorText" -ErrorState
            if ($main.SelectedTab -eq $tabLogs) { Set-OpsNotice $logNotice "Stale data | $errorText" -ErrorState }
            if ($main.SelectedTab -eq $tabQueue) { Set-OpsNotice $queueNotice "Stale data | $errorText" -ErrorState }
            return
        }
        Show-OpsSnapshot $result
    }
}

function Invoke-OpsUiSmokeTest {
    if ($main.TabCount -ne 7) { throw 'Expected seven operator tabs.' }
    $actionChoice.Text = 'Sync folder sheet index'
    if (-not $script:ActionFields.FolderPath.folder -or -not $script:ActionFields.OneLevelDeep) { throw 'Maintenance fields were not built.' }
    $actionChoice.Text = 'Audit event retention'
    if ($script:ActionFields.AuditEventsDays.Items.Count -lt 2) { throw 'Maintenance choices were not built.' }
    $main.SelectedTab = $tabLogs
    $first = [pscustomobject]@{ key = 'one'; severity = 'Warning'; message = 'First' }
    $second = [pscustomobject]@{ key = 'two'; severity = 'Error'; message = 'Second' }
    Set-OpsRows $logGrid @($first) key -ColorSeverity
    $logGrid.CurrentCell = $logGrid.Rows[0].Cells[0]
    Set-OpsRows $logGrid @($second,$first) key -ColorSeverity
    if ($logGrid.Rows[0].Tag.key -ne 'two' -or (Get-OpsSelected $logGrid).key -ne 'one') { throw 'Log insert lost order or selection.' }
    $main.SelectedTab = $tabSql
    $table = New-Object Data.DataTable
    [void]$table.Columns.Add('id', [int])
    [void]$table.Rows.Add(7)
    Set-OpsColumns $sqlGrid (Get-OpsFriendlyColumns @('id'))
    Set-OpsRows $sqlGrid @($table.Rows)
    if ($sqlGrid.Rows.Count -ne 1 -or $sqlGrid.Rows[0].Cells[0].Value -ne '7') { throw 'SQL grid did not render the fixture.' }
    $script:SmokeAsync = $null
    Start-OpsWork smoke {
        param($unused)
        $ErrorActionPreference = 'Stop'
        [pscustomobject]@{
            apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
            exports = @('Get-QCOpsSnapshot','Invoke-QCOpsMaintenanceScript','Invoke-QCOpsExplorerPage' | ForEach-Object { Get-Command $_ }).Count
        }
    } $null {
        param($result,$errorText)
        $script:SmokeAsync = @{ result = $result; error = $errorText }
    }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while (-not $script:SmokeAsync -and $clock.Elapsed.TotalSeconds -lt 20) {
        [Windows.Forms.Application]::DoEvents()
        Receive-OpsWork
        Start-Sleep -Milliseconds 30
    }
    if (-not $script:SmokeAsync -or $script:SmokeAsync.error -or $script:SmokeAsync.result.exports -ne 3 -or
        $script:SmokeAsync.result.apartment -ne 'MTA') {
        throw "Background runspace failed: $($script:SmokeAsync.error)"
    }
}

$main.Add_SelectedIndexChanged({
    Request-OpsSnapshot
    if ($this.SelectedTab -eq $tabHistory) { Request-OpsHistory }
    if ($this.SelectedTab -eq $tabSql -and -not $script:SqlDatasets.Count) { Request-OpsCatalog }
})
$pump = New-Object Windows.Forms.Timer
$pump.Interval = 200
$pump.Add_Tick({
    Receive-OpsWork
    Receive-OpsActionOutput
})
$refresh = New-Object Windows.Forms.Timer
$refresh.Interval = 10000
$refresh.Add_Tick({
    if ($form.WindowState -ne 'Minimized') { Request-OpsSnapshot }
})
$form.Add_Shown({
    if ($SmokeTest) {
        try { Invoke-OpsUiSmokeTest }
        catch { $script:SmokeFailure = $_ }
        $form.Close()
        return
    }
    $pump.Start()
    $refresh.Start()
    Request-OpsSnapshot
    Set-OpsNotice $statusText "Profile: $AppSettingsPath"
})
$form.Add_FormClosing({
    if ($script:Workers.ContainsKey('action') -or $script:Workers.ContainsKey('control')) {
        $args[1].Cancel = $true
        Set-OpsNotice $statusText 'An operation is still running. The window can close when it finishes.' -ErrorState
    }
})
try { [void]$form.ShowDialog() }
finally {
    $refresh.Stop()
    $pump.Stop()
    if (-not $SmokeTest) { Save-OpsPreferences }
    foreach ($entry in @($script:Workers.Values)) {
        try { $entry.pipeline.Stop() } catch { }
        $entry.pipeline.Dispose()
    }
    $script:Pool.Dispose()
    $refresh.Dispose()
    $pump.Dispose()
    $form.Dispose()
}
if ($script:SmokeFailure) { throw $script:SmokeFailure }
if ($SmokeTest) { Write-Output 'QC Ops WinForms smoke test passed.' }
