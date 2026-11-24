$tabActions = New-OpsTab 'Actions'
$actionBar = New-OpsToolbar $tabActions
$script:ActionCatalog = @(Get-QCOpsMaintenanceCatalog)
$actionChoice = New-OpsCombo $actionBar @($script:ActionCatalog.label) 310 -Searchable
$actionPreview = New-OpsButton $actionBar 'Run diagnostic'
$actionApply = New-OpsButton $actionBar 'Apply'
$actionReport = New-OpsButton $actionBar 'Enqueue QC reporting'
$actionWatermark = New-OpsButton $actionBar 'Rewind audit watermark...'
$actionNotice = New-OpsLabel $actionBar 'Choose an action.'
$actionBody = New-Object Windows.Forms.FlowLayoutPanel
$actionBody.Dock = 'Top'
$actionBody.AutoSize = $true
$actionBody.FlowDirection = 'TopDown'
$actionBody.WrapContents = $false
$actionBody.Padding = '10,4,10,10'
$tabActions.Controls.Add($actionBody)
$actionBar.BringToFront()
$actionOutput = New-OpsDetails $tabActions
$script:ActionFields = @{}
$script:ActionRun = $null

function Get-OpsChosenAction {
    return $script:ActionCatalog | Where-Object label -eq $actionChoice.Text | Select-Object -First 1
}
function Update-OpsActionButtons {
    $action = Get-OpsChosenAction
    $enabled = $null -ne $action -and -not $script:ActionRun
    $actionPreview.Enabled = $enabled
    $actionApply.Visible = $action -and $action.ContainsKey('apply')
    $actionApply.Enabled = $enabled
    $actionPreview.Text = if ($actionApply.Visible) { 'Preview' } else { 'Run diagnostic' }
    $actionReport.Enabled = -not $script:ActionRun
    $actionChoice.Enabled = -not $script:ActionRun
    $actionBody.Enabled = -not $script:ActionRun
}
function Set-OpsActionFields {
    $actionBody.SuspendLayout()
    try {
        $script:ActionFields = @{}
        foreach ($control in @($actionBody.Controls)) { $control.Dispose() }
        $action = Get-OpsChosenAction
        if (-not $action) { return }
        [void](New-OpsLabel $actionBody $action.help)
        foreach ($field in $action.fields) {
            $row = New-Object Windows.Forms.FlowLayoutPanel
            $row.AutoSize = $true
            $row.WrapContents = $false
            $actionBody.Controls.Add($row)
            if ($field.kind -ne 'switch' -and $field.kind -ne 'folder') { [void](New-OpsLabel $row $field.label) }
            $control = switch ($field.kind) {
                'folder' { New-OpsFolderPicker $row }
                'switch' { New-OpsCheck $row $field.label }
                'choice' { New-OpsCombo $row $field.choices 245 }
                default {
                    $box = New-OpsText $row 235
                    $box.Text = [string]$field.initial
                    if ($field.kind -eq 'date') { $box.Text = (Get-QCWallClockNow).ToString('yyyy-MM-dd') }
                    $box
                }
            }
            $script:ActionFields[$field.name] = $control
        }
    } finally {
        $actionBody.ResumeLayout()
        Update-OpsActionButtons
    }
}
function Start-OpsAction([string]$Title, [string]$ScriptPath, [hashtable]$Parameters, [int]$TimeoutMs, [switch]$ManualWatcher) {
    if ($script:Workers.ContainsKey('control') -or $script:ActionRun) { return }
    $main.SelectedTab = $tabActions
    $script:ActionRun = @{
        title = $Title; clock = [Diagnostics.Stopwatch]::StartNew()
        output = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    }
    $actionOutput.Text = "$Title$([Environment]::NewLine)$([IO.Path]::GetFileName($ScriptPath))$([Environment]::NewLine)" +
        (($Parameters.Keys | Sort-Object | ForEach-Object { "  $_ = $($Parameters[$_])" }) -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine
    Set-OpsNotice $actionNotice "$Title | starting..."
    Update-OpsActionButtons
    Start-OpsWork action {
        param($p)
        $ErrorActionPreference = 'Stop'
        if ($p.manual) {
            if (-not (Test-QCOpsHostAllowed)) { throw "Manual watcher passes must run on $(Get-QCOpsExpectedHost)." }
            $active = (Get-QCOpsPipelineStatus -Config $p.config -Light).Data
            if (-not $active.task.registered -or $active.task.enabled -or $active.task.state -eq 'Running' -or $active.lock.alive) {
                throw 'The scheduled pipeline can run. Stop and disable it before a manual watcher pass.'
            }
        }
        Invoke-QCOpsMaintenanceScript -ScriptPath $p.path -Parameters $p.parameters -TimeoutMs $p.timeout -OutputQueue $p.output
    } @{
        path = $ScriptPath; parameters = $Parameters; timeout = $TimeoutMs; output = $script:ActionRun.output
        config = $script:Cfg; manual = $ManualWatcher.IsPresent
    } {
        param($result, $errorText)
        Receive-OpsActionOutput
        $elapsed = $script:ActionRun.clock.Elapsed.ToString('hh\:mm\:ss')
        if (-not $errorText -and -not $result.IsSuccess) { $errorText = $result.Message }
        $message = if ($errorText) { "$elapsed | $errorText" } else { "$elapsed | $($result.Message)" }
        Set-OpsNotice $actionNotice $message -ErrorState:([bool]$errorText)
        $actionOutput.AppendText([Environment]::NewLine + $message + [Environment]::NewLine)
        $script:ActionRun = $null
        Update-OpsActionButtons
    }
}
function Receive-OpsActionOutput {
    if (-not $script:ActionRun) { return }
    $line = ''
    $batch = [Text.StringBuilder]::new()
    $queue = $script:ActionRun.output
    for ($i = 0; $i -lt 2000 -and $queue.TryDequeue([ref]$line); $i++) { [void]$batch.AppendLine($line) }
    if ($batch.Length) {
        $follow = $actionOutput.SelectionStart -ge ($actionOutput.TextLength - 1)
        $selection = $actionOutput.SelectionStart
        $length = $actionOutput.SelectionLength
        $actionOutput.AppendText($batch.ToString())
        if ($follow) { $actionOutput.ScrollToCaret() }
        else { $actionOutput.Select($selection, $length) }
        if ($actionOutput.TextLength -gt 700000) {
            $actionOutput.Select(0, $actionOutput.TextLength - 550000)
            $actionOutput.SelectedText = "[Earlier output omitted]$([Environment]::NewLine)"
        }
    }
    Set-OpsNotice $actionNotice "$($script:ActionRun.title) | running $($script:ActionRun.clock.Elapsed.ToString('hh\:mm\:ss'))"
}
function Invoke-OpsChosenAction([bool]$Apply) {
    $action = Get-OpsChosenAction
    if (-not $action) { return }
    try {
        $values = @{}
        foreach ($field in $action.fields) {
            $control = $script:ActionFields[$field.name]
            $values[$field.name] = switch ($field.kind) {
                'switch' { $control.Checked }
                'folder' { Get-OpsFolder $control -ExplicitOnly }
                default { $control.Text }
            }
        }
        $parameters = Get-QCOpsMaintenanceParameters -Action $action -Values $values -Apply:$Apply -AppSettingsPath $AppSettingsPath
        if ($Apply) {
            $context = ($parameters.Keys | Where-Object { $_ -ne 'AppSettingsPath' } | Sort-Object |
                ForEach-Object { "$_ = $($parameters[$_])" }) -join [Environment]::NewLine
            if (-not $context) { $context = 'All configured workspaces.' }
            if (-not (Confirm-Ops "Apply $($action.label)?$([Environment]::NewLine)$context")) { return }
        }
        $title = if ($Apply) { "Apply: $($action.label)" } else { $action.label }
        Start-OpsAction $title $action.script $parameters $action.timeoutMs
    } catch { Set-OpsNotice $actionNotice $_.Exception.Message -ErrorState }
}
function Invoke-OpsManualPass {
    if ($script:ActionRun -or $script:Workers.ContainsKey('control')) { return }
    if (Confirm-Ops 'Run one watcher pass using this profile? It can enqueue jobs and run due reconciliation.') {
        Start-OpsAction 'One watcher pass' (Join-Path $PSScriptRoot 'Watch-QCTrigger.ps1') @{
            AppSettingsPath = $AppSettingsPath; Once = $true
        } 1800000 -ManualWatcher
    }
}
$actionChoice.Add_TextChanged({ Set-OpsActionFields })
$actionPreview.Add_Click({ Invoke-OpsChosenAction $false })
$actionApply.Add_Click({ Invoke-OpsChosenAction $true })
$actionReport.Add_Click({
    if (Confirm-Ops 'Enqueue the configured QC reporting scan?') {
        Invoke-OpsControl Invoke-QCOpsEnqueueJobType @{ Config = $script:Cfg; JobType = 'QC_REPORTING_SCAN' } 'Enqueueing QC reporting...'
    }
})
$actionWatermark.Add_Click({
    if ($script:ActionRun -or $script:Workers.ContainsKey('control')) { return }
    if (-not $script:LatestStatus -or $script:LatestStatus.task.enabled -or $script:LatestStatus.lock.alive) {
        Set-OpsNotice $actionNotice 'Stop and disable the pipeline before rewinding its audit cursor.' -ErrorState
        return
    }
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = 'Rewind audit watermark'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object Drawing.Size(545, 185)
    $layout = New-OpsToolbar $dialog
    $layout.FlowDirection = 'LeftToRight'
    [void](New-OpsLabel $layout 'Target time (Arizona)')
    $when = New-OpsSqlDate $layout ((Get-QCWallClockNow).AddHours(-1))
    $layout.SetFlowBreak($when, $true)
    [void](New-OpsLabel $layout 'Type REWIND WATERMARK')
    $confirmText = New-OpsText $layout 205
    $layout.SetFlowBreak($confirmText, $true)
    $apply = New-OpsButton $layout 'Rewind'
    $apply.DialogResult = 'OK'
    $dialog.AcceptButton = $apply
    try {
        if ($dialog.ShowDialog($form) -ne 'OK') { return }
        $value = ConvertFrom-QCDisplayWallClock -WallClock $when.Value
        Invoke-OpsControl Set-QCOpsAuditWatermark @{ Config = $script:Cfg; WatermarkUtc = $value; ConfirmText = $confirmText.Text; WhileStopped = $true } 'Rewinding watermark...'
    } finally { $dialog.Dispose() }
})
Set-OpsChoices $actionChoice @($script:ActionCatalog.label) $script:Prefs.Action
Set-OpsActionFields
