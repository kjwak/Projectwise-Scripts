# QC.WatcherAlerts.psm1
# Responsibility: Operational email alerts when the QC watcher loses ProjectWise connectivity.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationGraph.psm1') -Force -ErrorAction SilentlyContinue

function _QCWA-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[[string]$k] = $Value[$k] }
        return $h
    }
    if ($Value -is [pscustomobject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCWA-GetRepoRoot {
    return (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
}

function Get-QCWatcherSessionAlertSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $defaults = @{
        enabled = $true
        recipients = @('jflint@aztec.us')
        importance = 'high'
        dedupeMinutes = 60
        probeIntervalTicks = 60
        staleAuditLagSeconds = 180
        enableAuditStallDetection = $false
        probeFolderPath = ''
    }

    $watcher = _QCWA-ToHashtable $Config.watcher
    $alerts = $null
    if ($watcher -and $watcher.ContainsKey('sessionAlerts')) {
        $alerts = _QCWA-ToHashtable $watcher.sessionAlerts
    }

    $settings = @{}
    foreach ($key in @($defaults.Keys)) {
        $settings[$key] = $defaults[$key]
    }
    if ($alerts) {
        if ($null -ne $alerts.enabled) { try { $settings.enabled = [bool]$alerts.enabled } catch { } }
        if ($alerts.recipients) {
            $settings.recipients = @($alerts.recipients | ForEach-Object { [string]$_ } | Where-Object { $_ -match '@' })
        }
        if ($alerts.importance) { $settings.importance = [string]$alerts.importance }
        if ($null -ne $alerts.dedupeMinutes) { try { $settings.dedupeMinutes = [int]$alerts.dedupeMinutes } catch { } }
        if ($null -ne $alerts.probeIntervalTicks) { try { $settings.probeIntervalTicks = [int]$alerts.probeIntervalTicks } catch { } }
        if ($null -ne $alerts.staleAuditLagSeconds) { try { $settings.staleAuditLagSeconds = [int]$alerts.staleAuditLagSeconds } catch { } }
        if ($null -ne $alerts.enableAuditStallDetection) { try { $settings.enableAuditStallDetection = [bool]$alerts.enableAuditStallDetection } catch { } }
        if ($alerts.probeFolderPath) { $settings.probeFolderPath = [string]$alerts.probeFolderPath }
    }

    if ($settings.recipients.Count -eq 0 -and $Config.notifications -and $Config.notifications.adminRecipients) {
        $settings.recipients = @($Config.notifications.adminRecipients | ForEach-Object { [string]$_ } | Where-Object { $_ -match '@' })
    }

    if ($settings.dedupeMinutes -lt 1) { $settings.dedupeMinutes = 1 }
    if ($settings.probeIntervalTicks -lt 1) { $settings.probeIntervalTicks = 1 }
    if ($settings.staleAuditLagSeconds -lt 30) { $settings.staleAuditLagSeconds = 30 }

    return $settings
}

function Get-QCWatcherSessionReconnectSettings {
    <#
    .SYNOPSIS
    Resolves watcher.sessionReconnect settings for proactive ProjectWise re-login.
    .DESCRIPTION
    On Bentley-hosted datasources, long-lived watcher sessions can still be severed even when
    CredentialExpirationPolicy is NoExpiration. Periodic disconnect/reconnect under the server
    login-token window (~10h default) prefers silent re-auth over interactive session-expired dialogs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $defaults = @{
        enabled = $false
        intervalMinutes = 360
        minIntervalMinutes = 15
    }

    $watcher = _QCWA-ToHashtable $Config.watcher
    $reconnect = $null
    if ($watcher -and $watcher.ContainsKey('sessionReconnect')) {
        $reconnect = _QCWA-ToHashtable $watcher.sessionReconnect
    }

    $settings = @{}
    foreach ($key in @($defaults.Keys)) {
        $settings[$key] = $defaults[$key]
    }
    if ($reconnect) {
        if ($null -ne $reconnect.enabled) { try { $settings.enabled = [bool]$reconnect.enabled } catch { } }
        if ($null -ne $reconnect.intervalMinutes) { try { $settings.intervalMinutes = [int]$reconnect.intervalMinutes } catch { } }
        if ($null -ne $reconnect.minIntervalMinutes) { try { $settings.minIntervalMinutes = [int]$reconnect.minIntervalMinutes } catch { } }
    }

    if ($settings.minIntervalMinutes -lt 1) { $settings.minIntervalMinutes = 1 }
    if ($settings.intervalMinutes -lt [int]$settings.minIntervalMinutes) {
        $settings.intervalMinutes = [int]$settings.minIntervalMinutes
    }

    return $settings
}

function Test-QCWatcherSessionReconnectDue {
    <#
    .SYNOPSIS
    Returns whether a proactive ProjectWise reconnect is due based on wall-clock age.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [AllowNull()]
        [object]$LastConnectUtc = $null,
        [datetime]$NowUtc = ([datetime]::UtcNow)
    )

    $settings = Get-QCWatcherSessionReconnectSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return @{ due = $false; reason = 'disabled'; settings = $settings; ageMinutes = $null }
    }
    if ($null -eq $LastConnectUtc -or [string]::IsNullOrWhiteSpace([string]$LastConnectUtc)) {
        return @{ due = $false; reason = 'no_prior_connect'; settings = $settings; ageMinutes = $null }
    }

    try {
        $last = [datetime]$LastConnectUtc
    } catch {
        return @{ due = $false; reason = 'invalid_last_connect'; settings = $settings; ageMinutes = $null }
    }
    if ($last.Kind -eq [System.DateTimeKind]::Unspecified) {
        $last = [datetime]::SpecifyKind($last, [System.DateTimeKind]::Utc)
    } elseif ($last.Kind -eq [System.DateTimeKind]::Local) {
        $last = $last.ToUniversalTime()
    }

    $now = $NowUtc
    if ($now.Kind -eq [System.DateTimeKind]::Unspecified) {
        $now = [datetime]::SpecifyKind($now, [System.DateTimeKind]::Utc)
    } elseif ($now.Kind -eq [System.DateTimeKind]::Local) {
        $now = $now.ToUniversalTime()
    }

    $age = $now - $last
    if ($age.TotalMinutes -lt 0) {
        return @{ due = $false; reason = 'clock_skew'; settings = $settings; ageMinutes = [math]::Round($age.TotalMinutes, 2) }
    }

    $interval = [int]$settings.intervalMinutes
    if ($age.TotalMinutes -ge $interval) {
        return @{
            due = $true
            reason = 'interval_elapsed'
            settings = $settings
            ageMinutes = [math]::Round($age.TotalMinutes, 2)
            intervalMinutes = $interval
        }
    }

    return @{
        due = $false
        reason = 'within_interval'
        settings = $settings
        ageMinutes = [math]::Round($age.TotalMinutes, 2)
        intervalMinutes = $interval
    }
}

function Get-QCWatcherSessionAlertStatePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $root = _QCWA-GetRepoRoot
    if ($Config.notifications -and $Config.notifications.outputRoot) {
        $out = [string]$Config.notifications.outputRoot
        if (-not [System.IO.Path]::IsPathRooted($out)) {
            $out = Join-Path $root $out
        }
        return (Join-Path $out 'session-alerts\last-sent.json')
    }
    return (Join-Path $root 'notifications\session-alerts\last-sent.json')
}

function Test-QCWatcherSessionAlertDue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Reason = 'pw_session_lost'
    )

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return @{ due = $false; reason = 'disabled' }
    }

    $path = Get-QCWatcherSessionAlertStatePath -Config $Config
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ due = $true; reason = $Reason; statePath = $path }
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $lastUtc = [string]$obj.lastSentUtc
        if ([string]::IsNullOrWhiteSpace($lastUtc)) {
            return @{ due = $true; reason = $Reason; statePath = $path }
        }
        $last = [DateTime]::Parse($lastUtc, $null, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
        $ageMin = ((Get-Date).ToUniversalTime() - $last).TotalMinutes
        if ($ageMin -ge [double]$settings.dedupeMinutes) {
            return @{ due = $true; reason = $Reason; statePath = $path; minutesSinceLast = [math]::Round($ageMin, 1) }
        }
        return @{ due = $false; reason = 'deduped'; statePath = $path; minutesSinceLast = [math]::Round($ageMin, 1) }
    } catch {
        return @{ due = $true; reason = $Reason; statePath = $path; readError = [string]$_.Exception.Message }
    }
}

function Set-QCWatcherSessionAlertSent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Reason = 'pw_session_lost',
        [hashtable]$Details = @{}
    )

    $path = Get-QCWatcherSessionAlertStatePath -Config $Config
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $payload = @{
        lastSentUtc = (Get-Date).ToUniversalTime().ToString('o')
        reason = $Reason
        details = $Details
    }
    Set-Content -LiteralPath $path -Value ($payload | ConvertTo-Json -Depth 8) -Encoding UTF8
}

function New-QCWatcherSessionLostEmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Details
    )

    $rows = @(
        @{ label = 'Detected (UTC)'; value = [string]$Details.detectedUtc }
        @{ label = 'Reason'; value = [string]$Details.reason }
        @{ label = 'Datasource'; value = [string]$Details.datasourceName }
        @{ label = 'Watcher tick'; value = [string]$Details.tick }
        @{ label = 'Probe folder'; value = [string]$Details.probeFolderPath }
        @{ label = 'Last PW audit activity'; value = [string]$Details.maxPwActTime }
        @{ label = 'Audit watermark'; value = [string]$Details.watermarkAfter }
        @{ label = 'Error'; value = [string]$Details.errorMessage }
    )

    $htmlRows = ($rows | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_.value)) { return }
        $label = [System.Net.WebUtility]::HtmlEncode([string]$_.label)
        $value = [System.Net.WebUtility]::HtmlEncode([string]$_.value)
        "<tr><td style=""padding:6px 12px;font-weight:600;vertical-align:top;"">$label</td><td style=""padding:6px 12px;"">$value</td></tr>"
    }) -join ''

    return @"
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1a1a1a;">
  <p style="margin:0 0 12px 0;"><strong>The QC watcher lost its ProjectWise session.</strong></p>
  <p style="margin:0 0 16px 0;">QC triggers may be missed until the watcher reconnects. Restart the watcher or re-authenticate the service account if the session was logged out.</p>
  <table style="border-collapse:collapse;border:1px solid #d0d0d0;">$htmlRows</table>
</div>
"@
}

function _QCWA-SendWatcherOperationalEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$HtmlBody,
        [Parameter(Mandatory)][string]$EventType,
        [Parameter(Mandatory)][string]$Importance,
        [Parameter(Mandatory)][string[]]$Recipients
    )

    $notif = _QCWA-ToHashtable $Config.notifications
    if (-not $notif) { $notif = @{} }
    $provider = if ($notif.provider) { [string]$notif.provider } else { 'Mock' }
    $dryRun = $false
    if ($null -ne $notif.dryRun) { try { $dryRun = [bool]$notif.dryRun } catch { } }
    if ($Config.ContainsKey('dryRun') -and [bool]$Config.dryRun) { $dryRun = $true }

    $payload = @{
        eventType = $EventType
        documentName = 'QC Watcher'
        subject = $Subject
        htmlBody = $HtmlBody
        importance = $Importance
        to = @($Recipients)
        cc = @()
        logoPath = if ($notif.email -and $notif.email.logoPath) { [string]$notif.email.logoPath } else { 'email/typsalogo.png.webp' }
    }

    if ($provider -eq 'MicrosoftGraph' -and (Get-Command -Name 'Send-QCNotificationGraph' -ErrorAction SilentlyContinue)) {
        $graph = _QCWA-ToHashtable $notif.graph
        if (-not $graph) { $graph = @{} }
        return Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:$dryRun
    }

    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationMock.psm1') -Force -ErrorAction SilentlyContinue | Out-Null
    if (Get-Command -Name 'Send-QCNotificationMock' -ErrorAction SilentlyContinue) {
        $outputRoot = if ($notif.outputRoot) { [string]$notif.outputRoot } else { 'notifications' }
        if (-not [System.IO.Path]::IsPathRooted($outputRoot)) {
            $outputRoot = Join-Path (_QCWA-GetRepoRoot) $outputRoot
        }
        $mockPayload = $payload.Clone()
        $mockPayload['body'] = $HtmlBody
        return Send-QCNotificationMock -Payload $mockPayload -OutputRoot $outputRoot -DryRun:$dryRun
    }

    return New-QCFailureResult -Code 'QC_WATCHER_ALERT_PROVIDER_UNAVAILABLE' -Message 'No notification provider available for watcher operational alert.' -Data @{ provider = $provider }
}

function Send-QCWatcherSessionLostAlert {
    <#
    .SYNOPSIS
    Sends a high-importance operational email when the watcher detects a lost ProjectWise session.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$Details,
        [switch]$Force
    )

    $alertReason = ([string]$Details.reason).Trim().ToLowerInvariant()
    if ($alertReason -eq 'watcher_child_stalled') {
        return New-QCSuccessResult -Code 'QC_WATCHER_ALERT_SKIPPED_WRONG_CHANNEL' -Message 'Stall recovery alerts must use Send-QCWatcherStallRecoveryAlert.' -Data @{ skipped = $true; reason = $alertReason }
    }

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'QC_WATCHER_ALERT_SKIPPED_DISABLED' -Message 'Watcher session alerts are disabled.' -Data @{ skipped = $true }
    }

    $recipients = @($settings.recipients)
    if ($recipients.Count -eq 0) {
        return New-QCFailureResult -Code 'QC_WATCHER_ALERT_NO_RECIPIENTS' -Message 'No session alert recipients configured.' -Data @{ skipped = $true }
    }

    if (-not $Force) {
        $due = Test-QCWatcherSessionAlertDue -Config $Config -Reason ([string]$Details.reason)
        if (-not $due.due) {
            return New-QCSuccessResult -Code 'QC_WATCHER_ALERT_SKIPPED_DEDUPED' -Message 'Session alert suppressed by dedupe window.' -Data $due
        }
    }

    $importance = [string]$settings.importance
    if ([string]::IsNullOrWhiteSpace($importance)) { $importance = 'high' }

    $subject = '[QC Watcher] URGENT: ProjectWise session lost'
    $htmlBody = New-QCWatcherSessionLostEmailBody -Details $Details
    $sendResult = _QCWA-SendWatcherOperationalEmail -Config $Config -Subject $subject -HtmlBody $htmlBody `
        -EventType 'QC_WATCHER_PW_SESSION_LOST' -Importance $importance -Recipients $recipients

    if ($sendResult.IsSuccess) {
        Set-QCWatcherSessionAlertSent -Config $Config -Reason ([string]$Details.reason) -Details $Details
        return New-QCSuccessResult -Code 'QC_WATCHER_ALERT_SENT' -Message 'Watcher session lost alert sent.' -Data @{
            provider = if ($Config.notifications -and $Config.notifications.provider) { [string]$Config.notifications.provider } else { 'Mock' }
            dryRun = [bool]$Config.dryRun
            to = $recipients
            importance = $importance
            send = $sendResult.Data
        }
    }

    return New-QCFailureResult -Code 'QC_WATCHER_ALERT_SEND_FAILED' -Message ([string]$sendResult.Message) -Data @{
        sendCode = [string]$sendResult.Code
    }
}

function Get-QCWatcherStallAlertStatePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $root = _QCWA-GetRepoRoot
    if ($Config.notifications -and $Config.notifications.outputRoot) {
        $out = [string]$Config.notifications.outputRoot
        if (-not [System.IO.Path]::IsPathRooted($out)) {
            $out = Join-Path $root $out
        }
        return (Join-Path $out 'stall-alerts\last-sent.json')
    }
    return (Join-Path $root 'notifications\stall-alerts\last-sent.json')
}

function Test-QCWatcherStallAlertDue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Reason = 'watcher_child_stalled'
    )

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return @{ due = $false; reason = 'disabled' }
    }

    $path = Get-QCWatcherStallAlertStatePath -Config $Config
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ due = $true; reason = $Reason; statePath = $path }
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $lastUtc = [string]$obj.lastSentUtc
        if ([string]::IsNullOrWhiteSpace($lastUtc)) {
            return @{ due = $true; reason = $Reason; statePath = $path }
        }
        $last = [DateTime]::Parse($lastUtc, $null, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
        $ageMin = ((Get-Date).ToUniversalTime() - $last).TotalMinutes
        if ($ageMin -ge [double]$settings.dedupeMinutes) {
            return @{ due = $true; reason = $Reason; statePath = $path; minutesSinceLast = [math]::Round($ageMin, 1) }
        }
        return @{ due = $false; reason = 'deduped'; statePath = $path; minutesSinceLast = [math]::Round($ageMin, 1) }
    } catch {
        return @{ due = $true; reason = $Reason; statePath = $path; readError = [string]$_.Exception.Message }
    }
}

function Set-QCWatcherStallAlertSent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Reason = 'watcher_child_stalled',
        [hashtable]$Details = @{}
    )

    $path = Get-QCWatcherStallAlertStatePath -Config $Config
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $payload = @{
        lastSentUtc = (Get-Date).ToUniversalTime().ToString('o')
        reason = $Reason
        details = $Details
    }
    Set-Content -LiteralPath $path -Value ($payload | ConvertTo-Json -Depth 8) -Encoding UTF8
}

function New-QCWatcherStallRecoveryEmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Details
    )

    $restartResult = if ($Details.restartResult) { [string]$Details.restartResult } else { '' }
    if ([string]::IsNullOrWhiteSpace($restartResult)) {
        if ($null -ne $Details.killed) {
            $restartResult = if ([bool]$Details.killed) { 'killed for respawn' } else { 'process already exited' }
        }
    }

    $rows = @(
        @{ label = 'Detected (UTC)'; value = [string]$Details.detectedUtc }
        @{ label = 'Recovery reason'; value = [string]$Details.stallReason }
        @{ label = 'No-log duration (seconds)'; value = [string]$Details.secondsSilent }
        @{ label = 'Last log activity (UTC)'; value = [string]$Details.lastLogActivityUtc }
        @{ label = 'Last event code'; value = [string]$Details.lastEventCode }
        @{ label = 'Watcher PID'; value = [string]$Details.watcherPid }
        @{ label = 'Restart result'; value = $restartResult }
        @{ label = 'Details'; value = [string]$Details.errorMessage }
    )

    $htmlRows = ($rows | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_.value)) { return }
        $label = [System.Net.WebUtility]::HtmlEncode([string]$_.label)
        $value = [System.Net.WebUtility]::HtmlEncode([string]$_.value)
        "<tr><td style=""padding:6px 12px;font-weight:600;vertical-align:top;"">$label</td><td style=""padding:6px 12px;"">$value</td></tr>"
    }) -join ''

    return @"
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1a1a1a;">
  <p style="margin:0 0 12px 0;"><strong>The QC watcher child was restarted after a stall.</strong></p>
  <p style="margin:0 0 16px 0;">The dashboard killed and respawned the watcher because JSONL progress stopped. ProjectWise may still be connected; verify the dashboard and watcher logs if triggers were missed during the stall window.</p>
  <table style="border-collapse:collapse;border:1px solid #d0d0d0;">$htmlRows</table>
</div>
"@
}

function Send-QCWatcherStallRecoveryAlert {
    <#
    .SYNOPSIS
    Sends an operational email when the dashboard kills and respawns a wedged watcher child.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$Details,
        [switch]$Force
    )

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'QC_WATCHER_STALL_ALERT_SKIPPED_DISABLED' -Message 'Watcher operational alerts are disabled.' -Data @{ skipped = $true }
    }

    $recipients = @($settings.recipients)
    if ($recipients.Count -eq 0) {
        return New-QCFailureResult -Code 'QC_WATCHER_STALL_ALERT_NO_RECIPIENTS' -Message 'No stall alert recipients configured.' -Data @{ skipped = $true }
    }

    if (-not $Force) {
        $due = Test-QCWatcherStallAlertDue -Config $Config -Reason ([string]$Details.reason)
        if (-not $due.due) {
            return New-QCSuccessResult -Code 'QC_WATCHER_STALL_ALERT_SKIPPED_DEDUPED' -Message 'Stall alert suppressed by dedupe window.' -Data $due
        }
    }

    $importance = [string]$settings.importance
    if ([string]::IsNullOrWhiteSpace($importance)) { $importance = 'high' }

    $subject = '[QC Watcher] Watcher child restarted after stall'
    $htmlBody = New-QCWatcherStallRecoveryEmailBody -Details $Details
    $sendResult = _QCWA-SendWatcherOperationalEmail -Config $Config -Subject $subject -HtmlBody $htmlBody `
        -EventType 'QC_WATCHER_STALL_RECOVERY' -Importance $importance -Recipients $recipients

    if ($sendResult.IsSuccess) {
        Set-QCWatcherStallAlertSent -Config $Config -Reason ([string]$Details.reason) -Details $Details
        return New-QCSuccessResult -Code 'QC_WATCHER_STALL_ALERT_SENT' -Message 'Watcher stall recovery alert sent.' -Data @{
            to = $recipients
            importance = $importance
            send = $sendResult.Data
        }
    }

    return New-QCFailureResult -Code 'QC_WATCHER_STALL_ALERT_SEND_FAILED' -Message ([string]$sendResult.Message) -Data @{
        sendCode = [string]$sendResult.Code
    }
}

function Test-QCWatcherAuditActivityStalled {
    <#
    .SYNOPSIS
    Diagnostic helper for audit watermark lag. Not used for session-loss alerts by default.
    Quiet periods without QC events can exceed the lag threshold while ProjectWise is healthy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$MaxPwActTimeUtc = '',
        [string]$PollUntilUtc = '',
        [datetime]$LastMaxPwActChangeUtc
    )

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
    if (-not [bool]$settings.enableAuditStallDetection) {
        return @{ stalled = $false; reason = 'disabled' }
    }
    if ([string]::IsNullOrWhiteSpace($MaxPwActTimeUtc) -or [string]::IsNullOrWhiteSpace($PollUntilUtc)) {
        return @{ stalled = $false; reason = 'missing_timestamps' }
    }

    try {
        $maxAct = [DateTime]::ParseExact(
            $MaxPwActTimeUtc.Trim().TrimEnd('Z'),
            'yyyy-MM-dd HH:mm:ss',
            $null,
            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        )
        $until = [DateTime]::ParseExact(
            $PollUntilUtc.Trim().TrimEnd('Z'),
            'yyyy-MM-dd HH:mm:ss',
            $null,
            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        )
    } catch {
        return @{ stalled = $false; reason = 'parse_failed'; error = [string]$_.Exception.Message }
    }

    $lagSeconds = ($until - $maxAct).TotalSeconds
    $sinceChangeSeconds = if ($LastMaxPwActChangeUtc -and $LastMaxPwActChangeUtc -gt [datetime]'1970-01-01') {
        ((Get-Date).ToUniversalTime() - $LastMaxPwActChangeUtc).TotalSeconds
    } else {
        $lagSeconds
    }

    $stalled = ($lagSeconds -ge [double]$settings.staleAuditLagSeconds) -and
        ($sinceChangeSeconds -ge [double]$settings.staleAuditLagSeconds)

    return @{
        stalled = [bool]$stalled
        lagSeconds = [math]::Round($lagSeconds, 1)
        sinceChangeSeconds = [math]::Round($sinceChangeSeconds, 1)
        thresholdSeconds = [int]$settings.staleAuditLagSeconds
    }
}

Export-ModuleMember -Function @(
    'Get-QCWatcherSessionAlertSettings',
    'Get-QCWatcherSessionReconnectSettings',
    'Test-QCWatcherSessionReconnectDue',
    'Get-QCWatcherSessionAlertStatePath',
    'Test-QCWatcherSessionAlertDue',
    'Set-QCWatcherSessionAlertSent',
    'Send-QCWatcherSessionLostAlert',
    'Get-QCWatcherStallAlertStatePath',
    'Test-QCWatcherStallAlertDue',
    'Set-QCWatcherStallAlertSent',
    'Send-QCWatcherStallRecoveryAlert',
    'Test-QCWatcherAuditActivityStalled',
    'New-QCWatcherSessionLostEmailBody',
    'New-QCWatcherStallRecoveryEmailBody'
)
