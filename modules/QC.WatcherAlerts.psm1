# QC.WatcherAlerts.psm1
# Responsibility: Operational email alerts when the QC watcher loses ProjectWise connectivity.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationGraph.psm1') -Force -ErrorAction SilentlyContinue

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
    $root = $PSScriptRoot
    if ($root -match '[\\/]modules$') { return (Split-Path -Parent $root) }
    return $root
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

    $notif = _QCWA-ToHashtable $Config.notifications
    if (-not $notif) { $notif = @{} }
    $provider = if ($notif.provider) { [string]$notif.provider } else { 'Mock' }
    $dryRun = $false
    if ($null -ne $notif.dryRun) { try { $dryRun = [bool]$notif.dryRun } catch { } }
    if ($Config.ContainsKey('dryRun') -and [bool]$Config.dryRun) { $dryRun = $true }

    $importance = [string]$settings.importance
    if ([string]::IsNullOrWhiteSpace($importance)) { $importance = 'high' }

    $subject = '[QC Watcher] URGENT: ProjectWise session lost'
    $htmlBody = New-QCWatcherSessionLostEmailBody -Details $Details
    $payload = @{
        eventType = 'QC_WATCHER_PW_SESSION_LOST'
        documentName = 'QC Watcher'
        subject = $subject
        htmlBody = $htmlBody
        importance = $importance
        to = $recipients
        cc = @()
        logoPath = if ($notif.email -and $notif.email.logoPath) { [string]$notif.email.logoPath } else { 'email/typsalogo.png.webp' }
    }

    $sendResult = $null
    if ($provider -eq 'MicrosoftGraph' -and (Get-Command -Name 'Send-QCNotificationGraph' -ErrorAction SilentlyContinue)) {
        $graph = _QCWA-ToHashtable $notif.graph
        if (-not $graph) { $graph = @{} }
        $sendResult = Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:$dryRun
    } else {
        Import-Module (Join-Path $PSScriptRoot 'QC.NotificationMock.psm1') -Force -ErrorAction SilentlyContinue | Out-Null
        if (Get-Command -Name 'Send-QCNotificationMock' -ErrorAction SilentlyContinue) {
            $outputRoot = if ($notif.outputRoot) { [string]$notif.outputRoot } else { 'notifications' }
            if (-not [System.IO.Path]::IsPathRooted($outputRoot)) {
                $outputRoot = Join-Path (_QCWA-GetRepoRoot) $outputRoot
            }
            $mockPayload = $payload.Clone()
            $mockPayload['body'] = $htmlBody
            $sendResult = Send-QCNotificationMock -Payload $mockPayload -OutputRoot $outputRoot -DryRun:$dryRun
        } else {
            return New-QCFailureResult -Code 'QC_WATCHER_ALERT_PROVIDER_UNAVAILABLE' -Message 'No notification provider available for watcher session alert.' -Data @{ provider = $provider }
        }
    }

    if ($sendResult.IsSuccess) {
        Set-QCWatcherSessionAlertSent -Config $Config -Reason ([string]$Details.reason) -Details $Details
        return New-QCSuccessResult -Code 'QC_WATCHER_ALERT_SENT' -Message 'Watcher session lost alert sent.' -Data @{
            provider = $provider
            dryRun = $dryRun
            to = $recipients
            importance = $importance
            send = $sendResult.Data
        }
    }

    return New-QCFailureResult -Code 'QC_WATCHER_ALERT_SEND_FAILED' -Message ([string]$sendResult.Message) -Data @{
        provider = $provider
        sendCode = [string]$sendResult.Code
    }
}

function Test-QCWatcherAuditActivityStalled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$MaxPwActTimeUtc = '',
        [string]$PollUntilUtc = '',
        [datetime]$LastMaxPwActChangeUtc
    )

    $settings = Get-QCWatcherSessionAlertSettings -Config $Config
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
    'Get-QCWatcherSessionAlertStatePath',
    'Test-QCWatcherSessionAlertDue',
    'Set-QCWatcherSessionAlertSent',
    'Send-QCWatcherSessionLostAlert',
    'Test-QCWatcherAuditActivityStalled',
    'New-QCWatcherSessionLostEmailBody'
)
