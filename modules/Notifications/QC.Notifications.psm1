# QC.Notifications.psm1
# Responsibility: Configurable QC workflow email notifications (Mock + future Microsoft Graph).

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Config.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationTemplates.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationMock.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationGraph.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Notifications/QC.NotificationThreads.psm1') -Force -Global -ErrorAction SilentlyContinue
if (-not (Get-Command -Name 'Get-PWDocName' -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectWise/PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue
}
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Workflow/QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue
# Core.Database must be imported by the caller. Re-importing with -Force
# here clobbers the caller's global-scope exports.

# Test-friendly stubs: allow unit tests to run without importing Core.Database.
if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
    function Test-QCDatabaseEnabled {
        [CmdletBinding()]
        param([Parameter(Mandatory)][hashtable]$Config)
        return $false
    }
}

function _QCN-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($key in @($Value.Keys)) { $h[[string]$key] = $Value[$key] }
        return $h
    }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return @{ value = $Value } }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCN-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCN-GetProp([object]$Object, [string[]]$Names) {
    foreach ($n in @($Names)) {
        try {
            if ($Object -is [hashtable]) {
                if ($Object.ContainsKey($n) -and $null -ne $Object[$n]) { return $Object[$n] }
                continue
            }
            if ($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n) { return $Object.$n }
        } catch { }
    }
    return $null
}

function _QCN-GetJobValue([hashtable]$Job, [string[]]$Keys) {
    if (-not $Job) { return $null }
    foreach ($k in @($Keys)) {
        if ($Job.ContainsKey($k) -and $null -ne $Job[$k]) { return $Job[$k] }
    }
    $md = $null
    if ($Job.ContainsKey('metadata') -and $Job.metadata) { $md = _QCN-ToHashtable $Job.metadata }
    if ($md) {
        foreach ($k in @($Keys)) {
            if ($md.ContainsKey($k) -and $null -ne $md[$k]) { return $md[$k] }
        }
        if ($md.ContainsKey('attributes') -and $md.attributes) {
            $attrs = _QCN-ToHashtable $md.attributes
            if ($attrs) {
                foreach ($k in @($Keys)) {
                    if ($attrs.ContainsKey($k) -and $null -ne $attrs[$k]) { return $attrs[$k] }
                }
            }
        }
        if ($md.ContainsKey('candidate') -and $md.candidate) {
            $cand = _QCN-ToHashtable $md.candidate
            if ($cand) {
                foreach ($k in @($Keys)) {
                    if ($cand.ContainsKey($k) -and $null -ne $cand[$k]) { return $cand[$k] }
                }
            }
        }
    }
    return $null
}

function _QCN-GetRepoRoot() {
    return (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
}

function _QCN-ParseEmailList([object]$Value) {
    if (_QCN-IsBlank $Value) { return @() }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $out = [System.Collections.Generic.List[string]]::new()
        foreach ($item in @($Value)) {
            foreach ($email in (_QCN-ParseEmailList $item)) {
                if ($email -and -not $out.Contains($email)) { $out.Add($email) | Out-Null }
            }
        }
        return @($out)
    }
    $text = [string]$Value
    $parts = $text -split '[;,`\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { -not (_QCN-IsBlank $_) }
    $unique = [System.Collections.Generic.List[string]]::new()
    foreach ($p in @($parts)) {
        if ($p -match '@' -and -not $unique.Contains($p)) { $unique.Add($p) | Out-Null }
    }
    return @($unique)
}

function _QCN-GetAttributeValue([object]$Document, [string]$AttributeName) {
    if (_QCN-IsBlank $AttributeName) { return $null }
    $containers = @()
    foreach ($prop in @('qcAttributes','attributes','Attributes','CustomAttributes','EnvironmentAttributes')) {
        try { if ($Document -and $Document.PSObject.Properties[$prop] -and $Document.$prop) { $containers += $Document.$prop } } catch { }
    }
    foreach ($bag in @($containers)) {
        if ($bag -is [System.Collections.IEnumerable] -and -not ($bag -is [string]) -and -not ($bag -is [System.Collections.IDictionary])) {
            foreach ($item in @($bag)) {
                if ($item -is [System.Collections.IDictionary]) {
                    if ($item.Contains($AttributeName)) { return $item[$AttributeName] }
                }
            }
        }
        $h = _QCN-ToHashtable $bag
        if ($h -and $h.ContainsKey($AttributeName)) { return $h[$AttributeName] } # $h is always [hashtable] here
    }
    try { if ($Document -and $Document.PSObject.Properties[$AttributeName]) { return $Document.$AttributeName } } catch { }
    return $null
}

function Test-QCNotificationsEnqueueAsJob {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    try {
        if ($Config -and $Config.ContainsKey('notifications') -and $Config.notifications) {
            $n = _QCN-ToHashtable $Config.notifications
            if ($n -and $n.ContainsKey('enqueueAsJob')) { return [bool]$n.enqueueAsJob }
        }
    } catch { }
    return $false
}

function Get-QCNotificationSettings {
    [CmdletBinding()]
    param([hashtable]$Config)

    $raw = @{}
    if ($Config -and $Config.ContainsKey('notifications') -and $Config.notifications) {
        $norm = _QCN-ToHashtable $Config.notifications
        if ($norm) { $raw = $norm }
    }

    $defaults = @{
        enabled = $false
        provider = 'Mock'
        dryRun = $true
        rollbackWhenEmailAttributesMissing = $true
        outputRoot = (Join-Path (_QCN-GetRepoRoot) 'notifications')
        dedupe = @{
            enabled = $true
            storePath = (Join-Path (_QCN-GetRepoRoot) 'notifications\dedupe\sent-keys.jsonl')
            # Logical sheet-transition identity.  Do not key by transitionId by default because
            # DGN, sheet PDF, and lane QC PDF sibling sync can each create their own transition row.
            keyFields = @('sheetStem', 'qcProcessType', 'previousState', 'currentState', 'transitionSource', 'logicalTransitionAnchor', 'recipientKey')
            sheetPackageKeyFields = @('sheetStem', 'currentState', 'cycleId')
        }
        graph = @{
            tenantId = ''
            clientId = ''
            clientSecret = ''
            senderMailbox = ''
            certificateThumbprint = ''
            certificatePath = ''
        }
        email = @{
            bodyFormat = 'Text'
            templatePath = 'email/templates/qc_notification.html'
            logoPath = 'email/typsalogo.png.webp'
            environment = 'Production'
            subjectTemplate = '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}'
            qcPdfUrlTemplate = ''
            pwLinkBaseUrl = 'https://connect-projectwisewac.bentley.com/pwlink/'
            pwLinkApp = 'pwe'
        }
        attributes = @{
            reviewerEmailField = 'EM_Reviewer_Email'
            designerEmailField = 'EM_Designer_Email'
            checkerEmailField = 'EM_Checker_Email'
            ccEmailField = 'CcEmails'
            qcPdfUrlField = ''
            projectNumberField = ''
            reviewTypeField = ''
        }
        events = @{
            'QC Received' = @{
                enabled = $true
                eventType = 'QC_RECEIVED'
                to = @('reviewers')
                cc = @('designers')
                actionRequired = 'Reviewer to begin QC review.'
            }
            'Ready for QC' = @{
                enabled = $true
                eventType = 'READY_FOR_QC'
                to = @('reviewers')
                cc = @('designers')
                actionRequired = 'Reviewer to begin QC review.'
            }
            'Redlines Received' = @{
                enabled = $true
                eventType = 'REDLINES_RECEIVED'
                to = @('designers')
                cc = @('reviewers')
                actionRequired = 'Designer to address QC comments and return corrections.'
            }
            'Error Needs Attention' = @{
                enabled = $true
                eventType = 'QC_ERROR'
                to = @('reviewers', 'designers')
                cc = @()
                actionRequired = 'Manual review required.'
            }
            'QC Complete' = @{
                enabled = $true
                eventType = 'QC_COMPLETE'
                to = @('designers', 'reviewers')
                cc = @()
                actionRequired = 'QC review cycle is complete.'
            }
        }
    }

    $settings = @{}
    foreach ($k in $defaults.Keys) { $settings[$k] = $defaults[$k] }
    foreach ($k in $raw.Keys) {
        if ($k -in @('dedupe','graph','attributes','events','email')) {
            $merged = _QCN-ToHashtable $defaults[$k]
            $incoming = _QCN-ToHashtable $raw[$k]
            if ($incoming) {
                foreach ($ik in $incoming.Keys) { $merged[$ik] = $incoming[$ik] }
            }
            $settings[$k] = $merged
        } else {
            $settings[$k] = $raw[$k]
        }
    }
    foreach ($boolKey in @('enabled','dryRun','rollbackWhenEmailAttributesMissing')) {
        try { $settings[$boolKey] = [bool]$settings[$boolKey] } catch { $settings[$boolKey] = [bool]$defaults[$boolKey] }
    }
    if ($settings.dedupe) {
        try { $settings.dedupe.enabled = [bool]$settings.dedupe.enabled } catch { }
    }
    if ($settings.outputRoot) {
        $settings.outputRoot = _QCN-ResolveRepoPath -Path ([string]$settings.outputRoot)
    }
    if ($settings.dedupe -and $settings.dedupe.storePath) {
        $settings.dedupe.storePath = _QCN-ResolveRepoPath -Path ([string]$settings.dedupe.storePath)
    }
    _QCN-NormalizeNotificationSubjectTemplates -Settings $settings
    return $settings
}

function _QCN-NormalizeNotificationSubjectTemplates {
    param([hashtable]$Settings)

    if (-not $Settings) { return }
    $emailCfg = _QCN-ToHashtable $Settings.email
    $globalTemplate = ''
    if ($emailCfg -and $emailCfg.subjectTemplate) {
        $globalTemplate = [string]$emailCfg.subjectTemplate
    }
    if (_QCN-IsBlank $globalTemplate) { return }

    if (-not $Settings.events) { return }
    foreach ($stateKey in @($Settings.events.Keys)) {
        $ev = _QCN-ToHashtable $Settings.events[$stateKey]
        if (-not $ev -or -not $ev.ContainsKey('subjectTemplate')) { continue }
        $ev.Remove('subjectTemplate')
        $Settings.events[$stateKey] = $ev
    }
}

function _QCN-GetDefaultNotificationSubjectTemplate {
    return '[{ReviewType}] {ProjectName} | {DocumentName} | {WorkflowState}'
}

function _QCN-NewNotificationSubjectTokens {
    param(
        [string]$DocumentName,
        [string]$DocumentPath = '',
        [string]$Project = '',
        [string]$PreviousState = '',
        [string]$CurrentState = '',
        [string]$EventType = '',
        [string]$ReviewType = ''
    )

    $tokens = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $tokens['documentName'] = [string]$DocumentName
    $tokens['DocumentName'] = [string]$DocumentName
    $tokens['DocumentID'] = [string]$DocumentName
    $tokens['documentId'] = [string]$DocumentName
    $tokens['documentPath'] = [string]$DocumentPath
    $tokens['project'] = [string]$Project
    $tokens['ProjectName'] = [string]$Project
    $tokens['previousState'] = [string]$PreviousState
    $tokens['currentState'] = [string]$CurrentState
    $tokens['WorkflowState'] = [string]$CurrentState
    $tokens['eventType'] = [string]$EventType
    $tokens['reviewType'] = [string]$ReviewType
    $tokens['ReviewType'] = [string]$ReviewType
    $tokens['qc_review_type'] = [string]$ReviewType
    $tokens['qcReviewType'] = [string]$ReviewType
    return $tokens
}

function _QCN-ResolveNotificationSubjectTemplate {
    param(
        [hashtable]$EventCfg,
        [hashtable]$Settings
    )

    $template = ''
    if ($Settings -and $Settings.email) {
        $emailCfg = _QCN-ToHashtable $Settings.email
        if ($emailCfg -and $emailCfg.subjectTemplate) {
            $template = [string]$emailCfg.subjectTemplate
        }
    }
    if (_QCN-IsBlank $template -and $EventCfg -and $EventCfg.subjectTemplate) {
        $template = [string]$EventCfg.subjectTemplate
    }
    if (_QCN-IsBlank $template) {
        $template = _QCN-GetDefaultNotificationSubjectTemplate
    }
    return $template
}

function _QCN-GetAutomationActorLabel {
    param([hashtable]$Config)
    if (-not $Config) { return 'QC Automation' }
    $ap = _QCN-ToHashtable $Config.auditPoller
    if (-not $ap) { return 'QC Automation' }
    $wt = _QCN-ToHashtable $ap.workflowTriggers
    if (-not $wt) { return 'QC Automation' }
    if ($wt.automationPwUsernames) {
        foreach ($name in @($wt.automationPwUsernames)) {
            if (-not (_QCN-IsBlank $name)) { return [string]$name }
        }
    }
    return 'QC Automation'
}

function _QCN-FormatPwUserIdentityDisplay {
    param([hashtable]$Identity)
    if (-not $Identity) { return '' }
    foreach ($key in @('pw_user_email', 'display_name', 'pw_username')) {
        if ($Identity.ContainsKey($key) -and -not (_QCN-IsBlank $Identity[$key])) {
            return [string]$Identity[$key]
        }
    }
    if ($Identity.ContainsKey('pw_userno')) {
        try {
            $n = [int]$Identity.pw_userno
            if ($n -gt 0) { return ('PW User ' + [string]$n) }
        } catch { }
    }
    return ''
}

function _QCN-TrySyncPwUserIdentity {
    param(
        [hashtable]$Config,
        [int]$UserNumber
    )
    if ($UserNumber -le 0) { return $null }
    if (-not (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) { return $null }
    try {
        Sync-PWUserDirectory -Config $Config -UserNumbers @($UserNumber) -MaxUsers 1 | Out-Null
    } catch { }
    if (-not (Get-Command -Name 'Get-QCPWUserIdentity' -ErrorAction SilentlyContinue)) { return $null }
    try {
        $res = Get-QCPWUserIdentity -Config $Config -UserNumber $UserNumber
        if ($res -and $res.IsSuccess -and $res.Data -and $res.Data.identity) {
            return _QCN-ToHashtable $res.Data.identity
        }
    } catch { }
    return $null
}

function Resolve-QCNotificationSubmittedBy {
    <#
    .SYNOPSIS
    Resolves the display label for the user who triggered a workflow state change.
    Prefers email, then display name, then PW username.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config = $null,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$SubmittedBy = ''
    )

    if (-not (_QCN-IsBlank $SubmittedBy)) { return [string]$SubmittedBy.Trim() }

    $username = if ($ChangedByUsername) { [string]$ChangedByUsername.Trim() } else { '' }
    $userno = $null
    if ($null -ne $ChangedByUser) {
        try {
            $n = [int]$ChangedByUser
            if ($n -gt 0) { $userno = $n }
        } catch { }
    }

    $isAutomation = $false
    if ($Config -and (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue)) {
        try {
            $isAutomation = [bool](Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $userno -ChangedByUsername $username)
        } catch { }
    }
    if ($isAutomation) {
        if (-not (_QCN-IsBlank $username)) { return $username }
        return _QCN-GetAutomationActorLabel -Config $Config
    }

    $identity = $null
    if ($userno -and $Config -and (Get-Command -Name 'Get-QCPWUserIdentity' -ErrorAction SilentlyContinue)) {
        try {
            $res = Get-QCPWUserIdentity -Config $Config -UserNumber $userno
            if ($res -and $res.IsSuccess -and $res.Data -and $res.Data.identity) {
                $identity = _QCN-ToHashtable $res.Data.identity
            }
        } catch { }
        if (-not $identity) {
            $identity = _QCN-TrySyncPwUserIdentity -Config $Config -UserNumber $userno
        }
    }
    $fromIdentity = _QCN-FormatPwUserIdentityDisplay -Identity $identity
    if (-not (_QCN-IsBlank $fromIdentity)) { return $fromIdentity }
    if (-not (_QCN-IsBlank $username)) { return $username }
    if ($userno) { return ('PW User ' + [string]$userno) }
    return '(unknown)'
}

function Resolve-QCNotificationStateChangeActor {
    <#
    .SYNOPSIS
    Resolves PW user number and username for the actor who caused a workflow state change.
    Prefers transition_events / audit_events when stateTransitionKey references a DB row.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config = $null,
        [string]$StateTransitionKey = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [hashtable]$Job = $null
    )

    $user = $null
    $username = ''
    $key = ([string]$StateTransitionKey).Trim()

    if ($Config -and $key.Length -gt 0) {
        if ($key -match '^transition:(\d+)$' -and (Get-Command -Name 'Get-QCTransitionEventActor' -ErrorAction SilentlyContinue)) {
            try {
                $tid = [int]$Matches[1]
                $fromDb = Get-QCTransitionEventActor -Config $Config -TransitionId $tid
                if ($fromDb) {
                    if ($null -ne $fromDb.changedByUser) { $user = $fromDb.changedByUser }
                    if (-not (_QCN-IsBlank $fromDb.changedByUsername)) { $username = [string]$fromDb.changedByUsername }
                }
            } catch { }
        }
        elseif ($key -match '^audit:(\d+)$' -and (Get-Command -Name 'Get-QCAuditEventActor' -ErrorAction SilentlyContinue)) {
            try {
                $aid = [long]$Matches[1]
                $fromDb = Get-QCAuditEventActor -Config $Config -AuditEventId $aid
                if ($fromDb) {
                    if ($null -eq $user -and $null -ne $fromDb.changedByUser) { $user = $fromDb.changedByUser }
                    if (_QCN-IsBlank $username -and -not (_QCN-IsBlank $fromDb.changedByUsername)) {
                        $username = [string]$fromDb.changedByUsername
                    }
                }
            } catch { }
        }
    }

    if ($null -eq $user -and $null -ne $ChangedByUser) {
        try {
            $n = [int]$ChangedByUser
            if ($n -gt 0) { $user = $n }
        } catch { }
    }
    if (_QCN-IsBlank $username -and -not (_QCN-IsBlank $ChangedByUsername)) {
        $username = [string]$ChangedByUsername.Trim()
    }

    if (($null -eq $user) -or (_QCN-IsBlank $username)) {
        $fromJob = _QCN-ResolveStateChangeActorFromJob -Job $Job
        if ($null -eq $user -and $null -ne $fromJob.changedByUser) { $user = $fromJob.changedByUser }
        if (_QCN-IsBlank $username -and -not (_QCN-IsBlank $fromJob.changedByUsername)) {
            $username = [string]$fromJob.changedByUsername
        }
    }

    return @{ changedByUser = $user; changedByUsername = $username }
}

function _QCN-ResolveStateChangeActorFromJob {
    param([hashtable]$Job)

    $changedByUser = $null
    $changedByUsername = ''
    $submittedBy = ''
    if (-not $Job) {
        return @{ changedByUser = $null; changedByUsername = ''; submittedBy = '' }
    }
    $md = $null
    if ($Job.ContainsKey('metadata') -and $Job.metadata) {
        $md = _QCN-ToHashtable $Job.metadata
    }
    if ($md) {
        if ($md.ContainsKey('changedByUser') -and $null -ne $md.changedByUser) {
            try { $changedByUser = [int]$md.changedByUser } catch { }
        }
        if ($md.ContainsKey('changedByUsername') -and -not (_QCN-IsBlank $md.changedByUsername)) {
            $changedByUsername = [string]$md.changedByUsername
        }
        if ($md.ContainsKey('submittedBy') -and -not (_QCN-IsBlank $md.submittedBy)) {
            $submittedBy = [string]$md.submittedBy
        }
    }
    return @{
        changedByUser = $changedByUser
        changedByUsername = $changedByUsername
        submittedBy = $submittedBy
    }
}

function New-QCNotificationEvent {
    <#
    .SYNOPSIS
    Builds a QC notification event hashtable for a QC PDF workflow state change.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EventType,
        [string]$Project = '',
        [Parameter(Mandatory)]
        [string]$DocumentName,
        [string]$DocumentPath = '',
        [string]$DocumentGuid = '',
        [string]$PreviousState = '',
        [Parameter(Mandatory)]
        [string]$CurrentState,
        [string[]]$Reviewers = @(),
        [string[]]$Designers = @(),
        [string[]]$Cc = @(),
        [string]$ActionRequired = '',
        [string]$SourceJobId = '',
        [string]$QcPdfUrl = '',
        [string]$SubmittedBy = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    return @{
        eventType = $EventType
        project = $Project
        documentName = $DocumentName
        documentPath = $DocumentPath
        documentGuid = $DocumentGuid
        previousState = $PreviousState
        currentState = $CurrentState
        reviewers = @($Reviewers)
        designers = @($Designers)
        cc = @($Cc)
        actionRequired = $ActionRequired
        sourceJobId = $SourceJobId
        qcPdfUrl = $QcPdfUrl
        submittedBy = $SubmittedBy
        changedByUser = $ChangedByUser
        changedByUsername = $ChangedByUsername
    }
}

function _QCN-ResolveRepoPath([string]$Path) {
    if (_QCN-IsBlank $Path) { return '' }
    $p = [string]$Path
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    return (Join-Path (_QCN-GetRepoRoot) $p)
}

function _QCN-GetEncodedPwDatasource([string]$DatasourceName) {
    if (_QCN-IsBlank $DatasourceName) { return '' }
    $ds = [string]$DatasourceName.Trim()
    $encoded = $ds.Replace(':', '~3A')
    return 'Bentley.PW--' + $encoded
}

function Resolve-QCNotificationQcPdfUrl {
    <#
    .SYNOPSIS
    Resolves an HTTPS URL to open the QC PDF in ProjectWise (web or Explorer link).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [hashtable]$Settings,
        [object]$Document = $null,
        [hashtable]$Config = $null
    )

    if ($Event.qcPdfUrl -and -not (_QCN-IsBlank $Event.qcPdfUrl)) {
        return [string]$Event.qcPdfUrl.Trim()
    }

    $trustedLaneGuid = _QCN-ResolveTrustedLaneDocumentGuid -Event $Event
    if (-not (_QCN-IsBlank $trustedLaneGuid)) {
        $trustedUrl = _QCN-BuildPwDocumentLinkUrl -DocumentGuid $trustedLaneGuid -Settings $Settings -Config $Config
        if (-not (_QCN-IsBlank $trustedUrl)) {
            $Event['linkResolutionSource'] = 'trusted_lane_metadata_guid'
            return $trustedUrl.Trim()
        }
    }

    $attrs = @{}
    if ($Settings -and $Settings.attributes) {
        $a = _QCN-ToHashtable $Settings.attributes
        if ($a) { $attrs = $a }
    }

    if ($attrs.ContainsKey('qcPdfUrlField') -and -not (_QCN-IsBlank $attrs['qcPdfUrlField']) -and $Document) {
        $fromAttr = _QCN-GetAttributeValue -Document $Document -AttributeName ([string]$attrs['qcPdfUrlField'])
        if ($fromAttr -and -not (_QCN-IsBlank $fromAttr)) {
            return [string]$fromAttr.Trim()
        }
    }

    $folderForGuid = if ($Event.folderPath) { [string]$Event.folderPath } elseif ($Event.documentPath -and ($Event.documentPath -match '\\')) {
        [System.IO.Path]::GetDirectoryName([string]$Event.documentPath)
    } else { '' }
    $sheetStem = if ($Event.sheetStem) { [string]$Event.sheetStem } else { '' }
    $srcForGuid = if ($Event.documentName) { [string]$Event.documentName } else { '' }
    if (-not (_QCN-IsBlank $srcForGuid) -and (Test-QCIsQcPdfDocumentName -DocumentName $srcForGuid)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            try {
                $stem = Get-PWSheetStemFromDocumentName -DocumentName $srcForGuid
                if (-not (_QCN-IsBlank $stem)) { $srcForGuid = $stem + '.pdf' }
            } catch { }
        } else {
            $srcForGuid = [System.IO.Path]::GetFileNameWithoutExtension($srcForGuid) + '.pdf'
        }
    }
    $qcPdfName = _QCN-NormalizeQcPdfDocumentName -DocumentName ([string]$Event.documentName) -SheetStem $sheetStem `
        -ProcessType ([string]$Event.qcProcessType) -Config $Config -Event $Event
    $hintGuid = if ($Event.documentGuid) { [string]$Event.documentGuid.Trim() } else { '' }
    $sheetPackageIdForLink = _QCN-ResolveNotificationSheetPackageId -Event $Event
    $linkResolutionSource = ''
    $docGuid = _QCN-ResolveLiveQcPdfDocumentGuid -Config $Config -FolderPath $folderForGuid `
        -QcPdfName $qcPdfName -SheetStem $sheetStem -HintGuid $hintGuid -SourceSheetPdfName $srcForGuid `
        -SheetPackageId $sheetPackageIdForLink -ResolutionSource ([ref]$linkResolutionSource) -Event $Event
    if (-not (_QCN-IsBlank $linkResolutionSource)) { $Event['linkResolutionSource'] = $linkResolutionSource }
    if ($docGuid -and (Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue)) {
        try {
            $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs @($docGuid) -ErrorAction SilentlyContinue)
            if ($docs.Count -gt 0) {
                $doc = $docs[0]
                foreach ($prop in @('ProjectWiseWebLink', 'projectWiseWebLink')) {
                    try {
                        if ($doc.PSObject.Properties[$prop] -and -not (_QCN-IsBlank $doc.$prop)) {
                            return [string]$doc.$prop.Trim()
                        }
                    } catch { }
                }
            }
        } catch { }
    }

    $emailCfg = @{}
    if ($Settings -and $Settings.email) {
        $e = _QCN-ToHashtable $Settings.email
        if ($e) { $emailCfg = $e }
    }

    $template = if ($emailCfg.qcPdfUrlTemplate) { [string]$emailCfg.qcPdfUrlTemplate } else { '' }
    if (-not (_QCN-IsBlank $template)) {
        $tokens = @{
            documentGuid = $docGuid
            documentName = if ($Event.documentName) { [string]$Event.documentName } else { '' }
            documentPath = if ($Event.documentPath) { [string]$Event.documentPath } else { '' }
            project = if ($Event.project) { [string]$Event.project } else { '' }
            currentState = if ($Event.currentState) { [string]$Event.currentState } else { '' }
            eventType = if ($Event.eventType) { [string]$Event.eventType } else { '' }
        }
        $expanded = Expand-QCNotificationTemplate -Template $template -Tokens $tokens
        if (-not (_QCN-IsBlank $expanded)) { return $expanded.Trim() }
    }

    $builtUrl = _QCN-BuildPwDocumentLinkUrl -DocumentGuid $docGuid -Settings $Settings -Config $Config
    if (-not (_QCN-IsBlank $builtUrl)) { return $builtUrl }

    return ''
}

function _QCN-UsesHtmlEmailBody([hashtable]$Settings, [hashtable]$EventCfg) {
    $emailCfg = @{}
    if ($Settings -and $Settings.email) {
        $e = _QCN-ToHashtable $Settings.email
        if ($e) { $emailCfg = $e }
    }
    $format = if ($emailCfg.bodyFormat) { ([string]$emailCfg.bodyFormat).Trim() } else { 'Text' }
    if ($EventCfg -and $EventCfg.ContainsKey('emailTemplate') -and -not (_QCN-IsBlank $EventCfg.emailTemplate)) {
        return $true
    }
    return ($format.Equals('Html', [StringComparison]::OrdinalIgnoreCase))
}

function _QCN-GetIndependentCheckReviewTypeLabel {
    param([hashtable]$Config)

    $label = 'Independent Check'
    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf) {
            $reviewTypes = _QCN-ToHashtable $wf['reviewTypes']
            if ($reviewTypes -and $reviewTypes['independentCheck']) {
                $label = [string]$reviewTypes['independentCheck']
            }
        }
    }
    return $label.Trim()
}

function _QCN-TestNotificationIndependentCheckReviewType {
    param(
        [string]$ReviewType,
        [hashtable]$Config
    )

    if (_QCN-IsBlank $ReviewType) { return $false }
    $expected = _QCN-GetIndependentCheckReviewTypeLabel -Config $Config
    return ([string]$ReviewType).Trim() -eq $expected
}

function _QCN-GetQcReviewTypeAttributeName {
    param([hashtable]$Config)

    $col = 'QC_Review_Type'
    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf) {
            $attrMap = _QCN-ToHashtable $wf['attributeMap']
            if ($attrMap -and $attrMap['reviewType']) { $col = [string]$attrMap['reviewType'] }
        }
    }
    return $col
}

function _QCN-GetQcCheckerEmailAttributeName {
    param([hashtable]$Config)

    $col = 'QC_Checker_Email'
    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf) {
            $attrMap = _QCN-ToHashtable $wf['attributeMap']
            if ($attrMap -and $attrMap['checkerEmail']) { $col = [string]$attrMap['checkerEmail'] }
        }
    }
    return $col
}

function _QCN-ResolveNotificationProcessType {
    param(
        [hashtable]$Event = $null,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    if ($Job -and $Job.metadata) {
        $jobMd = _QCN-ToHashtable $Job.metadata
        if ($jobMd -and $jobMd.ContainsKey('qcProcessType') -and -not (_QCN-IsBlank $jobMd['qcProcessType'])) {
            $fromJob = [string]$jobMd['qcProcessType']
            if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
                $norm = Normalize-QCProcessType -ProcessType $fromJob -AllowNullOnEmpty
                if ($norm) { return $norm }
            }
            if (-not (_QCN-IsBlank $fromJob)) { return $fromJob.Trim() }
        }
    }

    if ($Event) {
        if ($Event.qcProcessType) {
            if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
                $norm = Normalize-QCProcessType -ProcessType ([string]$Event.qcProcessType) -AllowNullOnEmpty
                if ($norm) { return $norm }
            }
            return ([string]$Event.qcProcessType).Trim()
        }
        if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
            $norm = Normalize-QCProcessType -ProcessType ([string]$Event.qcProcessType) -ReviewType ([string]$Event.qcReviewType) `
                -Context $Event -AllowNullOnEmpty
            if ($norm) { return $norm }
        }
    }

    return ''
}

function _QCN-ResolveReviewTypeLabelFromProcessType {
    param([string]$ProcessType = '')

    if (_QCN-IsBlank $ProcessType) { return '' }
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType ([string]$ProcessType) -AllowNullOnEmpty
        if ($norm -and (Get-Command -Name 'Get-QCProcessTypeDisplayLabel' -ErrorAction SilentlyContinue)) {
            return Get-QCProcessTypeDisplayLabel -ProcessType $norm
        }
    }
    $trimmed = ([string]$ProcessType).Trim()
    if (_QCN-IsBlank $trimmed) { return '' }
    return $trimmed.Substring(0, 1).ToUpperInvariant() + $trimmed.Substring(1).ToLowerInvariant()
}

function _QCN-TestUsableNotificationReviewTypeValue {
    param([string]$Value = '')

    if (_QCN-IsBlank $Value) { return $false }
    $trimmed = ([string]$Value).Trim()
    if ($trimmed -eq '0') { return $false }
    if ($trimmed -match '^\d+$') { return $false }
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType $trimmed -AllowNullOnEmpty
        if ($norm) { return $true }
    }
    return $true
}

function _QCN-ResolveNotificationReviewType {
    param(
        [object]$Document,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null,
        [string]$FolderPath = '',
        [string]$SourceName = '',
        [string]$DocumentName = ''
    )

    if ($Job -and $Job.metadata) {
        $jobMd = _QCN-ToHashtable $Job.metadata
        if ($jobMd -and $jobMd.ContainsKey('qcProcessType') -and -not (_QCN-IsBlank $jobMd['qcProcessType'])) {
            $fromJobProcess = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$jobMd['qcProcessType'])
            if (-not (_QCN-IsBlank $fromJobProcess)) { return $fromJobProcess }
        }
    }

    $laneName = [string]$DocumentName
    if (_QCN-IsBlank $laneName) { $laneName = [string]$SourceName }
    if ($Job -and $Job.metadata) {
        $jobMdLane = _QCN-ToHashtable $Job.metadata
        if ($jobMdLane -and $jobMdLane.expectedLanePdfName) { $laneName = [string]$jobMdLane.expectedLanePdfName }
        elseif ($jobMdLane -and $jobMdLane.notificationLaneDocumentName) { $laneName = [string]$jobMdLane.notificationLaneDocumentName }
    }
    if (-not (_QCN-IsBlank $laneName) -and (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue)) {
        $lane = Get-PWQcPdfLaneFromDocumentName -DocumentName $laneName
        if ($lane) {
            $fromLane = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType $lane
            if (-not (_QCN-IsBlank $fromLane)) { return $fromLane }
        }
    }

    $attr = _QCN-ToHashtable $Settings.attributes
    if (-not $attr) { $attr = @{} }

    if ($attr.ContainsKey('reviewTypeField') -and -not (_QCN-IsBlank $attr['reviewTypeField'])) {
        $fromNotify = _QCN-GetAttributeValue -Document $Document -AttributeName ([string]$attr['reviewTypeField'])
        if (_QCN-TestUsableNotificationReviewTypeValue -Value $fromNotify) { return ([string]$fromNotify).Trim() }
    }

    $processCol = 'QC_Process_Type'
    if (Get-Command -Name 'Get-PWQcProcessTypeAttributeName' -ErrorAction SilentlyContinue) {
        $processCol = Get-PWQcProcessTypeAttributeName -Config $Config
    }
    $fromProcess = _QCN-GetAttributeValue -Document $Document -AttributeName $processCol
    if (-not (_QCN-IsBlank $fromProcess)) {
        $fromProcessLabel = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$fromProcess)
        if (-not (_QCN-IsBlank $fromProcessLabel)) { return $fromProcessLabel }
        return ([string]$fromProcess).Trim()
    }

    $stemSource = [string]$SourceName
    if (-not (_QCN-IsBlank $stemSource) -and (Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) `
            -and (Test-QCIsQcPdfDocumentName -DocumentName $stemSource)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $stemSource = (Get-PWSheetStemFromDocumentName -DocumentName $stemSource) + '.pdf'
        }
    }
    if (_QCN-IsBlank $stemSource) { $stemSource = [string](_QCN-GetProp -Object $Document -Names @('Name', 'DocumentName', 'FileName')) }
    if (-not (_QCN-IsBlank $stemSource) -and (Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) `
            -and (Test-QCIsQcPdfDocumentName -DocumentName $stemSource)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $stemSource = (Get-PWSheetStemFromDocumentName -DocumentName $stemSource) + '.pdf'
        }
    }
    if ($Config -and -not (_QCN-IsBlank $FolderPath) -and -not (_QCN-IsBlank $stemSource)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $stemSource -Config $Config
            if ($pw.found -and -not (_QCN-IsBlank $pw.qcProcessType)) {
                $fromPwProcess = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$pw.qcProcessType)
                if (-not (_QCN-IsBlank $fromPwProcess)) { return $fromPwProcess }
            }
            if ($pw.found -and (_QCN-TestUsableNotificationReviewTypeValue -Value $pw.qcReviewType)) {
                return ([string]$pw.qcReviewType).Trim()
            }
        }
    }

    $qcReviewCol = _QCN-GetQcReviewTypeAttributeName -Config $Config
    $fromDoc = _QCN-GetAttributeValue -Document $Document -AttributeName $qcReviewCol
    if (_QCN-TestUsableNotificationReviewTypeValue -Value $fromDoc) {
        $fromLegacy = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$fromDoc)
        if (-not (_QCN-IsBlank $fromLegacy)) { return $fromLegacy }
        return ([string]$fromDoc).Trim()
    }

    if ($Job) {
        $fromJob = [string](_QCN-GetJobValue -Job $Job -Keys @('reviewType', 'qcReviewType'))
        if (_QCN-TestUsableNotificationReviewTypeValue -Value $fromJob) {
            $fromJobLabel = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType $fromJob
            if (-not (_QCN-IsBlank $fromJobLabel)) { return $fromJobLabel }
            return $fromJob.Trim()
        }
        if ($Job.metadata) {
            $jobMdRt = _QCN-ToHashtable $Job.metadata
            if ($jobMdRt -and $jobMdRt.attributes) {
                $attrs = _QCN-ToHashtable $jobMdRt.attributes
                if ($attrs) {
                    foreach ($k in @('reviewType', 'qcReviewType')) {
                        if ($attrs.ContainsKey($k) -and (_QCN-TestUsableNotificationReviewTypeValue -Value ([string]$attrs[$k]))) {
                            $fromAttr = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$attrs[$k])
                            if (-not (_QCN-IsBlank $fromAttr)) { return $fromAttr }
                            return ([string]$attrs[$k]).Trim()
                        }
                    }
                }
            }
        }
    }

    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf -and $wf.defaultReviewType -and (_QCN-TestUsableNotificationReviewTypeValue -Value ([string]$wf.defaultReviewType))) {
            $fromDefault = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType ([string]$wf.defaultReviewType)
            if (-not (_QCN-IsBlank $fromDefault)) { return $fromDefault }
            return ([string]$wf.defaultReviewType).Trim()
        }
    }

    return ''
}

function _QCN-ApplyIndependentCheckReviewerEmail {
    param(
        [object]$Document,
        [string]$CheckerEmail,
        [string]$CheckerField,
        [string]$QcCheckerField
    )

    if (-not (_QCN-IsBlank $CheckerEmail)) { return [string]$CheckerEmail }
    $fromEm = _QCN-GetAttributeValue -Document $Document -AttributeName $CheckerField
    if (-not (_QCN-IsBlank $fromEm)) { return ([string]$fromEm).Trim() }
    $fromQc = _QCN-GetAttributeValue -Document $Document -AttributeName $QcCheckerField
    if (-not (_QCN-IsBlank $fromQc)) { return ([string]$fromQc).Trim() }
    return ''
}

function _QCN-GetQcDesignerEmailAttributeName {
    param([hashtable]$Config)
    $col = 'QC_Designer_Email'
    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf) {
            $attrMap = _QCN-ToHashtable $wf['attributeMap']
            if ($attrMap -and $attrMap['designerEmail']) { $col = [string]$attrMap['designerEmail'] }
        }
    }
    return $col
}

function _QCN-GetQcReviewerEmailAttributeName {
    param([hashtable]$Config)
    $col = 'QC_Reviewer_Email'
    if ($Config) {
        $wf = _QCN-ToHashtable $Config['qcWorkflow']
        if ($wf) {
            $attrMap = _QCN-ToHashtable $wf['attributeMap']
            if ($attrMap -and $attrMap['reviewerEmail']) { $col = [string]$attrMap['reviewerEmail'] }
        }
    }
    return $col
}

function _QCN-GetRoleEmailsFromSheetIndex {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$SourceDocumentName = ''
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return $null }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }

    try {
        if (-not (_QCN-IsBlank $DocumentGuid)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, checker_email
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = [string]$DocumentGuid.Trim() }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $r = $res.Data.table.Rows[0]
                $roles = @{
                    designerEmail = if ($r.designer_email -is [DBNull]) { '' } else { [string]$r.designer_email }
                    reviewerEmail = if ($r.reviewer_email -is [DBNull]) { '' } else { [string]$r.reviewer_email }
                    checkerEmail = if ($r.checker_email -is [DBNull]) { '' } else { [string]$r.checker_email }
                }
                if (-not (_QCN-IsBlank $roles.reviewerEmail) -or -not (_QCN-IsBlank $roles.designerEmail)) {
                    return $roles
                }
            }
        }
        if (-not (_QCN-IsBlank $FolderPath) -and -not (_QCN-IsBlank $SourceDocumentName)) {
            $srcName = [string]$SourceDocumentName
            if (Test-QCIsQcPdfDocumentName -DocumentName $srcName) {
                if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                    try {
                        $stem = Get-PWSheetStemFromDocumentName -DocumentName $srcName
                        if (-not (_QCN-IsBlank $stem)) { $srcName = $stem + '.pdf' }
                    } catch { }
                } else {
                    $srcName = [System.IO.Path]::GetFileNameWithoutExtension($srcName) + '.pdf'
                }
            }
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, checker_email
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@srcName)
"@ -Parameters @{ folderPath = [string]$FolderPath; srcName = $srcName }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $r2 = $res2.Data.table.Rows[0]
                return @{
                    designerEmail = if ($r2.designer_email -is [DBNull]) { '' } else { [string]$r2.designer_email }
                    reviewerEmail = if ($r2.reviewer_email -is [DBNull]) { '' } else { [string]$r2.reviewer_email }
                    checkerEmail = if ($r2.checker_email -is [DBNull]) { '' } else { [string]$r2.checker_email }
                }
            }
        }
    } catch { return $null }
    return $null
}

function _QCN-ResolveQcPdfGuidFromSheetIndex {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$DocumentGuid = ''
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    try {
        if (-not (_QCN-IsBlank $DocumentGuid)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT qc_pdf_guid
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = [string]$DocumentGuid.Trim() }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $g = if ($res.Data.table.Rows[0].qc_pdf_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].qc_pdf_guid }
                if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
            }
        }
        if (-not (_QCN-IsBlank $FolderPath) -and -not (_QCN-IsBlank $SourceDocumentName)) {
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT qc_pdf_guid
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@srcName)
"@ -Parameters @{ folderPath = [string]$FolderPath; srcName = [string]$SourceDocumentName }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $g2 = if ($res2.Data.table.Rows[0].qc_pdf_guid -is [DBNull]) { '' } else { [string]$res2.Data.table.Rows[0].qc_pdf_guid }
                if (-not (_QCN-IsBlank $g2)) { return $g2.Trim() }
            }
        }
    } catch { }
    return ''
}

function _QCN-GetPwSearchFolderCandidates {
    param([string]$FolderPath)

    $candidates = New-Object 'System.Collections.Generic.List[string]'
    function Add-Candidate([string]$Candidate) {
        if (_QCN-IsBlank $Candidate) { return }
        $trimmed = $Candidate.Trim().TrimEnd('\')
        $exists = $false
        foreach ($existing in @($candidates)) {
            if ($existing -eq $trimmed) { $exists = $true; break }
        }
        if (-not $exists) { [void]$candidates.Add($trimmed) }
    }

    Add-Candidate $FolderPath
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        Add-Candidate (ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath)
    }
    if (Get-Command -Name 'ConvertTo-PWCanonicalDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        Add-Candidate (ConvertTo-PWCanonicalDocumentsFolderPath -FolderPathProperty $FolderPath)
    }
    if (Get-Command -Name '_PWD-GetSheetRoleFolderCandidates' -ErrorAction SilentlyContinue) {
        foreach ($fp in @(_PWD-GetSheetRoleFolderCandidates -FolderPath $FolderPath)) {
            Add-Candidate $fp
            if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
                Add-Candidate (ConvertTo-PWCmdletFolderPath -InternalFolderPath $fp)
            }
        }
    }
    return @($candidates.ToArray())
}

function _QCN-TryResolveQcPdfGuidFromPwSearch {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$QcPdfName
    )

    if (_QCN-IsBlank $FolderPath) { return '' }
    if (_QCN-IsBlank $QcPdfName) { return '' }

    $nameCandidates = [System.Collections.Generic.List[string]]::new()
    $seenNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    function _AddPwSearchName([string]$Candidate) {
        if (_QCN-IsBlank $Candidate) { return }
        $trimmed = $Candidate.Trim()
        if ($seenNames.Add($trimmed)) { [void]$nameCandidates.Add($trimmed) }
    }
    _AddPwSearchName $QcPdfName
    if ($QcPdfName -match '(?i)^(.+)-(prod|chk|rev)\.pdf$') {
        $stemPart = [string]$Matches[1]
        $suffixPart = [string]$Matches[2].ToLowerInvariant()
        _AddPwSearchName ($stemPart + '-' + $suffixPart + '.pdf')
    }

    $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
    if ($searchCmd) {
        foreach ($searchFolder in @(_QCN-GetPwSearchFolderCandidates -FolderPath $FolderPath)) {
            if (_QCN-IsBlank $searchFolder) { continue }
            foreach ($searchName in @($nameCandidates)) {
                try {
                    $docs = @(Get-PWDocumentsBySearch -FolderPath $searchFolder -DocumentName $searchName -JustThisFolder -ErrorAction SilentlyContinue)
                    if ($docs.Count -le 0) { continue }
                    $bestGuid = ''
                    $bestTicks = [long]::MinValue
                    foreach ($doc in $docs) {
                        $g = [string](_QCN-GetProp -Object $doc -Names @('DocumentGUID', 'DocumentGuid', 'GUID'))
                        if (_QCN-IsBlank $g) { continue }
                        $ticks = [long]::MinValue
                        foreach ($n in @('FileUpdatedDate', 'FileUpdateDate', 'DocumentUpdateDate', 'VersionModifiedDate')) {
                            $raw = _QCN-GetProp -Object $doc -Names @($n)
                            if (_QCN-IsBlank $raw) { continue }
                            try {
                                $dt = [datetime]$raw
                                if ($dt.Ticks -gt $ticks) { $ticks = $dt.Ticks }
                            } catch { }
                        }
                        if ((_QCN-IsBlank $bestGuid) -or ($ticks -gt $bestTicks)) {
                            $bestGuid = $g.Trim()
                            $bestTicks = $ticks
                        }
                    }
                    if (-not (_QCN-IsBlank $bestGuid)) { return $bestGuid }
                } catch { }
            }
        }
    }

    if (Get-Command -Name 'Get-PWDocumentsInFolder' -ErrorAction SilentlyContinue) {
        foreach ($searchFolder in @(_QCN-GetPwSearchFolderCandidates -FolderPath $FolderPath)) {
            if (_QCN-IsBlank $searchFolder) { continue }
            try {
                $all = @(Get-PWDocumentsInFolder -FolderPath $searchFolder -ErrorAction SilentlyContinue)
                $bestGuid = ''
                $bestTicks = [long]::MinValue
                foreach ($doc in $all) {
                    $actualName = [string](_QCN-GetProp -Object $doc -Names @('Name', 'DocumentName', 'FileName'))
                    if (_QCN-IsBlank $actualName) { continue }
                    $nameMatch = $false
                    foreach ($searchName in @($nameCandidates)) {
                        if ($actualName.Equals($searchName, [StringComparison]::OrdinalIgnoreCase)) {
                            $nameMatch = $true
                            break
                        }
                    }
                    if (-not $nameMatch) { continue }
                    $g = [string](_QCN-GetProp -Object $doc -Names @('DocumentGUID', 'DocumentGuid', 'GUID'))
                    if (_QCN-IsBlank $g) { continue }
                    $ticks = [long]::MinValue
                    foreach ($n in @('FileUpdatedDate', 'FileUpdateDate', 'DocumentUpdateDate', 'VersionModifiedDate')) {
                        $raw = _QCN-GetProp -Object $doc -Names @($n)
                        if (_QCN-IsBlank $raw) { continue }
                        try {
                            $dt = [datetime]$raw
                            if ($dt.Ticks -gt $ticks) { $ticks = $dt.Ticks }
                        } catch { }
                    }
                    if ((_QCN-IsBlank $bestGuid) -or ($ticks -gt $bestTicks)) {
                        $bestGuid = $g.Trim()
                        $bestTicks = $ticks
                    }
                }
                if (-not (_QCN-IsBlank $bestGuid)) { return $bestGuid }
            } catch { }
        }
    }
    return ''
}

function _QCN-ResolveTrustedLaneDocumentGuid {
    param(
        [hashtable]$Event = $null,
        [hashtable]$Job = $null
    )

    foreach ($source in @($Event, $(if ($Job -and $Job.metadata) { _QCN-ToHashtable $Job.metadata } else { $null }))) {
        if (-not $source) { continue }
        foreach ($key in @('notificationLaneDocumentGuid', 'laneQcPdfDocumentGuid')) {
            if ($source.ContainsKey($key) -and -not (_QCN-IsBlank $source[$key])) {
                return [string]$source[$key].Trim()
            }
        }
    }
    return ''
}

function _QCN-BuildPwDocumentLinkUrl {
    param(
        [string]$DocumentGuid,
        [hashtable]$Settings,
        [hashtable]$Config = $null
    )

    if (_QCN-IsBlank $DocumentGuid) { return '' }
    $emailCfg = @{}
    if ($Settings -and $Settings.email) {
        $e = _QCN-ToHashtable $Settings.email
        if ($e) { $emailCfg = $e }
    }
    $baseUrl = if ($emailCfg.pwLinkBaseUrl) { [string]$emailCfg.pwLinkBaseUrl } else { '' }
    if (_QCN-IsBlank $baseUrl) { return '' }
    $datasourceName = ''
    if ($Config -and $Config.projectWise) {
        $pw = _QCN-ToHashtable $Config.projectWise
        if ($pw) {
            if ($pw.datasourceName) { $datasourceName = [string]$pw.datasourceName }
            elseif ($pw.datasource) { $datasourceName = [string]$pw.datasource }
        }
    }
    if (_QCN-IsBlank $datasourceName -and (Get-Command -Name 'Get-PWCurrentDatasource' -ErrorAction SilentlyContinue)) {
        try { $datasourceName = [string](Get-PWCurrentDatasource) } catch { }
    }
    if (_QCN-IsBlank $datasourceName) { return '' }
    $dsParam = _QCN-GetEncodedPwDatasource -DatasourceName $datasourceName
    $app = if ($emailCfg.pwLinkApp) { [string]$emailCfg.pwLinkApp } else { 'pwe' }
    $sep = if ($baseUrl.Contains('?')) { '&' } else { '?' }
    return ('{0}{1}objectId={2}&objectType=doc&datasource={3}&app={4}' -f $baseUrl.TrimEnd('/'), $sep, $DocumentGuid.Trim(), $dsParam, $app)
}

function _QCN-NormalizeQcPdfDocumentName {
    param(
        [string]$DocumentName = '',
        [string]$SheetStem = '',
        [string]$ProcessType = '',
        [hashtable]$Config = $null,
        [hashtable]$Event = $null
    )

    $resolvedProcessType = [string]$ProcessType
    if (_QCN-IsBlank $resolvedProcessType -and $Event) {
        if ($Event.qcProcessType) { $resolvedProcessType = [string]$Event.qcProcessType }
        elseif (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
            $norm = Normalize-QCProcessType -ProcessType ([string]$Event.qcProcessType) -ReviewType ([string]$Event.qcReviewType) -Context $Event -AllowNullOnEmpty
            if ($norm) { $resolvedProcessType = $norm }
        }
    }
    if (_QCN-IsBlank $resolvedProcessType) {
        $resolvedProcessType = 'production'
    }

    $name = [string]$DocumentName
    if ((_QCN-IsBlank $name) -and (-not (_QCN-IsBlank $SheetStem))) {
        if (Get-Command -Name 'Get-QCLaneQcPdfExpectedName' -ErrorAction SilentlyContinue) {
            $expected = Get-QCLaneQcPdfExpectedName -SheetBaseName $SheetStem -ProcessType $resolvedProcessType -Config $Config
            if (-not (_QCN-IsBlank $expected)) { return $expected.Trim() }
        }
        return ([string]$SheetStem + '-prod.pdf')
    }
    if (_QCN-IsBlank $name) { return '' }
    if ($name -match '(?i)-(prod|chk|rev)\.pdf$') { return $name.Trim() }

    $base = [string]$SheetStem
    if (_QCN-IsBlank $base) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $stem = Get-PWSheetStemFromDocumentName -DocumentName $name
            if (-not (_QCN-IsBlank $stem)) { $base = $stem }
        }
        if (_QCN-IsBlank $base) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        }
    }
    if (_QCN-IsBlank $base) { return '' }

    if (Get-Command -Name 'Get-QCLaneQcPdfExpectedName' -ErrorAction SilentlyContinue) {
        $expected = Get-QCLaneQcPdfExpectedName -SheetBaseName $base -ProcessType $resolvedProcessType -Config $Config
        if (-not (_QCN-IsBlank $expected)) { return $expected.Trim() }
    }
    return ($base + '-prod.pdf')
}

function _QCN-GetQcPdfLookupCandidateNames {
    param(
        [string]$PrimaryName = '',
        [string]$SheetStem = '',
        [string]$ProcessType = '',
        [hashtable]$Config = $null,
        [hashtable]$Event = $null
    )

    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ordered = [System.Collections.Generic.List[string]]::new()
    function _AddCandidate([string]$Candidate) {
        if (_QCN-IsBlank $Candidate) { return }
        $trimmed = $Candidate.Trim()
        if ($seen.Add($trimmed)) { [void]$ordered.Add($trimmed) }
    }

    _AddCandidate $PrimaryName
    _AddCandidate (_QCN-NormalizeQcPdfDocumentName -DocumentName $PrimaryName -SheetStem $SheetStem `
        -ProcessType $ProcessType -Config $Config -Event $Event)

    $base = [string]$SheetStem
    if (_QCN-IsBlank $base -and -not (_QCN-IsBlank $PrimaryName)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $base = [string](Get-PWSheetStemFromDocumentName -DocumentName $PrimaryName)
        }
        if (_QCN-IsBlank $base) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($PrimaryName)
        }
    }
    if (-not (_QCN-IsBlank $base)) {
        foreach ($suffix in @('prod', 'chk', 'rev')) {
            _AddCandidate ($base + '-' + $suffix + '.pdf')
        }
        if (Get-Command -Name 'Get-QCLaneQcPdfExpectedName' -ErrorAction SilentlyContinue) {
            foreach ($pt in @($ProcessType, 'production', 'review', 'check')) {
                if (_QCN-IsBlank $pt) { continue }
                _AddCandidate (Get-QCLaneQcPdfExpectedName -SheetBaseName $base -ProcessType $pt -Config $Config)
            }
        }
    }
    return @($ordered.ToArray())
}

function _QCN-LookupQcPdfGuidFromSheetDocuments {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$QcPdfName = '',
        [string]$SheetStem = ''
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    if (_QCN-IsBlank $FolderPath) { return '' }
    try {
        if (-not (_QCN-IsBlank $QcPdfName)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sd.document_guid
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE sd.document_role = 'qc_pdf'
  AND sp.folder_path = @folderPath
  AND LOWER(sd.document_name) = LOWER(@qcPdfName)
ORDER BY sp.last_updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; qcPdfName = [string]$QcPdfName }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $g = if ($res.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].document_guid }
                if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
            }
        }
        if (-not (_QCN-IsBlank $SheetStem)) {
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 sd.document_guid
FROM sheet_documents sd
INNER JOIN sheet_packages sp ON sp.sheet_package_id = sd.sheet_package_id
WHERE sd.document_role = 'qc_pdf'
  AND sp.folder_path = @folderPath
  AND LOWER(sp.sheet_stem) = LOWER(@sheetStem)
ORDER BY sp.last_updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; sheetStem = [string]$SheetStem }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $g2 = if ($res2.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$res2.Data.table.Rows[0].document_guid }
                if (-not (_QCN-IsBlank $g2)) { return $g2.Trim() }
            }
        }
    } catch { }
    return ''
}

function _QCN-NormalizeNotificationProcessType {
    param(
        [string]$ProcessType = '',
        [hashtable]$Event = $null
    )
    if (-not (_QCN-IsBlank $ProcessType)) {
        if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
            $n = Normalize-QCProcessType -ProcessType $ProcessType
            if ($n) { return $n }
        }
    }
    if ($Event -and (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue)) {
        $n = Normalize-QCProcessType -ProcessType ([string]$Event.qcProcessType) -ReviewType ([string]$Event.qcReviewType) -Context $Event
        if ($n) { return $n }
    }
    return 'production'
}

function _QCN-LookupLaneQcPdfGuidFromPackageQcPdfs {
    param(
        [hashtable]$Config,
        [string]$SheetPackageId = '',
        [string]$FolderPath = '',
        [string]$QcPdfName = '',
        [string]$ProcessType = ''
    )
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    $lane = _QCN-NormalizeNotificationProcessType -ProcessType $ProcessType
    try {
        if (-not (_QCN-IsBlank $SheetPackageId)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid
FROM sheet_package_qc_pdfs
WHERE sheet_package_id = @sheetPackageId
  AND qc_process_type = @qcProcessType
  AND is_active = 1
ORDER BY updated_at DESC
"@ -Parameters @{ sheetPackageId = [string]$SheetPackageId.Trim(); qcProcessType = $lane }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $g = if ($res.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].document_guid }
                if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
            }
        }
        if ((-not (_QCN-IsBlank $FolderPath)) -and (-not (_QCN-IsBlank $QcPdfName))) {
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid
FROM sheet_package_qc_pdfs
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@qcPdfName)
  AND qc_process_type = @qcProcessType
  AND is_active = 1
ORDER BY updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; qcPdfName = [string]$QcPdfName; qcProcessType = $lane }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $g2 = if ($res2.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$res2.Data.table.Rows[0].document_guid }
                if (-not (_QCN-IsBlank $g2)) { return $g2.Trim() }
            }
        }
    } catch { }
    return ''
}

function _QCN-LookupQcPdfGuidBySheetPackageId {
    param(
        [hashtable]$Config,
        [string]$SheetPackageId = '',
        [string]$ProcessType = '',
        [hashtable]$Event = $null
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    if (_QCN-IsBlank $SheetPackageId) { return '' }
    $lane = _QCN-NormalizeNotificationProcessType -ProcessType $ProcessType -Event $Event
    $fromLaneTable = _QCN-LookupLaneQcPdfGuidFromPackageQcPdfs -Config $Config -SheetPackageId $SheetPackageId -ProcessType $lane
    if (-not (_QCN-IsBlank $fromLaneTable)) { return $fromLaneTable }
    try {
        $col = switch ($lane) {
            'check' { 'qc_chk_pdf_guid' }
            'review' { 'qc_rev_pdf_guid' }
            default { 'qc_pdf_guid' }
        }
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 $col AS lane_guid
FROM sheet_packages
WHERE sheet_package_id = @sheetPackageId
  AND $col IS NOT NULL
"@ -Parameters @{ sheetPackageId = [string]$SheetPackageId.Trim() }
        if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
            $g = if ($res.Data.table.Rows[0].lane_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].lane_guid }
            if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
        }
    } catch { }
    return ''
}

function _QCN-LookupQcPdfGuidFromSheetPackages {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$QcPdfName = '',
        [string]$SheetStem = '',
        [string]$ProcessType = '',
        [hashtable]$Event = $null
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    if (_QCN-IsBlank $FolderPath) { return '' }
    $lane = _QCN-NormalizeNotificationProcessType -ProcessType $ProcessType -Event $Event
    $fromLaneTable = _QCN-LookupLaneQcPdfGuidFromPackageQcPdfs -Config $Config -FolderPath $FolderPath `
        -QcPdfName $QcPdfName -ProcessType $lane
    if (-not (_QCN-IsBlank $fromLaneTable)) { return $fromLaneTable }
    $guidCol = switch ($lane) {
        'check' { 'qc_chk_pdf_guid' }
        'review' { 'qc_rev_pdf_guid' }
        default { 'qc_pdf_guid' }
    }
    $nameCol = switch ($lane) {
        'check' { 'qc_chk_pdf_name' }
        'review' { 'qc_rev_pdf_name' }
        default { 'qc_pdf_name' }
    }
    try {
        if (-not (_QCN-IsBlank $QcPdfName)) {
            $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 $guidCol AS lane_guid
FROM sheet_packages
WHERE folder_path = @folderPath
  AND $guidCol IS NOT NULL
  AND LOWER($nameCol) = LOWER(@qcPdfName)
ORDER BY last_updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; qcPdfName = [string]$QcPdfName }
            if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
                $g = if ($res.Data.table.Rows[0].lane_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].lane_guid }
                if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
            }
        }
        if (-not (_QCN-IsBlank $SheetStem)) {
            $res2 = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 $guidCol AS lane_guid
FROM sheet_packages
WHERE folder_path = @folderPath
  AND $guidCol IS NOT NULL
  AND LOWER(sheet_stem) = LOWER(@sheetStem)
ORDER BY last_updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; sheetStem = [string]$SheetStem }
            if ($res2.IsSuccess -and $res2.Data.table -and $res2.Data.table.Rows.Count -gt 0) {
                $g2 = if ($res2.Data.table.Rows[0].lane_guid -is [DBNull]) { '' } else { [string]$res2.Data.table.Rows[0].lane_guid }
                if (-not (_QCN-IsBlank $g2)) { return $g2.Trim() }
            }
        }
    } catch { }
    return ''
}

function _QCN-LookupQcPdfGuidInSheetIndex {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$QcPdfName = ''
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    if ((_QCN-IsBlank $FolderPath) -or (_QCN-IsBlank $QcPdfName)) { return '' }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_guid
FROM sheet_index
WHERE folder_path = @folderPath
  AND LOWER(document_name) = LOWER(@qcPdfName)
ORDER BY
  CASE WHEN sheet_package_id IS NOT NULL THEN 0 ELSE 1 END,
  last_updated_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath; qcPdfName = [string]$QcPdfName }
        if ($res.IsSuccess -and $res.Data.table -and $res.Data.table.Rows.Count -gt 0) {
            $g = if ($res.Data.table.Rows[0].document_guid -is [DBNull]) { '' } else { [string]$res.Data.table.Rows[0].document_guid }
            if (-not (_QCN-IsBlank $g)) { return $g.Trim() }
        }
    } catch { }
    return ''
}

function _QCN-TestPwDocumentGuidMatchesName {
    param(
        [string]$DocumentGuid,
        [string]$ExpectedName
    )

    if ((_QCN-IsBlank $DocumentGuid) -or (_QCN-IsBlank $ExpectedName)) { return $false }
    if (-not (Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue)) { return $false }
    try {
        $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs @([string]$DocumentGuid.Trim()) -ErrorAction SilentlyContinue)
        if ($docs.Count -le 0) { return $false }
        $actual = [string](_QCN-GetProp -Object $docs[0] -Names @('Name', 'DocumentName', 'FileName'))
        if (_QCN-IsBlank $actual) { return $false }
        return ($actual.Trim().Equals([string]$ExpectedName.Trim(), [System.StringComparison]::OrdinalIgnoreCase))
    } catch { }
    return $false
}

function _QCN-AcceptGuidForQcPdfLink {
    param(
        [string]$DocumentGuid,
        [string]$ExpectedName
    )

    if (_QCN-IsBlank $DocumentGuid) { return $false }
    if (_QCN-IsBlank $ExpectedName) { return $true }
    return [bool](_QCN-TestPwDocumentGuidMatchesName -DocumentGuid $DocumentGuid -ExpectedName $ExpectedName)
}

function _QCN-ExtractPwLinkDocumentGuid {
    param([string]$Url = '')

    if (_QCN-IsBlank $Url) { return '' }
    $m = [regex]::Match([string]$Url, '(?:[?&]objectId=|objectid=)([0-9a-fA-F-]{36})', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($m.Success) { return [string]$m.Groups[1].Value.Trim() }
    return ''
}

function _QCN-WriteQcPdfNotificationGuidResolutionLog {
    param(
        [hashtable]$Config = $null,
        [string]$TriggerDocumentGuid = '',
        [string]$ResolvedQcPdfGuid = '',
        [string]$NotificationDocumentGuid = '',
        [string]$LinkDocumentGuid = '',
        [string]$SheetPackageId = '',
        [string]$QcReviewType = '',
        [string]$ResolutionSource = '',
        [string]$LinkResolutionSource = '',
        [hashtable]$Event = $null,
        [hashtable]$Job = $null
    )

    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    $data = @{
        trigger_document_guid = [string]$TriggerDocumentGuid
        resolved_qc_pdf_guid = [string]$ResolvedQcPdfGuid
        notification_document_guid = [string]$NotificationDocumentGuid
        link_document_guid = [string]$LinkDocumentGuid
        sheet_package_id = [string]$SheetPackageId
        qc_review_type = [string]$QcReviewType
        resolution_source = [string]$ResolutionSource
        link_resolution_source = [string]$LinkResolutionSource
    }
    if ($Event) {
        if ($Event.ContainsKey('documentName')) { $data['document_name'] = [string]$Event.documentName }
        if ($Event.ContainsKey('eventType')) { $data['event_type'] = [string]$Event.eventType }
        if ($Event.ContainsKey('currentState')) { $data['current_state'] = [string]$Event.currentState }
        if ($Event.ContainsKey('dedupeKey')) { $data['dedupe_key'] = [string]$Event.dedupeKey }
    }
    if ($Job -and $Job.ContainsKey('id')) { $data['job_id'] = [string]$Job.id }
    Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_QC_PDF_GUID_RESOLVED' `
        -Message 'QC PDF notification GUID resolution for email link.' -Data $data | Out-Null
}

function _QCN-ResolveLiveQcPdfDocumentGuid {
    <#
    Resolves the current ProjectWise GUID for a sheet's lane QC PDF (*-prod/-chk/-rev.pdf).
    Prefers package member tables over PW search and stale sheet_index rows.
    #>
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$QcPdfName = '',
        [string]$SheetStem = '',
        [string]$HintGuid = '',
        [string]$SourceSheetPdfName = '',
        [string]$SheetPackageId = '',
        [ref]$ResolutionSource = $null,
        [hashtable]$Event = $null
    )

    $resolved = _QCN-ResolveLiveQcPdfDocumentGuidResult -Config $Config -FolderPath $FolderPath `
        -QcPdfName $QcPdfName -SheetStem $SheetStem -HintGuid $HintGuid -SourceSheetPdfName $SourceSheetPdfName `
        -SheetPackageId $SheetPackageId -Event $Event
    if ($null -ne $ResolutionSource) { $ResolutionSource.Value = [string]$resolved.resolutionSource }
    return [string]$resolved.documentGuid
}

function _QCN-ResolveLiveQcPdfDocumentGuidResult {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$QcPdfName = '',
        [string]$SheetStem = '',
        [string]$HintGuid = '',
        [string]$SourceSheetPdfName = '',
        [string]$SheetPackageId = '',
        [hashtable]$Event = $null
    )

    $processType = if ($Event -and $Event.qcProcessType) { [string]$Event.qcProcessType } else { '' }
    $qcName = _QCN-NormalizeQcPdfDocumentName -DocumentName $QcPdfName -SheetStem $SheetStem `
        -ProcessType $processType -Config $Config -Event $Event
    $stem = if (-not (_QCN-IsBlank $SheetStem)) { [string]$SheetStem } else {
        if (-not (_QCN-IsBlank $qcName)) {
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                [string](Get-PWSheetStemFromDocumentName -DocumentName $qcName)
            } else {
                [System.IO.Path]::GetFileNameWithoutExtension($qcName)
            }
        } else { '' }
    }
    if (_QCN-IsBlank $qcName) {
        $guid = if (-not (_QCN-IsBlank $HintGuid)) { [string]$HintGuid.Trim() } else { '' }
        $source = if (-not (_QCN-IsBlank $guid)) { 'hint_guid_no_qc_name' } else { '' }
        return @{ documentGuid = $guid; resolutionSource = $source }
    }

    $candidateNames = _QCN-GetQcPdfLookupCandidateNames -PrimaryName $qcName -SheetStem $stem `
        -ProcessType $processType -Config $Config -Event $Event

    $trustedLaneGuid = _QCN-ResolveTrustedLaneDocumentGuid -Event $Event
    if (-not (_QCN-IsBlank $trustedLaneGuid)) {
        foreach ($candidate in @($candidateNames)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $trustedLaneGuid -ExpectedName $candidate) {
                return @{ documentGuid = $trustedLaneGuid; resolutionSource = 'trusted_lane_metadata_guid' }
            }
        }
        if ($qcName -match '(?i)-(prod|chk|rev)\.pdf$') {
            return @{ documentGuid = $trustedLaneGuid; resolutionSource = 'trusted_lane_metadata_guid' }
        }
    }

    $packageGuid = ''
    if (-not (_QCN-IsBlank $SheetPackageId)) {
        $packageGuid = _QCN-LookupQcPdfGuidBySheetPackageId -Config $Config -SheetPackageId $SheetPackageId `
            -ProcessType $processType -Event $Event
        if (-not (_QCN-IsBlank $packageGuid)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $packageGuid -ExpectedName $qcName) {
                return @{ documentGuid = $packageGuid; resolutionSource = 'sheet_package_qc_pdfs' }
            }
            return @{ documentGuid = $packageGuid; resolutionSource = 'sheet_package_qc_pdfs' }
        }
    }

    foreach ($candidate in @($candidateNames)) {
        if (_QCN-IsBlank $candidate) { continue }

        $laneGuid = _QCN-LookupLaneQcPdfGuidFromPackageQcPdfs -Config $Config -FolderPath $FolderPath `
            -QcPdfName $candidate -ProcessType $processType
        if (-not (_QCN-IsBlank $laneGuid)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $laneGuid -ExpectedName $candidate) {
                return @{ documentGuid = $laneGuid; resolutionSource = 'sheet_package_qc_pdfs' }
            }
            return @{ documentGuid = $laneGuid; resolutionSource = 'sheet_package_qc_pdfs' }
        }

        $pkgGuid = _QCN-LookupQcPdfGuidFromSheetPackages -Config $Config -FolderPath $FolderPath `
            -QcPdfName $candidate -SheetStem $stem -ProcessType $processType -Event $Event
        if (-not (_QCN-IsBlank $pkgGuid)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $pkgGuid -ExpectedName $candidate) {
                return @{ documentGuid = $pkgGuid; resolutionSource = 'sheet_packages_lane_alias' }
            }
            return @{ documentGuid = $pkgGuid; resolutionSource = 'sheet_packages_lane_alias' }
        }

        if (-not (_QCN-IsBlank $FolderPath)) {
            $pwGuid = _QCN-TryResolveQcPdfGuidFromPwSearch -Config $Config -FolderPath $FolderPath -QcPdfName $candidate
            if (-not (_QCN-IsBlank $pwGuid)) {
                if (_QCN-TestPwDocumentGuidMatchesName -DocumentGuid $pwGuid -ExpectedName $candidate) {
                    return @{ documentGuid = $pwGuid; resolutionSource = 'pw_search' }
                }
                return @{ documentGuid = $pwGuid; resolutionSource = 'pw_search' }
            }
        }

        $idxGuid = _QCN-LookupQcPdfGuidInSheetIndex -Config $Config -FolderPath $FolderPath -QcPdfName $candidate
        if (-not (_QCN-IsBlank $idxGuid)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $idxGuid -ExpectedName $candidate) {
                return @{ documentGuid = $idxGuid; resolutionSource = 'sheet_index' }
            }
        }

        $pkgDocGuid = _QCN-LookupQcPdfGuidFromSheetDocuments -Config $Config -FolderPath $FolderPath `
            -QcPdfName $candidate -SheetStem $stem
        if (-not (_QCN-IsBlank $pkgDocGuid)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $pkgDocGuid -ExpectedName $candidate) {
                return @{ documentGuid = $pkgDocGuid; resolutionSource = 'sheet_documents' }
            }
            return @{ documentGuid = $pkgDocGuid; resolutionSource = 'sheet_documents' }
        }
    }

    $srcPdf = [string]$SourceSheetPdfName
    if (_QCN-IsBlank $srcPdf) {
        if (-not (_QCN-IsBlank $stem)) { $srcPdf = $stem + '.pdf' }
    }
    $linkedGuid = _QCN-ResolveQcPdfGuidFromSheetIndex -Config $Config -FolderPath $FolderPath `
        -SourceDocumentName $srcPdf -DocumentGuid $HintGuid
    if (-not (_QCN-IsBlank $linkedGuid)) {
        foreach ($candidate in @($candidateNames)) {
            if (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $linkedGuid -ExpectedName $candidate) {
                return @{ documentGuid = $linkedGuid; resolutionSource = 'sheet_index_qc_pdf_guid' }
            }
        }
    }

    foreach ($candidate in @($candidateNames)) {
        if (_QCN-TestPwDocumentGuidMatchesName -DocumentGuid $HintGuid -ExpectedName $candidate) {
            return @{ documentGuid = [string]$HintGuid.Trim(); resolutionSource = 'hint_guid_pw_verified' }
        }
    }
    return @{ documentGuid = ''; resolutionSource = '' }
}

function _QCN-ResolveNotificationRoleEmails {
    param(
        [object]$Document,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null,
        [hashtable]$RoleOverrides = $null,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$DocumentGuid = ''
    )

    $attr = _QCN-ToHashtable $Settings.attributes
    if (-not $attr) { $attr = @{} }
    $reviewerField = if ($attr.reviewerEmailField) { [string]$attr.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attr.designerEmailField) { [string]$attr.designerEmailField } else { 'EM_Designer_Email' }
    $checkerField = if ($attr.checkerEmailField) { [string]$attr.checkerEmailField } else { 'EM_Checker_Email' }
    $qcDesignerField = _QCN-GetQcDesignerEmailAttributeName -Config $Config
    $qcReviewerField = _QCN-GetQcReviewerEmailAttributeName -Config $Config
    $qcCheckerField = _QCN-GetQcCheckerEmailAttributeName -Config $Config

    $folderPath = [string]$FolderPath
    $sourceName = [string]$SourceDocumentName
    if ($Job) {
        if (_QCN-IsBlank $folderPath) {
            $folderPath = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder','folderPath','incomingFolderPath'))
        }
        if (_QCN-IsBlank $sourceName) {
            $sourceName = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName','sourceDocumentName','incomingDocName'))
            if (_QCN-IsBlank $sourceName) {
                $sp = [string](_QCN-GetJobValue -Job $Job -Keys @('sourcePath'))
                if (-not (_QCN-IsBlank $sp) -and $sp -match '\\') { $sourceName = [System.IO.Path]::GetFileName($sp) }
            }
        }
    }
    if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetProp -Object $Document -Names @('FolderPath','folderPath')) }
    if (_QCN-IsBlank $sourceName) { $sourceName = [string](_QCN-GetProp -Object $Document -Names @('Name','DocumentName','FileName')) }
    if (_QCN-IsBlank $DocumentGuid) { $DocumentGuid = [string](_QCN-GetProp -Object $Document -Names @('DocumentGUID','DocumentGuid','GUID','Id')) }
    if ($Job) {
        $roleFromJob = [string](_QCN-GetJobValue -Job $Job -Keys @('roleSourceDocumentName'))
        if (-not (_QCN-IsBlank $roleFromJob)) { $sourceName = $roleFromJob }
        elseif ($Job.metadata) {
            $mdRole = _QCN-ToHashtable $Job.metadata
            if ($mdRole -and $mdRole.roleSourceDocumentName) { $sourceName = [string]$mdRole.roleSourceDocumentName }
        }
    }
    if (-not (_QCN-IsBlank $sourceName) -and (Test-QCIsQcPdfDocumentName -DocumentName $sourceName)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            try {
                $stem = Get-PWSheetStemFromDocumentName -DocumentName $sourceName
                if (-not (_QCN-IsBlank $stem)) { $sourceName = $stem + '.pdf' }
            } catch { }
        } else {
            $sourceName = [System.IO.Path]::GetFileNameWithoutExtension($sourceName) + '.pdf'
        }
    }

    $designer = ''
    $reviewer = ''
    $checker = ''
    if ($Config -and -not (_QCN-IsBlank $folderPath) -and -not (_QCN-IsBlank $sourceName)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $folderPath -SourceDocumentName $sourceName -Config $Config
            if ($pw.found) {
                if (-not (_QCN-IsBlank $pw.designerEmail)) { $designer = [string]$pw.designerEmail }
                if (-not (_QCN-IsBlank $pw.reviewerEmail)) { $reviewer = [string]$pw.reviewerEmail }
                if (-not (_QCN-IsBlank $pw.checkerEmail)) { $checker = [string]$pw.checkerEmail }
            }
        }
    }
    if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $designerField) }
    if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $reviewerField) }
    if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $checkerField) }
    if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcDesignerField) }
    if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcReviewerField) }
    if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcCheckerField) }
    if ($Job) {
        if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetJobValue -Job $Job -Keys @('designerEmail','qcDesignerEmail')) }
        if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetJobValue -Job $Job -Keys @('reviewerEmail','qcReviewerEmail')) }
        if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetJobValue -Job $Job -Keys @('checkerEmail','qcCheckerEmail')) }
    }
    if ($RoleOverrides) {
        if (_QCN-IsBlank $designer -and $RoleOverrides.ContainsKey('designerEmail')) { $designer = [string]$RoleOverrides.designerEmail }
        if (_QCN-IsBlank $reviewer -and $RoleOverrides.ContainsKey('reviewerEmail')) { $reviewer = [string]$RoleOverrides.reviewerEmail }
        if (_QCN-IsBlank $checker -and $RoleOverrides.ContainsKey('checkerEmail')) { $checker = [string]$RoleOverrides.checkerEmail }
    }

    if ($Config -and ((-not (_QCN-IsBlank $DocumentGuid)) -or ((-not (_QCN-IsBlank $folderPath)) -and (-not (_QCN-IsBlank $sourceName))))) {
        $needIndex = (_QCN-IsBlank $designer) -or (_QCN-IsBlank $reviewer) -or (_QCN-IsBlank $checker)
        if ($needIndex) {
            $idx = _QCN-GetRoleEmailsFromSheetIndex -Config $Config -DocumentGuid $DocumentGuid `
                -FolderPath $folderPath -SourceDocumentName $sourceName
            if ($idx) {
                if (_QCN-IsBlank $designer) { $designer = [string]$idx.designerEmail }
                if (_QCN-IsBlank $reviewer) { $reviewer = [string]$idx.reviewerEmail }
                if (_QCN-IsBlank $checker) { $checker = [string]$idx.checkerEmail }
            }
        }
    }

    $reviewType = _QCN-ResolveNotificationReviewType -Document $Document -Settings $Settings -Config $Config -Job $Job `
        -FolderPath $folderPath -SourceName $sourceName
    if (_QCN-TestNotificationIndependentCheckReviewType -ReviewType $reviewType -Config $Config) {
        $reviewer = _QCN-ApplyIndependentCheckReviewerEmail -Document $Document -CheckerEmail $checker `
            -CheckerField $checkerField -QcCheckerField $qcCheckerField
        if (-not (_QCN-IsBlank $reviewer) -and (_QCN-IsBlank $checker)) { $checker = $reviewer }
    }

    return @{ designerEmail = $designer; reviewerEmail = $reviewer; checkerEmail = $checker }
}

function _QCN-ResolveNotificationSheetPackageId {
    param(
        [hashtable]$Job = $null,
        [hashtable]$Event = $null
    )

    if ($Job -and ($Job.metadata -is [hashtable])) {
        $md = _QCN-ToHashtable $Job.metadata
        if ($md -and $md.ContainsKey('sheetPackageId') -and -not (_QCN-IsBlank $md.sheetPackageId)) {
            return [string]$md.sheetPackageId.Trim()
        }
    }
    if ($Job -and $Job.ContainsKey('sheetPackageId') -and -not (_QCN-IsBlank $Job.sheetPackageId)) {
        return [string]$Job.sheetPackageId.Trim()
    }
    if ($Event -and $Event.ContainsKey('sheetPackageId') -and -not (_QCN-IsBlank $Event.sheetPackageId)) {
        return [string]$Event.sheetPackageId.Trim()
    }
    return ''
}

function _QCN-ResolveQcPdfNotificationTarget {
    param(
        [object]$Document,
        [hashtable]$Config,
        [hashtable]$Job,
        [string]$DocumentName,
        [string]$DocumentGuid,
        [string]$DocumentPath
    )

    $sheetPackageId = _QCN-ResolveNotificationSheetPackageId -Job $Job
    $out = @{
        document = $Document
        documentName = $DocumentName
        documentGuid = $DocumentGuid
        documentPath = $DocumentPath
        resolutionSource = ''
    }
    $folderPath = $DocumentPath
    if ($folderPath -and ($folderPath -match '\\')) {
        $folderPath = [System.IO.Path]::GetDirectoryName($folderPath)
    }
    if ($Job -and (_QCN-IsBlank $folderPath)) {
        $folderPath = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath'))
    }
    if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetProp -Object $Document -Names @('FolderPath', 'folderPath')) }

    $trustedLaneGuid = _QCN-ResolveTrustedLaneDocumentGuid -Job $Job
    if (-not (_QCN-IsBlank $trustedLaneGuid)) {
        $trustedLaneName = $DocumentName
        if ($Job -and $Job.metadata) {
            $mdTrusted = _QCN-ToHashtable $Job.metadata
            if ($mdTrusted -and $mdTrusted.expectedLanePdfName) { $trustedLaneName = [string]$mdTrusted.expectedLanePdfName }
            elseif ($mdTrusted -and $mdTrusted.notificationLaneDocumentName) { $trustedLaneName = [string]$mdTrusted.notificationLaneDocumentName }
        }
        if (-not (_QCN-IsBlank $trustedLaneName)) {
            $out.documentName = $trustedLaneName
            $out.documentGuid = $trustedLaneGuid
            if (-not (_QCN-IsBlank $folderPath)) {
                $out.documentPath = $folderPath.TrimEnd('\') + '\' + $trustedLaneName
            }
            $out.resolutionSource = 'trusted_lane_metadata_guid'
            return $out
        }
    }

    $qcProcessType = ''
    if ($Job) {
        $qcProcessType = [string](_QCN-GetJobValue -Job $Job -Keys @('qcProcessType'))
        if (_QCN-IsBlank $qcProcessType -and $Job.metadata) {
            $md = _QCN-ToHashtable $Job.metadata
            if ($md -and $md.qcProcessType) { $qcProcessType = [string]$md.qcProcessType }
            if (_QCN-IsBlank $qcProcessType -and $md -and $md.expectedLanePdfName) {
                $laneName = [string]$md.expectedLanePdfName
                $out.documentName = $laneName
                $out.documentPath = if ($folderPath) { ($folderPath.TrimEnd('\') + '\' + $laneName) } else { $DocumentPath }
                $out.resolutionSource = 'job_expected_lane_pdf'
                $guidOk = (-not (_QCN-IsBlank $out.documentGuid)) -and `
                    (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $out.documentGuid -ExpectedName $laneName)
                if (-not $guidOk -and $Config -and -not (_QCN-IsBlank $folderPath)) {
                    $pwGuid = _QCN-TryResolveQcPdfGuidFromPwSearch -Config $Config -FolderPath $folderPath -QcPdfName $laneName
                    if (-not (_QCN-IsBlank $pwGuid)) {
                        $out.documentGuid = $pwGuid.Trim()
                        $out.resolutionSource = 'job_expected_lane_pdf_pw_search'
                    }
                }
                return $out
            }
        }
    }

    if ($DocumentName -match '(?i)-(prod|chk|rev)\.pdf$') {
        $out.documentName = $DocumentName
        if (-not (_QCN-IsBlank $folderPath)) {
            $out.documentPath = $folderPath.TrimEnd('\') + '\' + $DocumentName
        }
        $out.resolutionSource = 'lane_pdf_name'
        $guidOk = (-not (_QCN-IsBlank $out.documentGuid)) -and `
            (_QCN-AcceptGuidForQcPdfLink -DocumentGuid $out.documentGuid -ExpectedName $DocumentName)
        if (-not $guidOk -and $Config -and -not (_QCN-IsBlank $folderPath)) {
            $pwGuid = _QCN-TryResolveQcPdfGuidFromPwSearch -Config $Config -FolderPath $folderPath -QcPdfName $DocumentName
            if (-not (_QCN-IsBlank $pwGuid)) {
                $out.documentGuid = $pwGuid.Trim()
                $out.resolutionSource = 'lane_pdf_pw_search'
            }
        }
        return $out
    }

    $triggerName = $DocumentName
    if ($Job) {
        if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder','folderPath','incomingFolderPath')) }
        if (_QCN-IsBlank $triggerName) { $triggerName = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName','sourceDocumentName','incomingDocName')) }
    }
    if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetProp -Object $Document -Names @('FolderPath','folderPath')) }
    if (_QCN-IsBlank $triggerName) { $triggerName = [string](_QCN-GetProp -Object $Document -Names @('Name','DocumentName')) }

    if ($Config -and -not (_QCN-IsBlank $folderPath) -and -not (_QCN-IsBlank $triggerName)) {
        $guid = $DocumentGuid
        if (_QCN-IsBlank $guid) { $guid = [string](_QCN-GetProp -Object $Document -Names @('DocumentGUID','DocumentGuid','GUID')) }
        if (Get-Command -Name 'Get-PWAssociatedSheetMembers' -ErrorAction SilentlyContinue) {
            try {
                $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $folderPath -DocumentName $triggerName -DocumentGuid $guid)
                foreach ($m in $members) {
                    $dn = [string]$m.documentName
                    if ($dn -match '(?i)-(prod|chk|rev)\.pdf$') {
                        if (-not (_QCN-IsBlank $qcProcessType) -and (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue)) {
                            $memberLane = Get-PWQcPdfLaneFromDocumentName -DocumentName $dn
                            if ($memberLane -and $memberLane -ne (Normalize-QCProcessType -ProcessType $qcProcessType)) { continue }
                        }
                        $out.documentName = $dn
                        $resolvedGuid = [string]$m.documentGuid
                        if (-not (_QCN-IsBlank $resolvedGuid)) {
                            $out.documentGuid = $resolvedGuid.Trim()
                        }
                        $out.documentPath = if ($folderPath) { ($folderPath.TrimEnd('\') + '\' + $dn) } else { $DocumentPath }
                        if ($m.document) { $out.document = $m.document }
                        $out.resolutionSource = 'associated_lane_pdf'
                        return $out
                    }
                }
            } catch { }
        }
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($triggerName)
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $resolvedStem = Get-PWSheetStemFromDocumentName -DocumentName $triggerName
            if (-not (_QCN-IsBlank $resolvedStem)) { $stem = $resolvedStem }
        }
        $normalized = _QCN-NormalizeQcPdfDocumentName -DocumentName $triggerName -SheetStem $stem `
            -ProcessType $qcProcessType -Config $Config
        if (-not (_QCN-IsBlank $normalized)) {
            $out.documentName = $normalized
            $out.documentPath = if ($folderPath) { ($folderPath.TrimEnd('\') + '\' + $normalized) } else { $DocumentPath }
            $out.resolutionSource = 'expected_lane_pdf_name'
        }
    }
    return $out
}

function Resolve-QCNotificationRecipients {
    <#
    .SYNOPSIS
    Resolves reviewer/designer/cc email lists from a ProjectWise document object and notification settings.
    #>
    [CmdletBinding()]
    param(
        [object]$Document,
        [hashtable]$Settings,
        [string[]]$ToRoles = @(),
        [string[]]$CcRoles = @(),
        [string[]]$ExplicitCc = @(),
        [hashtable]$Config = $null,
        [hashtable]$Job = $null,
        [hashtable]$RoleOverrides = $null,
        [string]$FolderPath = '',
        [string]$SourceDocumentName = '',
        [string]$DocumentGuid = ''
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $attr = _QCN-ToHashtable $Settings.attributes
    if (-not $attr) { $attr = @{} }

    $reviewerField = if ($attr.reviewerEmailField) { [string]$attr.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attr.designerEmailField) { [string]$attr.designerEmailField } else { 'EM_Designer_Email' }
    $ccField = if ($attr.ccEmailField) { [string]$attr.ccEmailField } else { 'CcEmails' }

    $roles = _QCN-ResolveNotificationRoleEmails -Document $Document -Settings $Settings -Config $Config -Job $Job `
        -RoleOverrides $RoleOverrides -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -DocumentGuid $DocumentGuid
    $reviewers = _QCN-ParseEmailList $roles.reviewerEmail
    $designers = _QCN-ParseEmailList $roles.designerEmail
    $checkers = _QCN-ParseEmailList $roles.checkerEmail
    $ccFromAttr = _QCN-ParseEmailList (_QCN-GetAttributeValue -Document $Document -AttributeName $ccField)

    $to = [System.Collections.Generic.List[string]]::new()
    $cc = [System.Collections.Generic.List[string]]::new()

    foreach ($role in @($ToRoles)) {
        switch -Regex ($role) {
            '^reviewers?$' { foreach ($e in $reviewers) { if (-not $to.Contains($e)) { $to.Add($e) | Out-Null } } }
            '^designers?$' { foreach ($e in $designers) { if (-not $to.Contains($e)) { $to.Add($e) | Out-Null } } }
            '^checkers?$' { foreach ($e in $checkers) { if (-not $to.Contains($e)) { $to.Add($e) | Out-Null } } }
            default {
                foreach ($e in (_QCN-ParseEmailList $role)) {
                    if (-not $to.Contains($e)) { $to.Add($e) | Out-Null }
                }
            }
        }
    }
    foreach ($role in @($CcRoles)) {
        switch -Regex ($role) {
            '^reviewers?$' { foreach ($e in $reviewers) { if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null } } }
            '^designers?$' { foreach ($e in $designers) { if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null } } }
            '^checkers?$' { foreach ($e in $checkers) { if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null } } }
            default {
                foreach ($e in (_QCN-ParseEmailList $role)) {
                    if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null }
                }
            }
        }
    }
    foreach ($e in @($ExplicitCc)) { if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null } }
    foreach ($e in @($ccFromAttr)) { if (-not $cc.Contains($e)) { $cc.Add($e) | Out-Null } }

    foreach ($e in @($cc)) { if ($to.Contains($e)) { [void]$cc.Remove($e) } }

    return @{
        reviewers = @($reviewers)
        designers = @($designers)
        checkers = @($checkers)
        to = @($to)
        cc = @($cc)
    }
}

function _QCN-BuildRecipientKey {
    param(
        [string[]]$To = @(),
        [string[]]$Cc = @()
    )

    $set = [System.Collections.Generic.SortedSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($email in @($To) + @($Cc)) {
        if (_QCN-IsBlank $email) { continue }
        [void]$set.Add(([string]$email).Trim().ToLowerInvariant())
    }
    if ($set.Count -eq 0) { return '' }
    return ('recipients:' + ((@($set)) -join ','))
}

function _QCN-IsPlaceholderNotificationSheetStem {
    param([string]$Stem)
    if (_QCN-IsBlank $Stem) { return $true }
    return $Stem.Equals('unknown-document', [System.StringComparison]::OrdinalIgnoreCase)
}

function _QCN-IsPlaceholderNotificationDocumentName {
    param([string]$DocumentName)
    if (_QCN-IsBlank $DocumentName) { return $true }
    $name = [System.IO.Path]::GetFileName([string]$DocumentName)
    if ($name -match '(?i)^unknown-document(-qc)?\.pdf$') { return $true }
    if ($name.Equals('unknown-document', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $false
}

function _QCN-ResolveNotificationDisplayDocumentName {
    param(
        [string]$DocumentName = '',
        [hashtable]$Job = $null,
        [hashtable]$Event = $null,
        [hashtable]$Config = $null
    )

    if ($Job) {
        $src = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
        if (-not (_QCN-IsBlank $src)) {
            if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
                $src = [string](Get-PWSheetStemFromDocumentName -DocumentName ($src + '.pdf'))
            } else {
                $src = [System.IO.Path]::GetFileNameWithoutExtension($src)
            }
            if (-not (_QCN-IsPlaceholderNotificationDocumentName -DocumentName $src)) { return $src }
        }
    }

    $stem = ''
    if ($Event) {
        if ($Event.ContainsKey('sheetStem') -and -not (_QCN-IsBlank $Event.sheetStem) `
                -and -not (_QCN-IsPlaceholderNotificationSheetStem -Stem ([string]$Event.sheetStem))) {
            $stem = [string]$Event.sheetStem
        }
        else {
            $stem = _QCN-NormalizeNotificationSheetStemForDedupe -Event $Event -Config $Config
        }
    }
    if (-not (_QCN-IsBlank $stem) -and -not (_QCN-IsPlaceholderNotificationSheetStem -Stem $stem)) {
        return $stem
    }

    $name = [string]$DocumentName
    if (-not (_QCN-IsBlank $name)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            $name = [string](Get-PWSheetStemFromDocumentName -DocumentName ($name + '.pdf'))
        } else {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($name)
        }
    }
    if (-not (_QCN-IsPlaceholderNotificationDocumentName -DocumentName $name)) { return $name }
    return $name
}

function _QCN-ResolveSheetStemFromFolderForDedupe {
    param(
        [hashtable]$Config,
        [string]$FolderPath = '',
        [string]$DocumentGuid = ''
    )

    if (_QCN-IsBlank $FolderPath) { return '' }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }

    try {
        if (-not (_QCN-IsBlank $DocumentGuid)) {
            $byGuid = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = [string]$DocumentGuid.Trim() }
            if ($byGuid.IsSuccess -and $byGuid.Data.table -and $byGuid.Data.table.Rows.Count -gt 0) {
                $dn = if ($byGuid.Data.table.Rows[0].document_name -is [DBNull]) { '' } else { [string]$byGuid.Data.table.Rows[0].document_name }
                if (-not (_QCN-IsBlank $dn) -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                    $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName $dn)
                    if (-not (_QCN-IsPlaceholderNotificationSheetStem -Stem $stem)) { return $stem }
                }
            }
            $byQcGuid = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name
FROM sheet_index
WHERE qc_pdf_guid = @docGuid
"@ -Parameters @{ docGuid = [string]$DocumentGuid.Trim() }
            if ($byQcGuid.IsSuccess -and $byQcGuid.Data.table -and $byQcGuid.Data.table.Rows.Count -gt 0) {
                $dn = if ($byQcGuid.Data.table.Rows[0].document_name -is [DBNull]) { '' } else { [string]$byQcGuid.Data.table.Rows[0].document_name }
                if (-not (_QCN-IsBlank $dn) -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                    $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName $dn)
                    if (-not (_QCN-IsPlaceholderNotificationSheetStem -Stem $stem)) { return $stem }
                }
            }
        }

        $byFolder = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name
FROM sheet_index
WHERE folder_path = @folderPath
  AND document_name LIKE '%.pdf'
  AND document_name NOT LIKE '%-prod.pdf'
  AND document_name NOT LIKE '%-chk.pdf'
  AND document_name NOT LIKE '%-rev.pdf'
  AND document_name NOT LIKE '%-qc.pdf'
ORDER BY last_audit_event_at DESC
"@ -Parameters @{ folderPath = [string]$FolderPath.Trim() }
        if ($byFolder.IsSuccess -and $byFolder.Data.table -and $byFolder.Data.table.Rows.Count -gt 0) {
            $dn = if ($byFolder.Data.table.Rows[0].document_name -is [DBNull]) { '' } else { [string]$byFolder.Data.table.Rows[0].document_name }
            if (-not (_QCN-IsBlank $dn) -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                return [string](Get-PWSheetStemFromDocumentName -DocumentName $dn)
            }
        }
    } catch { }

    return ''
}

function _QCN-NormalizeNotificationSheetStemForDedupe {
    param(
        [hashtable]$Event,
        [hashtable]$Config = $null
    )

    if (-not $Event) { return '' }
    $stem = if ($Event.ContainsKey('sheetStem') -and -not (_QCN-IsBlank $Event.sheetStem)) {
        [string]$Event.sheetStem
    } else { '' }

    if (-not (_QCN-IsBlank $stem) -and -not (_QCN-IsPlaceholderNotificationSheetStem -Stem $stem)) { return $stem }

    if ((_QCN-IsBlank $stem) -and $Event.documentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
        try { $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName ([string]$Event.documentName)) } catch { }
    }

    if (-not (_QCN-IsPlaceholderNotificationSheetStem -Stem $stem)) { return $stem }

    $folderPath = ''
    if ($Event.folderPath) { $folderPath = [string]$Event.folderPath }
    elseif ($Event.documentPath -and ([string]$Event.documentPath -match '\\')) {
        $folderPath = [System.IO.Path]::GetDirectoryName([string]$Event.documentPath)
    }

    $docGuid = if ($Event.documentGuid) { [string]$Event.documentGuid } else { '' }
    if ($Config -and -not (_QCN-IsBlank $folderPath)) {
        $resolved = _QCN-ResolveSheetStemFromFolderForDedupe -Config $Config -FolderPath $folderPath -DocumentGuid $docGuid
        if (-not (_QCN-IsBlank $resolved)) { return $resolved }
    }

    return $stem
}

function _QCN-NormalizeNotificationDedupePreviousState {
    <#
    Collapses stale sheet_index baselines (often empty) with the nominal prior workflow state
    so prepend writeback and audit stale-index paths share one notification dedupe key.
    Applies to every configured workflow notification target state, not only QC Complete.
    #>
    param(
        [string]$PreviousState = '',
        [string]$CurrentState = '',
        [hashtable]$Config = $null
    )

    $prev = ([string]$PreviousState).Trim()
    if ($prev.Length -gt 0) { return $prev }

    $curr = ([string]$CurrentState).Trim()
    if ($curr.Length -eq 0) { return $prev }

    $priorByTarget = _QCN-GetWorkflowNotificationPriorStateMap -Config $Config
    if ($null -ne $priorByTarget -and $priorByTarget.ContainsKey($curr)) {
        $nominal = [string]$priorByTarget[$curr]
        if (-not (_QCN-IsBlank $nominal)) { return $nominal }
    }
    return $prev
}

function _QCN-GetWorkflowNotificationPriorStateMap {
    param([hashtable]$Config = $null)

    $map = @{
        'Verified' = 'Ready for Verification'
        'Ready for Verification' = 'Initiate Verification'
        'QC Complete' = 'QC Finalizing'
        'Initiate Verification' = 'Redlines Received'
        'QC Finalizing' = 'Ready for QC'
        'Originated' = 'Initiate Origination'
        'Redlines Received' = 'Originated'
        'Ready for QC' = 'QC Initiated'
        'QC Received' = 'In Development'
    }

    if ($Config -and $Config.qcWorkflow -and $Config.qcWorkflow.states) {
        $states = _QCN-ToHashtable $Config.qcWorkflow.states
        if ($states) {
            $resolved = @{}
            $pairs = @(
                @('complete', 'readyForVerification')
                @('readyForVerification', 'qcFinalizing')
                @('qcFinalizing', 'redlinesReceived')
                @('redlinesReceived', 'readyForQc')
                @('readyForQc', 'qcInitiated')
                @('qcReceived', 'production')
            )
            foreach ($pair in @($pairs)) {
                $targetKey = [string]$pair[0]
                $priorKey = [string]$pair[1]
                if (-not $states.ContainsKey($targetKey)) { continue }
                $targetName = [string]$states[$targetKey]
                if (_QCN-IsBlank $targetName) { continue }
                $priorName = if ($states.ContainsKey($priorKey)) { [string]$states[$priorKey] } else { '' }
                if (-not (_QCN-IsBlank $priorName)) {
                    $resolved[$targetName] = $priorName
                }
            }
            if ($resolved.Count -gt 0) { return $resolved }
        }
    }

    return $map
}

function Get-QCNotificationSheetTransitionKey {
    <#
    Stable sheet-package transition identity for notification dedupe.
    Intentionally ignores audit:{id} / transition:{id} echo events from sibling sync.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [hashtable]$Config = $null
    )

    $stem = _QCN-NormalizeNotificationSheetStemForDedupe -Event $Event -Config $Config
    if (_QCN-IsBlank $stem) { return '' }

    $curr = if ($Event.currentState) { [string]$Event.currentState } elseif ($Event.targetState) { [string]$Event.targetState } else { '' }
    $from = _QCN-NormalizeNotificationDedupePreviousState -PreviousState ([string]$Event.previousState) `
        -CurrentState $curr -Config $Config
    $to = ([string]$curr).Trim()
    if ($to.Length -eq 0) { return '' }

    return ('sheet:' + $stem.Trim().ToLowerInvariant() + '|from:' + $from.Trim().ToLowerInvariant() + '|to:' + $to.ToLowerInvariant())
}

function _QCN-ResolveNotificationTransitionSource {
    param(
        [hashtable]$Event,
        [hashtable]$Job = $null
    )

    if ($Event) {
        foreach ($key in @('transitionSource', 'notificationStateSource', 'triggerSource')) {
            if ($Event.ContainsKey($key) -and -not (_QCN-IsBlank $Event[$key])) {
                return ([string]$Event[$key]).Trim().ToLowerInvariant()
            }
        }
    }
    if ($Job) {
        $jobMd = _QCN-ToHashtable $Job.metadata
        if ($jobMd) {
            foreach ($key in @('transitionSource', 'notificationStateSource')) {
                if ($jobMd.ContainsKey($key) -and -not (_QCN-IsBlank $jobMd[$key])) {
                    return ([string]$jobMd[$key]).Trim().ToLowerInvariant()
                }
            }
        }
    }
    return ''
}

function _QCN-ResolveNotificationLogicalTransitionAnchor {
    <#
    Stable anchor for one logical sheet-package transition.
    Uses sheet|from|to plus cycle and audit id. transitionGroupId is intentionally excluded here
    because each Invoke-QCSheetGroupWorkflowTransition call mints a new group id; sibling triggers
    for the same audit event must share one dedupe anchor instead.
    #>
    param(
        [hashtable]$Event,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    $sheetKey = if ($Event -and $Event.sheetTransitionKey -and -not (_QCN-IsBlank $Event.sheetTransitionKey)) {
        [string]$Event.sheetTransitionKey
    } elseif ($Event) {
        Get-QCNotificationSheetTransitionKey -Event $Event -Config $Config
    } else { '' }

    $cycleId = _QCN-ResolveNotificationCycleId -Event $Event -Config $Config -Job $Job
    $cyclePart = if (-not (_QCN-IsBlank $cycleId)) { '|cycle:' + ([string]$cycleId).Trim() } else { '' }

    $auditId = _QCN-GetNotificationAuditEventId -Event $Event
    $auditPart = if ($null -ne $auditId -and $auditId -gt 0) { '|audit:' + [string]$auditId } else { '' }

    if (-not (_QCN-IsBlank $sheetKey)) { return $sheetKey + $cyclePart + $auditPart }

    if ($Event -and $Event.ContainsKey('transitionGroupId') -and -not (_QCN-IsBlank $Event.transitionGroupId)) {
        return 'transitionGroup:' + ([string]$Event.transitionGroupId).Trim().ToLowerInvariant()
    }
    if ($auditPart.Length -gt 0) { return $auditPart.TrimStart('|') }
    return ''
}

function Test-QCNotificationResultSent {
    [CmdletBinding()]
    param([object]$Result)

    if ($null -eq $Result) { return $false }
    $data = _QCN-ToHashtable $Result.Data
    if ($data -and $data.ContainsKey('skipped') -and [bool]$data.skipped) { return $false }
    $code = if ($Result.Code) { [string]$Result.Code } else { '' }
    if ($code -in @('QC_NOTIFICATION_SKIPPED_DUPLICATE', 'QC_NOTIFICATION_ENQUEUE_SKIPPED_DUPLICATE', 'QC_NOTIFICATION_ENQUEUED')) { return $false }
    if ($code -match '^QC_NOTIFICATION_(SENT|MOCK|GRAPH|JOB_OK)$') { return $true }
    if ($Result.IsSuccess -and $data -and $data.success -eq $true) { return $true }
    return $false
}

function _QCN-ResolveNotificationCycleId {
    param(
        [hashtable]$Event,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    $sources = [System.Collections.Generic.List[object]]::new()
    if ($Event) { $sources.Add($Event) | Out-Null }
    if ($Event) {
        $eventAttrs = _QCN-ToHashtable $Event.attributes
        if ($eventAttrs) { $sources.Add($eventAttrs) | Out-Null }
    }
    if ($Job) {
        $jobMd = _QCN-ToHashtable $Job.metadata
        if ($jobMd) {
            $sources.Add($jobMd) | Out-Null
            $jobAttrs = _QCN-ToHashtable $jobMd.attributes
            if ($jobAttrs) { $sources.Add($jobAttrs) | Out-Null }
        }
    }

    foreach ($source in @($sources)) {
        foreach ($key in @('cycleId', 'qcCycleId', 'QC_Cycle_ID')) {
            if (-not $source.ContainsKey($key)) { continue }
            $value = [string]$source[$key]
            if (_QCN-IsBlank $value) { continue }
            if ($value -match '^(audit:|transition:)') { continue }
            return $value.Trim()
        }
    }

    if ($Config -and (Get-Command -Name 'Get-QCSheetIndexCycle' -ErrorAction SilentlyContinue)) {
        try {
            $folderPath = ''
            $sheetStem = ''
            $documentGuid = ''
            if ($Event) {
                if ($Event.folderPath) { $folderPath = [string]$Event.folderPath }
                if ($Event.sheetStem) { $sheetStem = [string]$Event.sheetStem }
                if ($Event.documentGuid) { $documentGuid = [string]$Event.documentGuid }
            }
            if ((_QCN-IsBlank $sheetStem) -and $Event -and $Event.documentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                try { $sheetStem = [string](Get-PWSheetStemFromDocumentName -DocumentName ([string]$Event.documentName)) } catch { }
            }
            if ((_QCN-IsBlank $folderPath) -and $Job -and $Job.sourceFolder) { $folderPath = [string]$Job.sourceFolder }
            $cycle = Get-QCSheetIndexCycle -Config $Config -DocumentGuid $documentGuid -FolderPath $folderPath -SheetStem $sheetStem
            if ($cycle -and -not (_QCN-IsBlank $cycle.cycleId)) { return [string]$cycle.cycleId.Trim() }
        } catch { }
    }
    return ''
}

function Get-QCNotificationDedupeKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }

    $dedupe = _QCN-ToHashtable $Settings.dedupe
    # Default to one durable notification per logical sheet transition + recipient set.  A
    # transition_events id is intentionally *not* authoritative here: sibling sync can create
    # one transition row for the DGN, sheet PDF, and lane QC PDF for the same logical sheet action.
    $fields = if ($dedupe -and $dedupe.keyFields) { @($dedupe.keyFields) } else {
        @('sheetStem', 'documentGuid', 'previousState', 'currentState', 'transitionSource', 'logicalTransitionAnchor', 'recipientKey')
    }

    $folderPath = ''
    if ($Event.folderPath) { $folderPath = [string]$Event.folderPath }
    elseif ($Event.documentPath -and ([string]$Event.documentPath -match '\\')) {
        $folderPath = [System.IO.Path]::GetDirectoryName([string]$Event.documentPath)
    }
    if (-not (_QCN-IsBlank $folderPath)) { $Event['folderPath'] = $folderPath }

    $normalizedStem = _QCN-NormalizeNotificationSheetStemForDedupe -Event $Event -Config $Config
    if (-not (_QCN-IsBlank $normalizedStem)) { $Event['sheetStem'] = $normalizedStem }
    elseif (-not $Event.ContainsKey('sheetStem') -or (_QCN-IsBlank $Event.sheetStem)) {
        if ($Event.documentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
            try { $Event['sheetStem'] = [string](Get-PWSheetStemFromDocumentName -DocumentName ([string]$Event.documentName)) } catch { }
        }
    }

    if ((-not $Event.ContainsKey('recipientKey') -or (_QCN-IsBlank $Event.recipientKey)) -and ($Event.ContainsKey('to') -or $Event.ContainsKey('cc'))) {
        $Event['recipientKey'] = _QCN-BuildRecipientKey -To @($Event.to) -Cc @($Event.cc)
    }
    if (-not $Event.ContainsKey('notificationType') -or (_QCN-IsBlank $Event.notificationType)) {
        $Event['notificationType'] = [string]$Event.eventType
    }
    if (-not $Event.ContainsKey('targetState') -or (_QCN-IsBlank $Event.targetState)) {
        $Event['targetState'] = [string]$Event.currentState
    }
    $resolvedCycleId = _QCN-ResolveNotificationCycleId -Event $Event -Config $Config -Job $Job
    if (-not (_QCN-IsBlank $resolvedCycleId)) {
        $Event['cycleId'] = $resolvedCycleId
    } elseif (-not $Event.ContainsKey('cycleId')) {
        $Event['cycleId'] = ''
    }
    $sheetTransitionKey = Get-QCNotificationSheetTransitionKey -Event $Event -Config $Config
    if (-not (_QCN-IsBlank $sheetTransitionKey)) {
        $Event['sheetTransitionKey'] = $sheetTransitionKey
    }
    $transitionSource = _QCN-ResolveNotificationTransitionSource -Event $Event -Job $Job
    if (-not (_QCN-IsBlank $transitionSource)) { $Event['transitionSource'] = $transitionSource }
    $logicalAnchor = _QCN-ResolveNotificationLogicalTransitionAnchor -Event $Event -Config $Config -Job $Job
    if (-not (_QCN-IsBlank $logicalAnchor)) { $Event['logicalTransitionAnchor'] = $logicalAnchor }
    $auditIdForKey = _QCN-GetNotificationAuditEventId -Event $Event
    if ($null -ne $auditIdForKey -and $auditIdForKey -gt 0) { $Event['auditEventId'] = $auditIdForKey }

    $resolvedProcessForDedupe = _QCN-ResolveNotificationProcessType -Event $Event -Config $Config -Job $Job
    if (-not (_QCN-IsBlank $resolvedProcessForDedupe)) {
        $Event['qcProcessType'] = $resolvedProcessForDedupe
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($field in @($fields)) {
        $value = ''
        switch ([string]$field) {
            'sheetStem' {
                if ($Event.sheetStem) { $value = [string]$Event.sheetStem }
                elseif ($Event.documentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                    try { $value = [string](Get-PWSheetStemFromDocumentName -DocumentName ([string]$Event.documentName)) } catch { }
                }
                if (_QCN-IsPlaceholderNotificationSheetStem -Stem $value) {
                    $resolvedStem = _QCN-NormalizeNotificationSheetStemForDedupe -Event $Event -Config $Config
                    if (-not (_QCN-IsBlank $resolvedStem)) { $value = $resolvedStem }
                }
            }
            'qcProcessType' {
                $value = ''
                if (Get-Command -Name '_QCN-ResolveNotificationProcessType' -ErrorAction SilentlyContinue) {
                    $value = _QCN-ResolveNotificationProcessType -Event $Event -Config $Config -Job $Job
                }
                if (_QCN-IsBlank $value) {
                    if ($Event.qcProcessType) {
                        if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
                            $norm = Normalize-QCProcessType -ProcessType ([string]$Event.qcProcessType) -ReviewType ([string]$Event.qcReviewType) -Context $Event -AllowNullOnEmpty
                            if ($norm) { $value = $norm }
                        }
                        if (_QCN-IsBlank $value) { $value = ([string]$Event.qcProcessType).Trim().ToLowerInvariant() }
                    }
                }
            }
            'folderPath' { $value = [string]$Event.folderPath }
            'documentGuid' {
                $value = if ($Event.documentGuid) { [string]$Event.documentGuid }
                elseif ($Event.documentPath) { [string]$Event.documentPath }
                else { [string]$Event.documentName }
            }
            'documentName' { $value = [string]$Event.documentName }
            'documentPath' { $value = [string]$Event.documentPath }
            'eventType' { $value = [string]$Event.eventType }
            'notificationType' { $value = [string]$Event.notificationType }
            'currentState' { $value = [string]$Event.currentState }
            'targetState' { $value = [string]$Event.targetState }
            'previousState' {
                $value = _QCN-NormalizeNotificationDedupePreviousState -PreviousState ([string]$Event.previousState) `
                    -CurrentState ([string]$Event.currentState) -Config $Config
            }
            'project' { $value = [string]$Event.project }
            'stateTransitionKey' {
                # Legacy config field name: value is sheet-package transition, not audit:{id}.
                $value = if ($Event.sheetTransitionKey) { [string]$Event.sheetTransitionKey } else {
                    Get-QCNotificationSheetTransitionKey -Event $Event -Config $Config
                }
            }
            'sheetTransitionKey' {
                $value = if ($Event.sheetTransitionKey) { [string]$Event.sheetTransitionKey } else {
                    Get-QCNotificationSheetTransitionKey -Event $Event -Config $Config
                }
            }
            'cycleId' { $value = [string]$Event.cycleId }
            'recipientKey' { $value = [string]$Event.recipientKey }
            'transitionSource' {
                $value = if ($Event.transitionSource) { [string]$Event.transitionSource } else {
                    _QCN-ResolveNotificationTransitionSource -Event $Event -Job $Job
                }
            }
            'logicalTransitionAnchor' {
                $value = if ($Event.logicalTransitionAnchor) { [string]$Event.logicalTransitionAnchor } else {
                    _QCN-ResolveNotificationLogicalTransitionAnchor -Event $Event -Config $Config -Job $Job
                }
            }
            'auditEventId' {
                $auditVal = _QCN-GetNotificationAuditEventId -Event $Event
                if ($null -ne $auditVal -and $auditVal -gt 0) { $value = [string]$auditVal }
            }
            'transitionGroupId' {
                if ($Event.ContainsKey('transitionGroupId') -and -not (_QCN-IsBlank $Event.transitionGroupId)) {
                    $value = ([string]$Event.transitionGroupId).Trim().ToLowerInvariant()
                }
            }
            'sheetPackageId' {
                if ($Event.ContainsKey('sheetPackageId') -and -not (_QCN-IsBlank $Event.sheetPackageId)) {
                    $value = ([string]$Event.sheetPackageId).Trim().ToLowerInvariant()
                }
            }
            'transitionId' {
                # Backward-compatible opt-in only. Not part of the default key because it is
                # per-document-row rather than per logical sheet transition.
                if ($Event.ContainsKey('transitionId') -and $null -ne $Event.transitionId) {
                    try { $value = [string][int]$Event.transitionId } catch { $value = [string]$Event.transitionId }
                }
            }
            default {
                if ($Event.ContainsKey($field)) { $value = [string]$Event[$field] }
            }
        }
        $parts.Add(('{0}={1}' -f $field, $value)) | Out-Null
    }
    return ($parts -join '|')
}

function _QCN-GetSheetPackageDedupeSettings {
    param(
        [hashtable]$Settings
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $base = _QCN-ToHashtable $Settings
    if (-not $base) { $base = @{} }
    $pkgSettings = @{}
    foreach ($k in @($base.Keys)) { $pkgSettings[$k] = $base[$k] }
    $dedupe = _QCN-ToHashtable $base.dedupe
    if (-not $dedupe) { $dedupe = @{} }
    $pkgDedupe = @{}
    foreach ($k in @($dedupe.Keys)) { $pkgDedupe[$k] = $dedupe[$k] }
    $pkgFields = if ($dedupe.sheetPackageKeyFields) { @($dedupe.sheetPackageKeyFields) } else { @('sheetStem', 'currentState', 'cycleId') }
    $pkgDedupe['keyFields'] = $pkgFields
    $pkgSettings['dedupe'] = $pkgDedupe
    return $pkgSettings
}

function _QCN-FormatSheetPackageDedupeStoreKey {
    param([Parameter(Mandatory)][string]$PackageKey)
    if (_QCN-IsBlank $PackageKey) { return '' }
    return 'spkg|' + $PackageKey
}

function Get-QCNotificationSheetPackageDedupeKey {
    <#
    .SYNOPSIS
    Dedupe identity for one logical sheet-package notification in a QC cycle (ignores per-file GUID and audit echo ids).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Event,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    $pkgSettings = _QCN-GetSheetPackageDedupeSettings -Settings $Settings
    return Get-QCNotificationDedupeKey -Event $Event -Settings $pkgSettings -Config $Config -Job $Job
}

function Test-QCNotificationSheetPackageAlreadySent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Event,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config $Config }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return $false }

    $packageKey = Get-QCNotificationSheetPackageDedupeKey -Event $Event -Settings $Settings -Config $Config -Job $Job
    if (_QCN-IsBlank $packageKey) { return $false }
    $storeKey = _QCN-FormatSheetPackageDedupeStoreKey -PackageKey $packageKey
    $storePath = _QCN-GetNotificationDedupeStorePath -Settings $Settings
    if (_QCN-TestNotificationDedupeInStore -DedupeKey $storeKey -StorePath $storePath) { return $true }

    if ($Config -and (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue)) {
        try {
            $dup = Test-QCDuplicateJob -DedupeKey $storeKey -Config $Config
            if ($dup.IsSuccess -and [bool]$dup.Data.isDuplicate) {
                foreach ($match in @($dup.Data.matches)) {
                    if (-not $match) { continue }
                    $matchState = ''
                    try { $matchState = [string]$match.state } catch { }
                    if ($matchState -in @('pending', 'running')) { return $true }
                    if ($matchState -eq 'succeeded') {
                        $jobObj = $null
                        if ($match.ContainsKey('path') -and $match.path -and (Test-Path -LiteralPath ([string]$match.path))) {
                            try { $jobObj = Get-Content -LiteralPath ([string]$match.path) -Raw | ConvertFrom-Json -ErrorAction Stop } catch { }
                        }
                        if ($null -ne $jobObj -and (_QCN-TestNotificationJobActuallySent -Job (_QCN-ToHashtable $jobObj))) { return $true }
                    }
                }
            }
        } catch { }
    }

    if ($Config -and (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
        try {
            if (Test-QCDatabaseEnabled -Config $Config) {
                $stem = ''
                if ($Event.sheetStem) { $stem = [string]$Event.sheetStem }
                elseif ($Event.documentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                    try { $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName ([string]$Event.documentName)) } catch { }
                }
                $curr = if ($Event.currentState) { [string]$Event.currentState } else { '' }
                $cycleId = if ($Event.cycleId) { [string]$Event.cycleId } else { '' }
                if (-not (_QCN-IsBlank $stem) -and -not (_QCN-IsBlank $curr)) {
                    $stemPattern = '%sheetStem=' + $stem + '%'
                    $statePattern = '%currentState=' + $curr + '%'
                    $cyclePattern = if (-not (_QCN-IsBlank $cycleId)) { '%' + $cycleId + '%' } else { '' }
                    $sql = @"
SELECT TOP 1 1 AS found
FROM notification_log
WHERE success = 1
  AND dedupe_key LIKE @stemPattern
  AND dedupe_key LIKE @statePattern
"@
                    $params = @{ stemPattern = $stemPattern; statePattern = $statePattern }
                    if (-not (_QCN-IsBlank $cyclePattern)) {
                        $sql += "  AND dedupe_key LIKE @cyclePattern`n"
                        $params['cyclePattern'] = $cyclePattern
                    }
                    $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters $params
                    if ($dbRes.IsSuccess -and $dbRes.Data.table -and $dbRes.Data.table.Rows.Count -gt 0) {
                        return $true
                    }
                }
            }
        } catch { }
    }

    return $false
}

function Register-QCNotificationSheetPackageDedupe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Event,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null,
        [hashtable]$ResultData = $null
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config $Config }
    $packageKey = Get-QCNotificationSheetPackageDedupeKey -Event $Event -Settings $Settings -Config $Config -Job $Job
    if (_QCN-IsBlank $packageKey) { return }
    $storeKey = _QCN-FormatSheetPackageDedupeStoreKey -PackageKey $packageKey
    $data = if ($ResultData) { $ResultData } else { @{ eventType = ''; documentName = ''; provider = '' } }
    Register-QCNotificationDedupe -DedupeKey $storeKey -Settings $Settings -ResultData $data
}

function _QCN-GetNotificationDedupeStorePath {
    param([hashtable]$Settings)
    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if ($dedupe -and $dedupe.storePath) { return (_QCN-ResolveRepoPath -Path ([string]$dedupe.storePath)) }
    return (Join-Path (_QCN-GetRepoRoot) 'notifications\dedupe\sent-keys.jsonl')
}

function _QCN-TestNotificationDedupeInStore {
    param(
        [Parameter(Mandatory)][string]$DedupeKey,
        [Parameter(Mandatory)][string]$StorePath
    )

    if (_QCN-IsBlank $DedupeKey) { return $false }
    if (-not (Test-Path -LiteralPath $StorePath)) { return $false }
    try {
        $lines = Get-Content -LiteralPath $StorePath -ErrorAction Stop
        foreach ($line in @($lines)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $row = $line | ConvertFrom-Json -ErrorAction Stop
                if ($row.key -eq $DedupeKey) {
                    $status = if ($row.status) { ([string]$row.status).Trim().ToLowerInvariant() } else { 'sent' }
                    if ($status -eq 'pending') { continue }
                    return $true
                }
            } catch { }
        }
    } catch { }
    return $false
}

function Remove-QCNotificationDedupeKey {
    <#
    .SYNOPSIS
    Removes a dedupe key from the local store (e.g. after a failed send that claimed pending).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DedupeKey,
        [hashtable]$Settings
    )

    if (_QCN-IsBlank $DedupeKey) { return $false }
    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return $false }

    $storePath = _QCN-GetNotificationDedupeStorePath -Settings $Settings
    if (-not (Test-Path -LiteralPath $storePath)) { return $false }

    $mutexName = 'Global\QCNotifyDedupe_' + ([string][BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($storePath.ToLowerInvariant())))).Replace('-', '')
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(15000)
        if (-not $acquired) { return $false }
        $remaining = [System.Collections.Generic.List[string]]::new()
        $removed = $false
        foreach ($line in @(Get-Content -LiteralPath $storePath -ErrorAction Stop)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $row = $line | ConvertFrom-Json -ErrorAction Stop
                if ($row.key -eq $DedupeKey) {
                    $removed = $true
                    continue
                }
            } catch { }
            [void]$remaining.Add($line)
        }
        if ($removed) {
            if ($remaining.Count -eq 0) {
                Set-Content -LiteralPath $storePath -Value '' -Encoding UTF8
            } else {
                Set-Content -LiteralPath $storePath -Value ($remaining.ToArray()) -Encoding UTF8
            }
        }
        return $removed
    } catch {
        return $false
    } finally {
        if ($acquired) {
            try { $mutex.ReleaseMutex() | Out-Null } catch { }
        }
        try { $mutex.Dispose() } catch { }
    }
}

function _QCN-TestNotificationJobActuallySent {
    param([hashtable]$Job)

    if (-not $Job) { return $false }
    $result = _QCN-ToHashtable $Job.result
    if (-not $result) { return $false }

    $data = _QCN-ToHashtable $result.data
    if ($data -and $data.notification) {
        $notif = _QCN-ToHashtable $data.notification
        if ($notif) {
            if ($notif.ContainsKey('skipped') -and [bool]$notif.skipped) { return $false }
            if ($notif.ContainsKey('success') -and $notif.success -eq $true) { return $true }
            return $false
        }
    }

    $code = if ($result.code) { [string]$result.code } else { '' }
    if ($code -match '^QC_NOTIFICATION_(SENT|MOCK|GRAPH|JOB_OK)$') {
        if ($data -and $data.skipped) { return $false }
        return $true
    }
    return $false
}

function Test-QCNotificationDedupe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DedupeKey,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [Nullable[int]]$TransitionId = $null,
        [string]$ExcludeJobId = ''
    )

    if (_QCN-IsBlank $DedupeKey) { return $false }

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return $false }

    $transitionAlreadySent = $false
    if ($null -ne $TransitionId -and $TransitionId -gt 0) {
        if ($Config -and (Get-Command -Name 'Test-QCTransitionEventNotificationSent' -ErrorAction SilentlyContinue)) {
            try {
                if (Test-QCTransitionEventNotificationSent -Config $Config -TransitionId $TransitionId) { $transitionAlreadySent = $true }
            } catch { }
        }
        if (-not $transitionAlreadySent -and $Config -and (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
            try {
                if (Test-QCDatabaseEnabled -Config $Config) {
                    $tidRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 1 AS found
FROM notification_log
WHERE transition_id = @transitionId AND success = 1
"@ -Parameters @{ transitionId = $TransitionId }
                    if ($tidRes.IsSuccess -and $tidRes.Data.table -and $tidRes.Data.table.Rows.Count -gt 0) {
                        $transitionAlreadySent = $true
                    }
                }
            } catch { }
        }
        if ($transitionAlreadySent) { return $true }
    }

    if ($Config -and (Get-Command -Name 'Test-QCDuplicateJob' -ErrorAction SilentlyContinue)) {
        try {
            $dup = Test-QCDuplicateJob -DedupeKey $DedupeKey -Config $Config
            if ($dup.IsSuccess -and $dup.Data -and [bool]$dup.Data.isDuplicate) {
                foreach ($match in @($dup.Data.matches)) {
                    if (-not $match) { continue }
                    $matchJobId = ''
                    $matchState = ''
                    try { $matchJobId = [string]$match.jobId } catch { }
                    try { $matchState = [string]$match.state } catch { }
                    if (-not (_QCN-IsBlank $ExcludeJobId) -and $matchJobId -eq $ExcludeJobId) { continue }
                    if ($matchState -in @('pending', 'running')) { return $true }
                    if ($matchState -eq 'succeeded') {
                        $jobObj = $null
                        if ($match.ContainsKey('path') -and $match.path -and (Test-Path -LiteralPath ([string]$match.path))) {
                            try { $jobObj = Get-Content -LiteralPath ([string]$match.path) -Raw | ConvertFrom-Json -ErrorAction Stop } catch { }
                        }
                        if ($null -eq $jobObj) { continue }
                        if (_QCN-TestNotificationJobActuallySent -Job (_QCN-ToHashtable $jobObj)) { return $true }
                    }
                }
            }
        } catch { }
    }

    if ($Config -and (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
        try {
            if (Test-QCDatabaseEnabled -Config $Config) {
                $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 1 AS found
FROM notification_log
WHERE dedupe_key = @dedupeKey AND success = 1
"@ -Parameters @{ dedupeKey = $DedupeKey }
                if ($dbRes.IsSuccess -and $dbRes.Data.table -and $dbRes.Data.table.Rows.Count -gt 0) {
                    return $true
                }
            }
        } catch { }
    }

    $storePath = _QCN-GetNotificationDedupeStorePath -Settings $Settings
    if (_QCN-TestNotificationDedupeInStore -DedupeKey $DedupeKey -StorePath $storePath) { return $true }
    return $false
}

function Register-QCNotificationDedupe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DedupeKey,
        [hashtable]$Settings,
        [hashtable]$ResultData
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return }

    $storePath = _QCN-GetNotificationDedupeStorePath -Settings $Settings
    $status = 'sent'
    if ($ResultData -and $ResultData.status) { $status = ([string]$ResultData.status).Trim().ToLowerInvariant() }
    if ($status -eq 'sent' -and (_QCN-TestNotificationDedupeInStore -DedupeKey $DedupeKey -StorePath $storePath)) { return }

    Remove-QCNotificationDedupeKey -DedupeKey $DedupeKey -Settings $Settings | Out-Null

    $dir = Split-Path -Parent $storePath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $entry = @{
        key = $DedupeKey
        timestampUtc = Get-QCTimestamp
        eventType = $ResultData.eventType
        documentName = $ResultData.documentName
        provider = $ResultData.provider
        status = $status
    } | ConvertTo-Json -Compress

    Add-Content -LiteralPath $storePath -Value $entry -Encoding UTF8
}

function Register-QCNotificationDedupeClaim {
    <#
    Atomically reserves a dedupe key before send so concurrent notification jobs cannot both pass Test-QCNotificationDedupe.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DedupeKey,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [Nullable[int]]$TransitionId = $null,
        [string]$ExcludeJobId = '',
        [hashtable]$ResultData = $null
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return $true }
    if (_QCN-IsBlank $DedupeKey) { return $false }

    $storePath = _QCN-GetNotificationDedupeStorePath -Settings $Settings
    $mutexName = 'Global\QCNotifyDedupe_' + ([string][BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($storePath.ToLowerInvariant())))).Replace('-', '')
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $acquired = $false
    try {
        $acquired = $mutex.WaitOne(15000)
        if (-not $acquired) { return $false }
        if (Test-QCNotificationDedupe -DedupeKey $DedupeKey -Settings $Settings -Config $Config -TransitionId $TransitionId -ExcludeJobId $ExcludeJobId) {
            return $false
        }
        $claimData = if ($ResultData) { $ResultData } else { @{ eventType = ''; documentName = ''; provider = '' } }
        Register-QCNotificationDedupe -DedupeKey $DedupeKey -Settings $Settings -ResultData (@{
            eventType = $claimData.eventType
            documentName = $claimData.documentName
            provider = $claimData.provider
            status = 'pending'
        })
        return $true
    } finally {
        if ($acquired) {
            try { $mutex.ReleaseMutex() | Out-Null } catch { }
        }
        try { $mutex.Dispose() } catch { }
    }
}

function _QCN-EnsureSingleResult {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Array]) {
        foreach ($item in @($Value)) {
            if ($null -eq $item) { continue }
            try {
                $c = [string]$item.Code
                if ($c -match '^QC_NOTIFICATION') { return $item }
            } catch { }
        }
        return $Value[-1]
    }
    return $Value
}

function Write-QCNotificationResult {
    <#
    .SYNOPSIS
    Logs a notification attempt/result as structured JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Code,
        [Parameter(Mandatory)]
        [string]$Level,
        [Parameter(Mandatory)]
        [string]$Message,
        [hashtable]$Result,
        [hashtable]$Event,
        [hashtable]$Job
    )

    $data = @{}
    if ($Result) {
        foreach ($k in @($Result.Keys)) { $data[$k] = $Result[$k] }
    }
    if ($Event) {
        $data['event'] = $Event
    }
    if ($Job -and $Job.ContainsKey('id')) {
        $data['jobId'] = [string]$Job.id
    }

    if (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $data | Out-Null
    }
}

function _QCN-GetNotificationAuditEventId {
    param([hashtable]$Event)
    if (-not $Event) { return $null }
    foreach ($key in @('auditEventId', 'triggerAuditId', 'sourceAuditId')) {
        if ($Event.ContainsKey($key) -and $null -ne $Event[$key]) {
            try {
                $v = [long]$Event[$key]
                if ($v -gt 0) { return $v }
            } catch { }
        }
    }
    if ($Event.ContainsKey('stateTransitionKey') -and -not (_QCN-IsBlank $Event.stateTransitionKey)) {
        $m = [regex]::Match([string]$Event.stateTransitionKey, 'audit:(\d+)')
        if ($m.Success) {
            try { return [long]$m.Groups[1].Value } catch { }
        }
    }
    return $null
}

function _QCN-TestNotificationActorIsAutomation {
    param(
        [hashtable]$Config,
        [hashtable]$Event
    )
    if (-not $Config -or -not $Event) { return $false }
    if (-not (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue)) { return $false }
    $userno = $null
    if ($Event.ContainsKey('changedByUser') -and $null -ne $Event.changedByUser) {
        try {
            $n = [int]$Event.changedByUser
            if ($n -gt 0) { $userno = $n }
        } catch { }
    }
    $username = if ($Event.ContainsKey('changedByUsername')) { [string]$Event.changedByUsername } else { '' }
    try { return [bool](Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $userno -ChangedByUsername $username) } catch { }
    return $false
}

function _QCN-ResolveNotificationTriggerSource {
    param(
        [hashtable]$Event,
        [hashtable]$Job
    )
    if ($Event) {
        foreach ($key in @('triggerSource', 'source', 'origin')) {
            if ($Event.ContainsKey($key) -and -not (_QCN-IsBlank $Event[$key])) { return [string]$Event[$key] }
        }
        if ($Event.ContainsKey('stateTransitionKey') -and -not (_QCN-IsBlank $Event.stateTransitionKey)) {
            $st = [string]$Event.stateTransitionKey
            if ($st.StartsWith('audit:', [StringComparison]::OrdinalIgnoreCase)) { return 'audit' }
            if ($st.StartsWith('workflow:', [StringComparison]::OrdinalIgnoreCase)) { return 'workflow' }
        }
    }
    if ($Job) {
        foreach ($key in @('type', 'jobType')) {
            if ($Job.ContainsKey($key) -and -not (_QCN-IsBlank $Job[$key])) { return [string]$Job[$key] }
        }
    }
    return ''
}

function _QCN-WriteNotificationLifecycleLog {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Level = 'Information',
        [string]$Message = '',
        [hashtable]$Event,
        [hashtable]$Config = $null,
        [hashtable]$Job,
        [string[]]$To = @(),
        [string[]]$Cc = @(),
        [string]$Subject = ''
    )

    if (-not (Get-Command -Name Write-QCJsonLog -ErrorAction SilentlyContinue)) { return }
    if (_QCN-IsBlank $Message) { $Message = $Code }
    $actor = ''
    if ($Event -and $Event.ContainsKey('changedByUsername') -and -not (_QCN-IsBlank $Event.changedByUsername)) { $actor = [string]$Event.changedByUsername }
    elseif ($Event -and $Event.ContainsKey('submittedBy') -and -not (_QCN-IsBlank $Event.submittedBy)) { $actor = [string]$Event.submittedBy }
    $recipientKey = ''
    if ($Event -and $Event.ContainsKey('recipientKey') -and -not (_QCN-IsBlank $Event.recipientKey)) { $recipientKey = [string]$Event.recipientKey }
    else { $recipientKey = _QCN-BuildRecipientKey -To @($To) -Cc @($Cc) }
    $data = @{
        jobId = if ($Job -and $Job.ContainsKey('id')) { [string]$Job.id } elseif ($Event -and $Event.ContainsKey('sourceJobId')) { [string]$Event.sourceJobId } else { '' }
        auditEventId = _QCN-GetNotificationAuditEventId -Event $Event
        documentName = if ($Event -and $Event.ContainsKey('documentName')) { [string]$Event.documentName } else { '' }
        sheetStem = if ($Event -and $Event.ContainsKey('sheetStem')) { [string]$Event.sheetStem } else { '' }
        targetState = if ($Event -and $Event.ContainsKey('targetState')) { [string]$Event.targetState } elseif ($Event -and $Event.ContainsKey('currentState')) { [string]$Event.currentState } else { '' }
        previousState = if ($Event -and $Event.ContainsKey('previousState')) { [string]$Event.previousState } else { '' }
        recipient = $recipientKey
        to = @($To)
        cc = @($Cc)
        notificationType = if ($Event -and $Event.ContainsKey('notificationType')) { [string]$Event.notificationType } elseif ($Event -and $Event.ContainsKey('eventType')) { [string]$Event.eventType } else { '' }
        dedupeKey = if ($Event -and $Event.ContainsKey('dedupeKey')) { [string]$Event.dedupeKey } else { '' }
        actor = $actor
        actorIsAutomation = _QCN-TestNotificationActorIsAutomation -Config $Config -Event $Event
        triggerSource = _QCN-ResolveNotificationTriggerSource -Event $Event -Job $Job
        subject = $Subject
    }
    Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data $data | Out-Null
}

function Send-QCNotification {
    <#
    .SYNOPSIS
    Sends a QC notification using the configured provider (Mock or MicrosoftGraph).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Subject = '',
        [string]$Body = '',
        [string[]]$To = @(),
        [string[]]$Cc = @()
    )

    $settings = Get-QCNotificationSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        $result = @{
            success = $false
            skipped = $true
            provider = [string]$settings.provider
            dryRun = [bool]$settings.dryRun
            eventType = $Event.eventType
            documentName = $Event.documentName
            to = @($To)
            cc = @($Cc)
            message = 'Notifications are disabled in configuration.'
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_DISABLED' -Level 'Information' -Message $result.message -Result $result -Event $Event
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DISABLED' -Message $result.message -Data $result
    }

    if ($To.Count -eq 0) {
        $result = @{
            success = $false
            skipped = $true
            provider = [string]$settings.provider
            dryRun = [bool]$settings.dryRun
            eventType = $Event.eventType
            documentName = $Event.documentName
            to = @()
            cc = @($Cc)
            message = 'Notification skipped: no To recipients resolved.'
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS' -Level 'Warning' -Message $result.message -Result $result -Event $Event
        return New-QCFailureResult -Code 'QC_NOTIFICATION_SKIPPED_NO_RECIPIENTS' -Message $result.message -Data $result
    }

    $eventCfg = @{}
    if ($Event._eventCfg) {
        $ec = _QCN-ToHashtable $Event._eventCfg
        if ($ec) { $eventCfg = $ec }
    }

    if (_QCN-IsBlank $Subject) {
        $reviewType = ''
        if ($Event.reviewType) { $reviewType = [string]$Event.reviewType }
        elseif ($Event.qcReviewType) { $reviewType = [string]$Event.qcReviewType }
        $subjectDocumentName = if ($Event.displayDocumentName) { [string]$Event.displayDocumentName } else { [string]$Event.documentName }
        $tokens = _QCN-NewNotificationSubjectTokens -DocumentName $subjectDocumentName `
            -DocumentPath ([string]$Event.documentPath) -Project ([string]$Event.project) `
            -PreviousState ([string]$Event.previousState) -CurrentState ([string]$Event.currentState) `
            -EventType ([string]$Event.eventType) -ReviewType $reviewType
        $template = _QCN-ResolveNotificationSubjectTemplate -EventCfg $eventCfg -Settings $settings
        $Subject = Expand-QCNotificationTemplate -Template $template -Tokens $tokens
    }
    if (_QCN-IsBlank $Body) {
        $Body = New-QCNotificationEmailBody -Event $Event
    }

    $htmlBody = ''
    $bodyContentType = 'Text'
    $logoPath = 'email/typsalogo.png.webp'
    if (_QCN-UsesHtmlEmailBody -Settings $settings -EventCfg $eventCfg) {
        $emailCfg = _QCN-ToHashtable $settings.email
        if ($emailCfg) {
            if ($emailCfg.logoPath) { $logoPath = [string]$emailCfg.logoPath }
        }
        if (-not (_QCN-IsBlank $Event.qcPdfUrl)) {
            $qcUrl = [string]$Event.qcPdfUrl
        }
        else {
            $qcUrl = Resolve-QCNotificationQcPdfUrl -Event $Event -Settings $settings -Document $Event._document -Config $Config
            if (-not (_QCN-IsBlank $qcUrl)) { $Event['qcPdfUrl'] = $qcUrl }
        }
        if (-not (_QCN-IsBlank $qcUrl)) {
            $linkGuid = _QCN-ExtractPwLinkDocumentGuid -Url $qcUrl
            $linkResolutionSource = if ($Event.linkResolutionSource) { [string]$Event.linkResolutionSource } else { '' }
            if (-not $Event.ContainsKey('qcPdfGuidResolutionLogged')) {
                _QCN-WriteQcPdfNotificationGuidResolutionLog -Config $Config `
                    -TriggerDocumentGuid $(if ($Event.triggerDocumentGuid) { [string]$Event.triggerDocumentGuid } else { [string]$Event.documentGuid }) `
                    -ResolvedQcPdfGuid ([string]$Event.documentGuid) -NotificationDocumentGuid ([string]$Event.documentGuid) `
                    -LinkDocumentGuid $linkGuid -SheetPackageId $(if ($Event.sheetPackageId) { [string]$Event.sheetPackageId } else { '' }) `
                    -QcReviewType $(if ($Event.qcReviewType) { [string]$Event.qcReviewType } elseif ($Event.reviewType) { [string]$Event.reviewType } else { '' }) `
                    -ResolutionSource $(if ($Event.resolutionSource) { [string]$Event.resolutionSource } else { '' }) `
                    -LinkResolutionSource $linkResolutionSource -Event $Event
                $Event['qcPdfGuidResolutionLogged'] = $true
            }
        }
        if (_QCN-IsBlank $qcUrl) {
            $result = @{
                success = $false
                skipped = $false
                provider = [string]$settings.provider
                dryRun = [bool]$settings.dryRun
                eventType = $Event.eventType
                documentName = $Event.documentName
                to = @($To)
                cc = @($Cc)
                message = 'HTML notification requires QCPdfUrl; none could be resolved.'
                timestampUtc = Get-QCTimestamp
            }
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_MISSING_QC_PDF_URL' -Level 'Warning' -Message $result.message -Result $result -Event $Event
            return New-QCFailureResult -Code 'QC_NOTIFICATION_MISSING_QC_PDF_URL' -Message $result.message -Data $result
        }

        $templatePath = 'email/templates/qc_notification.html'
        if ($eventCfg -and $eventCfg.emailTemplate) { $templatePath = [string]$eventCfg.emailTemplate }
        elseif ($emailCfg -and $emailCfg.templatePath) { $templatePath = [string]$emailCfg.templatePath }

        $folderForEmail = ''
        if ($Event.folderPath) { $folderForEmail = [string]$Event.folderPath }
        elseif ($Event.documentPath -and ($Event.documentPath -match '\\')) {
            $folderForEmail = [System.IO.Path]::GetDirectoryName([string]$Event.documentPath)
        }
        $templateData = New-QCNotificationEmailTemplateData -Event $Event -EventCfg $eventCfg -Settings $settings `
            -Config $Config -FolderPath $folderForEmail -Document $Event._document -Subject $Subject
        try {
            $htmlBody = ConvertTo-QCEmailHtml -TemplatePath $templatePath -Data $templateData
            $bodyContentType = 'HTML'
            if (_QCN-IsBlank $Body) {
                $Body = New-QCNotificationEmailBody -Event $Event
            }
        } catch {
            $result = @{
                success = $false
                provider = [string]$settings.provider
                dryRun = [bool]$settings.dryRun
                eventType = $Event.eventType
                documentName = $Event.documentName
                to = @($To)
                cc = @($Cc)
                message = $_.Exception.Message
                timestampUtc = Get-QCTimestamp
            }
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_HTML_RENDER_FAILED' -Level 'Warning' -Message $result.message -Result $result -Event $Event
            return New-QCFailureResult -Code 'QC_NOTIFICATION_HTML_RENDER_FAILED' -Message $result.message -Data $result
        }
    }

    $payload = @{
        eventType = $Event.eventType
        project = $Event.project
        documentName = $Event.documentName
        documentPath = $Event.documentPath
        documentGuid = $Event.documentGuid
        previousState = $Event.previousState
        currentState = $Event.currentState
        actionRequired = $Event.actionRequired
        sourceJobId = $Event.sourceJobId
        qcPdfUrl = $Event.qcPdfUrl
        subject = $Subject
        body = $Body
        bodyContentType = $bodyContentType
        htmlBody = $htmlBody
        logoPath = $logoPath
        to = @($To)
        cc = @($Cc)
        reviewers = @($Event.reviewers)
        designers = @($Event.designers)
    }

    $provider = ([string]$settings.provider).Trim()
    if (_QCN-IsBlank $provider) { $provider = 'Mock' }

    $threadPlan = $null
    if (Get-Command -Name 'Invoke-QCNotificationThreadedSend' -ErrorAction SilentlyContinue) {
        $threadPlan = Invoke-QCNotificationThreadedSend -Event $Event -Config $Config -Settings $settings `
            -Payload $payload -Provider $provider -DryRun:([bool]$settings.dryRun)
    }

    _QCN-WriteNotificationLifecycleLog -Code 'QC_NOTIFICATION_SEND_ATTEMPT' -Level 'Information' `
        -Message 'Sending QC workflow notification.' -Event $Event -Config $Config -To @($To) -Cc @($Cc) -Subject $Subject
    if ($threadPlan -and $threadPlan.threadKey) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_THREAD_DISPATCH' `
                -Message 'Threaded notification dispatch.' -Data @{
                sheetPackageId = $threadPlan.sheetPackageId
                reviewType = $threadPlan.reviewType
                sendMode = $threadPlan.sendMode
                parentGraphMessageId = $threadPlan.parentMessageId
                workflowEvent = $threadPlan.workflowEvent
                threadKey = $threadPlan.threadKey
            } | Out-Null
        }
    }

    $sendResult = $null
    switch ($provider.ToLowerInvariant()) {
        'microsoftgraph' {
            $graph = _QCN-ToHashtable $settings.graph
            if (-not $graph) { $graph = @{} }
            $sendResult = Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:([bool]$settings.dryRun)
        }
        default {
            $outputRoot = if ($settings.outputRoot) {
                _QCN-ResolveRepoPath -Path ([string]$settings.outputRoot)
            } else {
                (Join-Path (_QCN-GetRepoRoot) 'notifications')
            }
            $sendResult = Send-QCNotificationMock -Payload $payload -OutputRoot $outputRoot -DryRun:([bool]$settings.dryRun)
        }
    }

    $result = @{}
    if ($sendResult.Data) {
        $rd = _QCN-ToHashtable $sendResult.Data
        if ($rd) { foreach ($k in $rd.Keys) { $result[$k] = $rd[$k] } }
    }
    if (-not $result.ContainsKey('timestampUtc')) {
        $result['timestampUtc'] = Get-QCTimestamp
    }

    $code = if ($sendResult.IsSuccess) {
        if ($sendResult.Code) { [string]$sendResult.Code } else { 'QC_NOTIFICATION_SENT' }
    } else { 'QC_NOTIFICATION_FAILED' }
    $level = if ($sendResult.IsSuccess) { 'Information' } else { 'Warning' }
    Write-QCNotificationResult -Code $code -Level $level -Message $sendResult.Message -Result $result -Event $Event
    if ($sendResult.IsSuccess) {
        _QCN-WriteNotificationLifecycleLog -Code 'QC_NOTIFICATION_SENT' -Level 'Information' `
            -Message 'QC workflow notification sent.' -Event $Event -Config $Config -To @($To) -Cc @($Cc) -Subject $Subject
    }

    $notificationLogId = $null
    if (Get-Command -Name Write-QCNotificationTelemetry -ErrorAction SilentlyContinue) {
        $telemetryFolder = [string]$Event.folderPath
        if (_QCN-IsBlank $telemetryFolder) { $telemetryFolder = [string]$Event.documentPath }
        $dedupeForTelemetry = ''
        if ($Event.ContainsKey('dedupeKey') -and -not (_QCN-IsBlank $Event.dedupeKey)) {
            $dedupeForTelemetry = [string]$Event.dedupeKey
        }
        $transitionForTelemetry = $null
        if ($Event.ContainsKey('transitionId') -and $null -ne $Event.transitionId) {
            try { $transitionForTelemetry = [int]$Event.transitionId } catch { $transitionForTelemetry = $null }
        }
        $sheetPkgForTelemetry = $null
        if ($threadPlan -and -not (_QCN-IsBlank $threadPlan.sheetPackageId)) {
            try { $sheetPkgForTelemetry = [guid]$threadPlan.sheetPackageId } catch { $sheetPkgForTelemetry = $null }
        }
        $telemetryRes = Write-QCNotificationTelemetry -Config $Config -EventType ([string]$Event.eventType) `
            -DocumentGuid ([string]$Event.documentGuid) -DocumentName ([string]$Event.documentName) `
            -FolderPath $telemetryFolder `
            -Recipients ((@($To) + @($Cc)) -join ';') -Subject $Subject `
            -DedupeKey $dedupeForTelemetry -Provider ([string]$provider) -Success $sendResult.IsSuccess `
            -ErrorMessage $(if (-not $sendResult.IsSuccess) { [string]$sendResult.Message } else { $null }) `
            -TransitionId $transitionForTelemetry -SheetPackageId $sheetPkgForTelemetry
        if ($telemetryRes -and $telemetryRes.Data -and $telemetryRes.Data.notificationLogId) {
            try { $notificationLogId = [int]$telemetryRes.Data.notificationLogId } catch { }
        }
    }

    if ($sendResult.IsSuccess -and $threadPlan -and (Get-Command -Name 'Complete-QCNotificationThreadedSend' -ErrorAction SilentlyContinue)) {
        $parentInvalid = $false
        if ($result.parentInvalid -eq $true) { $parentInvalid = $true }
        Complete-QCNotificationThreadedSend -Config $Config -ThreadPlan $threadPlan -SendResultData $result `
            -NotificationLogId $notificationLogId -WorkflowEvent $threadPlan.workflowEvent -ParentInvalid:$parentInvalid
    }

    if ($sendResult.IsSuccess) { return New-QCSuccessResult -Code $code -Message $sendResult.Message -Data $result }
    return New-QCFailureResult -Code $code -Message $sendResult.Message -Data $result
}

function _QCN-GetSheetIndexPwStateName {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid
    )
    if (_QCN-IsBlank $DocumentGuid) { return '' }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    try {
        $siRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT pw_state_name FROM sheet_index WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if ($siRes.IsSuccess -and $siRes.Data.table -and $siRes.Data.table.Rows.Count -gt 0) {
            $r = $siRes.Data.table.Rows[0]
            if (-not ($r.pw_state_name -is [DBNull])) { return [string]$r.pw_state_name }
        }
    } catch { }
    return ''
}

function _QCN-GetRequiredEmailRolesForStateEvent {
    param(
        [hashtable]$EventCfg,
        [hashtable]$Config
    )

    $roles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $EventCfg) { return @() }
    foreach ($role in @($EventCfg.to) + @($EventCfg.cc)) {
        if (_QCN-IsBlank $role) { continue }
        switch -Regex ([string]$role) {
            '^reviewers?$' { [void]$roles.Add('reviewer') }
            '^designers?$' { [void]$roles.Add('designer') }
            '^checkers?$' { [void]$roles.Add('checker') }
        }
    }
    return @($roles)
}

function Get-QCStateChangeMissingEmailFields {
    <#
    .SYNOPSIS
    Returns ProjectWise attribute column names that are required but empty for a workflow state notification.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TargetStateName,
        [object]$Document = $null,
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = ''
    )

    $settings = Get-QCNotificationSettings -Config $Config
    if (-not [bool]$settings.enabled) { return @() }
    if (-not [bool]$settings.rollbackWhenEmailAttributesMissing) { return @() }

    $state = ([string]$TargetStateName).Trim()
    if ([string]::IsNullOrWhiteSpace($state)) { return @() }

    $events = _QCN-ToHashtable $settings.events
    if (-not $events -or -not $events.ContainsKey($state)) { return @() }
    $eventCfg = _QCN-ToHashtable $events[$state]
    if (-not $eventCfg -or ($eventCfg.ContainsKey('enabled') -and -not [bool]$eventCfg.enabled)) { return @() }

    $requiredRoles = _QCN-GetRequiredEmailRolesForStateEvent -EventCfg $eventCfg -Config $Config
    if ($requiredRoles.Count -eq 0) { return @() }

    $attr = _QCN-ToHashtable $settings.attributes
    if (-not $attr) { $attr = @{} }
    $designerField = if ($attr.designerEmailField) { [string]$attr.designerEmailField } else { 'EM_Designer_Email' }
    $reviewerField = if ($attr.reviewerEmailField) { [string]$attr.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $checkerField = if ($attr.checkerEmailField) { [string]$attr.checkerEmailField } else { 'EM_Checker_Email' }

    $qcTarget = _QCN-ResolveQcPdfNotificationTarget -Document $Document -Config $Config -Job $null `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -DocumentPath ''
    $notifyDoc = $qcTarget.document
    $notifyName = [string]$qcTarget.documentName
    $notifyGuid = [string]$qcTarget.documentGuid
    $folder = [string]$FolderPath
    if (_QCN-IsBlank $folder) { $folder = [string](_QCN-GetProp -Object $notifyDoc -Names @('FolderPath', 'folderPath')) }

    $roles = _QCN-ResolveNotificationRoleEmails -Document $notifyDoc -Settings $settings -Config $Config `
        -FolderPath $folder -SourceDocumentName $notifyName -DocumentGuid $notifyGuid
    $reviewType = _QCN-ResolveNotificationReviewType -Document $notifyDoc -Settings $settings -Config $Config `
        -FolderPath $folder -SourceName $notifyName
    $useCheckerForReviewer = _QCN-TestNotificationIndependentCheckReviewType -ReviewType $reviewType -Config $Config

    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($role in $requiredRoles) {
        switch ($role) {
            'designer' {
                if (_QCN-IsBlank $roles.designerEmail) { $missing.Add($designerField) | Out-Null }
            }
            'reviewer' {
                if ($useCheckerForReviewer) {
                    if (_QCN-IsBlank $roles.checkerEmail) { $missing.Add($checkerField) | Out-Null }
                } elseif (_QCN-IsBlank $roles.reviewerEmail) {
                    $missing.Add($reviewerField) | Out-Null
                }
            }
            'checker' {
                if (_QCN-IsBlank $roles.checkerEmail) { $missing.Add($checkerField) | Out-Null }
            }
        }
    }
    return @($missing | Select-Object -Unique)
}

function _QCN-ResolveSheetNotificationIdentity {
    <#
    .SYNOPSIS
    Resolves canonical sheet-level identity for notifications (siblings share one sheet stem).
    #>
    param(
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [hashtable]$Config = $null,
        [array]$Members = $null
    )

    $stem = ''
    if (-not (_QCN-IsBlank $DocumentName) -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
        try { $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName $DocumentName) } catch { }
    }

    $sheetPdfName = if ($stem) { ($stem + '.pdf') } else { $DocumentName }
    $sheetPdfGuid = $DocumentGuid

    foreach ($member in @($Members)) {
        if (-not $member) { continue }
        $dn = [string]$member.documentName
        if (Test-QCIsSheetPdfDocumentName -DocumentName $dn) {
            $sheetPdfName = $dn
            if ($member.documentGuid) { $sheetPdfGuid = [string]$member.documentGuid }
            if (_QCN-IsBlank $stem -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
                try { $stem = [string](Get-PWSheetStemFromDocumentName -DocumentName $dn) } catch { }
            }
            break
        }
    }

    if (_QCN-IsBlank $stem -and -not (_QCN-IsBlank $sheetPdfName)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($sheetPdfName)
    }

    return @{
        sheetStem = $stem
        sheetPdfName = $sheetPdfName
        sheetPdfGuid = $sheetPdfGuid
    }
}

function _QCN-GetFinalizingPrependPostStateName {
    param([hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if (Get-Command -Name 'Resolve-QCWorkflowStateAfterPrepend' -ErrorAction SilentlyContinue) {
                return [string](Resolve-QCWorkflowStateAfterPrepend -Settings $wf -Context @{ prependTrigger = 'finalQcComplete' })
            }
        } catch { }
    }
    return 'Ready for Verification'
}

function _QCN-GetQcFinalizingWorkflowStateName {
    param([hashtable]$Config)

    $name = 'Initiate Verification'
    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcFinalizing'
            if (-not (_QCN-IsBlank $resolved)) { $name = [string]$resolved }
        } catch { }
    }
    return $name
}

function _QCN-TestWorkflowStateNameIsQcFinalizing {
    param(
        [string]$StateName,
        [hashtable]$Config
    )

    $state = ([string]$StateName).Trim()
    if ([string]::IsNullOrWhiteSpace($state)) { return $false }
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue) {
        try { return [bool](Test-QCWorkflowStateIsQcFinalizing -StateName $state -Config $Config) } catch { }
    }
    return ($state.ToLowerInvariant() -eq 'initiate verification')
}

function _QCN-GetInitialPrependPostStateName {
    param([hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if (Get-Command -Name 'Resolve-QCWorkflowStateAfterPrepend' -ErrorAction SilentlyContinue) {
                return [string](Resolve-QCWorkflowStateAfterPrepend -Settings $wf -Context @{ prependTrigger = 'initialQcPdf' })
            }
            if ($wf.stateAfterSuccessfulPrepend) { return [string]$wf.stateAfterSuccessfulPrepend }
        } catch { }
    }
    return 'Ready for QC'
}

function _QCN-GetQcInitiatedWorkflowStateName {
    param([hashtable]$Config)

    $name = 'QC Initiated'
    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'qcInitiated'
            if (-not (_QCN-IsBlank $resolved)) { $name = [string]$resolved }
        } catch { }
    }
    return $name
}

function _QCN-TestWorkflowStateNameIsQcInitiated {
    param(
        [string]$StateName,
        [hashtable]$Config
    )

    $state = ([string]$StateName).Trim()
    if ([string]::IsNullOrWhiteSpace($state)) { return $false }
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue) {
        try { return [bool](Test-QCWorkflowStateIsQcInitiated -StateName $state -Config $Config) } catch { }
    }
    return ($state.ToLowerInvariant() -eq 'qc initiated')
}

function Get-QCWorkflowTransitionMissingEmailFields {
    <#
    .SYNOPSIS
    Returns missing email attribute columns for a workflow transition, including prerequisite states
    (e.g. QC Initiated also requires emails for the post-prepend Ready for QC notification).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TargetStateName,
        [object]$Document = $null,
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = ''
    )

    $statesToCheck = [System.Collections.Generic.List[string]]::new()
    $target = ([string]$TargetStateName).Trim()
    if ($target.Length -gt 0) { $statesToCheck.Add($target) | Out-Null }

    if (_QCN-TestWorkflowStateNameIsQcInitiated -StateName $target -Config $Config) {
        $postState = _QCN-GetInitialPrependPostStateName -Config $Config
        if (-not (_QCN-IsBlank $postState) -and ($postState.ToLowerInvariant() -ne $target.ToLowerInvariant())) {
            $statesToCheck.Add([string]$postState) | Out-Null
        }
    }
    if (_QCN-TestWorkflowStateNameIsQcFinalizing -StateName $target -Config $Config) {
        $postState = _QCN-GetFinalizingPrependPostStateName -Config $Config
        if (-not (_QCN-IsBlank $postState) -and ($postState.ToLowerInvariant() -ne $target.ToLowerInvariant())) {
            $statesToCheck.Add([string]$postState) | Out-Null
        }
    }

    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($stateName in @($statesToCheck | Select-Object -Unique)) {
        $fields = @(Get-QCStateChangeMissingEmailFields -Config $Config -TargetStateName $stateName `
            -Document $Document -DocumentName $DocumentName -DocumentGuid $DocumentGuid -FolderPath $FolderPath)
        foreach ($field in $fields) {
            if (-not $missing.Contains([string]$field)) { $missing.Add([string]$field) | Out-Null }
        }
    }
    return @($missing)
}

function Test-QCPrependBlockedByMissingEmailAttributes {
    <#
    .SYNOPSIS
    True when initial QC prepend must not run because required notification email attributes are missing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$FolderPath = '',
        [string]$SheetPdfName = '',
        [string]$DocumentGuid = '',
        [object]$Document = $null
    )

    $initiated = _QCN-GetQcInitiatedWorkflowStateName -Config $Config
    $missing = @(Get-QCWorkflowTransitionMissingEmailFields -Config $Config -TargetStateName $initiated `
        -Document $Document -DocumentName $SheetPdfName -DocumentGuid $DocumentGuid -FolderPath $FolderPath)

    return @{
        blocked = ($missing.Count -gt 0)
        missingFields = @($missing)
        postPrependState = _QCN-GetInitialPrependPostStateName -Config $Config
    }
}

function _QCN-GetProductionWorkflowStateName {
    param([hashtable]$Config)

    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            $resolved = Get-QCWorkflowStateName -Settings $wf -StateKey 'production'
            if (-not (_QCN-IsBlank $resolved)) { return [string]$resolved }
        } catch { }
    }
    return 'In Production'
}

function Resolve-QCWorkflowRollbackPreviousState {
    <#
    .SYNOPSIS
    Resolves the workflow state to restore when a blocked transition is rolled back.
    Prefers sheet PDF sheet_index; falls back to configured production when index is missing or already advanced.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$TargetStateName,
        [array]$Members = @(),
        [string]$SheetPdfGuid = ''
    )

    $targetNorm = ([string]$TargetStateName).Trim().ToLowerInvariant()
    $previous = ''

    foreach ($member in @($Members)) {
        if (-not $member) { continue }
        $dn = [string]$member.documentName
        if (Test-QCIsSheetPdfDocumentName -DocumentName $dn) {
            $dg = [string]$member.documentGuid
            if ($dg) { $previous = _QCN-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg }
            break
        }
    }
    if (_QCN-IsBlank $previous -and -not (_QCN-IsBlank $SheetPdfGuid)) {
        $previous = _QCN-GetSheetIndexPwStateName -Config $Config -DocumentGuid $SheetPdfGuid
    }

    $prevNorm = ([string]$previous).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($prevNorm) -or $prevNorm -eq $targetNorm) {
        $previous = _QCN-GetProductionWorkflowStateName -Config $Config
    }
    return [string]$previous
}

function _QCN-GetStateChangeBlockedDedupeKey {
    param(
        [string]$SheetStem,
        [string]$TargetStateName,
        [Nullable[int]]$ChangedByUser = $null
    )

    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add(('sheetStem={0}' -f $(if ($SheetStem) { $SheetStem } else { '' }))) | Out-Null
    $parts.Add('eventType=STATE_CHANGE_BLOCKED') | Out-Null
    $parts.Add(('targetState={0}' -f $(if ($TargetStateName) { $TargetStateName } else { '' }))) | Out-Null
    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0) {
        $parts.Add(('changedByUser={0}' -f $ChangedByUser)) | Out-Null
    }
    return ($parts -join '|')
}

function Resolve-QCStateChangeActorEmailAddress {
    <#
    .SYNOPSIS
    Resolves the SMTP address of the ProjectWise user who committed a workflow state change.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config = $null,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
        try { Sync-PWUserDirectory -Config $Config -UserNumbers @([int]$ChangedByUser) -MaxUsers 1 | Out-Null } catch { }
    }

    if ($null -ne $ChangedByUser -and $ChangedByUser -gt 0 -and $Config -and (Get-Command -Name 'Get-QCPWUserIdentity' -ErrorAction SilentlyContinue)) {
        try {
            $res = Get-QCPWUserIdentity -Config $Config -UserNumber ([int]$ChangedByUser)
            if ($res -and $res.IsSuccess -and $res.Data -and $res.Data.identity) {
                $identity = _QCN-ToHashtable $res.Data.identity
                if ($identity -and $identity.pw_user_email -and -not (_QCN-IsBlank $identity.pw_user_email)) {
                    return [string]$identity.pw_user_email.Trim()
                }
            }
        } catch { }
    }

    if (-not (_QCN-IsBlank $ChangedByUsername) -and $ChangedByUsername -match '@') {
        return [string]$ChangedByUsername.Trim()
    }
    return ''
}

function Send-QCStateChangeBlockedNotification {
    <#
    .SYNOPSIS
    Emails the user who attempted a workflow state change when required document email attributes are missing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string[]]$MissingFields,
        [Parameter(Mandatory)][string]$TargetStateName,
        [string]$PreviousStateName = '',
        [string]$DocumentName = '',
        [string]$DocumentPath = '',
        [string]$DocumentGuid = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [bool]$DryRun = $false
    )

    $settings = Get-QCNotificationSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        return New-QCSuccessResult -Code 'QC_STATE_CHANGE_BLOCKED_NOTIFY_SKIPPED' -Message 'Notifications disabled.' -Data @{ skipped = $true }
    }

    $actorEmail = Resolve-QCStateChangeActorEmailAddress -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername
    if (_QCN-IsBlank $actorEmail) {
        $msg = 'State change blocked for missing email attributes, but actor email could not be resolved.'
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_STATE_CHANGE_BLOCKED_NO_ACTOR_EMAIL' -Message $msg -Data @{
                documentName = $DocumentName; targetState = $TargetStateName; missingFields = @($MissingFields)
                changedByUser = $ChangedByUser; changedByUsername = $ChangedByUsername
            } | Out-Null
        }
        return New-QCFailureResult -Code 'QC_STATE_CHANGE_BLOCKED_NO_ACTOR_EMAIL' -Message $msg -Data @{
            missingFields = @($MissingFields); actorEmail = ''
        }
    }

    $fieldList = ($MissingFields | ForEach-Object { [string]$_ }) -join ', '
    $intro = @(
        'Your ProjectWise QC workflow state change was not applied because required email attributes are missing on the sheet documents.',
        '',
        ('Attempted transition: {0} -> {1}' -f $(if ($PreviousStateName) { $PreviousStateName } else { '(unknown)' }), $TargetStateName),
        ('Document: {0}' -f $(if ($DocumentName) { $DocumentName } else { '(unknown)' })),
        ('Missing attribute(s): {0}' -f $fieldList),
        '',
        'Workflow states for this sheet and its associated files (DGN, sheet PDF, QC PDF) were restored to their previous values.',
        'Enter the missing email values on the sheet PDF (or QC PDF) in ProjectWise, then change the workflow state again.'
    ) -join "`r`n"

    $event = @{
        eventType = 'STATE_CHANGE_BLOCKED'
        documentName = $DocumentName
        documentPath = $DocumentPath
        documentGuid = $DocumentGuid
        previousState = $PreviousStateName
        currentState = $TargetStateName
        actionRequired = ('Missing attributes: ' + $fieldList)
        missingFields = @($MissingFields)
        changedByUser = $ChangedByUser
        changedByUsername = $ChangedByUsername
    }
    $subject = ('QC state change blocked: missing email attributes on {0}' -f $(if ($DocumentName) { $DocumentName } else { 'sheet' }))

    if ($DryRun -or [bool]$settings.dryRun) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_STATE_CHANGE_BLOCKED_NOTIFY_PLANNED' -Message 'Planned blocked-state notice to state-change actor.' -Data @{
                to = $actorEmail; subject = $subject; missingFields = @($MissingFields); documentName = $DocumentName
            } | Out-Null
        }
        return New-QCSuccessResult -Code 'QC_STATE_CHANGE_BLOCKED_NOTIFY_PLANNED' -Message 'Dry-run: would notify state-change actor.' -Data @{
            to = @($actorEmail); subject = $subject; missingFields = @($MissingFields); planned = $true
        }
    }

    return Send-QCNotification -Event $event -Config $Config -Subject $subject -Body $intro -To @($actorEmail) -Cc @()
}

function Invoke-QCWorkflowStateEmailAttributeGate {
    <#
    .SYNOPSIS
    Blocks audit-driven sheet state sync when notification email attributes are missing; rolls back and notifies the actor.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [Parameter(Mandatory)][string]$TargetStateName,
        [Parameter(Mandatory)][array]$Members,
        [hashtable]$StateByGuid = @{},
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [bool]$DryRun = $false
    )

    $result = @{ blocked = $false; missingFields = @(); rollback = $null; notification = $null }
    $settings = Get-QCNotificationSettings -Config $Config
    if (-not [bool]$settings.enabled -or -not [bool]$settings.rollbackWhenEmailAttributesMissing) {
        return $result
    }

    if ($Config -and (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue)) {
        try {
            if (Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername) {
                return $result
            }
        } catch { }
    }

    $sheetIdentity = _QCN-ResolveSheetNotificationIdentity -DocumentName $DocumentName -DocumentGuid $DocumentGuid `
        -Config $Config -Members $Members
    $notifyDocName = [string]$sheetIdentity.sheetPdfName
    $notifyDocGuid = [string]$sheetIdentity.sheetPdfGuid
    if (_QCN-IsBlank $notifyDocName) { $notifyDocName = $DocumentName }
    if (_QCN-IsBlank $notifyDocGuid) { $notifyDocGuid = $DocumentGuid }

    $missing = @(Get-QCWorkflowTransitionMissingEmailFields -Config $Config -TargetStateName $TargetStateName `
        -DocumentName $notifyDocName -DocumentGuid $notifyDocGuid -FolderPath $FolderPath)
    if ($missing.Count -eq 0) { return $result }

    $result.blocked = $true
    $result.missingFields = @($missing)

    $previousState = Resolve-QCWorkflowRollbackPreviousState -Config $Config -TargetStateName $TargetStateName `
        -Members $Members -SheetPdfGuid $notifyDocGuid

    $statesToRevert = [System.Collections.Generic.List[string]]::new()
    $statesToRevert.Add([string]$TargetStateName) | Out-Null
    if (_QCN-TestWorkflowStateNameIsQcInitiated -StateName $TargetStateName -Config $Config) {
        $postState = _QCN-GetInitialPrependPostStateName -Config $Config
        if (-not (_QCN-IsBlank $postState)) { $statesToRevert.Add([string]$postState) | Out-Null }
    }

    if (Get-Command -Name 'Revert-PWAssociatedSheetWorkflowStates' -ErrorAction SilentlyContinue) {
        $result.rollback = Revert-PWAssociatedSheetWorkflowStates -Config $Config -Members $Members `
            -StateByGuid $StateByGuid -FolderPath $FolderPath -TargetStateName $TargetStateName `
            -FallbackPreviousState $previousState -AdditionalStatesToRevert @($statesToRevert | Select-Object -Unique) `
            -DryRun:$DryRun
    }

    $docPath = if ($FolderPath -and $notifyDocName) { ($FolderPath.TrimEnd('\') + '\' + $notifyDocName) } else { '' }
    $blockedDedupeKey = _QCN-GetStateChangeBlockedDedupeKey -SheetStem ([string]$sheetIdentity.sheetStem) `
        -TargetStateName $TargetStateName -ChangedByUser $ChangedByUser
    $notifySkippedDuplicate = $false
    if (Test-QCNotificationDedupe -DedupeKey $blockedDedupeKey -Settings $settings -Config $Config) {
        $notifySkippedDuplicate = $true
        $result.notification = New-QCSuccessResult -Code 'QC_STATE_CHANGE_BLOCKED_NOTIFY_DEDUPED' `
            -Message 'Blocked-state notice already sent for this sheet transition.' -Data @{
            dedupeKey = $blockedDedupeKey; sheetStem = [string]$sheetIdentity.sheetStem; skipped = $true
        }
    } else {
        $result.notification = Send-QCStateChangeBlockedNotification -Config $Config -MissingFields $missing `
            -TargetStateName $TargetStateName -PreviousStateName $previousState -DocumentName $notifyDocName `
            -DocumentPath $docPath -DocumentGuid $notifyDocGuid -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -DryRun:$DryRun
        if (-not $DryRun -and -not [bool]$settings.dryRun -and $result.notification -and $result.notification.IsSuccess) {
            Register-QCNotificationDedupe -DedupeKey $blockedDedupeKey -Settings $settings -ResultData @{
                eventType = 'STATE_CHANGE_BLOCKED'
                documentName = $notifyDocName
                provider = [string]$settings.provider
            }
        }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_STATE_CHANGE_BLOCKED_MISSING_EMAIL' `
            -Message 'Workflow state change rolled back because required email attributes are missing.' -Data @{
            documentName = $notifyDocName; documentGuid = $notifyDocGuid; folderPath = $FolderPath
            triggerDocumentName = $DocumentName; triggerDocumentGuid = $DocumentGuid
            sheetStem = [string]$sheetIdentity.sheetStem; targetState = $TargetStateName
            previousState = $previousState; missingFields = @($missing)
            changedByUser = $ChangedByUser; rollback = $result.rollback; dryRun = [bool]$DryRun
            notifySkippedDuplicate = $notifySkippedDuplicate; blockedDedupeKey = $blockedDedupeKey
        } | Out-Null
    }

    return $result
}

function Invoke-QCNotificationForStateChange {
    <#
    .SYNOPSIS
    Detects a QC PDF workflow state transition and sends the matching configured notification.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$PreviousState,
        [Parameter(Mandatory)]
        [string]$CurrentState,
        [object]$Document,
        [AllowNull()][AllowEmptyString()][string]$DocumentName = $null,
        [string]$DocumentPath = '',
        [string]$DocumentGuid = '',
        [string]$Project = '',
        [hashtable]$Job,
        [string]$StateTransitionKey = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$SubmittedBy = '',
        [Nullable[int]]$TransitionId = $null,
        [string]$NotificationStateSource = '',
        [switch]$Force
    )

    $settings = Get-QCNotificationSettings -Config $Config
    if (-not [bool]$settings.enabled) {
        $skipped = @{
            success = $false
            skipped = $true
            provider = [string]$settings.provider
            dryRun = [bool]$settings.dryRun
            eventType = $null
            documentName = $DocumentName
            to = @()
            cc = @()
            message = 'Notifications are disabled in configuration.'
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_DISABLED' -Level 'Information' -Message $skipped.message -Result $skipped -Event @{ currentState = $CurrentState; previousState = $PreviousState }
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DISABLED' -Message $skipped.message -Data $skipped
    }

    $prev = if ($PreviousState) { ([string]$PreviousState).Trim() } else { '' }
    $curr = ([string]$CurrentState).Trim()
    if (-not $Force -and $prev -eq $curr) {
        $skipped = @{
            success = $false
            skipped = $true
            message = 'No workflow state change detected.'
            previousState = $prev
            currentState = $curr
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_NO_CHANGE' -Level 'Information' -Message $skipped.message -Result $skipped
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_NO_CHANGE' -Message $skipped.message -Data $skipped
    }

    $events = _QCN-ToHashtable $settings.events
    if (-not $events -or -not $events.ContainsKey($curr)) {
        $skipped = @{
            success = $false
            skipped = $true
            message = "No notification event configured for state '$curr'."
            currentState = $curr
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_NO_EVENT' -Level 'Information' -Message $skipped.message -Result $skipped
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_NO_EVENT' -Message $skipped.message -Data $skipped
    }

    if (-not $Force -and (Get-Command -Name 'Test-QCShouldDeferReadyForQcNotification' -ErrorAction SilentlyContinue)) {
        if (Test-QCShouldDeferReadyForQcNotification -Config $Config -CurrentState $curr) {
            $skipped = @{
                success = $false
                skipped = $true
                message = 'Ready for QC notification deferred until prepend and rendition complete.'
                currentState = $curr
                timestampUtc = Get-QCTimestamp
            }
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_DEFERRED_READY_FOR_QC' -Level 'Information' -Message $skipped.message -Result $skipped
            return New-QCSuccessResult -Code 'QC_NOTIFICATION_DEFERRED_READY_FOR_QC' -Message $skipped.message -Data $skipped
        }
    }

    $eventCfg = _QCN-ToHashtable $events[$curr]
    if (-not $eventCfg -or ($eventCfg.ContainsKey('enabled') -and -not [bool]$eventCfg.enabled)) {
        $skipped = @{
            success = $false
            skipped = $true
            message = "Notification event for state '$curr' is disabled."
            currentState = $curr
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_EVENT_DISABLED' -Level 'Information' -Message $skipped.message -Result $skipped
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_EVENT_DISABLED' -Message $skipped.message -Data $skipped
    }

    if (_QCN-IsBlank $DocumentName) {
        if ($Job) {
            $DocumentName = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
        }
        if (_QCN-IsBlank $DocumentName) {
            $DocumentName = _QCN-GetProp -Object $Document -Names @('Name','DocumentName','FileName')
        }
        if (_QCN-IsBlank $DocumentName) { $DocumentName = 'unknown-document' }
    }
    if (_QCN-IsBlank $DocumentGuid) {
        $DocumentGuid = _QCN-GetProp -Object $Document -Names @('DocumentGUID','DocumentGuid','GUID','Id','DocumentID')
        if ($DocumentGuid) { $DocumentGuid = [string]$DocumentGuid }
    }
    if (_QCN-IsBlank $DocumentPath) {
        $DocumentPath = _QCN-GetProp -Object $Document -Names @('DocumentPath','FullPath','Path')
        if ($DocumentPath) { $DocumentPath = [string]$DocumentPath }
    }
    $eventType = if ($eventCfg.eventType) { [string]$eventCfg.eventType } else { $curr.ToUpperInvariant().Replace(' ', '_') }
    $actionRequired = if ($eventCfg.actionRequired) { [string]$eventCfg.actionRequired } else { '' }
    $sourceJobId = if ($Job -and $Job.ContainsKey('id')) { [string]$Job.id } else { '' }

    $resolvedStKey = $StateTransitionKey
    if ((_QCN-IsBlank $resolvedStKey) -and $Job -and ($Job.metadata -is [hashtable]) -and $Job.metadata.ContainsKey('stateTransitionKey') -and $Job.metadata.stateTransitionKey) {
        $resolvedStKey = [string]$Job.metadata.stateTransitionKey
    }
    $resolvedActor = Resolve-QCNotificationStateChangeActor -Config $Config -StateTransitionKey $resolvedStKey `
        -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername -Job $Job
    if ($resolvedActor) {
        if ($null -ne $resolvedActor.changedByUser) { $ChangedByUser = $resolvedActor.changedByUser }
        if (-not (_QCN-IsBlank $resolvedActor.changedByUsername)) { $ChangedByUsername = [string]$resolvedActor.changedByUsername }
    }

    $actorFromJob = _QCN-ResolveStateChangeActorFromJob -Job $Job
    if (_QCN-IsBlank $SubmittedBy -and -not (_QCN-IsBlank $actorFromJob.submittedBy)) {
        $SubmittedBy = [string]$actorFromJob.submittedBy
    }

    $roleOverrides = $null
    if ($Job -and $Job.ContainsKey('metadata') -and $Job.metadata) {
        $md = _QCN-ToHashtable $Job.metadata
        if ($md -and ($md.ContainsKey('attributes') -or $md.ContainsKey('designerEmail'))) {
            $attrs = if ($md.attributes) { _QCN-ToHashtable $md.attributes } else { $md }
            if ($attrs) {
                $roleOverrides = @{
                    designerEmail = if ($attrs.designerEmail) { [string]$attrs.designerEmail } else { '' }
                    reviewerEmail = if ($attrs.reviewerEmail) { [string]$attrs.reviewerEmail } else { '' }
                    checkerEmail = if ($attrs.checkerEmail) { [string]$attrs.checkerEmail } else { '' }
                }
            }
        }
    }

    $triggerDocumentGuid = [string]$DocumentGuid
    $triggerDocumentName = [string]$DocumentName
    $qcTarget = _QCN-ResolveQcPdfNotificationTarget -Document $Document -Config $Config -Job $Job `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -DocumentPath $DocumentPath
    $Document = $qcTarget.document
    $DocumentName = [string]$qcTarget.documentName
    $DocumentGuid = [string]$qcTarget.documentGuid
    $qcPdfResolutionSource = if ($qcTarget.resolutionSource) { [string]$qcTarget.resolutionSource } else { '' }
    if (-not (_QCN-IsBlank $qcTarget.documentPath)) { $DocumentPath = [string]$qcTarget.documentPath }

    if (-not $Force -and -not (_QCN-IsBlank $DocumentName) -and (Get-Command -Name 'Test-QCShouldNotifyForSheetPackageMember' -ErrorAction SilentlyContinue)) {
        if (-not (Test-QCShouldNotifyForSheetPackageMember -Config $Config -DocumentName $DocumentName)) {
            $skipped = @{
                success = $false
                skipped = $true
                message = 'Notification skipped: sheet package notifies from QC PDF only.'
                documentName = $DocumentName
                currentState = $curr
                timestampUtc = Get-QCTimestamp
            }
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_PACKAGE_MEMBER' -Level 'Information' -Message $skipped.message -Result $skipped -Event @{
                documentName = $DocumentName; currentState = $curr; previousState = $prev
            }
            return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_PACKAGE_MEMBER' -Message $skipped.message -Data $skipped
        }
    }

    $folderForRoles = ''
    $sourceForRoles = $DocumentName
    if ($Job) {
        $folderForRoles = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath'))
        $sourceForRoles = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
        $roleFromJob = [string](_QCN-GetJobValue -Job $Job -Keys @('roleSourceDocumentName'))
        if (-not (_QCN-IsBlank $roleFromJob)) { $sourceForRoles = $roleFromJob }
        elseif ($Job.metadata) {
            $mdRole = _QCN-ToHashtable $Job.metadata
            if ($mdRole -and $mdRole.roleSourceDocumentName) { $sourceForRoles = [string]$mdRole.roleSourceDocumentName }
        }
    }
    if (_QCN-IsBlank $folderForRoles) { $folderForRoles = [string](_QCN-GetProp -Object $Document -Names @('FolderPath', 'folderPath')) }
    if ((_QCN-IsBlank $folderForRoles) -and (-not (_QCN-IsBlank $DocumentPath)) -and ($DocumentPath -match '\\')) {
        $folderForRoles = [System.IO.Path]::GetDirectoryName($DocumentPath)
    }
    if (_QCN-IsBlank $Project) {
        if ($Config -and -not (_QCN-IsBlank $folderForRoles) -and (Get-Command -Name 'Get-QCProjectNameFromFolderPath' -ErrorAction SilentlyContinue)) {
            try {
                $pn = Get-QCProjectNameFromFolderPath -Config $Config -FolderPath $folderForRoles
                if ($pn) { $Project = [string]$pn }
            } catch { }
        }
        if (_QCN-IsBlank $Project) {
            $pwProj = _QCN-GetProp -Object $Document -Names @('ProjectName', 'Project')
            if ($pwProj) { $Project = [string]$pwProj }
        }
    }
    if (_QCN-IsBlank $sourceForRoles) { $sourceForRoles = [string](_QCN-GetProp -Object $Document -Names @('Name', 'DocumentName', 'FileName')) }
    if (-not (_QCN-IsBlank $sourceForRoles) -and (Test-QCIsQcPdfDocumentName -DocumentName $sourceForRoles)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            try {
                $stem = Get-PWSheetStemFromDocumentName -DocumentName $sourceForRoles
                if (-not (_QCN-IsBlank $stem)) { $sourceForRoles = $stem + '.pdf' }
            } catch { }
        } else {
            $sourceForRoles = [System.IO.Path]::GetFileNameWithoutExtension($sourceForRoles) + '.pdf'
        }
    }

    $resolved = Resolve-QCNotificationRecipients -Document $Document -Settings $settings -ToRoles @($eventCfg.to) -CcRoles @($eventCfg.cc) `
        -Config $Config -Job $Job -RoleOverrides $roleOverrides -FolderPath $folderForRoles -SourceDocumentName $sourceForRoles `
        -DocumentGuid $DocumentGuid
    $recipientKey = _QCN-BuildRecipientKey -To @($resolved.to) -Cc @($resolved.cc)
    $event = New-QCNotificationEvent -EventType $eventType -Project $Project -DocumentName $DocumentName `
        -DocumentPath $DocumentPath -DocumentGuid ([string]$DocumentGuid) -PreviousState $prev -CurrentState $curr `
        -Reviewers $resolved.reviewers -Designers $resolved.designers -Cc $resolved.cc -ActionRequired $actionRequired -SourceJobId $sourceJobId
    if (-not (_QCN-IsBlank $folderForRoles)) { $event['folderPath'] = $folderForRoles }
    $event['notificationType'] = $eventType
    $event['targetState'] = $curr
    $event['recipientKey'] = $recipientKey
    $event['to'] = @($resolved.to)
    $event['cc'] = @($resolved.cc)
    if (-not (_QCN-IsBlank $StateTransitionKey)) { $event['stateTransitionKey'] = [string]$StateTransitionKey }
    elseif ($Job -and ($Job.metadata -is [hashtable]) -and $Job.metadata.ContainsKey('stateTransitionKey') -and $Job.metadata.stateTransitionKey) {
        $event['stateTransitionKey'] = [string]$Job.metadata.stateTransitionKey
    }
    if ($event.ContainsKey('stateTransitionKey') -and -not (_QCN-IsBlank $event.stateTransitionKey)) {
        $auditId = _QCN-GetNotificationAuditEventId -Event $event
        if ($null -ne $auditId) { $event['auditEventId'] = $auditId }
    }
    $resolvedCycleId = _QCN-ResolveNotificationCycleId -Event $event -Config $Config -Job $Job
    if (-not (_QCN-IsBlank $resolvedCycleId)) { $event['cycleId'] = $resolvedCycleId }

    $folderForRt = ''
    $sourceForRt = $DocumentName
    if ($Job) {
        $folderForRt = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath'))
        $sourceForRt = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
    }
    if (_QCN-IsBlank $folderForRt) { $folderForRt = [string](_QCN-GetProp -Object $Document -Names @('FolderPath', 'folderPath')) }
    if (_QCN-IsBlank $sourceForRt) { $sourceForRt = [string](_QCN-GetProp -Object $Document -Names @('Name', 'DocumentName', 'FileName')) }
    $resolvedReviewType = _QCN-ResolveNotificationReviewType -Document $Document -Settings $settings -Config $Config -Job $Job `
        -FolderPath $folderForRt -SourceName $sourceForRt -DocumentName $DocumentName
    $resolvedProcessType = ''
    if ($Job -and $Job.metadata) {
        $jobMdForPt = _QCN-ToHashtable $Job.metadata
        if ($jobMdForPt -and $jobMdForPt.ContainsKey('qcProcessType') -and -not (_QCN-IsBlank $jobMdForPt['qcProcessType'])) {
            $resolvedProcessType = [string]$jobMdForPt['qcProcessType']
        }
    }
    if (_QCN-IsBlank $resolvedProcessType) {
        $resolvedProcessType = _QCN-ResolveNotificationProcessType -Event $event -Config $Config -Job $Job
    }
    if (-not (_QCN-IsBlank $resolvedProcessType)) {
        $event['qcProcessType'] = $resolvedProcessType
    }
    if ((_QCN-IsBlank $resolvedReviewType) -or (-not (_QCN-TestUsableNotificationReviewTypeValue -Value $resolvedReviewType))) {
        $fromProcess = _QCN-ResolveReviewTypeLabelFromProcessType -ProcessType $resolvedProcessType
        if (-not (_QCN-IsBlank $fromProcess)) { $resolvedReviewType = $fromProcess }
    }
    if (-not (_QCN-IsBlank $resolvedReviewType)) {
        $event['reviewType'] = $resolvedReviewType
        $event['qcReviewType'] = $resolvedReviewType
    }

    $resolvedSheetPackageId = _QCN-ResolveNotificationSheetPackageId -Job $Job -Event $event
    if (_QCN-IsBlank $resolvedSheetPackageId -and -not (_QCN-IsBlank $DocumentGuid) `
            -and (Get-Command -Name 'Get-SheetPackageIdForDocument' -ErrorAction SilentlyContinue) `
            -and (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) `
            -and (Test-QCDatabaseEnabled -Config $Config)) {
        try {
            $pkgFromDoc = Get-SheetPackageIdForDocument -Config $Config -DocumentGuid $DocumentGuid
            if ($null -ne $pkgFromDoc) { $resolvedSheetPackageId = $pkgFromDoc.ToString() }
        } catch { }
    }
    if (-not (_QCN-IsBlank $resolvedSheetPackageId)) {
        $event['sheetPackageId'] = $resolvedSheetPackageId
    }

    $resolvedSubmittedBy = Resolve-QCNotificationSubmittedBy -Config $Config -ChangedByUser $ChangedByUser `
        -ChangedByUsername $ChangedByUsername -SubmittedBy $SubmittedBy
    $event['submittedBy'] = $resolvedSubmittedBy
    if ($null -ne $ChangedByUser) { $event['changedByUser'] = $ChangedByUser }
    if (-not (_QCN-IsBlank $ChangedByUsername)) { $event['changedByUsername'] = [string]$ChangedByUsername }

    $sheetIdentity = _QCN-ResolveSheetNotificationIdentity -DocumentName $DocumentName -DocumentGuid $DocumentGuid -Config $Config
    if (-not (_QCN-IsBlank $sheetIdentity.sheetStem)) {
        $event['sheetStem'] = [string]$sheetIdentity.sheetStem
    }

    $resolvedTransitionId = $TransitionId
    if (($null -eq $resolvedTransitionId -or $resolvedTransitionId -le 0) -and $Job -and ($Job.metadata -is [hashtable])) {
        $jobMd = _QCN-ToHashtable $Job.metadata
        if ($jobMd -and $jobMd.ContainsKey('transitionId') -and $null -ne $jobMd.transitionId) {
            try { $resolvedTransitionId = [int]$jobMd.transitionId } catch { }
        }
    }
    if ($null -ne $resolvedTransitionId -and $resolvedTransitionId -gt 0) {
        $event['transitionId'] = $resolvedTransitionId
    }
    if ($Job -and ($Job.metadata -is [hashtable])) {
        $jobMdForEvent = _QCN-ToHashtable $Job.metadata
        if ($jobMdForEvent) {
            foreach ($laneKey in @('notificationLaneDocumentGuid', 'laneQcPdfDocumentGuid', 'expectedLanePdfName', 'notificationLaneDocumentName')) {
                if ($jobMdForEvent.ContainsKey($laneKey) -and -not (_QCN-IsBlank $jobMdForEvent[$laneKey])) {
                    $event[$laneKey] = [string]$jobMdForEvent[$laneKey]
                }
            }
            if ($jobMdForEvent.ContainsKey('transitionSource') -and -not (_QCN-IsBlank $jobMdForEvent.transitionSource)) {
                $event['transitionSource'] = [string]$jobMdForEvent.transitionSource
            } elseif ($jobMdForEvent.ContainsKey('notificationStateSource') -and -not (_QCN-IsBlank $jobMdForEvent.notificationStateSource)) {
                $event['transitionSource'] = [string]$jobMdForEvent.notificationStateSource
            }
            if ($jobMdForEvent.ContainsKey('transitionGroupId') -and -not (_QCN-IsBlank $jobMdForEvent.transitionGroupId)) {
                $event['transitionGroupId'] = [string]$jobMdForEvent.transitionGroupId
            }
            if ($jobMdForEvent.ContainsKey('sheetPackageId') -and -not (_QCN-IsBlank $jobMdForEvent.sheetPackageId)) {
                $event['sheetPackageId'] = [string]$jobMdForEvent.sheetPackageId
            }
            if ($jobMdForEvent.ContainsKey('auditEventId') -and $null -ne $jobMdForEvent.auditEventId) {
                try {
                    $jobAudit = [long]$jobMdForEvent.auditEventId
                    if ($jobAudit -gt 0) { $event['auditEventId'] = $jobAudit }
                } catch { }
            }
        }
    }
    if (-not (_QCN-IsBlank $NotificationStateSource)) {
        $event['notificationStateSource'] = [string]$NotificationStateSource
        if (-not $event.ContainsKey('transitionSource') -or (_QCN-IsBlank $event.transitionSource)) {
            $event['transitionSource'] = [string]$NotificationStateSource
        }
    }

    $resolvedNotificationStateSource = if (-not (_QCN-IsBlank $NotificationStateSource)) {
        [string]$NotificationStateSource
    } else {
        'currentStateParameter'
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_NOTIFICATION_STATE_RESOLVED' `
            -Message 'Notification template state resolved before send.' -Data @{
            documentGuid = $DocumentGuid
            documentName = $DocumentName
            previousState = $prev
            notificationStateSource = $resolvedNotificationStateSource
            notificationStateValue = $curr
            eventType = $eventType
            transitionId = $resolvedTransitionId
        } | Out-Null
    }

    $dedupeKey = Get-QCNotificationDedupeKey -Event $event -Settings $settings -Config $Config -Job $Job
    $event['dedupeKey'] = $dedupeKey
    if ($Job -and $Job.ContainsKey('id') -and -not (_QCN-IsBlank $Job.id)) {
        $Job['dedupeKey'] = $dedupeKey
    }
    $displayDocumentName = _QCN-ResolveNotificationDisplayDocumentName -DocumentName $DocumentName -Job $Job -Event $event -Config $Config
    if (-not (_QCN-IsBlank $displayDocumentName)) {
        $event['displayDocumentName'] = $displayDocumentName
    }
    $excludeJobId = if ($Job -and $Job.ContainsKey('id')) { [string]$Job.id } else { '' }
    if (-not $Force -and (Test-QCNotificationDedupe -DedupeKey $dedupeKey -Settings $settings -Config $Config -TransitionId $resolvedTransitionId -ExcludeJobId $excludeJobId)) {
        $skipCode = if (_QCN-TestNotificationActorIsAutomation -Config $Config -Event $event) { 'QC_NOTIFICATION_SKIPPED_AUTOMATION_ECHO' } else { 'QC_NOTIFICATION_DEDUPED' }
        _QCN-WriteNotificationLifecycleLog -Code $skipCode -Level 'Information' `
            -Message 'Duplicate workflow notification suppressed for logical sheet transition.' `
            -Event $event -Config $Config -Job $Job -To @($resolved.to) -Cc @($resolved.cc)
        $skipped = @{
            success = $false
            skipped = $true
            dedupeKey = $dedupeKey
            message = 'Duplicate notification suppressed.'
            eventType = $eventType
            documentName = $DocumentName
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code $skipCode -Level 'Information' -Message $skipped.message -Result $skipped -Event $event
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Message $skipped.message -Data $skipped
    }
    if (-not $Force) {
        $claimData = @{
            eventType = $eventType
            documentName = $DocumentName
            provider = [string]$settings.provider
        }
        if (-not (Register-QCNotificationDedupeClaim -DedupeKey $dedupeKey -Settings $settings -Config $Config `
                -TransitionId $resolvedTransitionId -ExcludeJobId $excludeJobId -ResultData $claimData)) {
            _QCN-WriteNotificationLifecycleLog -Code 'QC_NOTIFICATION_DEDUPED' -Level 'Information' `
                -Message 'Duplicate workflow notification suppressed during dedupe claim.' `
                -Event $event -Config $Config -Job $Job -To @($resolved.to) -Cc @($resolved.cc)
            $skipped = @{
                success = $false
                skipped = $true
                dedupeKey = $dedupeKey
                message = 'Duplicate notification suppressed during dedupe claim.'
                eventType = $eventType
                documentName = $DocumentName
                timestampUtc = Get-QCTimestamp
            }
            Write-QCNotificationResult -Code 'QC_NOTIFICATION_DEDUPED' -Level 'Information' -Message $skipped.message -Result $skipped -Event $event
            return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Message $skipped.message -Data $skipped
        }
    }

    $subjectDocumentName = if (-not (_QCN-IsBlank $displayDocumentName)) { $displayDocumentName } else { $DocumentName }
    $tokens = _QCN-NewNotificationSubjectTokens -DocumentName $subjectDocumentName -DocumentPath $DocumentPath `
        -Project $Project -PreviousState $prev -CurrentState $curr -EventType $eventType -ReviewType $resolvedReviewType
    $subjectTemplate = _QCN-ResolveNotificationSubjectTemplate -EventCfg $eventCfg -Settings $settings
    $subject = Expand-QCNotificationTemplate -Template $subjectTemplate -Tokens $tokens

    $qcUrl = Resolve-QCNotificationQcPdfUrl -Event $event -Settings $settings -Document $Document -Config $Config
    if (-not (_QCN-IsBlank $qcUrl)) {
        $event['qcPdfUrl'] = $qcUrl
    }
    $linkResolutionSource = if ($event.linkResolutionSource) { [string]$event.linkResolutionSource } else { '' }
    $linkDocumentGuid = _QCN-ExtractPwLinkDocumentGuid -Url $qcUrl
    if (_QCN-IsBlank $linkDocumentGuid) { $linkDocumentGuid = [string]$DocumentGuid }
    $sheetPackageIdForLog = if ($event.sheetPackageId) { [string]$event.sheetPackageId } else { '' }
    _QCN-WriteQcPdfNotificationGuidResolutionLog -Config $Config `
        -TriggerDocumentGuid $triggerDocumentGuid -ResolvedQcPdfGuid $DocumentGuid `
        -NotificationDocumentGuid $DocumentGuid -LinkDocumentGuid $linkDocumentGuid `
        -SheetPackageId $sheetPackageIdForLog -QcReviewType $resolvedReviewType `
        -ResolutionSource $qcPdfResolutionSource -LinkResolutionSource $linkResolutionSource `
        -Event $event -Job $Job
    if ((-not (_QCN-IsBlank $linkDocumentGuid)) -and (-not (_QCN-IsBlank $DocumentGuid)) `
            -and ($linkDocumentGuid -ne $DocumentGuid) -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Level 'Warning' -Code 'QC_NOTIFICATION_QC_PDF_LINK_GUID_MISMATCH' `
            -Message 'Email link GUID differs from resolved notification QC PDF GUID.' -Data @{
            notification_document_guid = [string]$DocumentGuid
            link_document_guid = [string]$linkDocumentGuid
            trigger_document_guid = [string]$triggerDocumentGuid
            resolution_source = [string]$qcPdfResolutionSource
            link_resolution_source = [string]$linkResolutionSource
            sheet_package_id = [string]$sheetPackageIdForLog
            document_name = [string]$DocumentName
        } | Out-Null
    }
    $event['_eventCfg'] = $eventCfg
    $event['_document'] = $Document

    $send = _QCN-EnsureSingleResult (Send-QCNotification -Event $event -Config $Config -Subject $subject -To $resolved.to -Cc $resolved.cc)

    $resultData = _QCN-ToHashtable $send.Data
    if ($send.IsSuccess -and $resultData -and $resultData.success -eq $true) {
        Register-QCNotificationDedupe -DedupeKey $dedupeKey -Settings $settings -ResultData $resultData
        Register-QCNotificationSheetPackageDedupe -Event $event -Settings $settings -Config $Config -Job $Job -ResultData $resultData
        if ($null -ne $resolvedTransitionId -and $resolvedTransitionId -gt 0 -and (Get-Command -Name 'Update-QCTransitionEventNotification' -ErrorAction SilentlyContinue)) {
            try {
                Update-QCTransitionEventNotification -Config $Config -TransitionId $resolvedTransitionId -NotificationSent $true -NotificationId $dedupeKey
            } catch { }
        }
    } elseif (-not $Force) {
        Remove-QCNotificationDedupeKey -DedupeKey $dedupeKey -Settings $settings | Out-Null
    }
    if ($resultData) {
        $resultData['dedupeKey'] = $dedupeKey
        if ($null -ne $resolvedTransitionId -and $resolvedTransitionId -gt 0) {
            $resultData['transitionId'] = $resolvedTransitionId
        }
    }
    $notifyCode = if ($send.IsSuccess) { 'QC_NOTIFICATION_SENT' } elseif ($send.Code) { [string]$send.Code } else { 'QC_NOTIFICATION_FAILED' }
    $notifyMessage = [string]$send.Message
    if ($send.IsSuccess) {
        return New-QCSuccessResult -Code $notifyCode -Message $notifyMessage -Data $resultData
    }
    return New-QCFailureResult -Code $notifyCode -Message $notifyMessage -Data $resultData
}

Export-ModuleMember -Function Get-QCNotificationSettings, Test-QCNotificationsEnqueueAsJob, New-QCNotificationEvent, Resolve-QCNotificationRecipients, `
    Resolve-QCNotificationQcPdfUrl, Resolve-QCNotificationSubmittedBy, Resolve-QCNotificationStateChangeActor, Get-QCNotificationSheetTransitionKey, Get-QCNotificationDedupeKey, Get-QCNotificationSheetPackageDedupeKey, Test-QCNotificationDedupe, Test-QCNotificationSheetPackageAlreadySent, Register-QCNotificationDedupe, Register-QCNotificationSheetPackageDedupe, Register-QCNotificationDedupeClaim, Remove-QCNotificationDedupeKey, `
    Get-QCStateChangeMissingEmailFields, Get-QCWorkflowTransitionMissingEmailFields, Test-QCPrependBlockedByMissingEmailAttributes, `
    Resolve-QCWorkflowRollbackPreviousState, Resolve-QCStateChangeActorEmailAddress, Send-QCStateChangeBlockedNotification, Invoke-QCWorkflowStateEmailAttributeGate, `
    Send-QCNotification, Invoke-QCNotificationForStateChange, Write-QCNotificationResult, Test-QCNotificationResultSent
