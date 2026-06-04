# QC.Notifications.psm1
# Responsibility: Configurable QC workflow email notifications (Mock + future Microsoft Graph).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Config.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationTemplates.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationMock.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationGraph.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PW.Discovery.psm1') -Force -ErrorAction SilentlyContinue
# Core.Database must be imported by the caller. Re-importing with -Force
# here clobbers the caller's global-scope exports.

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
        try { if ($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n) { return $Object.$n } } catch { }
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
    $root = $PSScriptRoot
    if ($root -match '[\\/]modules$') { return Split-Path -Parent $root }
    return $root
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
            keyFields = @('documentGuid', 'eventType', 'currentState')
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
            'Corrections Received' = @{
                enabled = $true
                eventType = 'CORRECTIONS_RECEIVED'
                to = @('reviewers')
                cc = @('designers')
                actionRequired = 'Reviewer to verify corrections and approve or return redlines.'
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
        foreach ($key in @('changedByUsername', 'lastActionBy', 'userName', 'triggeredBy')) {
            if (_QCN-IsBlank $changedByUsername -and $md.ContainsKey($key) -and -not (_QCN-IsBlank $md[$key])) {
                $changedByUsername = [string]$md[$key]
            }
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

    $docGuid = if ($Event.documentGuid) { [string]$Event.documentGuid.Trim() } else { '' }
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

    $baseUrl = if ($emailCfg.pwLinkBaseUrl) { [string]$emailCfg.pwLinkBaseUrl } else { '' }
    if (-not (_QCN-IsBlank $baseUrl) -and -not (_QCN-IsBlank $docGuid)) {
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
        if (-not (_QCN-IsBlank $datasourceName)) {
            $dsParam = _QCN-GetEncodedPwDatasource -DatasourceName $datasourceName
            $app = if ($emailCfg.pwLinkApp) { [string]$emailCfg.pwLinkApp } else { 'pwe' }
            $sep = if ($baseUrl.Contains('?')) { '&' } else { '?' }
            return ('{0}{1}objectId={2}&objectType=doc&datasource={3}&app={4}' -f $baseUrl.TrimEnd('/'), $sep, $docGuid, $dsParam, $app)
        }
    }

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

function _QCN-ResolveNotificationReviewType {
    param(
        [object]$Document,
        [hashtable]$Settings,
        [hashtable]$Config = $null,
        [hashtable]$Job = $null,
        [string]$FolderPath = '',
        [string]$SourceName = ''
    )

    $attr = _QCN-ToHashtable $Settings.attributes
    if (-not $attr) { $attr = @{} }

    if ($attr.ContainsKey('reviewTypeField') -and -not (_QCN-IsBlank $attr['reviewTypeField'])) {
        $fromNotify = _QCN-GetAttributeValue -Document $Document -AttributeName ([string]$attr['reviewTypeField'])
        if (-not (_QCN-IsBlank $fromNotify)) { return ([string]$fromNotify).Trim() }
    }

    $qcReviewCol = _QCN-GetQcReviewTypeAttributeName -Config $Config
    $fromDoc = _QCN-GetAttributeValue -Document $Document -AttributeName $qcReviewCol
    if (-not (_QCN-IsBlank $fromDoc)) { return ([string]$fromDoc).Trim() }

    if ($Job) {
        $fromJob = [string](_QCN-GetJobValue -Job $Job -Keys @('reviewType', 'qcReviewType'))
        if (-not (_QCN-IsBlank $fromJob)) { return $fromJob.Trim() }
    }

    if ($Config -and -not (_QCN-IsBlank $FolderPath) -and -not (_QCN-IsBlank $SourceName)) {
        if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
            $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceName -Config $Config
            if ($pw.found -and -not (_QCN-IsBlank $pw.qcReviewType)) { return ([string]$pw.qcReviewType).Trim() }
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
        [string]$DocumentGuid
    )

    if (_QCN-IsBlank $DocumentGuid) { return $null }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return $null }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, checker_email
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if (-not $res.IsSuccess -or -not $res.Data.table -or $res.Data.table.Rows.Count -eq 0) { return $null }
        $r = $res.Data.table.Rows[0]
        return @{
            designerEmail = if ($r.designer_email -is [DBNull]) { '' } else { [string]$r.designer_email }
            reviewerEmail = if ($r.reviewer_email -is [DBNull]) { '' } else { [string]$r.reviewer_email }
            checkerEmail = if ($r.checker_email -is [DBNull]) { '' } else { [string]$r.checker_email }
        }
    } catch { return $null }
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
    if (-not (_QCN-IsBlank $sourceName) -and $sourceName -match '(?i)-qc\.pdf$') {
        $sourceName = [string]([System.IO.Path]::GetFileNameWithoutExtension($sourceName)) + '.pdf'
    }

    $designer = ''
    $reviewer = ''
    $checker = ''
    if ($RoleOverrides) {
        if ($RoleOverrides.ContainsKey('designerEmail')) { $designer = [string]$RoleOverrides.designerEmail }
        if ($RoleOverrides.ContainsKey('reviewerEmail')) { $reviewer = [string]$RoleOverrides.reviewerEmail }
        if ($RoleOverrides.ContainsKey('checkerEmail')) { $checker = [string]$RoleOverrides.checkerEmail }
    }
    if ($Job) {
        if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetJobValue -Job $Job -Keys @('designerEmail','qcDesignerEmail')) }
        if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetJobValue -Job $Job -Keys @('reviewerEmail','qcReviewerEmail')) }
        if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetJobValue -Job $Job -Keys @('checkerEmail','qcCheckerEmail')) }
    }
    if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $designerField) }
    if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $reviewerField) }
    if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $checkerField) }
    if (_QCN-IsBlank $designer) { $designer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcDesignerField) }
    if (_QCN-IsBlank $reviewer) { $reviewer = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcReviewerField) }
    if (_QCN-IsBlank $checker) { $checker = [string](_QCN-GetAttributeValue -Document $Document -AttributeName $qcCheckerField) }

    if ($Config -and -not (_QCN-IsBlank $folderPath) -and -not (_QCN-IsBlank $sourceName)) {
        $needSource = (_QCN-IsBlank $designer) -or (_QCN-IsBlank $reviewer) -or (_QCN-IsBlank $checker)
        if ($needSource) {
            if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
                $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $folderPath -SourceDocumentName $sourceName -Config $Config
                if ($pw.found) {
                    if (_QCN-IsBlank $designer) { $designer = [string]$pw.designerEmail }
                    if (_QCN-IsBlank $reviewer) { $reviewer = [string]$pw.reviewerEmail }
                    if (_QCN-IsBlank $checker) { $checker = [string]$pw.checkerEmail }
                }
            }
        }
    }

    if ($Config -and (-not (_QCN-IsBlank $DocumentGuid))) {
        $needIndex = (_QCN-IsBlank $designer) -or (_QCN-IsBlank $reviewer) -or (_QCN-IsBlank $checker)
        if ($needIndex) {
            $idx = _QCN-GetRoleEmailsFromSheetIndex -Config $Config -DocumentGuid $DocumentGuid
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

function _QCN-ResolveQcPdfNotificationTarget {
    param(
        [object]$Document,
        [hashtable]$Config,
        [hashtable]$Job,
        [string]$DocumentName,
        [string]$DocumentGuid,
        [string]$DocumentPath
    )

    $out = @{
        document = $Document
        documentName = $DocumentName
        documentGuid = $DocumentGuid
        documentPath = $DocumentPath
    }
    if ($DocumentName -match '(?i)-qc\.pdf$') { return $out }

    $folderPath = $DocumentPath
    $triggerName = $DocumentName
    if ($Job) {
        if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder','folderPath','incomingFolderPath')) }
        if (_QCN-IsBlank $triggerName) { $triggerName = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName','sourceDocumentName','incomingDocName')) }
    }
    if (_QCN-IsBlank $folderPath) { $folderPath = [string](_QCN-GetProp -Object $Document -Names @('FolderPath','folderPath')) }
    if (_QCN-IsBlank $triggerName) { $triggerName = [string](_QCN-GetProp -Object $Document -Names @('Name','DocumentName')) }

    if ($Config -and -not (_QCN-IsBlank $folderPath) -and -not (_QCN-IsBlank $triggerName) -and (Get-Command -Name 'Get-PWAssociatedSheetMembers' -ErrorAction SilentlyContinue)) {
        $guid = $DocumentGuid
        if (_QCN-IsBlank $guid) { $guid = [string](_QCN-GetProp -Object $Document -Names @('DocumentGUID','DocumentGuid','GUID')) }
        try {
            $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $folderPath -DocumentName $triggerName -DocumentGuid $guid)
            foreach ($m in $members) {
                $dn = [string]$m.documentName
                if ($dn -match '(?i)-qc\.pdf$') {
                    $out.documentName = $dn
                    $out.documentGuid = [string]$m.documentGuid
                    $out.documentPath = if ($folderPath) { ($folderPath.TrimEnd('\') + '\' + $dn) } else { $DocumentPath }
                    if ($m.document) { $out.document = $m.document }
                    return $out
                }
            }
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($triggerName)
            if ($stem) {
                $qcName = $stem + '-qc.pdf'
                $out.documentName = $qcName
                $out.documentPath = if ($folderPath) { ($folderPath.TrimEnd('\') + '\' + $qcName) } else { $DocumentPath }
            }
        } catch { }
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
    $ccFromAttr = _QCN-ParseEmailList (_QCN-GetAttributeValue -Document $Document -AttributeName $ccField)

    $to = [System.Collections.Generic.List[string]]::new()
    $cc = [System.Collections.Generic.List[string]]::new()

    foreach ($role in @($ToRoles)) {
        switch -Regex ($role) {
            '^reviewers?$' { foreach ($e in $reviewers) { if (-not $to.Contains($e)) { $to.Add($e) | Out-Null } } }
            '^designers?$' { foreach ($e in $designers) { if (-not $to.Contains($e)) { $to.Add($e) | Out-Null } } }
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
        to = @($to)
        cc = @($cc)
    }
}

function Get-QCNotificationDedupeKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [hashtable]$Settings
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    $fields = if ($dedupe -and $dedupe.keyFields) { @($dedupe.keyFields) } else { @('documentGuid', 'eventType', 'currentState') }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($field in @($fields)) {
        $value = ''
        switch ([string]$field) {
            'documentGuid' {
                $value = if ($Event.documentGuid) { [string]$Event.documentGuid }
                elseif ($Event.documentPath) { [string]$Event.documentPath }
                else { [string]$Event.documentName }
            }
            'documentName' { $value = [string]$Event.documentName }
            'documentPath' { $value = [string]$Event.documentPath }
            'eventType' { $value = [string]$Event.eventType }
            'currentState' { $value = [string]$Event.currentState }
            'previousState' { $value = [string]$Event.previousState }
            'project' { $value = [string]$Event.project }
            'stateTransitionKey' { $value = [string]$Event.stateTransitionKey }
            default {
                if ($Event.ContainsKey($field)) { $value = [string]$Event[$field] }
            }
        }
        $parts.Add(('{0}={1}' -f $field, $value)) | Out-Null
    }
    # Per-transition dedupe when audit/processor supplies a key (all configured state emails).
    if (-not ($fields -contains 'stateTransitionKey') -and $Event.ContainsKey('stateTransitionKey') -and -not (_QCN-IsBlank $Event.stateTransitionKey)) {
        $parts.Add(('stateTransitionKey=' + [string]$Event.stateTransitionKey)) | Out-Null
    }
    return ($parts -join '|')
}

function Test-QCNotificationDedupe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DedupeKey,
        [hashtable]$Settings
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $dedupe = _QCN-ToHashtable $Settings.dedupe
    if (-not $dedupe -or -not [bool]$dedupe.enabled) { return $false }

    $storePath = if ($dedupe.storePath) { [string]$dedupe.storePath } else { (Join-Path (_QCN-GetRepoRoot) 'notifications\dedupe\sent-keys.jsonl') }
    if (-not (Test-Path -LiteralPath $storePath)) { return $false }

    try {
        $lines = Get-Content -LiteralPath $storePath -ErrorAction Stop
        foreach ($line in @($lines)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $row = $line | ConvertFrom-Json -ErrorAction Stop
                if ($row.key -eq $DedupeKey) { return $true }
            } catch { }
        }
    } catch { }
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

    $storePath = if ($dedupe.storePath) { [string]$dedupe.storePath } else { (Join-Path (_QCN-GetRepoRoot) 'notifications\dedupe\sent-keys.jsonl') }
    $dir = Split-Path -Parent $storePath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $entry = @{
        key = $DedupeKey
        timestampUtc = Get-QCTimestamp
        eventType = $ResultData.eventType
        documentName = $ResultData.documentName
        provider = $ResultData.provider
    } | ConvertTo-Json -Compress

    Add-Content -LiteralPath $storePath -Value $entry -Encoding UTF8
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
        $tokens = _QCN-NewNotificationSubjectTokens -DocumentName ([string]$Event.documentName) `
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

    $sendResult = $null
    switch ($provider.ToLowerInvariant()) {
        'microsoftgraph' {
            $graph = _QCN-ToHashtable $settings.graph
            if (-not $graph) { $graph = @{} }
            $sendResult = Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:([bool]$settings.dryRun)
        }
        default {
            $outputRoot = if ($settings.outputRoot) { [string]$settings.outputRoot } else { (Join-Path (_QCN-GetRepoRoot) 'notifications') }
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

    $code = if ($sendResult.IsSuccess) { 'QC_NOTIFICATION_SENT' } else { 'QC_NOTIFICATION_FAILED' }
    $level = if ($sendResult.IsSuccess) { 'Information' } else { 'Warning' }
    Write-QCNotificationResult -Code $code -Level $level -Message $sendResult.Message -Result $result -Event $Event

    if (Get-Command -Name Write-QCNotificationTelemetry -ErrorAction SilentlyContinue) {
        [void](Write-QCNotificationTelemetry -Config $Config -EventType ([string]$Event.eventType) `
            -DocumentGuid ([string]$Event.documentGuid) -DocumentName ([string]$Event.documentName) `
            -FolderPath ([string]$Event.documentPath) `
            -Recipients ((@($To) + @($Cc)) -join ';') -Subject $Subject `
            -Provider ([string]$provider) -Success $sendResult.IsSuccess `
            -ErrorMessage $(if (-not $sendResult.IsSuccess) { [string]$sendResult.Message } else { $null }))
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

    $missing = @(Get-QCStateChangeMissingEmailFields -Config $Config -TargetStateName $TargetStateName `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -FolderPath $FolderPath)
    if ($missing.Count -eq 0) { return $result }

    $result.blocked = $true
    $result.missingFields = @($missing)

    $previousState = ''
    foreach ($member in $Members) {
        $dn = [string]$member.documentName
        if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$')) {
            $dg = [string]$member.documentGuid
            if ($dg) { $previousState = _QCN-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg }
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($previousState) -and $DocumentGuid) {
        $previousState = _QCN-GetSheetIndexPwStateName -Config $Config -DocumentGuid $DocumentGuid
    }

    if (Get-Command -Name 'Revert-PWAssociatedSheetWorkflowStates' -ErrorAction SilentlyContinue) {
        $result.rollback = Revert-PWAssociatedSheetWorkflowStates -Config $Config -Members $Members `
            -StateByGuid $StateByGuid -FolderPath $FolderPath -TargetStateName $TargetStateName -DryRun:$DryRun
    }

    $docPath = if ($FolderPath -and $DocumentName) { ($FolderPath.TrimEnd('\') + '\' + $DocumentName) } else { '' }
    $result.notification = Send-QCStateChangeBlockedNotification -Config $Config -MissingFields $missing `
        -TargetStateName $TargetStateName -PreviousStateName $previousState -DocumentName $DocumentName `
        -DocumentPath $docPath -DocumentGuid $DocumentGuid -ChangedByUser $ChangedByUser `
        -ChangedByUsername $ChangedByUsername -DryRun:$DryRun

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_STATE_CHANGE_BLOCKED_MISSING_EMAIL' `
            -Message 'Workflow state change rolled back because required email attributes are missing.' -Data @{
            documentName = $DocumentName; documentGuid = $DocumentGuid; folderPath = $FolderPath
            targetState = $TargetStateName; previousState = $previousState; missingFields = @($missing)
            changedByUser = $ChangedByUser; rollback = $result.rollback; dryRun = [bool]$DryRun
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
        [string]$DocumentName = '',
        [string]$DocumentPath = '',
        [string]$DocumentGuid = '',
        [string]$Project = '',
        [hashtable]$Job,
        [string]$StateTransitionKey = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$SubmittedBy = '',
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
        $DocumentName = _QCN-GetProp -Object $Document -Names @('Name','DocumentName','FileName')
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

    $actorFromJob = _QCN-ResolveStateChangeActorFromJob -Job $Job
    if ($null -eq $ChangedByUser -and $null -ne $actorFromJob.changedByUser) { $ChangedByUser = $actorFromJob.changedByUser }
    if (_QCN-IsBlank $ChangedByUsername -and -not (_QCN-IsBlank $actorFromJob.changedByUsername)) {
        $ChangedByUsername = [string]$actorFromJob.changedByUsername
    }
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

    $qcTarget = _QCN-ResolveQcPdfNotificationTarget -Document $Document -Config $Config -Job $Job `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -DocumentPath $DocumentPath
    $Document = $qcTarget.document
    $DocumentName = [string]$qcTarget.documentName
    $DocumentGuid = [string]$qcTarget.documentGuid
    if (-not (_QCN-IsBlank $qcTarget.documentPath)) { $DocumentPath = [string]$qcTarget.documentPath }

    $folderForRoles = ''
    $sourceForRoles = $DocumentName
    if ($Job) {
        $folderForRoles = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath'))
        $sourceForRoles = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
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
    if (-not (_QCN-IsBlank $sourceForRoles) -and $sourceForRoles -match '(?i)-qc\.pdf$') {
        $sourceForRoles = [string]([System.IO.Path]::GetFileNameWithoutExtension($sourceForRoles)) + '.pdf'
    }

    $resolved = Resolve-QCNotificationRecipients -Document $Document -Settings $settings -ToRoles @($eventCfg.to) -CcRoles @($eventCfg.cc) `
        -Config $Config -Job $Job -RoleOverrides $roleOverrides -FolderPath $folderForRoles -SourceDocumentName $sourceForRoles `
        -DocumentGuid $DocumentGuid
    $event = New-QCNotificationEvent -EventType $eventType -Project $Project -DocumentName $DocumentName `
        -DocumentPath $DocumentPath -DocumentGuid ([string]$DocumentGuid) -PreviousState $prev -CurrentState $curr `
        -Reviewers $resolved.reviewers -Designers $resolved.designers -Cc $resolved.cc -ActionRequired $actionRequired -SourceJobId $sourceJobId
    if (-not (_QCN-IsBlank $folderForRoles)) { $event['folderPath'] = $folderForRoles }
    if (-not (_QCN-IsBlank $StateTransitionKey)) { $event['stateTransitionKey'] = [string]$StateTransitionKey }
    elseif ($Job -and $Job.metadata -is [hashtable] -and $Job.metadata.ContainsKey('stateTransitionKey') -and $Job.metadata.stateTransitionKey) {
        $event['stateTransitionKey'] = [string]$Job.metadata.stateTransitionKey
    }

    $folderForRt = ''
    $sourceForRt = $DocumentName
    if ($Job) {
        $folderForRt = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceFolder', 'folderPath', 'incomingFolderPath'))
        $sourceForRt = [string](_QCN-GetJobValue -Job $Job -Keys @('sourceName', 'sourceDocumentName', 'incomingDocName'))
    }
    if (_QCN-IsBlank $folderForRt) { $folderForRt = [string](_QCN-GetProp -Object $Document -Names @('FolderPath', 'folderPath')) }
    if (_QCN-IsBlank $sourceForRt) { $sourceForRt = [string](_QCN-GetProp -Object $Document -Names @('Name', 'DocumentName', 'FileName')) }
    $resolvedReviewType = _QCN-ResolveNotificationReviewType -Document $Document -Settings $settings -Config $Config -Job $Job `
        -FolderPath $folderForRt -SourceName $sourceForRt
    if (-not (_QCN-IsBlank $resolvedReviewType)) {
        $event['reviewType'] = $resolvedReviewType
        $event['qcReviewType'] = $resolvedReviewType
    }

    $resolvedSubmittedBy = Resolve-QCNotificationSubmittedBy -Config $Config -ChangedByUser $ChangedByUser `
        -ChangedByUsername $ChangedByUsername -SubmittedBy $SubmittedBy
    $event['submittedBy'] = $resolvedSubmittedBy
    if ($null -ne $ChangedByUser) { $event['changedByUser'] = $ChangedByUser }
    if (-not (_QCN-IsBlank $ChangedByUsername)) { $event['changedByUsername'] = [string]$ChangedByUsername }

    $dedupeKey = Get-QCNotificationDedupeKey -Event $event -Settings $settings
    if (-not $Force -and (Test-QCNotificationDedupe -DedupeKey $dedupeKey -Settings $settings)) {
        $skipped = @{
            success = $false
            skipped = $true
            dedupeKey = $dedupeKey
            message = 'Duplicate notification suppressed.'
            eventType = $eventType
            documentName = $DocumentName
            timestampUtc = Get-QCTimestamp
        }
        Write-QCNotificationResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Level 'Information' -Message $skipped.message -Result $skipped -Event $event
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_SKIPPED_DUPLICATE' -Message $skipped.message -Data $skipped
    }

    $tokens = _QCN-NewNotificationSubjectTokens -DocumentName $DocumentName -DocumentPath $DocumentPath `
        -Project $Project -PreviousState $prev -CurrentState $curr -EventType $eventType -ReviewType $resolvedReviewType
    $subjectTemplate = _QCN-ResolveNotificationSubjectTemplate -EventCfg $eventCfg -Settings $settings
    $subject = Expand-QCNotificationTemplate -Template $subjectTemplate -Tokens $tokens

    $qcUrl = Resolve-QCNotificationQcPdfUrl -Event $event -Settings $settings -Document $Document -Config $Config
    if (-not (_QCN-IsBlank $qcUrl)) {
        $event['qcPdfUrl'] = $qcUrl
    }
    $event['_eventCfg'] = $eventCfg
    $event['_document'] = $Document

    $send = _QCN-EnsureSingleResult (Send-QCNotification -Event $event -Config $Config -Subject $subject -To $resolved.to -Cc $resolved.cc)

    $resultData = _QCN-ToHashtable $send.Data
    if ($send.IsSuccess -and $resultData -and $resultData.success -eq $true) {
        Register-QCNotificationDedupe -DedupeKey $dedupeKey -Settings $settings -ResultData $resultData
    }
    if ($resultData) {
        $resultData['dedupeKey'] = $dedupeKey
    }
    $notifyCode = if ($send.IsSuccess) { 'QC_NOTIFICATION_SENT' } elseif ($send.Code) { [string]$send.Code } else { 'QC_NOTIFICATION_FAILED' }
    $notifyMessage = [string]$send.Message
    if ($send.IsSuccess) {
        return New-QCSuccessResult -Code $notifyCode -Message $notifyMessage -Data $resultData
    }
    return New-QCFailureResult -Code $notifyCode -Message $notifyMessage -Data $resultData
}

Export-ModuleMember -Function Get-QCNotificationSettings, New-QCNotificationEvent, Resolve-QCNotificationRecipients, `
    Resolve-QCNotificationQcPdfUrl, Resolve-QCNotificationSubmittedBy, Get-QCNotificationDedupeKey, Test-QCNotificationDedupe, Register-QCNotificationDedupe, `
    Get-QCStateChangeMissingEmailFields, Resolve-QCStateChangeActorEmailAddress, Send-QCStateChangeBlockedNotification, Invoke-QCWorkflowStateEmailAttributeGate, `
    Send-QCNotification, Invoke-QCNotificationForStateChange, Write-QCNotificationResult
