# QC.Notifications.psm1
# Responsibility: Configurable QC workflow email notifications (Mock + future Microsoft Graph).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationTemplates.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationMock.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.NotificationGraph.psm1') -Force
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
            qcPdfUrlTemplate = ''
            pwLinkBaseUrl = 'https://connect-projectwisewac.bentley.com/pwlink/'
            pwLinkApp = 'pwe'
        }
        attributes = @{
            reviewerEmailField = 'EM_Reviewer_Email'
            designerEmailField = 'EM_Designer_Email'
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
                subjectTemplate = 'QC Received - {documentName}'
                actionRequired = 'Reviewer to begin QC review.'
            }
            'Corrections In Progress' = @{
                enabled = $true
                eventType = 'CORRECTIONS_IN_PROGRESS'
                to = @('designers')
                cc = @('reviewers')
                subjectTemplate = 'QC Corrections Required - {documentName}'
                actionRequired = 'Designer to address QC comments.'
            }
            'Backcheck In Progress' = @{
                enabled = $true
                eventType = 'BACKCHECK_IN_PROGRESS'
                to = @('reviewers')
                cc = @('designers')
                subjectTemplate = 'QC Backcheck Required - {documentName}'
                actionRequired = 'Reviewer to backcheck corrections.'
            }
            'Error Needs Attention' = @{
                enabled = $true
                eventType = 'QC_ERROR'
                to = @('reviewers', 'designers')
                cc = @()
                subjectTemplate = 'QC Automation Error - {documentName}'
                actionRequired = 'Manual review required.'
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
    foreach ($boolKey in @('enabled','dryRun')) {
        try { $settings[$boolKey] = [bool]$settings[$boolKey] } catch { $settings[$boolKey] = [bool]$defaults[$boolKey] }
    }
    if ($settings.dedupe) {
        try { $settings.dedupe.enabled = [bool]$settings.dedupe.enabled } catch { }
    }
    return $settings
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
        [string]$QcPdfUrl = ''
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
        [string[]]$ExplicitCc = @()
    )

    if (-not $Settings) { $Settings = Get-QCNotificationSettings -Config @{} }
    $attr = _QCN-ToHashtable $Settings.attributes
    if (-not $attr) { $attr = @{} }

    $reviewerField = if ($attr.reviewerEmailField) { [string]$attr.reviewerEmailField } else { 'EM_Reviewer_Email' }
    $designerField = if ($attr.designerEmailField) { [string]$attr.designerEmailField } else { 'EM_Designer_Email' }
    $ccField = if ($attr.ccEmailField) { [string]$attr.ccEmailField } else { 'CcEmails' }

    $reviewers = _QCN-ParseEmailList (_QCN-GetAttributeValue -Document $Document -AttributeName $reviewerField)
    $designers = _QCN-ParseEmailList (_QCN-GetAttributeValue -Document $Document -AttributeName $designerField)
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
            default {
                if ($Event.ContainsKey($field)) { $value = [string]$Event[$field] }
            }
        }
        $parts.Add(('{0}={1}' -f $field, $value)) | Out-Null
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

    if (_QCN-IsBlank $Subject) {
        $Subject = ('QC Notification - {0}' -f $Event.documentName)
    }
    if (_QCN-IsBlank $Body) {
        $Body = New-QCNotificationEmailBody -Event $Event
    }

    $eventCfg = @{}
    if ($Event._eventCfg) {
        $ec = _QCN-ToHashtable $Event._eventCfg
        if ($ec) { $eventCfg = $ec }
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

        $templateData = New-QCNotificationEmailTemplateData -Event $Event -EventCfg $eventCfg -Settings $settings `
            -Document $Event._document -Subject $Subject
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
                message = 'QC Received notification deferred until prepend and rendition complete.'
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
    if (_QCN-IsBlank $Project) {
        $Project = _QCN-GetProp -Object $Document -Names @('ProjectName','Project')
        if ($Project) { $Project = [string]$Project }
    }

    $eventType = if ($eventCfg.eventType) { [string]$eventCfg.eventType } else { $curr.ToUpperInvariant().Replace(' ', '_') }
    $actionRequired = if ($eventCfg.actionRequired) { [string]$eventCfg.actionRequired } else { '' }
    $sourceJobId = if ($Job -and $Job.ContainsKey('id')) { [string]$Job.id } else { '' }

    $resolved = Resolve-QCNotificationRecipients -Document $Document -Settings $settings -ToRoles @($eventCfg.to) -CcRoles @($eventCfg.cc)
    $event = New-QCNotificationEvent -EventType $eventType -Project $Project -DocumentName $DocumentName `
        -DocumentPath $DocumentPath -DocumentGuid ([string]$DocumentGuid) -PreviousState $prev -CurrentState $curr `
        -Reviewers $resolved.reviewers -Designers $resolved.designers -Cc $resolved.cc -ActionRequired $actionRequired -SourceJobId $sourceJobId

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

    $tokens = @{
        documentName = $DocumentName
        documentPath = $DocumentPath
        project = $Project
        previousState = $prev
        currentState = $curr
        eventType = $eventType
    }
    $subject = Expand-QCNotificationTemplate -Template ([string]$eventCfg.subjectTemplate) -Tokens $tokens

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
    $notifyCode = if ($send.IsSuccess) { 'QC_NOTIFICATION_SENT' } else { 'QC_NOTIFICATION_FAILED' }
    $notifyMessage = [string]$send.Message
    if ($send.IsSuccess) {
        return New-QCSuccessResult -Code $notifyCode -Message $notifyMessage -Data $resultData
    }
    return New-QCFailureResult -Code $notifyCode -Message $notifyMessage -Data $resultData
}

Export-ModuleMember -Function Get-QCNotificationSettings, New-QCNotificationEvent, Resolve-QCNotificationRecipients, `
    Resolve-QCNotificationQcPdfUrl, Get-QCNotificationDedupeKey, Test-QCNotificationDedupe, Register-QCNotificationDedupe, `
    Send-QCNotification, Invoke-QCNotificationForStateChange, Write-QCNotificationResult
