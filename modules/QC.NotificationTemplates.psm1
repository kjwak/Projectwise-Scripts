# QC.NotificationTemplates.psm1
# Responsibility: Subject/body template expansion for QC workflow notifications.

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'Core.Config.psm1') -Force -ErrorAction SilentlyContinue

function _QCNT-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCNT-GetRepoRoot() {
    $root = $PSScriptRoot
    if ($root -match '[\\/]modules$') { return Split-Path -Parent $root }
    return $root
}

function _QCNT-ResolveRepoPath([string]$Path) {
    if (_QCNT-IsBlank $Path) { return '' }
    $p = [string]$Path
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    return (Join-Path (_QCNT-GetRepoRoot) $p)
}

function _QCNT-GetRequiredEmailPlaceholders {
    return @(
        'NotificationTitle', 'NotificationMessage', 'ProjectName',
        'DocumentName', 'ReviewType', 'WorkflowState', 'AssignedTo', 'SubmittedBy',
        'SubmittedDate', 'QCPdfUrl', 'GeneratedTimestamp'
    )
}

function _QCNT-NormalizeDataKey([string]$Key) {
    if (_QCNT-IsBlank $Key) { return '' }
    $k = [string]$Key
    if ($k.Length -eq 0) { return '' }
    return ($k.Substring(0, 1).ToUpper() + $k.Substring(1))
}

function _QCNT-NormalizeEmailData([hashtable]$Data) {
    $normalized = @{}
    foreach ($key in @($Data.Keys)) {
        $canon = _QCNT-NormalizeDataKey -Key ([string]$key)
        if (-not $canon) { continue }
        $normalized[$canon] = $Data[$key]
    }
    return $normalized
}

function _QCNT-EncodeEmailValue([string]$Value, [switch]$IsUrl) {
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ($IsUrl) {
        return [System.Net.WebUtility]::HtmlEncode($text)
    }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function _QCNT-ValidateEmailUrl([string]$FieldName, [string]$Url) {
    if (_QCNT-IsBlank $Url) {
        throw "QC_EMAIL_INVALID_URL: '$FieldName' must be a non-empty absolute URL."
    }
    $uri = $null
    if (-not [Uri]::TryCreate([string]$Url.Trim(), [UriKind]::Absolute, [ref]$uri)) {
        throw "QC_EMAIL_INVALID_URL: '$FieldName' is not a valid absolute URL: $Url"
    }
}

function _QCNT-HideRowStyle([string]$Value) {
    if (_QCNT-IsBlank $Value) { return 'display:none;' }
    return ''
}

function _QCNT-GetDocumentAttribute([object]$Document, [string]$AttributeName) {
    if (_QCNT-IsBlank $AttributeName -or -not $Document) { return $null }
    $containers = @()
    foreach ($prop in @('qcAttributes', 'attributes', 'Attributes', 'CustomAttributes', 'EnvironmentAttributes')) {
        try {
            if ($Document.PSObject.Properties[$prop] -and $Document.$prop) { $containers += $Document.$prop }
        } catch { }
    }
    foreach ($bag in @($containers)) {
        if ($bag -is [System.Collections.IDictionary]) {
            if ($bag.Contains($AttributeName)) { return $bag[$AttributeName] }
        }
        if ($bag -is [System.Collections.IEnumerable] -and -not ($bag -is [string])) {
            foreach ($item in @($bag)) {
                if ($item -is [System.Collections.IDictionary] -and $item.Contains($AttributeName)) {
                    return $item[$AttributeName]
                }
            }
        }
    }
    try {
        if ($Document.PSObject.Properties[$AttributeName]) { return $Document.$AttributeName }
    } catch { }
    return $null
}

function Expand-QCNotificationTemplate {
    <#
    .SYNOPSIS
    Replaces {token} placeholders in a notification template string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Template,
        [Parameter(Mandatory)]
        [hashtable]$Tokens
    )

    if (_QCNT-IsBlank $Template) { return '' }
    $result = [string]$Template
    foreach ($key in @($Tokens.Keys)) {
        $placeholder = '{' + [string]$key + '}'
        $value = if ($null -eq $Tokens[$key]) { '' } else { [string]$Tokens[$key] }
        $result = $result.Replace($placeholder, $value)
    }
    return $result
}

function ConvertTo-QCEmailHtml {
    <#
    .SYNOPSIS
    Loads an HTML email template and replaces {Placeholder} tokens with encoded values.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TemplatePath,
        [Parameter(Mandatory)]
        [hashtable]$Data,
        [string[]]$Required = @()
    )

    $resolvedPath = _QCNT-ResolveRepoPath -Path $TemplatePath
    if (-not $resolvedPath) { $resolvedPath = $TemplatePath }
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "QC_EMAIL_TEMPLATE_NOT_FOUND: Template file not found: $resolvedPath"
    }

    $requiredKeys = if ($Required -and $Required.Count -gt 0) { @($Required) } else { @(_QCNT-GetRequiredEmailPlaceholders) }
    $data = _QCNT-NormalizeEmailData -Data $Data

    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $requiredKeys) {
        if (-not $data.ContainsKey($key)) { $missing.Add($key) | Out-Null }
    }
    if ($missing.Count -gt 0) {
        throw ('QC_EMAIL_MISSING_PLACEHOLDER: Missing required placeholder data: ' + ($missing -join ', '))
    }

    if (_QCNT-IsBlank $data['QCPdfUrl']) {
        throw 'QC_EMAIL_MISSING_PLACEHOLDER: QCPdfUrl is required and cannot be empty.'
    }

    $urlKeys = @($data.Keys | Where-Object { $_ -eq 'QCPdfUrl' -or $_ -like '*Url' })
    foreach ($urlKey in $urlKeys) {
        $urlVal = [string]$data[$urlKey]
        if (-not (_QCNT-IsBlank $urlVal)) {
            _QCNT-ValidateEmailUrl -FieldName $urlKey -Url $urlVal
        }
    }

    $optionalKeys = @(
        'NotificationCategory', 'Environment', 'DocumentState', 'ReviewerComments'
    )
    foreach ($opt in $optionalKeys) {
        if (-not $data.ContainsKey($opt)) { $data[$opt] = '' }
    }

    $data['NotificationCategoryBannerStyle'] = _QCNT-HideRowStyle -Value ([string]$data['NotificationCategory'])
    $data['EnvironmentBadgeStyle'] = if (_QCNT-IsBlank $data['Environment']) { 'display:none;' } else { '' }
    $data['ReviewTypeRowStyle'] = _QCNT-HideRowStyle -Value ([string]$data['ReviewType'])
    $data['DocumentStateRowStyle'] = _QCNT-HideRowStyle -Value ([string]$data['DocumentState'])
    $data['ReviewerCommentsRowStyle'] = _QCNT-HideRowStyle -Value ([string]$data['ReviewerComments'])

    $template = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8

    $encoded = @{}
    foreach ($key in @($data.Keys)) {
        $raw = if ($null -eq $data[$key]) { '' } else { [string]$data[$key] }
        $isUrl = ($key -eq 'QCPdfUrl') -or ($key -like '*Url')
        $encoded[$key] = _QCNT-EncodeEmailValue -Value $raw -IsUrl:$isUrl
    }

    $result = [string]$template
    $matches = [regex]::Matches($result, '\{([A-Za-z][A-Za-z0-9]*)\}')
    $tokens = @{}
    foreach ($m in $matches) {
        $tokens[$m.Groups[1].Value] = $true
    }
    foreach ($token in @($tokens.Keys)) {
        $placeholder = '{' + $token + '}'
        $value = if ($encoded.ContainsKey($token)) { $encoded[$token] } else { '' }
        $result = $result.Replace($placeholder, $value)
    }
    return $result
}

function New-QCNotificationEmailTemplateData {
    <#
    .SYNOPSIS
    Builds placeholder data for QC HTML email templates from a notification event.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [hashtable]$EventCfg = @{},
        [hashtable]$Settings = @{},
        [hashtable]$Config = $null,
        [string]$FolderPath = '',
        [object]$Document = $null,
        [string]$Subject = ''
    )

    $attrs = @{}
    if ($Settings -and $Settings.attributes) {
        $a = $Settings.attributes
        if ($a -is [hashtable]) { $attrs = $a }
        elseif ($a -is [System.Collections.IDictionary]) {
            foreach ($k in $a.Keys) { $attrs[[string]$k] = $a[$k] }
        }
    }

    $emailCfg = @{}
    if ($Settings -and $Settings.email) {
        $e = $Settings.email
        if ($e -is [hashtable]) { $emailCfg = $e }
    }

    $reviewType = ''
    if ($Event -and $Event.ContainsKey('reviewType') -and -not (_QCNT-IsBlank $Event.reviewType)) {
        $reviewType = [string]$Event.reviewType
    }
    elseif ($Event -and $Event.ContainsKey('qcReviewType') -and -not (_QCNT-IsBlank $Event.qcReviewType)) {
        $reviewType = [string]$Event.qcReviewType
    }
    elseif ($EventCfg -and $EventCfg.ContainsKey('reviewType') -and $EventCfg.reviewType) {
        $reviewType = [string]$EventCfg.reviewType
    }
    elseif ($attrs.ContainsKey('reviewTypeField') -and -not (_QCNT-IsBlank $attrs['reviewTypeField'])) {
        $rt = _QCNT-GetDocumentAttribute -Document $Document -AttributeName ([string]$attrs['reviewTypeField'])
        if ($rt) { $reviewType = [string]$rt }
    }
    if (_QCNT-IsBlank $reviewType -and $Document) {
        $rt = _QCNT-GetDocumentAttribute -Document $Document -AttributeName 'QC_Review_Type'
        if ($rt) { $reviewType = [string]$rt }
    }

    $assignedTo = ''
    if ($Event.reviewers -and @($Event.reviewers).Count -gt 0) {
        $assignedTo = (@($Event.reviewers) | Where-Object { -not (_QCNT-IsBlank $_) } | Select-Object -First 1)
        if ($assignedTo) { $assignedTo = [string]$assignedTo }
    }
    if (_QCNT-IsBlank $assignedTo) { $assignedTo = '(not assigned)' }

    $submittedBy = ''
    if ($Event.designers -and @($Event.designers).Count -gt 0) {
        $submittedBy = (@($Event.designers) | Where-Object { -not (_QCNT-IsBlank $_) } | Select-Object -First 1)
        if ($submittedBy) { $submittedBy = [string]$submittedBy }
    }
    if (_QCNT-IsBlank $submittedBy) { $submittedBy = '(unknown)' }

    $title = ''
    if ($EventCfg -and $EventCfg.ContainsKey('emailTitle') -and $EventCfg.emailTitle) {
        $title = [string]$EventCfg.emailTitle
    }
    elseif (-not (_QCNT-IsBlank $Subject)) {
        $title = [string]$Subject
    }
    elseif ($Event.currentState) {
        $title = 'QC Notification - ' + [string]$Event.currentState
    }
    else {
        $title = 'QC Workflow Notification'
    }

    $message = ''
    if ($EventCfg -and $EventCfg.ContainsKey('emailMessage') -and $EventCfg.emailMessage) {
        $message = [string]$EventCfg.emailMessage
    }
    elseif ($Event.actionRequired) {
        $message = [string]$Event.actionRequired
    }
    else {
        $message = 'A ProjectWise QC workflow state change requires your attention.'
    }

    $category = ''
    if ($Event.currentState) { $category = [string]$Event.currentState }
    elseif ($Event.eventType) { $category = [string]$Event.eventType }

    $qcPdfUrl = ''
    if ($Event.qcPdfUrl) { $qcPdfUrl = [string]$Event.qcPdfUrl }

    $timestamp = ''
    if (Get-Command -Name 'Get-QCTimestamp' -ErrorAction SilentlyContinue) {
        $timestamp = Get-QCTimestamp
    }
    else {
        $timestamp = [DateTime]::UtcNow.ToString('o')
    }

    $environment = ''
    if ($emailCfg.ContainsKey('environment') -and $emailCfg.environment) {
        $environment = [string]$emailCfg.environment
    }

    $folderForProject = [string]$FolderPath
    if (_QCNT-IsBlank $folderForProject -and $Event.documentPath) {
        $dp = [string]$Event.documentPath
        if ($dp -match '\\') {
            $folderForProject = [System.IO.Path]::GetDirectoryName($dp)
        }
    }

    $projectName = ''
    if ($Config -and -not (_QCNT-IsBlank $folderForProject) -and (Get-Command -Name 'Get-QCProjectNameFromFolderPath' -ErrorAction SilentlyContinue)) {
        try {
            $fromPath = Get-QCProjectNameFromFolderPath -Config $Config -FolderPath $folderForProject
            if ($fromPath) { $projectName = [string]$fromPath }
        } catch { }
    }
    if (_QCNT-IsBlank $projectName -and $Event.project) {
        $projectName = [string]$Event.project
    }

    return @{
        NotificationTitle = $title
        NotificationMessage = $message
        NotificationCategory = $category
        ProjectName = if ($projectName) { $projectName } else { '(unknown)' }
        DocumentName = if ($Event.documentName) { [string]$Event.documentName } else { '(unknown)' }
        ReviewType = if ($reviewType) { $reviewType } else { '(not specified)' }
        WorkflowState = if ($Event.currentState) { [string]$Event.currentState } else { '(unknown)' }
        AssignedTo = $assignedTo
        SubmittedBy = $submittedBy
        SubmittedDate = $timestamp
        QCPdfUrl = $qcPdfUrl
        Environment = $environment
        GeneratedTimestamp = $timestamp
        DocumentState = if ($Event.currentState) { [string]$Event.currentState } else { '' }
        ReviewerComments = if ($Event.reviewerComments) { [string]$Event.reviewerComments } else { '' }
    }
}

function New-QCNotificationEmailBody {
    <#
    .SYNOPSIS
    Builds a plain-text email body from a notification event object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Event,
        [string]$Intro = 'A ProjectWise QC workflow state change requires your attention.'
    )

    $lines = @(
        $Intro,
        '',
        ('Project: {0}' -f $(if ($Event.project) { $Event.project } else { '(unknown)' })),
        ('Document: {0}' -f $(if ($Event.documentName) { $Event.documentName } else { '(unknown)' })),
        ('Path: {0}' -f $(if ($Event.documentPath) { $Event.documentPath } else { '(unknown)' })),
        ('Previous state: {0}' -f $(if ($Event.previousState) { $Event.previousState } else { '(none)' })),
        ('Current state: {0}' -f $(if ($Event.currentState) { $Event.currentState } else { '(unknown)' })),
        ('Event: {0}' -f $(if ($Event.eventType) { $Event.eventType } else { '(unknown)' })),
        ''
    )
    if (-not (_QCNT-IsBlank $Event.actionRequired)) {
        $lines += 'Action required:'
        $lines += [string]$Event.actionRequired
        $lines += ''
    }
    if ($Event.qcPdfUrl) {
        $lines += 'Open QC PDF:'
        $lines += [string]$Event.qcPdfUrl
        $lines += ''
    }
    if ($Event.sourceJobId) {
        $lines += ('Source job: {0}' -f [string]$Event.sourceJobId)
    }
    $lines += ''
    $lines += 'This message was generated by the ProjectWise QC automation system.'
    return ($lines -join [Environment]::NewLine)
}

Export-ModuleMember -Function Expand-QCNotificationTemplate, New-QCNotificationEmailBody, `
    ConvertTo-QCEmailHtml, New-QCNotificationEmailTemplateData
