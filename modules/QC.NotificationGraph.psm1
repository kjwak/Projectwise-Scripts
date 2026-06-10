# QC.NotificationGraph.psm1
# Responsibility: Microsoft Graph email provider (client-secret app auth).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

function _QCNG-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCNG-GetGraphSetting {
    param(
        [hashtable]$GraphSettings,
        [string]$Key
    )

    if (-not $GraphSettings -or -not $GraphSettings.ContainsKey($Key)) { return '' }
    $v = $GraphSettings[$Key]
    if ($null -eq $v) { return '' }
    if ($v -is [string]) { return $v.Trim() }
    if ($v -is [System.ValueType]) { return [string]$v }
    throw ("notifications.graph.{0} must be a string (got {1})." -f $Key, $v.GetType().FullName)
}

function Test-QCNotificationGraphConfigured {
    <#
    .SYNOPSIS
    Returns whether Microsoft Graph notification settings contain required non-secret identifiers and auth material.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$GraphSettings
    )

    $required = @('tenantId', 'clientId', 'senderMailbox')
    $missing = @()
    foreach ($key in @($required)) {
        try {
            if (_QCNG-IsBlank (_QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key $key)) { $missing += $key }
        } catch {
            $missing += $key
        }
    }
    $hasSecret = $false
    $hasCert = $false
    if ($GraphSettings) {
        try {
            if (-not (_QCNG-IsBlank (_QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'clientSecret'))) { $hasSecret = $true }
        } catch { }
        try {
            if (-not (_QCNG-IsBlank (_QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'certificateThumbprint'))) { $hasCert = $true }
        } catch { }
        try {
            if (-not (_QCNG-IsBlank (_QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'certificatePath'))) { $hasCert = $true }
        } catch { }
    }
    if (-not $hasSecret -and -not $hasCert) { $missing += 'clientSecret|certificateThumbprint|certificatePath' }

    return @{
        configured = ($missing.Count -eq 0)
        missing = @($missing)
        usesClientSecret = $hasSecret
        usesCertificate = $hasCert
    }
}

function _QCNG-GraphRecipientList([string[]]$Addresses) {
    $list = @()
    foreach ($addr in @($Addresses)) {
        if (_QCNG-IsBlank $addr) { continue }
        $trimmed = [string]$addr.Trim()
        if ($trimmed -notmatch '@') {
            throw "Invalid email recipient (missing @): $trimmed"
        }
        $list += @{ emailAddress = @{ address = $trimmed } }
    }
    return @($list)
}

function _QCNG-ResolveRepoPath([string]$Path) {
    if (_QCNG-IsBlank $Path) { return '' }
    $p = [string]$Path
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    $root = $PSScriptRoot
    if ($root -match '[\\/]modules$') { $root = Split-Path -Parent $root }
    return (Join-Path $root $p)
}

function New-QCGraphEmailMessage {
    <#
    .SYNOPSIS
    Builds a Microsoft Graph sendMail payload with HTML body and inline TYPSA logo attachment.
    Returns the payload object without sending.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ToRecipients,
        [Parameter(Mandatory)]
        [string]$Subject,
        [Parameter(Mandatory)]
        [string]$HtmlBody,
        [Parameter(Mandatory)]
        [string]$LogoPath,
        [string[]]$CcRecipients = @(),
        [string[]]$BccRecipients = @(),
        [string]$ContentId = 'typsa-logo',
        [ValidateSet('low', 'normal', 'high')]
        [string]$Importance = 'normal'
    )

    if (_QCNG-IsBlank $Subject) {
        throw 'New-QCGraphEmailMessage: Subject is required.'
    }
    if (_QCNG-IsBlank $HtmlBody) {
        throw 'New-QCGraphEmailMessage: HtmlBody is required.'
    }

    $to = @(_QCNG-GraphRecipientList -Addresses $ToRecipients)
    if ($to.Count -eq 0) {
        throw 'New-QCGraphEmailMessage: At least one To recipient is required.'
    }

    $resolvedLogo = _QCNG-ResolveRepoPath -Path $LogoPath
    if (-not $resolvedLogo) { $resolvedLogo = $LogoPath }
    if (-not (Test-Path -LiteralPath $resolvedLogo)) {
        throw "New-QCGraphEmailMessage: Logo file not found: $resolvedLogo"
    }

    $logoBytes = [System.IO.File]::ReadAllBytes($resolvedLogo)
    $logoBase64 = [Convert]::ToBase64String($logoBytes)
    $logoFileName = [System.IO.Path]::GetFileName($resolvedLogo)
    $logoContentType = 'image/webp'
    if ($logoFileName -match '\.png$') { $logoContentType = 'image/png' }
    elseif ($logoFileName -match '\.(jpe?g)$') { $logoContentType = 'image/jpeg' }
    elseif ($logoFileName -match '\.gif$') { $logoContentType = 'image/gif' }

    $message = @{
        subject = [string]$Subject
        body = @{
            contentType = 'HTML'
            content = [string]$HtmlBody
        }
        toRecipients = $to
        importance = [string]$Importance
        attachments = @(
            @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name = $logoFileName
                contentType = $logoContentType
                contentId = [string]$ContentId
                isInline = $true
                contentBytes = $logoBase64
            }
        )
    }

    $cc = @(_QCNG-GraphRecipientList -Addresses $CcRecipients)
    if ($cc.Count -gt 0) { $message['ccRecipients'] = $cc }
    $bcc = @(_QCNG-GraphRecipientList -Addresses $BccRecipients)
    if ($bcc.Count -gt 0) { $message['bccRecipients'] = $bcc }

    return @{
        message = $message
        saveToSentItems = $true
    }
}

function New-QCNotificationGraphSendMailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Payload
    )

    $message = @{
        subject = [string]$Payload.subject
        body = @{
            contentType = 'Text'
            content = [string]$Payload.body
        }
        toRecipients = @(_QCNG-GraphRecipientList -Addresses $Payload.to)
    }
    $cc = @(_QCNG-GraphRecipientList -Addresses $Payload.cc)
    if ($cc.Count -gt 0) {
        $message['ccRecipients'] = $cc
    }

    return @{
        message = $message
        saveToSentItems = $true
    }
}

function Get-QCNotificationGraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$GraphSettings
    )

    $tenantId = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'tenantId'
    $clientId = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'clientId'
    $clientSecret = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'clientSecret'
    $tokenUri = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $tenantId

    $body = @{
        grant_type = 'client_credentials'
        client_id = $clientId
        client_secret = $clientSecret
        scope = 'https://graph.microsoft.com/.default'
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    } catch {
        $detail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
        throw "Graph token request failed: $detail"
    }

    if (-not $response.access_token) {
        throw 'Graph token response did not include access_token.'
    }
    return [string]$response.access_token
}

function Invoke-QCNotificationGraphSendMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SenderMailbox,
        [Parameter(Mandatory)]
        [string]$AccessToken,
        [Parameter(Mandatory)]
        [hashtable]$SendMailBody
    )

    $encodedMailbox = [Uri]::EscapeDataString($SenderMailbox)
    $uri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/sendMail"
    $headers = @{
        Authorization = "Bearer $AccessToken"
    }
    $json = $SendMailBody | ConvertTo-Json -Depth 12 -Compress

    try {
        Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $json -ContentType 'application/json; charset=utf-8' -ErrorAction Stop | Out-Null
    } catch {
        $detail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
        throw "Graph sendMail failed: $detail"
    }
}

function Send-QCNotificationGraph {
    <#
    .SYNOPSIS
    Sends email via Microsoft Graph using client-secret app credentials (or dry-run preview).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$GraphSettings,
        [Parameter(Mandatory)]
        [hashtable]$Payload,
        [switch]$DryRun
    )

    $timestampUtc = Get-QCTimestamp
    $baseData = @{
        success = $false
        provider = 'MicrosoftGraph'
        dryRun = [bool]$DryRun
        eventType = $Payload.eventType
        documentName = $Payload.documentName
        to = @($Payload.to)
        cc = @($Payload.cc)
        timestampUtc = $timestampUtc
    }

    $validation = Test-QCNotificationGraphConfigured -GraphSettings $GraphSettings
    if (-not $validation.configured) {
        $data = $baseData.Clone()
        $data['message'] = 'Missing Graph configuration: ' + ($validation.missing -join ', ')
        $data['missing'] = @($validation.missing)
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_NOT_CONFIGURED' -Message 'Microsoft Graph provider is not configured.' -Data $data
    }

    $logoPath = ''
    if ($Payload.logoPath) { $logoPath = [string]$Payload.logoPath }
    if (_QCNG-IsBlank $logoPath) {
        $logoPath = 'email/typsalogo.png.webp'
    }

    $importance = 'normal'
    if ($Payload.importance) {
        $candidate = ([string]$Payload.importance).Trim().ToLowerInvariant()
        if ($candidate -in @('low', 'normal', 'high')) { $importance = $candidate }
    }

    if ($Payload.htmlBody) {
        $sendMailBody = New-QCGraphEmailMessage -ToRecipients @($Payload.to) -Subject ([string]$Payload.subject) `
            -HtmlBody ([string]$Payload.htmlBody) -LogoPath $logoPath -CcRecipients @($Payload.cc) -Importance $importance
    }
    else {
        $sendMailBody = New-QCNotificationGraphSendMailBody -Payload $Payload
    }

    if ($DryRun) {
        $data = $baseData.Clone()
        $data['success'] = $true
        $data['message'] = 'Microsoft Graph dry run: sendMail payload built; no API call made.'
        $data['senderMailbox'] = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'senderMailbox'
        $data['sendMailBody'] = $sendMailBody
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_GRAPH_DRY_RUN' -Message $data.message -Data $data
    }

    if ($validation.usesCertificate -and -not $validation.usesClientSecret) {
        $data = $baseData.Clone()
        $data['message'] = 'Certificate-based Graph auth is not implemented; set notifications.graph.clientSecret.'
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_NOT_IMPLEMENTED' -Message $data.message -Data $data
    }

    try {
        $token = Get-QCNotificationGraphAccessToken -GraphSettings $GraphSettings
        Invoke-QCNotificationGraphSendMail -SenderMailbox (_QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'senderMailbox') -AccessToken $token -SendMailBody $sendMailBody
    } catch {
        $data = $baseData.Clone()
        $data['message'] = $_.Exception.Message
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_SEND_FAILED' -Message $data.message -Data $data
    }

    $data = $baseData.Clone()
    $data['success'] = $true
    $data['message'] = 'Microsoft Graph sendMail completed.'
    $data['senderMailbox'] = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'senderMailbox'
    return New-QCSuccessResult -Code 'QC_NOTIFICATION_GRAPH_SENT' -Message $data.message -Data $data
}

Export-ModuleMember -Function Test-QCNotificationGraphConfigured, New-QCNotificationGraphSendMailBody, `
    New-QCGraphEmailMessage, Get-QCNotificationGraphAccessToken, Invoke-QCNotificationGraphSendMail, Send-QCNotificationGraph
