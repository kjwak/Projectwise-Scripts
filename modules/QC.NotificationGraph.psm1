# QC.NotificationGraph.psm1
# Responsibility: Microsoft Graph email provider (client-secret app auth) with conversation threading.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

# Optional test hook: scriptblock param(Method, Uri, Headers, Body) returns response object.
$script:QCNG_TestHttpHandler = $null
$script:QCNG_TestMessageRegistry = @{}

function Set-QCNotificationGraphTestHttpHandler {
    param([scriptblock]$Handler)
    $script:QCNG_TestHttpHandler = $Handler
}

function Clear-QCNotificationGraphTestHttpHandler {
    $script:QCNG_TestHttpHandler = $null
    $script:QCNG_TestMessageRegistry = @{}
}

function Register-QCNotificationGraphTestMessage {
    param([Parameter(Mandatory)][string]$MessageId)
    if (-not (_QCNG-IsBlank $MessageId)) {
        $script:QCNG_TestMessageRegistry[[string]$MessageId] = $true
    }
}

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

function _QCNG-NewGraphHeaders([string]$AccessToken, [switch]$ImmutableId) {
    $headers = @{
        Authorization = "Bearer $AccessToken"
    }
    if ($ImmutableId) {
        $headers['Prefer'] = 'IdType="ImmutableId"'
    }
    return $headers
}

function _QCNG-InvokeGraphRequest {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers = @{},
        [string]$Body = '',
        [string]$ContentType = 'application/json; charset=utf-8'
    )

    if ($script:QCNG_TestHttpHandler) {
        return & $script:QCNG_TestHttpHandler $Method $Uri $Headers $Body
    }

    $params = @{
        Method = $Method
        Uri = $Uri
        Headers = $Headers
        ErrorAction = 'Stop'
    }
    if (-not (_QCNG-IsBlank $Body)) {
        $params['Body'] = $Body
        $params['ContentType'] = $ContentType
    }
    return Invoke-RestMethod @params
}

function _QCNG-ExtractGraphMessageMetadata([object]$Message) {
    if ($null -eq $Message) { return @{} }
    $id = ''
    $conversationId = ''
    $internetMessageId = ''
    if ($Message -is [hashtable]) {
        if ($Message.ContainsKey('id')) { $id = [string]$Message.id }
        if ($Message.ContainsKey('conversationId')) { $conversationId = [string]$Message.conversationId }
        if ($Message.ContainsKey('internetMessageId')) { $internetMessageId = [string]$Message.internetMessageId }
    } else {
        if ($Message.PSObject.Properties['id']) { $id = [string]$Message.id }
        if ($Message.PSObject.Properties['conversationId']) { $conversationId = [string]$Message.conversationId }
        if ($Message.PSObject.Properties['internetMessageId']) { $internetMessageId = [string]$Message.internetMessageId }
    }
    return @{
        graphMessageId = $id
        graphImmutableMessageId = $id
        graphConversationId = $conversationId
        internetMessageId = $internetMessageId
    }
}

function New-QCGraphEmailMessage {
    <#
    .SYNOPSIS
    Builds a Microsoft Graph message object with HTML body and inline TYPSA logo attachment.
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

    return $message
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
        if ($script:QCNG_TestHttpHandler) {
            $response = & $script:QCNG_TestHttpHandler 'POST' $tokenUri @{} ($body | ConvertTo-Json -Compress)
        } else {
            $response = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        }
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
    $headers = _QCNG-NewGraphHeaders -AccessToken $AccessToken
    $json = $SendMailBody | ConvertTo-Json -Depth 12 -Compress

    try {
        _QCNG-InvokeGraphRequest -Method Post -Uri $uri -Headers $headers -Body $json | Out-Null
    } catch {
        $detail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
        throw "Graph sendMail failed: $detail"
    }
}

function Invoke-QCNotificationGraphCreateAndSendMessage {
    <#
    .SYNOPSIS
    Creates a draft message, sends it, and returns sent-message metadata (immutable ID when supported).
  Graph sendMail does not return message IDs; this path captures IDs for threading.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SenderMailbox,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][hashtable]$Message
    )

    $encodedMailbox = [Uri]::EscapeDataString($SenderMailbox)
    $createUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages"
    $headers = _QCNG-NewGraphHeaders -AccessToken $AccessToken -ImmutableId
    $json = $Message | ConvertTo-Json -Depth 12 -Compress

    $draft = _QCNG-InvokeGraphRequest -Method Post -Uri $createUri -Headers $headers -Body $json
    $meta = _QCNG-ExtractGraphMessageMetadata -Message $draft
    if (_QCNG-IsBlank $meta.graphMessageId) {
        throw 'Graph create message did not return a message id.'
    }
    if ($script:QCNG_TestHttpHandler) {
        Register-QCNotificationGraphTestMessage -MessageId $meta.graphMessageId
    }

    $sendUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$($meta.graphMessageId)/send"
    _QCNG-InvokeGraphRequest -Method Post -Uri $sendUri -Headers $headers | Out-Null

    try {
        $getUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$($meta.graphMessageId)?`$select=id,conversationId,internetMessageId"
        $sent = _QCNG-InvokeGraphRequest -Method Get -Uri $getUri -Headers $headers
        $sentMeta = _QCNG-ExtractGraphMessageMetadata -Message $sent
        foreach ($k in @('graphConversationId', 'internetMessageId', 'graphMessageId')) {
            if (-not (_QCNG-IsBlank $sentMeta[$k])) { $meta[$k] = $sentMeta[$k] }
        }
        $meta.graphImmutableMessageId = $meta.graphMessageId
    } catch {
        # Sent folder move can delay GET; keep create-time metadata.
    }

    return $meta
}

function Invoke-QCNotificationGraphCreateReplyAndSend {
    <#
    .SYNOPSIS
    Creates a reply draft via createReply, updates body/recipients, sends, and returns metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SenderMailbox,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$ParentMessageId,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyContent,
        [Parameter(Mandatory)][string]$BodyContentType,
        [Parameter(Mandatory)][string[]]$ToRecipients,
        [string[]]$CcRecipients = @(),
        [object[]]$Attachments = $null
    )

    $encodedMailbox = [Uri]::EscapeDataString($SenderMailbox)
    $encodedParent = [Uri]::EscapeDataString($ParentMessageId)
    $replyUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$encodedParent/createReply"
    $headers = _QCNG-NewGraphHeaders -AccessToken $AccessToken -ImmutableId

    try {
        $draft = _QCNG-InvokeGraphRequest -Method Post -Uri $replyUri -Headers $headers -Body '{}'
    } catch {
        $detail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = [string]$_.ErrorDetails.Message }
        throw "Graph createReply failed: $detail"
    }

    $draftMeta = _QCNG-ExtractGraphMessageMetadata -Message $draft
    if (_QCNG-IsBlank $draftMeta.graphMessageId) {
        throw 'Graph createReply did not return a draft message id.'
    }

    $patchBody = @{
        subject = [string]$Subject
        body = @{
            contentType = if ($BodyContentType -eq 'HTML') { 'HTML' } else { 'Text' }
            content = [string]$BodyContent
        }
        toRecipients = @(_QCNG-GraphRecipientList -Addresses $ToRecipients)
    }
    $cc = @(_QCNG-GraphRecipientList -Addresses $CcRecipients)
    if ($cc.Count -gt 0) { $patchBody['ccRecipients'] = $cc }
    if ($Attachments -and @($Attachments).Count -gt 0) { $patchBody['attachments'] = @($Attachments) }

    $patchUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$($draftMeta.graphMessageId)"
    $patched = _QCNG-InvokeGraphRequest -Method Patch -Uri $patchUri -Headers $headers -Body ($patchBody | ConvertTo-Json -Depth 12 -Compress)

    $meta = _QCNG-ExtractGraphMessageMetadata -Message $patched
    if (_QCNG-IsBlank $meta.graphMessageId) { $meta = $draftMeta.Clone() }

    $sendUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$($meta.graphMessageId)/send"
    _QCNG-InvokeGraphRequest -Method Post -Uri $sendUri -Headers $headers | Out-Null

    try {
        $getUri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$($meta.graphMessageId)?`$select=id,conversationId,internetMessageId"
        $sent = _QCNG-InvokeGraphRequest -Method Get -Uri $getUri -Headers $headers
        $sentMeta = _QCNG-ExtractGraphMessageMetadata -Message $sent
        foreach ($k in @('graphConversationId', 'internetMessageId', 'graphMessageId')) {
            if (-not (_QCNG-IsBlank $sentMeta[$k])) { $meta[$k] = $sentMeta[$k] }
        }
        $meta.graphImmutableMessageId = $meta.graphMessageId
    } catch { }

    return $meta
}

function Test-QCNotificationGraphParentMessageExists {
    param(
        [Parameter(Mandatory)][string]$SenderMailbox,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$MessageId
    )
    if ($script:QCNG_TestHttpHandler -and $script:QCNG_TestMessageRegistry.ContainsKey($MessageId)) {
        return $true
    }
    $encodedMailbox = [Uri]::EscapeDataString($SenderMailbox)
    $encodedId = [Uri]::EscapeDataString($MessageId)
    $uri = "https://graph.microsoft.com/v1.0/users/$encodedMailbox/messages/$encodedId?`$select=id"
    $headers = _QCNG-NewGraphHeaders -AccessToken $AccessToken -ImmutableId
    try {
        _QCNG-InvokeGraphRequest -Method Get -Uri $uri -Headers $headers | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Send-QCNotificationGraph {
    <#
    .SYNOPSIS
    Sends email via Microsoft Graph using client-secret app credentials (or dry-run preview).
    Supports root, reply, and replacement_root send modes for conversation threading.
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
    $sendMode = if ($Payload.threadSendMode) { [string]$Payload.threadSendMode } else { 'unthreaded' }
    $baseData = @{
        success = $false
        provider = 'MicrosoftGraph'
        dryRun = [bool]$DryRun
        eventType = $Payload.eventType
        documentName = $Payload.documentName
        to = @($Payload.to)
        cc = @($Payload.cc)
        timestampUtc = $timestampUtc
        sendMode = $sendMode
        threadKey = if ($Payload.threadKey) { [string]$Payload.threadKey } else { '' }
        parentGraphMessageId = if ($Payload.parentGraphMessageId) { [string]$Payload.parentGraphMessageId } else { '' }
        notificationThreadId = $Payload.notificationThreadId
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

    $usesHtml = [bool]$Payload.htmlBody
    $graphMessage = $null
    if ($usesHtml) {
        $graphMessage = New-QCGraphEmailMessage -ToRecipients @($Payload.to) -Subject ([string]$Payload.subject) `
            -HtmlBody ([string]$Payload.htmlBody) -LogoPath $logoPath -CcRecipients @($Payload.cc) -Importance $importance
    }

    if ($DryRun) {
        $data = $baseData.Clone()
        $data['success'] = $true
        $data['message'] = "Microsoft Graph dry run: $sendMode payload built; no API call made."
        $data['senderMailbox'] = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'senderMailbox'
        if ($graphMessage) { $data['graphMessage'] = $graphMessage }
        return New-QCSuccessResult -Code 'QC_NOTIFICATION_GRAPH_DRY_RUN' -Message $data.message -Data $data
    }

    if ($validation.usesCertificate -and -not $validation.usesClientSecret) {
        $data = $baseData.Clone()
        $data['message'] = 'Certificate-based Graph auth is not implemented; set notifications.graph.clientSecret.'
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_NOT_IMPLEMENTED' -Message $data.message -Data $data
    }

    $senderMailbox = _QCNG-GetGraphSetting -GraphSettings $GraphSettings -Key 'senderMailbox'
    $parentId = if ($Payload.parentGraphMessageId) { [string]$Payload.parentGraphMessageId } else { '' }
    $attemptSendMode = $sendMode
    $parentInvalid = $false

    try {
        $token = Get-QCNotificationGraphAccessToken -GraphSettings $GraphSettings
        $meta = @{}

        if ($sendMode -eq 'reply' -and -not (_QCNG-IsBlank $parentId)) {
            $parentExists = Test-QCNotificationGraphParentMessageExists -SenderMailbox $senderMailbox -AccessToken $token -MessageId $parentId
            if (-not $parentExists) {
                $parentInvalid = $true
                $attemptSendMode = 'replacement_root'
                $baseData['threadRecoveryReason'] = 'parent_message_not_found'
            }
        }

        if ($attemptSendMode -eq 'reply' -and -not (_QCNG-IsBlank $parentId)) {
            $bodyContent = if ($usesHtml) { [string]$Payload.htmlBody } else { [string]$Payload.body }
            $bodyType = if ($usesHtml) { 'HTML' } else { 'Text' }
            $attachments = $null
            if ($usesHtml -and $graphMessage -and $graphMessage.attachments) {
                $attachments = $graphMessage.attachments
            }
            $meta = Invoke-QCNotificationGraphCreateReplyAndSend -SenderMailbox $senderMailbox -AccessToken $token `
                -ParentMessageId $parentId -Subject ([string]$Payload.subject) -BodyContent $bodyContent `
                -BodyContentType $bodyType -ToRecipients @($Payload.to) -CcRecipients @($Payload.cc) -Attachments $attachments
        }
        elseif ($usesHtml -and $graphMessage) {
            $meta = Invoke-QCNotificationGraphCreateAndSendMessage -SenderMailbox $senderMailbox -AccessToken $token -Message $graphMessage
            if ($attemptSendMode -eq 'unthreaded') { $attemptSendMode = 'root' }
            if ($parentInvalid) { $attemptSendMode = 'replacement_root' }
            elseif ($sendMode -eq 'root' -or $sendMode -eq 'replacement_root') { $attemptSendMode = $sendMode }
            else { $attemptSendMode = 'root' }
        }
        else {
            $sendMailBody = New-QCNotificationGraphSendMailBody -Payload $Payload
            Invoke-QCNotificationGraphSendMail -SenderMailbox $senderMailbox -AccessToken $token -SendMailBody $sendMailBody
            if ($attemptSendMode -eq 'unthreaded') { $attemptSendMode = 'root' }
            if ($parentInvalid) { $attemptSendMode = 'replacement_root' }
            $meta = @{
                graphMessageId = ''
                graphImmutableMessageId = ''
                graphConversationId = ''
                internetMessageId = ''
                threadWarning = 'plain_text_sendmail_no_message_id'
            }
        }
    } catch {
        $data = $baseData.Clone()
        $data['message'] = $_.Exception.Message
        $data['sendMode'] = $attemptSendMode
        $data['threadError'] = $_.Exception.Message
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_SEND_FAILED' -Message $data.message -Data $data
    }

    $data = $baseData.Clone()
    $data['success'] = $true
    $data['message'] = "Microsoft Graph $attemptSendMode send completed."
    $data['senderMailbox'] = $senderMailbox
    $data['sendMode'] = $attemptSendMode
    $data['parentInvalid'] = $parentInvalid
    foreach ($k in @('graphMessageId', 'graphImmutableMessageId', 'graphConversationId', 'internetMessageId')) {
        if ($meta.ContainsKey($k)) { $data[$k] = $meta[$k] }
    }
    if ($meta.threadWarning) { $data['threadWarning'] = $meta.threadWarning }
    return New-QCSuccessResult -Code 'QC_NOTIFICATION_GRAPH_SENT' -Message $data.message -Data $data
}

Export-ModuleMember -Function Test-QCNotificationGraphConfigured, New-QCNotificationGraphSendMailBody, `
    New-QCGraphEmailMessage, Get-QCNotificationGraphAccessToken, Invoke-QCNotificationGraphSendMail, `
    Invoke-QCNotificationGraphCreateAndSendMessage, Invoke-QCNotificationGraphCreateReplyAndSend, `
    Test-QCNotificationGraphParentMessageExists, Register-QCNotificationGraphTestMessage, `
    Set-QCNotificationGraphTestHttpHandler, Clear-QCNotificationGraphTestHttpHandler, `
    Send-QCNotificationGraph
