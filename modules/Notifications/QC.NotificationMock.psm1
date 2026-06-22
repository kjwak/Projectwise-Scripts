# QC.NotificationMock.psm1
# Responsibility: Mock/dry-run notification delivery with deterministic email threading.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Runtime.psm1') -Force

function _QCNM-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCNM-EnsureDirectory([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function _QCNM-GetThreadStorePath([string]$OutputRoot) {
    return (Join-Path $OutputRoot 'mock\threads')
}

function _QCNM-LoadThreadState([string]$ThreadKey, [string]$OutputRoot) {
    $dir = _QCNM-GetThreadStorePath -OutputRoot $OutputRoot
    $safeKey = ($ThreadKey -replace '[\\/:*?"<>|]+', '_')
    $path = Join-Path $dir "$safeKey.json"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function _QCNM-SaveThreadState([string]$ThreadKey, [string]$OutputRoot, [hashtable]$State) {
    $dir = _QCNM-GetThreadStorePath -OutputRoot $OutputRoot
    _QCNM-EnsureDirectory -Path $dir
    $safeKey = ($ThreadKey -replace '[\\/:*?"<>|]+', '_')
    $path = Join-Path $dir "$safeKey.json"
    ($State | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $path -Encoding UTF8
    return $path
}

function _QCNM-BuildDeterministicMessageId([string]$ThreadKey, [int]$Sequence) {
    $hash = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($ThreadKey))).Replace('-', '').Substring(0, 12).ToLowerInvariant()
    return "mock-$hash-$Sequence"
}

function Send-QCNotificationMock {
    <#
    .SYNOPSIS
    Writes a mock notification payload to disk and returns a provider result object.
    Supports root, reply, and replacement_root threading when threadKey is present.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Payload,
        [Parameter(Mandatory)]
        [string]$OutputRoot,
        [switch]$DryRun
    )

    $timestampUtc = Get-QCTimestamp
    $mockDir = Join-Path $OutputRoot 'mock'
    _QCNM-EnsureDirectory -Path $mockDir

    $threadKey = if ($Payload.threadKey) { [string]$Payload.threadKey } else { '' }
    $sendMode = if ($Payload.threadSendMode) { [string]$Payload.threadSendMode } else { 'unthreaded' }
    $parentMessageId = if ($Payload.parentGraphMessageId) { [string]$Payload.parentGraphMessageId } else { '' }
    $threadId = $null
    $messageId = ''
    $conversationId = ''

    if (-not (_QCNM-IsBlank $threadKey) -and $sendMode -ne 'unthreaded') {
        $existing = _QCNM-LoadThreadState -ThreadKey $threadKey -OutputRoot $OutputRoot
        $sequence = 1
        if ($existing) {
            if ($existing.thread_id) { $threadId = [int]$existing.thread_id }
            if ($existing.message_count) { $sequence = [int]$existing.message_count + 1 }
            if ($existing.conversation_id) { $conversationId = [string]$existing.conversation_id }
        }
        if (-not $threadId) { $threadId = 1 }

        if ($sendMode -eq 'reply' -and -not (_QCNM-IsBlank $parentMessageId)) {
            $parentValid = $false
            if ($parentMessageId -match '(?i)^stale-') {
                $sendMode = 'replacement_root'
                $parentMessageId = ''
            }
            elseif ($existing -and $existing.latest_message_id -eq $parentMessageId) { $parentValid = $true }
            if (-not $parentValid -and $sendMode -eq 'reply') {
                $sendMode = 'replacement_root'
                $parentMessageId = ''
            }
        }

        if ($sendMode -in @('root', 'replacement_root') -and -not $conversationId) {
            $conversationId = "mock-conv-$threadKey"
        }
        $messageId = _QCNM-BuildDeterministicMessageId -ThreadKey $threadKey -Sequence $sequence

        $threadState = @{
            thread_id = $threadId
            thread_key = $threadKey
            conversation_id = $conversationId
            latest_message_id = $messageId
            message_count = $sequence
            updated_at = $timestampUtc
        }
        _QCNM-SaveThreadState -ThreadKey $threadKey -OutputRoot $OutputRoot -State $threadState | Out-Null
    }

    $safeDoc = if ($Payload.documentName) {
        (($Payload.documentName -replace '[\\/:*?"<>|]+', '_') -replace '\s+', '_')
    } else { 'document' }
    $safeEvent = if ($Payload.eventType) { $Payload.eventType } else { 'EVENT' }
    $fileName = ('{0}_{1}_{2}.json' -f $timestampUtc.Replace(':', '-'), $safeEvent, $safeDoc)
    $filePath = Join-Path $mockDir $fileName

    $filePayload = @{}
    foreach ($k in $Payload.Keys) { $filePayload[$k] = $Payload[$k] }
    $filePayload['timestampUtc'] = $timestampUtc
    $filePayload['dryRun'] = [bool]$DryRun
    $filePayload['thread_id'] = $threadId
    $filePayload['thread_key'] = $threadKey
    $filePayload['parent_message_id'] = $parentMessageId
    $filePayload['message_id'] = $messageId
    $filePayload['send_mode'] = $sendMode
    $filePayload['graphConversationId'] = $conversationId
    $filePayload['graphMessageId'] = $messageId
    $filePayload['graphImmutableMessageId'] = $messageId

    try {
        ($filePayload | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $filePath -Encoding UTF8 -ErrorAction Stop
    } catch {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_MOCK_WRITE_FAILED' -Message 'Failed to write mock notification payload.' -Data @{
            success = $false
            provider = 'Mock'
            dryRun = [bool]$DryRun
            eventType = $Payload.eventType
            documentName = $Payload.documentName
            to = @($Payload.to)
            cc = @($Payload.cc)
            message = $_.Exception.Message
            timestampUtc = $timestampUtc
            filePath = $filePath
            sendMode = $sendMode
        }
    }

    return New-QCSuccessResult -Code 'QC_NOTIFICATION_MOCK_SENT' -Message 'Mock notification written.' -Data @{
        success = $true
        provider = 'Mock'
        dryRun = [bool]$DryRun
        eventType = $Payload.eventType
        documentName = $Payload.documentName
        to = @($Payload.to)
        cc = @($Payload.cc)
        message = 'Mock notification written.'
        timestampUtc = $timestampUtc
        filePath = $filePath
        thread_id = $threadId
        thread_key = $threadKey
        parent_message_id = $parentMessageId
        message_id = $messageId
        send_mode = $sendMode
        sendMode = $sendMode
        graphMessageId = $messageId
        graphImmutableMessageId = $messageId
        graphConversationId = $conversationId
        notificationThreadId = $threadId
    }
}

Export-ModuleMember -Function Send-QCNotificationMock
