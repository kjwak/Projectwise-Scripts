# QC.NotificationGraph.psm1
# Responsibility: Microsoft Graph email provider (stub until tenant credentials are available).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

function _QCNG-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Test-QCNotificationGraphConfigured {
    <#
    .SYNOPSIS
    Returns whether Microsoft Graph notification settings contain required non-secret identifiers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$GraphSettings
    )

    $required = @('tenantId', 'clientId', 'senderMailbox')
    $missing = @()
    foreach ($key in @($required)) {
        if (-not $GraphSettings -or -not $GraphSettings.ContainsKey($key) -or (_QCNG-IsBlank $GraphSettings[$key])) {
            $missing += $key
        }
    }
    $hasCert = $false
    if ($GraphSettings) {
        if (-not (_QCNG-IsBlank $GraphSettings.certificateThumbprint)) { $hasCert = $true }
        if (-not (_QCNG-IsBlank $GraphSettings.certificatePath)) { $hasCert = $true }
    }
    if (-not $hasCert) { $missing += 'certificateThumbprint|certificatePath' }

    return @{
        configured = ($missing.Count -eq 0)
        missing = @($missing)
    }
}

function Send-QCNotificationGraph {
    <#
    .SYNOPSIS
    Sends email via Microsoft Graph (certificate-based app auth — not yet implemented).
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
    $validation = Test-QCNotificationGraphConfigured -GraphSettings $GraphSettings
    if (-not $validation.configured) {
        return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_NOT_CONFIGURED' -Message 'Microsoft Graph provider is not configured.' -Data @{
            success = $false
            provider = 'MicrosoftGraph'
            dryRun = [bool]$DryRun
            eventType = $Payload.eventType
            documentName = $Payload.documentName
            to = @($Payload.to)
            cc = @($Payload.cc)
            message = ('Missing Graph configuration: ' + ($validation.missing -join ', '))
            timestampUtc = $timestampUtc
            missing = @($validation.missing)
        }
    }

  # TODO: Acquire OAuth token using certificate-based client credentials (Entra app registration).
  # TODO: POST https://graph.microsoft.com/v1.0/users/{senderMailbox}/sendMail with message payload.
  # TODO: Honor DryRun by building the request body without calling Graph when -DryRun is set.

    return New-QCFailureResult -Code 'QC_NOTIFICATION_GRAPH_NOT_IMPLEMENTED' -Message 'Microsoft Graph sendMail is not implemented yet.' -Data @{
        success = $false
        provider = 'MicrosoftGraph'
        dryRun = [bool]$DryRun
        eventType = $Payload.eventType
        documentName = $Payload.documentName
        to = @($Payload.to)
        cc = @($Payload.cc)
        message = 'Microsoft Graph sendMail is not implemented yet.'
        timestampUtc = $timestampUtc
    }
}

Export-ModuleMember -Function Test-QCNotificationGraphConfigured, Send-QCNotificationGraph
