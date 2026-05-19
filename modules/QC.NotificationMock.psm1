# QC.NotificationMock.psm1
# Responsibility: Mock/dry-run notification delivery (log + local payload files).

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force

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

function Send-QCNotificationMock {
    <#
    .SYNOPSIS
    Writes a mock notification payload to disk and returns a provider result object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Payload,
        [Parameter(Mandatory)]
        [string]$OutputRoot,
        [switch]$DryRun
    )

    $timestampUtc = [DateTime]::UtcNow.ToString('o')
    $mockDir = Join-Path $OutputRoot 'mock'
    _QCNM-EnsureDirectory -Path $mockDir

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
    }
}

Export-ModuleMember -Function Send-QCNotificationMock
