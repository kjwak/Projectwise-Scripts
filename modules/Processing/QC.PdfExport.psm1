# QC.PdfExport.psm1
# Responsibility: Download/export ProjectWise PDFs to local staging for inspection.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.StatusSet.psm1') -Force -ErrorAction SilentlyContinue

function _QPE-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Get-QCCommentSyncSettings {
    param([hashtable]$Config)
    $defaults = @{
        enabled = $true
        stagingRoot = ''
        retainTempFiles = $false
        processorVersion = '1.0.0'
        failOnStateApply = $true
    }
    $out = @{}
    foreach ($k in $defaults.Keys) { $out[$k] = $defaults[$k] }
    if ($Config -and $Config.ContainsKey('qcCommentSync') -and $Config.qcCommentSync) {
        foreach ($k in $Config.qcCommentSync.Keys) { $out[$k] = $Config.qcCommentSync[$k] }
    }
    return $out
}

function Get-QCCommentSyncStagingPath {
    param(
        [hashtable]$Config,
        [string]$JobId
    )
    $settings = Get-QCCommentSyncSettings -Config $Config
    $root = [string]$settings.stagingRoot
    if (_QPE-IsNullOrWhiteSpace $root) {
        $root = Join-Path $env:TEMP 'qc_comment_sync'
    }
    if (_QPE-IsNullOrWhiteSpace $JobId) { $JobId = [guid]::NewGuid().ToString('N') }
    return (Join-Path $root $JobId)
}

function Export-QCPdfToStaging {
    <#
    .SYNOPSIS
    Exports a ProjectWise document PDF to a per-job staging folder.
    .PARAMETER MockExporter
    Optional scriptblock(Doc, TargetFolder) returning QCResult with Data.localPath for tests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $InputDocument,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$JobId,
        [scriptblock]$MockExporter = $null
    )

    $settings = Get-QCCommentSyncSettings -Config $Config
    $staging = Get-QCCommentSyncStagingPath -Config $Config -JobId $JobId
    if (-not (Test-Path -LiteralPath $staging)) {
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
    }

    if ($MockExporter) {
        try {
            $mr = & $MockExporter $InputDocument $staging
            if ($mr -is [hashtable] -and $mr.ContainsKey('IsSuccess')) {
                if ([bool]$mr.IsSuccess) {
                    $mr = New-QCSuccessResult -Code ([string]$mr.Code) -Message ([string]$mr.Message) -Data $mr.Data
                } else {
                    $mr = New-QCFailureResult -Code ([string]$mr.Code) -Message ([string]$mr.Message) -Data $mr.Data
                }
            }
            if ($mr -and $mr.IsSuccess -and $mr.Data.localPath) {
                return New-QCSuccessResult -Code 'QC_PDF_EXPORT_OK' -Message 'Mock export succeeded.' -Data @{
                    localPath = [string]$mr.Data.localPath
                    stagingFolder = $staging
                }
            }
            return New-QCFailureResult -Code 'QC_PDF_EXPORT_MOCK_FAILED' -Message 'Mock exporter did not return localPath.' -Data @{ stagingFolder = $staging }
        } catch {
            return New-QCFailureResult -Code 'QC_PDF_EXPORT_MOCK_THROW' -Message $_.Exception.Message -Data @{ error = $_.Exception.Message }
        }
    }

    if (-not (Get-Command -Name 'Export-StatusSetPdfToFolder' -ErrorAction SilentlyContinue)) {
        return New-QCFailureResult -Code 'QC_PDF_EXPORT_MODULE_MISSING' -Message 'Export-StatusSetPdfToFolder not available.' -Data @{}
    }

    $exportRes = Export-StatusSetPdfToFolder -InputDocument $InputDocument -TargetFolder $staging
    if (-not $exportRes.IsSuccess) { return $exportRes }
    return New-QCSuccessResult -Code 'QC_PDF_EXPORT_OK' -Message 'PDF exported to staging.' -Data @{
        localPath = [string]$exportRes.Data.localPath
        stagingFolder = $staging
    }
}

function Remove-QCCommentSyncStaging {
    param(
        [hashtable]$Config,
        [string]$StagingFolder
    )
    $settings = Get-QCCommentSyncSettings -Config $Config
    if ([bool]$settings.retainTempFiles) {
        return New-QCSuccessResult -Code 'QC_PDF_STAGING_RETAINED' -Message 'Staging retained per config.' -Data @{ path = $StagingFolder }
    }
    if (_QPE-IsNullOrWhiteSpace $StagingFolder) {
        return New-QCSuccessResult -Code 'QC_PDF_STAGING_SKIP' -Message 'No staging path to remove.' -Data @{}
    }
    try {
        if (Test-Path -LiteralPath $StagingFolder) {
            Remove-Item -LiteralPath $StagingFolder -Recurse -Force -ErrorAction Stop
        }
        return New-QCSuccessResult -Code 'QC_PDF_STAGING_REMOVED' -Message 'Staging folder removed.' -Data @{ path = $StagingFolder }
    } catch {
        return New-QCFailureResult -Code 'QC_PDF_STAGING_REMOVE_FAILED' -Message $_.Exception.Message -Data @{ path = $StagingFolder }
    }
}

Export-ModuleMember -Function Get-QCCommentSyncSettings, Get-QCCommentSyncStagingPath, Export-QCPdfToStaging, Remove-QCCommentSyncStaging
