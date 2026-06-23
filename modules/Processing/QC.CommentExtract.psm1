# QC.CommentExtract.psm1
# Responsibility: Invoke PDF comment parser and normalize to framework annotation objects.

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Processing/QC.PdfExport.psm1') -Force

function _QCE-IsNullOrWhiteSpace([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Get-QCCommentExtractSettings {
    param([hashtable]$Config)
    $settings = Get-QCCommentSyncSettings -Config $Config
    if (-not $settings.parserVersion) { $settings['parserVersion'] = '1.0.0' }
    if (-not $settings.pythonExecutable) { $settings['pythonExecutable'] = 'python' }
    if (-not $settings.commentExtractScript) { $settings['commentExtractScript'] = 'tools\overlay\qc_pdf_comments.py' }
    return $settings
}

function ConvertTo-QCNormalizedAnnotations {
    param([object[]]$RawAnnotations)
    $out = @()
    foreach ($a in @($RawAnnotations)) {
        if ($a -is [hashtable]) { $h = $a }
        elseif ($a.PSObject) {
            $h = @{}
            foreach ($p in $a.PSObject.Properties) { $h[$p.Name] = $p.Value }
        } else { continue }
        $out += @{
            annotation_id = if ($h.annotation_id) { [string]$h.annotation_id } else { '' }
            page_number = if ($null -ne $h.page_number) { [int]$h.page_number } else { 0 }
            author = if ($h.author) { [string]$h.author } else { '' }
            subject = if ($h.subject) { [string]$h.subject } else { '' }
            comment_text = if ($h.comment_text) { [string]$h.comment_text } else { '' }
            color = if ($h.color) { [string]$h.color } else { '' }
            status = if ($h.status) { [string]$h.status } else { 'Unknown' }
            status_author = if ($h.status_author) { [string]$h.status_author } else { '' }
            status_timestamp_utc = if ($h.status_timestamp_utc) { [string]$h.status_timestamp_utc } else { '' }
            created_utc = if ($h.created_utc) { [string]$h.created_utc } else { '' }
            modified_utc = if ($h.modified_utc) { [string]$h.modified_utc } else { '' }
            parent_annotation_id = if ($h.parent_annotation_id) { [string]$h.parent_annotation_id } else { '' }
            raw = $h.raw
        }
    }
    return @($out)
}

function Invoke-QCCommentExtract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LocalPdfPath,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [object[]]$MockAnnotations = $null
    )

    $settings = Get-QCCommentExtractSettings -Config $Config

    if ($null -ne $MockAnnotations) {
        $normalized = ConvertTo-QCNormalizedAnnotations -RawAnnotations $MockAnnotations
        return New-QCSuccessResult -Code 'QC_COMMENT_EXTRACT_OK' -Message 'Mock annotations returned.' -Data @{
            annotations = $normalized
            parserStatus = if ($normalized.Count -gt 0) { 'ok' } else { 'empty' }
            parserVersion = [string]$settings.parserVersion
            warnings = @()
        }
    }

    if (-not (Test-Path -LiteralPath $LocalPdfPath)) {
        return New-QCFailureResult -Code 'QC_COMMENT_EXTRACT_FILE_MISSING' -Message "PDF not found: $LocalPdfPath" -Data @{}
    }

    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $scriptPath = [string]$settings.commentExtractScript
    if (-not [System.IO.Path]::IsPathRooted($scriptPath)) {
        $scriptPath = Join-Path $repoRoot $scriptPath
    }
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        return New-QCFailureResult -Code 'QC_COMMENT_EXTRACT_SCRIPT_MISSING' -Message "Parser script not found: $scriptPath" -Data @{}
    }

    $python = [string]$settings.pythonExecutable
    $args = @($scriptPath, '--input', $LocalPdfPath, '--output', '-')
    try {
        $jsonText = & $python @args 2>&1 | Out-String
        if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
            return New-QCFailureResult -Code 'QC_COMMENT_EXTRACT_FAILED' -Message 'Python parser exited with error.' -Data @{ exitCode = $LASTEXITCODE; output = $jsonText }
        }
        $parsed = $jsonText | ConvertFrom-Json
        $annots = @()
        if ($parsed.annotations) { $annots = @($parsed.annotations) }
        $normalized = ConvertTo-QCNormalizedAnnotations -RawAnnotations $annots
        return New-QCSuccessResult -Code 'QC_COMMENT_EXTRACT_OK' -Message 'Comments extracted.' -Data @{
            annotations = $normalized
            parserStatus = [string]$parsed.parser_status
            parserVersion = if ($parsed.parser_version) { [string]$parsed.parser_version } else { [string]$settings.parserVersion }
            warnings = if ($parsed.warnings) { @($parsed.warnings) } else { @() }
        }
    } catch {
        return New-QCFailureResult -Code 'QC_COMMENT_EXTRACT_THROW' -Message $_.Exception.Message -Data @{ error = $_.Exception.Message }
    }
}

Export-ModuleMember -Function Get-QCCommentExtractSettings, ConvertTo-QCNormalizedAnnotations, Invoke-QCCommentExtract
