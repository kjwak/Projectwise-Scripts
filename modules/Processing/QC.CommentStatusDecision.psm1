# QC.CommentStatusDecision.psm1
# Responsibility: Pure workflow state decision from normalized annotations (no side effects).

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Core/Core.Results.psm1') -Force

function Get-QCCommentDecisionSettings {
    param([hashtable]$Config)
    $defaults = @{
        reviewerAuthorPatterns = @('(?i)reviewer|qc|typsa')
        statusMappings = @{
            open = @('Open', 'None', '', 'Unknown')
            resolved = @('Completed', 'Accepted', 'Checked')
            closed = @('Closed', 'Cancelled')
        }
        targetStates = @{
            redlinesReceived = 'Redlines Received'
            qcFinalizing = 'Initiate Verification'
            completed = 'Verified'
            error = 'Error Needs Attention'
        }
        strictMode = $false
    }
    $out = @{}
    foreach ($k in $defaults.Keys) { $out[$k] = $defaults[$k] }
    if ($Config -and $Config.ContainsKey('qcCommentSync') -and $Config.qcCommentSync) {
        $cs = $Config.qcCommentSync
        if ($cs.reviewerAuthorPatterns) { $out.reviewerAuthorPatterns = @($cs.reviewerAuthorPatterns) }
        if ($cs.statusMappings) { $out.statusMappings = $cs.statusMappings }
        if ($cs.targetStates) { $out.targetStates = $cs.targetStates }
        try { if ($null -ne $cs.strictMode) { $out.strictMode = [bool]$cs.strictMode } } catch { }
    }
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) {
        $wf = $Config.qcWorkflow
        $wfStates = $null
        if ($wf.states) { $wfStates = $wf.states }
        if ($wfStates) {
            if ($wfStates.redlinesReceived) { $out.targetStates.redlinesReceived = [string]$wfStates.redlinesReceived }
            elseif ($wfStates.redlinesIssued) { $out.targetStates.redlinesReceived = [string]$wfStates.redlinesIssued }
            if ($wfStates.qcFinalizing) { $out.targetStates.qcFinalizing = [string]$wfStates.qcFinalizing }
            if ($wfStates.complete) { $out.targetStates.completed = [string]$wfStates.complete }
            if ($wfStates.error) { $out.targetStates.error = [string]$wfStates.error }
        }
        if ($wf.correctionsInProgressStateName) { $out.targetStates.qcFinalizing = [string]$wf.correctionsInProgressStateName }
        if ($wf.backcheckInProgressStateName) { $out.targetStates.qcFinalizing = [string]$wf.backcheckInProgressStateName }
        if ($wf.errorStateName) { $out.targetStates.error = [string]$wf.errorStateName }
        try { if ($null -ne $wf.strictMode) { $out.strictMode = [bool]$wf.strictMode } } catch { }
    }
    return $out
}

function Test-QCAnnotationIsReviewer {
    param(
        [hashtable]$Annotation,
        [string[]]$Patterns
    )
    $author = if ($Annotation.author) { [string]$Annotation.author } else { '' }
    foreach ($rx in @($Patterns)) {
        if ($author -match [string]$rx) { return $true }
    }
    return $false
}

function Get-QCAnnotationStatusBucket {
    param(
        [string]$Status,
        [hashtable]$StatusMappings
    )
    $s = if ($Status) { $Status.Trim() } else { '' }
    foreach ($key in @('open', 'resolved', 'closed')) {
        $vals = @($StatusMappings[$key])
        foreach ($v in $vals) {
            if ([string]$v -eq $s) { return $key }
        }
    }
    return 'open'
}

function Resolve-QCCommentWorkflowState {
    <#
    .SYNOPSIS
    Pure decision: annotations + settings -> target workflow state (no I/O).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$Annotations = @(),
        [hashtable]$Settings = $null,
        [hashtable]$Config = $null,
        [string]$ParserStatus = 'ok'
    )

    if (-not $Settings) { $Settings = Get-QCCommentDecisionSettings -Config $(if ($Config) { $Config } else { @{} }) }
    $targets = $Settings.targetStates
    $warnings = [System.Collections.Generic.List[string]]::new()

    if ($ParserStatus -eq 'error') {
        return @{
            targetState = [string]$targets.error
            decisionCode = 'PARSE_ERROR'
            summary = 'Parser reported error.'
            reviewerOpenCount = 0
            reviewerResolvedCount = 0
            reviewerClosedCount = 0
            warnings = @('Parser status was error.')
        }
    }

    $reviewerAnnots = @($Annotations | Where-Object {
        $_ -is [hashtable] -and (Test-QCAnnotationIsReviewer -Annotation $_ -Patterns @($Settings.reviewerAuthorPatterns))
    })

    if ($reviewerAnnots.Count -eq 0 -and @($Annotations).Count -gt 0) {
        $reviewerAnnots = @($Annotations | Where-Object { $_ -is [hashtable] })
        $warnings.Add('No reviewer-pattern match; treating all annotations as reviewer scope.') | Out-Null
    }

    if ($reviewerAnnots.Count -eq 0) {
        if ([bool]$Settings.strictMode) {
            return @{
                targetState = [string]$targets.error
                decisionCode = 'NO_REVIEWER_ANNOTATIONS'
                summary = 'No reviewer annotations found.'
                reviewerOpenCount = 0
                reviewerResolvedCount = 0
                reviewerClosedCount = 0
                warnings = @($warnings)
            }
        }
        return @{
            targetState = [string]$targets.qcFinalizing
            decisionCode = 'NO_ANNOTATIONS'
            summary = 'No annotations; defaulting to QC Finalizing.'
            reviewerOpenCount = 0
            reviewerResolvedCount = 0
            reviewerClosedCount = 0
            warnings = @($warnings)
        }
    }

    $open = 0
    $resolved = 0
    $closed = 0
    $conflict = $false
    foreach ($a in @($reviewerAnnots)) {
        $bucket = Get-QCAnnotationStatusBucket -Status ([string]$a.status) -StatusMappings $Settings.statusMappings
        switch ($bucket) {
            'open' { $open++ }
            'resolved' { $resolved++ }
            'closed' { $closed++ }
            default { $open++ }
        }
    }

    if ($open -gt 0 -and $closed -gt 0 -and $resolved -eq 0) {
        $conflict = $true
        $warnings.Add('Mixed open and closed reviewer comments without resolved state.') | Out-Null
    }

    if ($conflict) {
        return @{
            targetState = [string]$targets.error
            decisionCode = 'STATUS_CONFLICT'
            summary = 'Conflicting comment statuses detected.'
            reviewerOpenCount = $open
            reviewerResolvedCount = $resolved
            reviewerClosedCount = $closed
            warnings = @($warnings)
        }
    }

    if ($open -gt 0) {
        $redlinesState = if ($targets.redlinesReceived) { [string]$targets.redlinesReceived } elseif ($targets.redlinesIssued) { [string]$targets.redlinesIssued } else { 'Redlines Received' }
        return @{
            targetState = $redlinesState
            decisionCode = 'REDLINES_RECEIVED'
            summary = "$open reviewer comment(s) still open; redlines returned to designer."
            reviewerOpenCount = $open
            reviewerResolvedCount = $resolved
            reviewerClosedCount = $closed
            warnings = @($warnings)
        }
    }

    if ($resolved -gt 0) {
        $finalizingState = if ($targets.qcFinalizing) { [string]$targets.qcFinalizing } else { 'Initiate Verification' }
        return @{
            targetState = $finalizingState
            decisionCode = 'QC_FINALIZING'
            summary = "$resolved reviewer comment(s) resolved; correction prepend will run via Initiate Verification."
            reviewerOpenCount = $open
            reviewerResolvedCount = $resolved
            reviewerClosedCount = $closed
            warnings = @($warnings)
        }
    }

    $finalizingState = if ($targets.qcFinalizing) { [string]$targets.qcFinalizing } else { 'QC Finalizing' }
    return @{
        targetState = $finalizingState
        decisionCode = 'QC_FINALIZING'
        summary = 'All reviewer comments closed or accepted; ready for final QC prepend.'
        reviewerOpenCount = $open
        reviewerResolvedCount = $resolved
        reviewerClosedCount = $closed
        warnings = @($warnings)
    }
}

Export-ModuleMember -Function Get-QCCommentDecisionSettings, Test-QCAnnotationIsReviewer, Get-QCAnnotationStatusBucket, Resolve-QCCommentWorkflowState
