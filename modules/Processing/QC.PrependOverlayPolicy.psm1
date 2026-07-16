# QC.PrependOverlayPolicy.psm1
# Shared policy for when QC_PREPEND should skip qc_overlay_prepend and use qpdf only.

function Test-QCOverlaySkipByIncomingSize {
    <#
    .SYNOPSIS
    Decide whether incoming PDF size should skip layered overlay.
    .DESCRIPTION
    Reads qcPrepend.overlaySkipIncomingAboveMb. When threshold is > 0 and
    IncomingBytes is strictly greater than Mb * 1024 * 1024, Skip is true.
    Threshold 0 / missing / null disables the gate (Skip = false).
    .PARAMETER QcPrepend
    The qcPrepend config hashtable (or empty).
    .PARAMETER IncomingBytes
    Byte length of the exported incoming PDF.
    .OUTPUTS
    PSCustomObject: Skip, ThresholdMb, ThresholdBytes, IncomingBytes, Reason
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$QcPrepend = @{},

        [Parameter(Mandatory = $true)]
        [ValidateRange(0, [long]::MaxValue)]
        [long]$IncomingBytes
    )

    $thresholdMb = 0.0
    if ($QcPrepend -and $QcPrepend.ContainsKey('overlaySkipIncomingAboveMb') -and $null -ne $QcPrepend.overlaySkipIncomingAboveMb) {
        try {
            $thresholdMb = [double]$QcPrepend.overlaySkipIncomingAboveMb
        } catch {
            $thresholdMb = 0.0
        }
    }

    if ($thresholdMb -le 0) {
        return [pscustomobject]@{
            Skip           = $false
            ThresholdMb    = 0.0
            ThresholdBytes = [long]0
            IncomingBytes  = $IncomingBytes
            Reason         = 'threshold_disabled'
        }
    }

    $thresholdBytes = [long]([math]::Floor($thresholdMb * 1024.0 * 1024.0))
    $skip = ($IncomingBytes -gt $thresholdBytes)
    return [pscustomobject]@{
        Skip           = [bool]$skip
        ThresholdMb    = $thresholdMb
        ThresholdBytes = $thresholdBytes
        IncomingBytes  = $IncomingBytes
        Reason         = if ($skip) { 'incoming_above_threshold' } else { 'incoming_within_threshold' }
    }
}

Export-ModuleMember -Function Test-QCOverlaySkipByIncomingSize
