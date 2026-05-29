# QC.ReviewStamp.psm1
# Applies editable peer-review stamps via qc_overlay_prepend.exe (--apply-review-stamp).
# No Python install required on the host; deploy dist\qc_overlay_prepend\ like the overlay step.

function _QCRS-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCRS-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCRS-GetRepoRoot {
    return (Split-Path -Parent $PSScriptRoot)
}

function Join-QCProcessArgumentList {
    param([string[]]$Args)
    ($Args | ForEach-Object {
        $t = [string]$_
        if ($t -match '[\s"]') { return ('"' + ($t -replace '"', '\"') + '"') }
        return $t
    }) -join ' '
}

function Resolve-QCReviewStampOverlayExe {
    <#
    .SYNOPSIS
    Resolves qc_overlay_prepend.exe (same binary as overlay; includes --apply-review-stamp).
    #>
    param(
        [string]$PreferredPath = '',
        [string]$RepoRoot = ''
    )
    if (_QCRS-IsBlank $RepoRoot) { $RepoRoot = _QCRS-GetRepoRoot }

    if (-not (_QCRS-IsBlank $PreferredPath)) {
        $p = [string]$PreferredPath
        if (-not [System.IO.Path]::IsPathRooted($p)) { $p = Join-Path $RepoRoot $p }
        if (Test-Path -LiteralPath $p) { return $p }
    }

    $candidates = @(
        (Join-Path $RepoRoot 'dist\qc_overlay_prepend\qc_overlay_prepend.exe'),
        (Join-Path $RepoRoot 'dist\qc_overlay_prepend.exe')
    )
    $legacyResolve = Join-Path (Join-Path $RepoRoot 'legacy') 'Resolve-OverlayExe.ps1'
    if (Test-Path -LiteralPath $legacyResolve) {
        . $legacyResolve
        $found = Select-ExistingOverlayExePath (Join-Path $RepoRoot 'legacy')
        if ($found) { $candidates = @([string]$found) + $candidates }
    }

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }
    return $null
}

function Get-QCReviewStampSettings {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$RepoRoot = ''
    )
    if (_QCRS-IsBlank $RepoRoot) { $RepoRoot = _QCRS-GetRepoRoot }

    $qc = _QCRS-ToHashtable $Config.qcPrepend
    if (-not $qc) { return $null }
    $rs = _QCRS-ToHashtable $qc.reviewStamps
    if (-not $rs) { return $null }

    $enabled = $true
    if ($rs.ContainsKey('enabled')) { try { $enabled = [bool]$rs.enabled } catch { $enabled = $true } }
    if (-not $enabled) { return $null }

    $peerType = 'Peer Review'
    $wf = _QCRS-ToHashtable $Config.qcWorkflow
    if ($wf) {
        $rt = _QCRS-ToHashtable $wf.reviewTypes
        if ($rt -and $rt.peerReview) { $peerType = [string]$rt.peerReview }
    }

    $stampPath = Join-Path $RepoRoot 'stamps\Peer_Review_Stamp.pdf'
    $stampHeight = 200.0
    $marginOutside = 12.0
    $pr = _QCRS-ToHashtable $rs.peerReview
    if ($rs.ContainsKey('stampHeightPt')) { try { $stampHeight = [double]$rs.stampHeightPt } catch { } }
    if ($rs.ContainsKey('marginOutsidePt')) { try { $marginOutside = [double]$rs.marginOutsidePt } catch { } }
    if ($pr) {
        if ($pr.stampPath) {
            $sp = [string]$pr.stampPath
            if (-not [System.IO.Path]::IsPathRooted($sp)) { $sp = Join-Path $RepoRoot $sp }
            $stampPath = $sp
        }
        if ($pr.reviewType) { $peerType = [string]$pr.reviewType }
    }
    if (-not (Test-Path -LiteralPath $stampPath)) { return $null }

    $overlayExe = Resolve-QCReviewStampOverlayExe -PreferredPath ([string]$qc.overlayExePath) -RepoRoot $RepoRoot
    if (-not $overlayExe) { return $null }

    return @{
        reviewType      = $peerType
        stampPath       = $stampPath
        stampHeightPt   = $stampHeight
        marginOutsidePt = $marginOutside
        overlayExe      = $overlayExe
    }
}

function Invoke-QCReviewStamp {
    <#
    .SYNOPSIS
    Stamps page 1 of a QC PDF using qc_overlay_prepend.exe --apply-review-stamp (editable AcroForm fields).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OverlayExe,
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][string]$StampPath,
        [string]$Originator = '',
        [string]$Checker = '',
        [string]$Backchecker = '',
        [string]$OriginatorDate = '',
        [double]$StampHeightPt = 200,
        [double]$MarginOutsidePt = 12
    )

    if (-not (Test-Path -LiteralPath $OverlayExe)) {
        return @{ applied = $false; reason = "Overlay exe not found: $OverlayExe" }
    }
    if (-not (Test-Path -LiteralPath $PdfPath)) {
        return @{ applied = $false; reason = "PDF not found: $PdfPath" }
    }
    if (-not (Test-Path -LiteralPath $StampPath)) {
        return @{ applied = $false; reason = "Stamp template not found: $StampPath" }
    }

    if (_QCRS-IsBlank $OriginatorDate) {
        $OriginatorDate = (Get-Date).ToString('MM/dd/yyyy')
    }

    $stampArgs = @(
        '--apply-review-stamp',
        $PdfPath,
        $StampPath,
        '--originator', [string]$Originator,
        '--checker', [string]$Checker,
        '--backchecker', [string]$Backchecker,
        '--originator-date', [string]$OriginatorDate,
        '--stamp-height-pt', [string]$StampHeightPt,
        '--margin-outside-pt', [string]$MarginOutsidePt
    )

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $argLine = Join-QCProcessArgumentList -Args $stampArgs
        $p = Start-Process -FilePath $OverlayExe -ArgumentList $argLine -Wait -NoNewWindow -PassThru `
            -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = [string](Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue)
        $stderr = [string](Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue)
        $combined = ($stdout + "`n" + $stderr).Trim()
        if ($p.ExitCode -ne 0) {
            return @{ applied = $false; reason = "review stamp exit $($p.ExitCode)"; stdout = $combined }
        }
        return @{ applied = $true; stdout = $combined }
    } catch {
        return @{ applied = $false; reason = $_.Exception.Message }
    } finally {
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-QCReviewStampIfPeerReview {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$RoleFields,
        [string]$OverlayExe = '',
        [scriptblock]$Log = $null
    )

    $stampCfg = Get-QCReviewStampSettings -Config $Config
    if (-not $stampCfg) {
        return @{ applied = $false; skipped = $true; reason = 'review stamps disabled or overlay exe / stamp template missing' }
    }

    $reviewType = [string]$RoleFields.qcReviewType
    if (_QCRS-IsBlank $reviewType -or ($reviewType -ne [string]$stampCfg.reviewType)) {
        return @{ applied = $false; skipped = $true; reason = "QC_Review_Type is '$reviewType' (need '$($stampCfg.reviewType)')" }
    }

    $exe = if (-not (_QCRS-IsBlank $OverlayExe)) { $OverlayExe } else { [string]$stampCfg.overlayExe }
    if ($Log) {
        & $Log "Applying peer review stamp (outside top-left) to page 1: $PdfPath"
    }

    $result = Invoke-QCReviewStamp -OverlayExe $exe -PdfPath $PdfPath -StampPath ([string]$stampCfg.stampPath) `
        -Originator ([string]$RoleFields.designerEmail) `
        -Checker ([string]$RoleFields.reviewerEmail) `
        -Backchecker ([string]$RoleFields.checkerEmail) `
        -StampHeightPt ([double]$stampCfg.stampHeightPt) `
        -MarginOutsidePt ([double]$stampCfg.marginOutsidePt)

    $result['reviewType'] = $reviewType
    if ($result.applied -and $Log) {
        & $Log 'Peer review stamp applied (editable fields).'
    }
    return $result
}

Export-ModuleMember -Function @(
    'Join-QCProcessArgumentList',
    'Resolve-QCReviewStampOverlayExe',
    'Get-QCReviewStampSettings',
    'Invoke-QCReviewStamp',
    'Invoke-QCReviewStampIfPeerReview'
)
