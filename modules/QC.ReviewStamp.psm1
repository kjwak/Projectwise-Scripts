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

function _QCRS-AppendOptionalCliFlags {
    param(
        [System.Collections.Generic.List[object]]$TokenList,
        [string]$Flag,
        [string]$Value
    )
    if (_QCRS-IsBlank $Value) { return }
    [void]$TokenList.Add($Flag)
    [void]$TokenList.Add([string]$Value)
}

function _QCRS-BuildProcessArgumentLine {
    param([object[]]$Tokens = @())
    $parts = @()
    foreach ($item in @($Tokens)) {
        if ($null -eq $item) { continue }
        $t = [string]$item
        if ($t.Length -eq 0) { continue }
        if ($t -match '[\s"]') { $parts += ('"' + ($t -replace '"', '\"') + '"') }
        else { $parts += $t }
    }
    return ($parts -join ' ')
}

function Join-QCProcessArgumentList {
    <#
    Single command-line string for Start-Process -ArgumentList (one string). Avoid parameter name $Args (conflicts with automatic $args).
    #>
    param(
        [AllowEmptyCollection()]
        [object[]]
        $ArgumentTokens = @()
    )
    return (_QCRS-BuildProcessArgumentLine -Tokens $ArgumentTokens)
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
    $stampX = $null
    $stampY = $null
    $pr = _QCRS-ToHashtable $rs.peerReview
    if ($rs.ContainsKey('stampHeightPt')) { try { $stampHeight = [double]$rs.stampHeightPt } catch { } }
    if ($rs.ContainsKey('marginOutsidePt')) { try { $marginOutside = [double]$rs.marginOutsidePt } catch { } }
    $pos = _QCRS-ToHashtable $rs.stampPositionPt
    if ($pos) {
        if ($pos.ContainsKey('x')) { try { $stampX = [double]$pos.x } catch { } }
        if ($pos.ContainsKey('y')) { try { $stampY = [double]$pos.y } catch { } }
    }
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
        stampXPt        = $stampX
        stampYPt        = $stampY
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
        [double]$MarginOutsidePt = 12,
        [Nullable[double]]$StampXPt = $null,
        [Nullable[double]]$StampYPt = $null
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

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $stampTokens = [System.Collections.Generic.List[object]]::new()
        [void]$stampTokens.Add('--apply-review-stamp')
        [void]$stampTokens.Add([string]$PdfPath)
        [void]$stampTokens.Add([string]$StampPath)
        _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--originator' -Value $Originator
        _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--checker' -Value $Checker
        _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--backchecker' -Value $Backchecker
        [void]$stampTokens.Add('--originator-date')
        [void]$stampTokens.Add([string]$OriginatorDate)
        [void]$stampTokens.Add('--stamp-height-pt')
        [void]$stampTokens.Add([string]$StampHeightPt)
        if ($null -ne $StampXPt -and $null -ne $StampYPt) {
            [void]$stampTokens.Add('--stamp-x-pt')
            [void]$stampTokens.Add([string]$StampXPt)
            [void]$stampTokens.Add('--stamp-y-pt')
            [void]$stampTokens.Add([string]$StampYPt)
        } else {
            [void]$stampTokens.Add('--margin-outside-pt')
            [void]$stampTokens.Add([string]$MarginOutsidePt)
        }

        $argLine = _QCRS-BuildProcessArgumentLine -Tokens $stampTokens.ToArray()
        if ([string]::IsNullOrWhiteSpace($argLine)) {
            return @{ applied = $false; reason = 'Review stamp command line is empty (internal argument build failed).' }
        }

        $p = Start-Process -FilePath $OverlayExe -ArgumentList $argLine -Wait -NoNewWindow -PassThru `
            -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -ErrorAction Stop
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
        $posHint = if ($null -ne $stampCfg.stampXPt -and $null -ne $stampCfg.stampYPt) {
            "at ($($stampCfg.stampXPt), $($stampCfg.stampYPt)) pt from page top-left"
        } else {
            "with margin $($stampCfg.marginOutsidePt) pt from page top-left"
        }
        & $Log "Applying peer review stamp ($posHint) on page 1: $PdfPath"
    }

    $stampParams = @{
        OverlayExe     = $exe
        PdfPath        = $PdfPath
        StampPath      = [string]$stampCfg.stampPath
        Originator     = [string]$RoleFields.designerEmail
        Checker        = [string]$RoleFields.reviewerEmail
        Backchecker    = [string]$RoleFields.checkerEmail
        StampHeightPt  = [double]$stampCfg.stampHeightPt
        MarginOutsidePt = [double]$stampCfg.marginOutsidePt
    }
    if ($null -ne $stampCfg.stampXPt -and $null -ne $stampCfg.stampYPt) {
        $stampParams['StampXPt'] = [double]$stampCfg.stampXPt
        $stampParams['StampYPt'] = [double]$stampCfg.stampYPt
    }
    $result = Invoke-QCReviewStamp @stampParams

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
