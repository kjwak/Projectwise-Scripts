# QC.ReviewStamp.psm1
# Applies editable QC review stamps via qc_overlay_prepend.exe (--apply-review-stamp).
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
    _QCRS-AppendCliFlagValue -TokenList $TokenList -Flag $Flag -Value $Value
}

function _QCRS-AppendCliFlagValue {
    param(
        [System.Collections.Generic.List[object]]$TokenList,
        [Parameter(Mandatory)][string]$Flag,
        [Parameter(Mandatory)][string]$Value
    )
    # --flag=value avoids Windows/argparse treating -400 as a switch.
    [void]$TokenList.Add("${Flag}=$Value")
}

function _QCRS-TestOverlaySupportsStampPositionPt {
    param([Parameter(Mandatory)][string]$OverlayExe)
    $dir = Split-Path -Parent $OverlayExe
    if (_QCRS-IsBlank $dir) { return $false }
    foreach ($rel in @('_internal\qc_review_stamp.py', 'qc_review_stamp.py')) {
        $py = Join-Path $dir $rel
        if (Test-Path -LiteralPath $py) {
            return [bool](Select-String -LiteralPath $py -Pattern '--stamp-x-pt' -Quiet)
        }
    }
    return $false
}

function _QCRS-NeedsShellCliQuoting {
    param([string]$Token)
    if ([string]::IsNullOrEmpty($Token)) { return $false }
    if ($Token -match '[\s"]') { return $true }
    if ($Token -match '^--') { return $false }
    if ($Token -match '^-') { return $true }
    return $false
}

function _QCRS-BuildProcessArgumentLine {
    param([object[]]$Tokens = @())
    $parts = @()
    foreach ($item in @($Tokens)) {
        if ($null -eq $item) { continue }
        $t = [string]$item
        if ($t.Length -eq 0) { continue }
        if (_QCRS-NeedsShellCliQuoting $t) { $parts += ('"' + ($t -replace '"', '\"') + '"') }
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

function _QCRS-ResolveStampPath {
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [string]$ConfiguredPath,
        [Parameter(Mandatory)][string]$DefaultRelativePath
    )
    $sp = if ($ConfiguredPath) { [string]$ConfiguredPath } else { $DefaultRelativePath }
    if (-not [System.IO.Path]::IsPathRooted($sp)) { $sp = Join-Path $RepoRoot $sp }
    return $sp
}

function _QCRS-NewReviewStampProfile {
    param(
        [Parameter(Mandatory)][string]$ProfileKey,
        [Parameter(Mandatory)][string]$ReviewTypeLabel,
        [Parameter(Mandatory)][string]$StampPath,
        [Parameter(Mandatory)][string]$LogLabel
    )
    return @{
        profileKey = $ProfileKey
        reviewType = $ReviewTypeLabel
        stampPath  = $StampPath
        logLabel   = $LogLabel
    }
}

function Get-QCReviewStampSettings {
    <#
    .SYNOPSIS
    Shared review-stamp settings plus one profile per configured stamp template (peer review, independent check).
    #>
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
    $icType = 'Independent Check'
    $productionType = 'Production QC'
    $wf = _QCRS-ToHashtable $Config.qcWorkflow
    if ($wf) {
        $rt = _QCRS-ToHashtable $wf.reviewTypes
        if ($rt) {
            if ($rt.peerReview) { $peerType = [string]$rt.peerReview }
            if ($rt.independentCheck) { $icType = [string]$rt.independentCheck }
            if ($rt.productionQc) { $productionType = [string]$rt.productionQc }
        }
    }

    $populateTextFields = $false
    if ($rs.ContainsKey('populateTextFields')) {
        try { $populateTextFields = [bool]$rs.populateTextFields } catch { $populateTextFields = $false }
    }

    $stampHeight = 200.0
    $marginOutside = 12.0
    $stampX = $null
    $stampY = $null
    if ($rs.ContainsKey('stampHeightPt')) { try { $stampHeight = [double]$rs.stampHeightPt } catch { } }
    if ($rs.ContainsKey('marginOutsidePt')) { try { $marginOutside = [double]$rs.marginOutsidePt } catch { } }
    $pos = _QCRS-ToHashtable $rs.stampPositionPt
    if ($pos) {
        if ($pos.ContainsKey('x')) { try { $stampX = [double]$pos.x } catch { } }
        if ($pos.ContainsKey('y')) { try { $stampY = [double]$pos.y } catch { } }
    }

    $profiles = [System.Collections.Generic.List[object]]::new()

    $pr = _QCRS-ToHashtable $rs.peerReview
    $peerPath = _QCRS-ResolveStampPath -RepoRoot $RepoRoot -ConfiguredPath $(if ($pr -and $pr.stampPath) { [string]$pr.stampPath } else { '' }) `
        -DefaultRelativePath 'stamps\Peer_Review_Stamp.pdf'
    if ($pr -and $pr.reviewType) { $peerType = [string]$pr.reviewType }
    if (Test-Path -LiteralPath $peerPath) {
        $profiles.Add((_QCRS-NewReviewStampProfile -ProfileKey 'peerReview' -ReviewTypeLabel $peerType -StampPath $peerPath -LogLabel 'peer review')) | Out-Null
    }

    $ic = _QCRS-ToHashtable $rs.independentCheck
    $icPath = _QCRS-ResolveStampPath -RepoRoot $RepoRoot -ConfiguredPath $(if ($ic -and $ic.stampPath) { [string]$ic.stampPath } else { '' }) `
        -DefaultRelativePath 'stamps\IC_Stamp.pdf'
    if ($ic -and $ic.reviewType) { $icType = [string]$ic.reviewType }
    if (Test-Path -LiteralPath $icPath) {
        $profiles.Add((_QCRS-NewReviewStampProfile -ProfileKey 'independentCheck' -ReviewTypeLabel $icType -StampPath $icPath -LogLabel 'independent check')) | Out-Null
    }

    if ($profiles.Count -eq 0) { return $null }

    $overlayExe = Resolve-QCReviewStampOverlayExe -PreferredPath ([string]$qc.overlayExePath) -RepoRoot $RepoRoot
    if (-not $overlayExe) { return $null }

    $peerProfile = @($profiles | Where-Object { $_.profileKey -eq 'peerReview' } | Select-Object -First 1)

    return @{
        profiles             = @($profiles)
        populateTextFields   = $populateTextFields
        productionReviewType = $productionType
        stampHeightPt        = $stampHeight
        marginOutsidePt      = $marginOutside
        stampXPt             = $stampX
        stampYPt             = $stampY
        overlayExe           = $overlayExe
        # Backward compatibility for callers that only checked peer review.
        reviewType           = if ($peerProfile) { [string]$peerProfile.reviewType } else { [string]$profiles[0].reviewType }
        stampPath            = if ($peerProfile) { [string]$peerProfile.stampPath } else { [string]$profiles[0].stampPath }
    }
}

function _QCRS-FindReviewStampProfile {
    param(
        [Parameter(Mandatory)][hashtable]$StampSettings,
        [string]$ReviewType
    )
    if (_QCRS-IsBlank $ReviewType) { return $null }
    $rt = $ReviewType.Trim()
    foreach ($p in @($StampSettings.profiles)) {
        if ($p.reviewType -eq $rt) { return $p }
    }
    return $null
}

function _QCRS-ResolveStampRoleValues {
    param(
        [Parameter(Mandatory)][hashtable]$RoleFields,
        [Parameter(Mandatory)][string]$ProfileKey
    )
    $designer = [string]$RoleFields.designerEmail
    $reviewer = [string]$RoleFields.reviewerEmail
    $checker = [string]$RoleFields.checkerEmail
    if ($ProfileKey -eq 'independentCheck') {
        return @{
            Originator  = $designer
            Checker     = $checker
            Backchecker = $reviewer
        }
    }
    return @{
        Originator  = $designer
        Checker     = $reviewer
        Backchecker = $checker
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
        [Nullable[double]]$StampYPt = $null,
        [bool]$PopulateTextFields = $false
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

    if ($PopulateTextFields -and (_QCRS-IsBlank $OriginatorDate)) {
        $OriginatorDate = (Get-Date).ToString('MM/dd/yyyy')
    }

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    try {
        $stampTokens = [System.Collections.Generic.List[object]]::new()
        [void]$stampTokens.Add('--apply-review-stamp')
        [void]$stampTokens.Add([string]$PdfPath)
        [void]$stampTokens.Add([string]$StampPath)
        if (-not $PopulateTextFields) {
            [void]$stampTokens.Add('--no-populate-text-fields')
        } else {
            _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--originator' -Value $Originator
            _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--checker' -Value $Checker
            _QCRS-AppendOptionalCliFlags -TokenList $stampTokens -Flag '--backchecker' -Value $Backchecker
            _QCRS-AppendCliFlagValue -TokenList $stampTokens -Flag '--originator-date' -Value ([string]$OriginatorDate)
        }
        _QCRS-AppendCliFlagValue -TokenList $stampTokens -Flag '--stamp-height-pt' -Value ([string]$StampHeightPt)
        $supportsXY = _QCRS-TestOverlaySupportsStampPositionPt -OverlayExe $OverlayExe
        if ($null -ne $StampXPt -and $null -ne $StampYPt -and $supportsXY) {
            _QCRS-AppendCliFlagValue -TokenList $stampTokens -Flag '--stamp-x-pt' -Value ([string]$StampXPt)
            _QCRS-AppendCliFlagValue -TokenList $stampTokens -Flag '--stamp-y-pt' -Value ([string]$StampYPt)
        } else {
            if ($null -ne $StampXPt -and $null -ne $StampYPt -and -not $supportsXY) {
                return @{ applied = $false; reason = 'Overlay exe is missing --stamp-x-pt/--stamp-y-pt support; copy overlay\qc_review_stamp.py to dist\qc_overlay_prepend\_internal\ or rebuild dist\qc_overlay_prepend.' }
            }
            _QCRS-AppendCliFlagValue -TokenList $stampTokens -Flag '--margin-outside-pt' -Value ([string]$MarginOutsidePt)
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

function Invoke-QCReviewStampForReviewType {
    <#
    .SYNOPSIS
    Applies the configured review stamp when QC_Review_Type matches a stamp profile (Peer Review or Independent Check).
    #>
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
    $productionType = ''
    if ($stampCfg.ContainsKey('productionReviewType') -and $stampCfg.productionReviewType) {
        $productionType = [string]$stampCfg.productionReviewType
    }

    if (-not (_QCRS-IsBlank $productionType) -and -not (_QCRS-IsBlank $reviewType) -and $reviewType.Trim() -eq $productionType.Trim()) {
        return @{ applied = $false; skipped = $true; reason = "QC_Review_Type is '$reviewType' (Production QC); review stamp not applicable." }
    }

    $profile = _QCRS-FindReviewStampProfile -StampSettings $stampCfg -ReviewType $reviewType
    if (-not $profile) {
        $expected = (@($stampCfg.profiles | ForEach-Object { [string]$_.reviewType }) -join "', '")
        return @{ applied = $false; skipped = $true; reason = "QC_Review_Type is '$reviewType' (need one of: '$expected')" }
    }

    $exe = if (-not (_QCRS-IsBlank $OverlayExe)) { $OverlayExe } else { [string]$stampCfg.overlayExe }
    if ($Log) {
        $posHint = if ($null -ne $stampCfg.stampXPt -and $null -ne $stampCfg.stampYPt) {
            "at ($($stampCfg.stampXPt), $($stampCfg.stampYPt)) pt from page top-left"
        } else {
            "with margin $($stampCfg.marginOutsidePt) pt from page top-left"
        }
        & $Log "Applying $($profile.logLabel) stamp ($posHint) on page 1: $PdfPath"
    }

    $populateTextFields = $false
    if ($stampCfg.ContainsKey('populateTextFields')) {
        try { $populateTextFields = [bool]$stampCfg.populateTextFields } catch { }
    }
    $roleValues = _QCRS-ResolveStampRoleValues -RoleFields $RoleFields -ProfileKey ([string]$profile.profileKey)
    $stampParams = @{
        OverlayExe           = $exe
        PdfPath              = $PdfPath
        StampPath            = [string]$profile.stampPath
        StampHeightPt        = [double]$stampCfg.stampHeightPt
        MarginOutsidePt      = [double]$stampCfg.marginOutsidePt
        PopulateTextFields   = $populateTextFields
    }
    if ($populateTextFields) {
        $stampParams['Originator'] = [string]$roleValues.Originator
        $stampParams['Checker'] = [string]$roleValues.Checker
        $stampParams['Backchecker'] = [string]$roleValues.Backchecker
    }
    if ($null -ne $stampCfg.stampXPt -and $null -ne $stampCfg.stampYPt) {
        $stampParams['StampXPt'] = [double]$stampCfg.stampXPt
        $stampParams['StampYPt'] = [double]$stampCfg.stampYPt
    }
    $result = Invoke-QCReviewStamp @stampParams

    $result['reviewType'] = $reviewType
    $result['profileKey'] = [string]$profile.profileKey
    if ($result.applied -and $Log) {
        $fieldNote = if ($populateTextFields) { 'with role/date fields' } else { 'blank template (no field population)' }
        & $Log "$($profile.logLabel) stamp applied ($fieldNote)."
    }
    return $result
}

function Invoke-QCReviewStampIfPeerReview {
    param(
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$RoleFields,
        [string]$OverlayExe = '',
        [scriptblock]$Log = $null
    )

    return Invoke-QCReviewStampForReviewType -PdfPath $PdfPath -Config $Config -RoleFields $RoleFields -OverlayExe $OverlayExe -Log $Log
}

Export-ModuleMember -Function @(
    'Join-QCProcessArgumentList',
    'Resolve-QCReviewStampOverlayExe',
    'Get-QCReviewStampSettings',
    'Invoke-QCReviewStamp',
    'Invoke-QCReviewStampForReviewType',
    'Invoke-QCReviewStampIfPeerReview'
)
