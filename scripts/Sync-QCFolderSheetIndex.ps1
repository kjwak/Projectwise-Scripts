<#
.SYNOPSIS
Scans a ProjectWise Sheets folder and refreshes sheet_index from live PW data.

.DESCRIPTION
Lists paired sheet PDF/DGN documents in the folder (same rules as status-set reconciliation),
reads workflow state and/or EM/QC attributes from ProjectWise, and upserts matching
sheet_index rows. Read-only against ProjectWise; writes only to SQL.

By default refreshes workflow states, role emails, and QC_Review_Type. Pass individual
-WorkflowStates, -EmailAttributes, or -QcReviewType switches to limit the scan.

Default is preview only. Pass -ConfirmWrites to apply database changes.

.EXAMPLE
.\scripts\Sync-QCFolderSheetIndex.ps1 -FolderPath 'Documents\Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
.\scripts\Sync-QCFolderSheetIndex.ps1 -FolderPath '...\Sheets\Seg_1' -ConfirmWrites
.\scripts\Sync-QCFolderSheetIndex.ps1 -FolderPath '...\Sheets\Seg_1' -ConfirmWrites -WorkflowStates
.\scripts\Sync-QCFolderSheetIndex.ps1 -FolderPath '...\Sheets\Seg_1' -ConfirmWrites -EmailAttributes -QcReviewType
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$FolderPath,

    [string]$AppSettingsPath = '',
    [switch]$ConfirmWrites,
    [switch]$DryRun,
    [switch]$WorkflowStates,
    [switch]$EmailAttributes,
    [switch]$QcReviewType,
    [switch]$OneLevelDeep,
    [switch]$SkipQcPdfLinks,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

. (Join-Path $PSScriptRoot 'Import-QCScriptModules.ps1') -RepoRoot $repoRoot -AdditionalModules @(
    'Core.Paths.psm1'
    'PW.Connection.psm1'
    'PW.Discovery.psm1'
    'QC.StatusSet.psm1'
)

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

if ($DryRun.IsPresent -and $ConfirmWrites.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmWrites (apply DB changes), not both.'
}
$doWrites = $ConfirmWrites.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmWrites.IsPresent) {
    Write-Host 'Preview only: pass -ConfirmWrites to update sheet_index, or -DryRun explicitly.' -ForegroundColor Yellow
    $DryRun = $true
}

$scanWorkflow = $WorkflowStates.IsPresent
$scanEmail = $EmailAttributes.IsPresent
$scanReviewType = $QcReviewType.IsPresent
if (-not ($scanWorkflow -or $scanEmail -or $scanReviewType)) {
    $scanWorkflow = $true
    $scanEmail = $true
    $scanReviewType = $true
}

$normRes = Normalize-QCDocumentsFolderPath -Path $FolderPath
if (-not $normRes.IsSuccess) { throw $normRes.Message }
$normFolder = [string]$normRes.Data.path

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config
if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true.'
}
Initialize-QCDatabaseSchema -Config $config | Out-Null

function _SIS-NormValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function _SIS-NormalizePathPattern {
    param([string]$Pattern)
    $norm = ($Pattern -replace '\\', '/').ToLowerInvariant()
    if ($norm -notmatch '[%_\[]') { return "%$norm%" }
    return $norm
}

function _SIS-ResolveWatchRoot {
    param(
        [hashtable]$Config,
        [string]$FolderPath
    )
    $roots = [System.Collections.Generic.List[string]]::new()
    try {
        $wl = $Config.projectWise.watchList
        if ($wl -and $wl.roots) {
            foreach ($r in @($wl.roots)) {
                $p = ''
                if ($r -is [hashtable] -and $r.ContainsKey('path')) { $p = [string]$r['path'] }
                elseif ($r -and $r.PSObject.Properties['path']) { $p = [string]$r.path }
                if (-not [string]::IsNullOrWhiteSpace($p)) { $roots.Add($p.Trim()) | Out-Null }
            }
        }
    } catch { }
    $folderNorm = ($FolderPath -replace '\\', '/').ToLowerInvariant()
    foreach ($root in @($roots | Sort-Object { $_.Length } -Descending)) {
        $rn = ($root -replace '\\', '/').ToLowerInvariant()
        if ($folderNorm.StartsWith($rn, [StringComparison]::Ordinal)) { return $root }
    }
    return ''
}

function _SIS-LoadDbRowsByGuid {
    param(
        [hashtable]$Config,
        [string]$FolderLike
    )
    $map = @{}
    $sql = @"
SELECT document_guid, document_name, designer_email, reviewer_email, checker_email, qc_review_type, pw_state_name
FROM sheet_index si
WHERE REPLACE(LOWER(LTRIM(RTRIM(ISNULL(si.folder_path, '')))), '\', '/') LIKE @folderLike
"@
    $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters @{ folderLike = $FolderLike }
    if (-not $res.IsSuccess) { throw $res.Message }
    foreach ($r in @($res.Data.table.Rows)) {
        $guid = [string]$r.document_guid
        if ([string]::IsNullOrWhiteSpace($guid)) { continue }
        $map[$guid.ToLowerInvariant()] = [pscustomobject]@{
            documentGuid   = $guid
            documentName   = if ($r.document_name -is [DBNull]) { '' } else { [string]$r.document_name }
            designerEmail  = if ($r.designer_email -is [DBNull]) { '' } else { [string]$r.designer_email }
            reviewerEmail  = if ($r.reviewer_email -is [DBNull]) { '' } else { [string]$r.reviewer_email }
            checkerEmail   = if ($r.checker_email -is [DBNull]) { '' } else { [string]$r.checker_email }
            qcReviewType   = if ($r.qc_review_type -is [DBNull]) { '' } else { [string]$r.qc_review_type }
            pwStateName    = if ($r.pw_state_name -is [DBNull]) { '' } else { [string]$r.pw_state_name }
        }
    }
    return $map
}

function _SIS-RowNeedsUpdate {
    param(
        [object]$PwRow,
        [object]$DbRow,
        [bool]$IncludeWorkflow,
        [bool]$IncludeEmail,
        [bool]$IncludeReviewType
    )
    $isNew = ($null -eq $DbRow)
    if ($IncludeWorkflow) {
        $pw = _SIS-NormValue $PwRow.pwStateName
        $db = if ($isNew) { '' } else { _SIS-NormValue $DbRow.pwStateName }
        if ($pw -ne $db) { return $true }
    }
    if ($IncludeEmail) {
        foreach ($pair in @(
            @{ pw = $PwRow.designerEmail; db = if ($isNew) { '' } else { $DbRow.designerEmail } }
            @{ pw = $PwRow.reviewerEmail; db = if ($isNew) { '' } else { $DbRow.reviewerEmail } }
            @{ pw = $PwRow.checkerEmail; db = if ($isNew) { '' } else { $DbRow.checkerEmail } }
        )) {
            if ((_SIS-NormValue $pair.pw) -ne (_SIS-NormValue $pair.db)) { return $true }
        }
    }
    if ($IncludeReviewType) {
        $pw = _SIS-NormValue $PwRow.qcReviewType
        $db = if ($isNew) { '' } else { _SIS-NormValue $DbRow.qcReviewType }
        if ($pw -ne $db) { return $true }
    }
    return $false
}

function _SIS-BuildBatchRow {
    param(
        [object]$PwRow,
        [bool]$IncludeWorkflow,
        [bool]$IncludeEmail,
        [bool]$IncludeReviewType
    )
    return @{
        documentGuid   = [string]$PwRow.documentGuid
        documentName   = [string]$PwRow.documentName
        folderPath     = [string]$PwRow.folderPath
        watchRoot      = [string]$PwRow.watchRoot
        sourceType     = [string]$PwRow.sourceType
        designerEmail  = if ($IncludeEmail) { [string]$PwRow.designerEmail } else { $null }
        reviewerEmail  = if ($IncludeEmail) { [string]$PwRow.reviewerEmail } else { $null }
        checkerEmail   = if ($IncludeEmail) { [string]$PwRow.checkerEmail } else { $null }
        qcReviewType   = if ($IncludeReviewType) { [string]$PwRow.qcReviewType } else { $null }
        qcAssignedTo   = $null
        qcStatus       = $null
        pwStateName    = if ($IncludeWorkflow) { [string]$PwRow.pwStateName } else { $null }
    }
}

$watchRoot = _SIS-ResolveWatchRoot -Config $config -FolderPath $normFolder
$folderLike = _SIS-NormalizePathPattern -Pattern $normFolder
$dbByGuid = _SIS-LoadDbRowsByGuid -Config $config -FolderLike $folderLike

$pwCfg = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $raw = $config.projectWise
    if ($raw -is [hashtable]) { $pwCfg = $raw }
    elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $pwCfg[$p.Name] = $p.Value } }
}
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { '' }
$ds = if ($pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }
if ([string]::IsNullOrWhiteSpace($credPath) -or [string]::IsNullOrWhiteSpace($ds)) {
    throw 'projectWise.datasourceName and projectWise.credentialPath are required in appsettings.'
}

$apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $normFolder
if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $normFolder }

$summary = [ordered]@{
    dryRun            = $DryRun.IsPresent
    folderPath        = $normFolder
    watchRoot         = $watchRoot
    scanWorkflow      = $scanWorkflow
    scanEmail         = $scanEmail
    scanReviewType    = $scanReviewType
    pairedCount       = 0
    pwRowsBuilt       = 0
    dbRowsMatched     = $dbByGuid.Count
    updatesPlanned    = 0
    insertsPlanned    = 0
    updatesApplied    = 0
    qcPdfLinksPlanned = 0
    qcPdfLinksApplied = 0
    samples           = @()
    errors            = @()
}

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw $connRes.Message }

try {
    if (-not (Get-Command -Name 'Get-StatusSetPWFolderState' -ErrorAction SilentlyContinue)) {
        throw 'Get-StatusSetPWFolderState is unavailable. Import QC.StatusSet.psm1.'
    }

    Write-Host '[ProjectWise] Scanning folder for sheet pairs...' -ForegroundColor Cyan
    $folderState = Get-StatusSetPWFolderState -FolderPath $apiPath -OneLevelDeep:$OneLevelDeep.IsPresent
    $summary.pairedCount = [int]$folderState.pairedCount

    $stateGuids = [System.Collections.Generic.List[string]]::new()
    foreach ($ps in @($folderState.pairedSheets)) {
        if ($ps.pdf -and $ps.pdf.documentGuid) { $stateGuids.Add([string]$ps.pdf.documentGuid) | Out-Null }
        if ($ps.dgn -and $ps.dgn.documentGuid) { $stateGuids.Add([string]$ps.dgn.documentGuid) | Out-Null }
    }
    $stateByGuid = @{}
    if ($scanWorkflow -and $stateGuids.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids @($stateGuids | Select-Object -Unique) } catch { }
    }

    $pwRows = @(Build-PWSheetIndexRowsForPairedSheets -Config $config `
        -FolderPath $normFolder `
        -WatchRoot $watchRoot `
        -PairedSheets @($folderState.pairedSheets) `
        -StateByGuid $stateByGuid)
    $summary.pwRowsBuilt = $pwRows.Count

    $batchRows = [System.Collections.Generic.List[object]]::new()
    foreach ($pwRow in $pwRows) {
        $guid = [string]$pwRow.documentGuid
        if ([string]::IsNullOrWhiteSpace($guid)) { continue }
        $dbRow = $null
        if ($dbByGuid.ContainsKey($guid.ToLowerInvariant())) {
            $dbRow = $dbByGuid[$guid.ToLowerInvariant()]
        }
        if (-not (_SIS-RowNeedsUpdate -PwRow $pwRow -DbRow $dbRow `
                -IncludeWorkflow $scanWorkflow -IncludeEmail $scanEmail -IncludeReviewType $scanReviewType)) {
            continue
        }
        if ($null -eq $dbRow) { $summary.insertsPlanned++ } else { $summary.updatesPlanned++ }
        if ($summary.samples.Count -lt 8) {
            $summary.samples += [pscustomobject]@{
                documentName = [string]$pwRow.documentName
                change       = if ($null -eq $dbRow) { 'insert' } else { 'update' }
                pwState      = if ($scanWorkflow) { [string]$pwRow.pwStateName } else { $null }
                dbState      = if ($scanWorkflow -and $dbRow) { [string]$dbRow.pwStateName } else { $null }
                pwReviewType = if ($scanReviewType) { [string]$pwRow.qcReviewType } else { $null }
                dbReviewType = if ($scanReviewType -and $dbRow) { [string]$dbRow.qcReviewType } else { $null }
            }
        }
        $batchRows.Add((_SIS-BuildBatchRow -PwRow $pwRow `
            -IncludeWorkflow $scanWorkflow -IncludeEmail $scanEmail -IncludeReviewType $scanReviewType)) | Out-Null
    }

    if (-not $SkipQcPdfLinks.IsPresent -and $folderState.qcPdfDocs) {
        foreach ($qc in @($folderState.qcPdfDocs)) {
            $summary.qcPdfLinksPlanned++
        }
    }

    Write-Host ("  Paired sheets: {0}; PW rows: {1}; planned updates: {2}; planned inserts: {3}" -f `
        $summary.pairedCount, $summary.pwRowsBuilt, $summary.updatesPlanned, $summary.insertsPlanned) -ForegroundColor Gray
    if ($summary.samples.Count -gt 0) {
        Write-Host '  Sample planned changes:' -ForegroundColor DarkGray
        $summary.samples | Format-Table -AutoSize
    }

    if ($doWrites -and $batchRows.Count -gt 0) {
        Write-Host '[Database] Upserting sheet_index rows...' -ForegroundColor Cyan
        if ($PSCmdlet.ShouldProcess($normFolder, 'Upsert sheet_index from PW scan')) {
            try {
                Write-QCSheetIndexBatch -Config $config -Rows @($batchRows)
                $summary.updatesApplied = $summary.updatesPlanned + $summary.insertsPlanned
            } catch {
                $summary.errors += $_.Exception.Message
                throw
            }
        }
    }

    if ($doWrites -and -not $SkipQcPdfLinks.IsPresent -and $folderState.qcPdfDocs) {
        foreach ($qc in @($folderState.qcPdfDocs)) {
            $qcGuid = [string]$qc.documentGuid
            $qcName = [string]$qc.name
            $qcStem = [string]$qc.stem
            if ([string]::IsNullOrWhiteSpace($qcGuid) -or [string]::IsNullOrWhiteSpace($qcStem)) { continue }
            $linked = $false
            foreach ($ps in @($folderState.pairedSheets)) {
                $stem = ''
                try { $stem = [string]$ps.stem } catch { }
                if ($stem -ne $qcStem) { continue }
                foreach ($member in @('pdf', 'dgn')) {
                    $src = $ps.$member
                    if (-not $src -or -not $src.documentGuid) { continue }
                    $srcGuid = [string]$src.documentGuid
                    if ($PSCmdlet.ShouldProcess("$normFolder\$($src.name)", "Link qc pdf $qcName")) {
                        try {
                            Update-QCSheetQcPdf -Config $config -SourceDocumentGuid $srcGuid `
                                -QcPdfGuid $qcGuid -QcPdfName $qcName | Out-Null
                            $linked = $true
                        } catch {
                            $summary.errors += "QC PDF link failed for $srcGuid : $($_.Exception.Message)"
                        }
                    }
                }
            }
            if ($linked) { $summary.qcPdfLinksApplied++ }
        }
    } elseif (-not $doWrites) {
        Write-Host 'Dry run: no sheet_index rows written. Pass -ConfirmWrites to apply.' -ForegroundColor Yellow
    }
} finally {
    Disconnect-PW | Out-Null
}

if ($Pretty) { $summary | ConvertTo-Json -Depth 6 }
