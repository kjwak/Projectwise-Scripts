<#
.SYNOPSIS
Reconciles designer/reviewer emails and workflow state across DGN, sheet PDF, and QC PDF.

.DESCRIPTION
Reads all rows from sheet_index, groups by folder + sheet stem, then applies:

  - Email source of truth: DGN (EM_Designer_Email / EM_Reviewer_Email)
    -> sheet PDF and *-qc.pdf are updated to match DGN when they differ.

  - State source of truth: *-qc.pdf workflow state
    -> DGN and sheet PDF are updated to match QC PDF when they differ.

ProjectWise documents are located by document_guid from the index (and qc_pdf_guid
when the QC PDF is linked but not indexed as its own row). sheet_index is updated
after successful PW reconciliation.

.EXAMPLE
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -DryRun
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -ConfirmWrites
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -ConfirmWrites -FolderPathFilter 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AppSettingsPath = '',
    [switch]$ConfirmWrites,
    [switch]$DryRun,
    [string]$FolderPathFilter = '',
    [int]$Limit = 0,
    [switch]$Pretty
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Connection.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\PW.Discovery.psm1') -Force

foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function _RSO-Norm([AllowNull()][string]$Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function _RSO-SheetStem([string]$DocumentName) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ($stem -match '(?i)-qc$') { $stem = $stem -replace '(?i)-qc$', '' }
    return $stem.ToLowerInvariant()
}

function _RSO-IsQcPdfName([string]$DocumentName) {
    return [bool]([string]$DocumentName -match '(?i)-qc\.pdf$')
}

function _RSO-GetEmailColumns {
    param([hashtable]$Config)
    $designer = 'EM_Designer_Email'
    $reviewer = 'EM_Reviewer_Email'
    try {
        $pw = $Config.projectWise
        if ($pw -and $pw.environmentEmailAttributes) {
            $ea = $pw.environmentEmailAttributes
            if ($ea.default) {
                if ($ea.default.designerEmailColumn) { $designer = [string]$ea.default.designerEmailColumn }
                if ($ea.default.reviewerEmailColumn) { $reviewer = [string]$ea.default.reviewerEmailColumn }
            }
        }
    } catch { }
    return @{ designer = $designer; reviewer = $reviewer }
}

function _RSO-BuildFolderEmailMap {
    param(
        [string]$FolderPath,
        [string]$DesignerColumn,
        [string]$ReviewerColumn
    )

    $map = @{}
    $cmd = Get-Command -Name 'Get-PWDocumentsBySearchWithReturnColumns' -ErrorAction SilentlyContinue
    if (-not $cmd) { return $map }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }

    try {
        $params = @{
            FolderPath     = $apiPath
            JustThisFolder = $true
            ErrorAction    = 'Stop'
        }
        if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
            $params['ColumnsToReturn'] = @($DesignerColumn, $ReviewerColumn)
        } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
            $params['ReturnColumns'] = @($DesignerColumn, $ReviewerColumn)
        }
        $rows = @(& $cmd @params)
        foreach ($row in $rows) {
            $name = $null
            try { $name = [string]$row.Name } catch { }
            if (-not $name) { try { $name = [string]$row.DocumentName } catch { } }
            if (-not $name) { continue }
            $guid = ''
            try { $guid = [string]$row.DocumentGUID } catch { }
            $attrs = Get-PWDocumentAttributeMap -DocRow $row
            $designer = if ($attrs.ContainsKey($DesignerColumn)) { [string]$attrs[$DesignerColumn] } else { '' }
            $reviewer = if ($attrs.ContainsKey($ReviewerColumn)) { [string]$attrs[$ReviewerColumn] } else { '' }
            $key = $name.ToLowerInvariant()
            $map[$key] = @{
                documentGuid  = $guid
                designerEmail = $designer.Trim()
                reviewerEmail = $reviewer.Trim()
                document      = $row
            }
        }
    } catch { }
    return $map
}

function _RSO-SetPwDocumentState {
    param(
        [object]$Document,
        [string]$StateName
    )
    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd -or -not $Document) { throw 'Set-PWDocumentState or document unavailable.' }
    $args = @{}
    $docParam = if ($cmd.Parameters.ContainsKey('InputDocuments')) { 'InputDocuments' }
        elseif ($cmd.Parameters.ContainsKey('InputDocument')) { 'InputDocument' }
        elseif ($cmd.Parameters.ContainsKey('Document')) { 'Document' }
        else { $null }
    $stateParam = if ($cmd.Parameters.ContainsKey('StateName')) { 'StateName' }
        elseif ($cmd.Parameters.ContainsKey('State')) { 'State' }
        else { $null }
    if ($docParam) { $args[$docParam] = @($Document) }
    if ($stateParam) { $args[$stateParam] = $StateName }
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) { $args['ReturnBoolean'] = $true }
    if ($docParam -and $stateParam) { & $cmd @args -ErrorAction Stop | Out-Null }
    elseif ($stateParam) { & $cmd $Document @args -ErrorAction Stop | Out-Null }
    else { & $cmd $Document $StateName -ErrorAction Stop | Out-Null }
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true for sheet_index reconciliation.'
}

$doWrites = $ConfirmWrites.IsPresent -and -not $DryRun.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmWrites.IsPresent) {
    Write-Host 'Refusing PW/DB writes: pass -ConfirmWrites to apply changes, or -DryRun to preview only.' -ForegroundColor Yellow
    $DryRun = $true
}

Initialize-QCDatabaseSchema -Config $config | Out-Null

$emailCols = _RSO-GetEmailColumns -Config $config
$designerCol = $emailCols.designer
$reviewerCol = $emailCols.reviewer

$sql = @"
SELECT document_guid, document_name, folder_path, extension, source_type,
       qc_pdf_guid, qc_pdf_name, designer_email, reviewer_email, pw_state_name
FROM sheet_index
"@
$sqlParams = @{}
if (-not [string]::IsNullOrWhiteSpace($FolderPathFilter)) {
    $sql += ' WHERE folder_path LIKE @folderFilter'
    $sqlParams['folderFilter'] = ('%' + $FolderPathFilter.Trim().Trim('\') + '%')
}
$sql += ' ORDER BY folder_path, document_name'

$qRes = Invoke-QCDatabaseQuery -Config $config -Sql $sql -Parameters $sqlParams
if (-not $qRes.IsSuccess) { throw $qRes.Message }

$indexRows = @()
foreach ($r in @($qRes.Data.table.Rows)) {
    $indexRows += [pscustomobject]@{
        documentGuid   = [string]$r.document_guid
        documentName   = [string]$r.document_name
        folderPath     = [string]$r.folder_path
        extension      = if ($r.extension -is [DBNull]) { '' } else { [string]$r.extension }
        sourceType     = if ($r.source_type -is [DBNull]) { '' } else { [string]$r.source_type }
        qcPdfGuid      = if ($r.qc_pdf_guid -is [DBNull]) { '' } else { [string]$r.qc_pdf_guid }
        qcPdfName      = if ($r.qc_pdf_name -is [DBNull]) { '' } else { [string]$r.qc_pdf_name }
        designerEmail  = if ($r.designer_email -is [DBNull]) { '' } else { [string]$r.designer_email }
        reviewerEmail  = if ($r.reviewer_email -is [DBNull]) { '' } else { [string]$r.reviewer_email }
        pwStateName    = if ($r.pw_state_name -is [DBNull]) { '' } else { [string]$r.pw_state_name }
    }
}

Write-Host "Loaded $($indexRows.Count) sheet_index rows." -ForegroundColor Cyan

$groups = @{}
foreach ($row in $indexRows) {
    if ([string]::IsNullOrWhiteSpace($row.folderPath) -or [string]::IsNullOrWhiteSpace($row.documentName)) { continue }
    $stem = _RSO-SheetStem -DocumentName $row.documentName
    if (-not $stem) { continue }
    $key = ($row.folderPath.ToLowerInvariant() + '|' + $stem)
    if (-not $groups.ContainsKey($key)) {
        $groups[$key] = @{
            folderPath = $row.folderPath
            stem       = $stem
            dgn        = $null
            pdf        = $null
            qcIndexed  = $null
            qcPdfGuid  = ''
            qcPdfName  = ''
            indexGuids = [System.Collections.Generic.List[string]]::new()
        }
    }
    $g = $groups[$key]
    if ($row.documentGuid) { $g.indexGuids.Add($row.documentGuid) | Out-Null }

    if ($row.sourceType -eq 'dgn' -or $row.extension -eq '.dgn') {
        $g.dgn = $row
    } elseif (_RSO-IsQcPdfName -DocumentName $row.documentName) {
        $g.qcIndexed = $row
        if (-not $g.qcPdfGuid) { $g.qcPdfGuid = $row.documentGuid }
        if (-not $g.qcPdfName) { $g.qcPdfName = $row.documentName }
    } elseif ($row.sourceType -eq 'pdf' -or $row.extension -eq '.pdf') {
        $g.pdf = $row
    }

    if ($row.qcPdfGuid -and -not $g.qcPdfGuid) { $g.qcPdfGuid = $row.qcPdfGuid }
    if ($row.qcPdfName -and -not $g.qcPdfName) { $g.qcPdfName = $row.qcPdfName }
}

$groupList = @($groups.Values | Sort-Object folderPath, stem)
if ($Limit -gt 0) { $groupList = @($groupList | Select-Object -First $Limit) }

Write-Host "Grouped into $($groupList.Count) sheet sets (folder + stem)." -ForegroundColor Cyan

$pwCfg = @{}
if ($config.ContainsKey('projectWise') -and $config.projectWise) {
    $raw = $config.projectWise
    if ($raw -is [hashtable]) { $pwCfg = $raw } elseif ($raw.PSObject) { foreach ($p in $raw.PSObject.Properties) { $pwCfg[$p.Name] = $p.Value } }
}
$credPath = if ($pwCfg.credentialPath) { [string]$pwCfg.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
$ds = if ($pwCfg.datasourceName) { [string]$pwCfg.datasourceName } else { '' }

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { throw $credRes.Message }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
if (-not $connRes.IsSuccess) { throw $connRes.Message }
Write-Host "Connected to ProjectWise: $ds" -ForegroundColor Green

$summary = @{
    groupsTotal       = $groupList.Count
    groupsProcessed   = 0
    groupsSkipped     = 0
    emailUpdates      = 0
    stateUpdates      = 0
    dbUpdates         = 0
    errors            = @()
    dryRun            = [bool]$DryRun.IsPresent
    details           = @()
}

try {
    $folderEmailCache = @{}
    $docByGuid = @{}

    $allGuids = [System.Collections.Generic.List[string]]::new()
    foreach ($g in $groupList) {
        foreach ($part in @($g.dgn, $g.pdf, $g.qcIndexed)) {
            if ($part -and $part.documentGuid) { $allGuids.Add($part.documentGuid) | Out-Null }
        }
        if ($g.qcPdfGuid) { $allGuids.Add($g.qcPdfGuid) | Out-Null }
    }
    $uniqueGuids = @($allGuids | Select-Object -Unique)
    if ($uniqueGuids.Count -gt 0) {
        $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
        if ($guidCmd) {
            $chunkSize = 200
            for ($i = 0; $i -lt $uniqueGuids.Count; $i += $chunkSize) {
                $chunk = @($uniqueGuids[$i..[Math]::Min($i + $chunkSize - 1, $uniqueGuids.Count - 1)])
                try {
                    foreach ($doc in @(& $guidCmd -DocumentGUIDs $chunk -ErrorAction Stop)) {
                        $dg = ''
                        try { $dg = [string]$doc.DocumentGUID } catch { }
                        if ($dg) { $docByGuid[$dg.ToLowerInvariant()] = $doc }
                    }
                } catch {
                    $summary.errors += "GUID batch read failed: $($_.Exception.Message)"
                }
            }
        }
    }
    $stateByGuid = @{}
    if ($uniqueGuids.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $uniqueGuids } catch { }
    }

    foreach ($g in $groupList) {
        $summary.groupsProcessed++

        if (-not $g.dgn) {
            $summary.groupsSkipped++
            $summary.details += [ordered]@{
                folder = $g.folderPath; stem = $g.stem; action = 'skip'; reason = 'no DGN row in sheet_index'
            }
            continue
        }

        $folderKey = $g.folderPath.ToLowerInvariant()
        if (-not $folderEmailCache.ContainsKey($folderKey)) {
            $folderEmailCache[$folderKey] = _RSO-BuildFolderEmailMap -FolderPath $g.folderPath `
                -DesignerColumn $designerCol -ReviewerColumn $reviewerCol
        }
        $emailMap = $folderEmailCache[$folderKey]

        $dgnNameKey = $g.dgn.documentName.ToLowerInvariant()
        $dgnEmails = if ($emailMap.ContainsKey($dgnNameKey)) { $emailMap[$dgnNameKey] } else { $null }
        $canonicalDesigner = if ($dgnEmails) { [string]$dgnEmails.designerEmail } else { $g.dgn.designerEmail }
        $canonicalReviewer = if ($dgnEmails) { [string]$dgnEmails.reviewerEmail } else { $g.dgn.reviewerEmail }

        $qcGuid = ''
        if ($g.qcIndexed) { $qcGuid = $g.qcIndexed.documentGuid }
        elseif ($g.qcPdfGuid) { $qcGuid = $g.qcPdfGuid }
        $canonicalState = ''
        if ($qcGuid) {
            $qk = $qcGuid.ToLowerInvariant()
            if ($stateByGuid.ContainsKey($qk)) { $canonicalState = [string]$stateByGuid[$qk] }
        }

        $detail = [ordered]@{
            folder            = $g.folderPath
            stem              = $g.stem
            dgn               = $g.dgn.documentName
            pdf               = if ($g.pdf) { $g.pdf.documentName } else { $null }
            qcPdf             = if ($g.qcPdfName) { $g.qcPdfName } elseif ($g.qcIndexed) { $g.qcIndexed.documentName } else { $null }
            canonicalDesigner = $canonicalDesigner
            canonicalReviewer = $canonicalReviewer
            canonicalState    = $canonicalState
            emailChanges      = @()
            stateChanges      = @()
            dbRowsUpdated     = 0
        }

        $emailTargets = @()
        if ($g.pdf) {
            $emailTargets += @{ role = 'pdf'; row = $g.pdf; nameKey = $g.pdf.documentName.ToLowerInvariant() }
        }
        if ($g.qcIndexed) {
            $emailTargets += @{ role = 'qc'; row = $g.qcIndexed; nameKey = $g.qcIndexed.documentName.ToLowerInvariant() }
        } elseif ($g.qcPdfGuid -and $g.qcPdfName) {
            $emailTargets += @{ role = 'qc'; row = $null; nameKey = $g.qcPdfName.ToLowerInvariant(); guid = $g.qcPdfGuid; name = $g.qcPdfName }
        }

        foreach ($target in $emailTargets) {
            $currentDesigner = ''
            $currentReviewer = ''
            if ($target.row) {
                $nk = $target.nameKey
                if ($emailMap.ContainsKey($nk)) {
                    $currentDesigner = [string]$emailMap[$nk].designerEmail
                    $currentReviewer = [string]$emailMap[$nk].reviewerEmail
                } else {
                    $currentDesigner = [string]$target.row.designerEmail
                    $currentReviewer = [string]$target.row.reviewerEmail
                }
            } elseif ($target.nameKey -and $emailMap.ContainsKey($target.nameKey)) {
                $currentDesigner = [string]$emailMap[$target.nameKey].designerEmail
                $currentReviewer = [string]$emailMap[$target.nameKey].reviewerEmail
            }

            $needsDesigner = (_RSO-Norm $currentDesigner) -ne (_RSO-Norm $canonicalDesigner)
            $needsReviewer = (_RSO-Norm $currentReviewer) -ne (_RSO-Norm $canonicalReviewer)
            if (-not $needsDesigner -and -not $needsReviewer) { continue }

            $docObj = $null
            if ($target.row) {
                $dg = $target.row.documentGuid.ToLowerInvariant()
                if ($docByGuid.ContainsKey($dg)) { $docObj = $docByGuid[$dg] }
                elseif ($emailMap.ContainsKey($target.nameKey)) { $docObj = $emailMap[$target.nameKey].document }
            } elseif ($target.guid) {
                $dg = [string]$target.guid.ToLowerInvariant()
                if ($docByGuid.ContainsKey($dg)) { $docObj = $docByGuid[$dg] }
                elseif ($emailMap.ContainsKey($target.nameKey)) { $docObj = $emailMap[$target.nameKey].document }
            }

            $change = [ordered]@{
                role     = $target.role
                document = if ($target.row) { $target.row.documentName } else { $target.name }
                from     = @{ designer = $currentDesigner; reviewer = $currentReviewer }
                to       = @{ designer = $canonicalDesigner; reviewer = $canonicalReviewer }
                applied  = $false
            }

            if ($docObj -and $doWrites) {
                $toWrite = @{}
                if ($needsDesigner) { $toWrite[$designerCol] = $canonicalDesigner }
                if ($needsReviewer) { $toWrite[$reviewerCol] = $canonicalReviewer }
                $targetPath = "$($g.folderPath)\$($change.document)"
                if ($PSCmdlet.ShouldProcess($targetPath, 'Sync email attributes from DGN')) {
                    try {
                        [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $docObj -Attributes $toWrite)
                        $change.applied = $true
                        $summary.emailUpdates++
                    } catch {
                        $summary.errors += "Email update failed for $($change.document): $($_.Exception.Message)"
                    }
                }
            } elseif ($docObj -and $DryRun.IsPresent) {
                $change.applied = $false
                $change.planned = $true
                $summary.emailUpdates++
            } elseif (-not $docObj) {
                $change.skipped = $true
                $change.reason = 'PW document object not found by GUID/name'
            }

            $detail.emailChanges += $change
        }

        if ($canonicalState) {
            $stateTargets = @()
            if ($g.dgn) { $stateTargets += @{ role = 'dgn'; row = $g.dgn } }
            if ($g.pdf) { $stateTargets += @{ role = 'pdf'; row = $g.pdf } }

            foreach ($target in $stateTargets) {
                $dg = $target.row.documentGuid.ToLowerInvariant()
                $currentState = if ($stateByGuid.ContainsKey($dg)) { [string]$stateByGuid[$dg] } else { [string]$target.row.pwStateName }
                if ((_RSO-Norm $currentState) -eq (_RSO-Norm $canonicalState)) { continue }

                $docObj = if ($docByGuid.ContainsKey($dg)) { $docByGuid[$dg] } else { $null }
                $change = [ordered]@{
                    role     = $target.role
                    document = $target.row.documentName
                    from     = $currentState
                    to       = $canonicalState
                    applied  = $false
                }

                if ($docObj -and $doWrites) {
                    $targetPath = "$($g.folderPath)\$($target.row.documentName)"
                    if ($PSCmdlet.ShouldProcess($targetPath, "Set workflow state to $canonicalState")) {
                        try {
                            _RSO-SetPwDocumentState -Document $docObj -StateName $canonicalState
                            $change.applied = $true
                            $summary.stateUpdates++
                        } catch {
                            $summary.errors += "State update failed for $($target.row.documentName): $($_.Exception.Message)"
                        }
                    }
                } elseif ($docObj -and $DryRun.IsPresent) {
                    $change.planned = $true
                    $summary.stateUpdates++
                } elseif (-not $docObj) {
                    $change.skipped = $true
                    $change.reason = 'PW document object not found by GUID'
                }

                $detail.stateChanges += $change
            }
        }

        $dbTargets = @($g.dgn)
        if ($g.pdf) { $dbTargets += $g.pdf }
        if ($g.qcIndexed) { $dbTargets += $g.qcIndexed }

        foreach ($dbRow in $dbTargets) {
            $needsDbEmail = (_RSO-Norm $dbRow.designerEmail) -ne (_RSO-Norm $canonicalDesigner) `
                -or (_RSO-Norm $dbRow.reviewerEmail) -ne (_RSO-Norm $canonicalReviewer)
            $needsDbState = $canonicalState -and ((_RSO-Norm $dbRow.pwStateName) -ne (_RSO-Norm $canonicalState))
            if (-not $needsDbEmail -and -not $needsDbState) { continue }

            if ($doWrites) {
                Write-QCSheetIndex -Config $config `
                    -DocumentGuid $dbRow.documentGuid `
                    -DocumentName $dbRow.documentName `
                    -FolderPath $dbRow.folderPath `
                    -Extension $dbRow.extension `
                    -SourceType $dbRow.sourceType `
                    -DesignerEmail $canonicalDesigner `
                    -ReviewerEmail $canonicalReviewer `
                    -PwStateName $(if ($canonicalState) { $canonicalState } else { $dbRow.pwStateName }) `
                    -SetOwnershipFromProjectWise
                $detail.dbRowsUpdated++
                $summary.dbUpdates++
            } elseif ($DryRun.IsPresent) {
                $detail.dbRowsUpdated++
                $summary.dbUpdates++
            }
        }

        if ($detail.emailChanges.Count -gt 0 -or $detail.stateChanges.Count -gt 0 -or $detail.dbRowsUpdated -gt 0) {
            $summary.details += $detail
        }
    }
} finally {
    Disconnect-PW | Out-Null
}

Write-Host "`n=== Reconcile summary ===" -ForegroundColor Cyan
Write-Host "  Groups processed: $($summary.groupsProcessed)"
Write-Host "  Groups skipped:   $($summary.groupsSkipped)"
Write-Host "  Email updates:    $($summary.emailUpdates)$(if ($DryRun) { ' (planned)' })"
Write-Host "  State updates:    $($summary.stateUpdates)$(if ($DryRun) { ' (planned)' })"
Write-Host "  DB row updates:   $($summary.dbUpdates)$(if ($DryRun) { ' (planned)' })"
Write-Host "  Errors:           $($summary.errors.Count)"
foreach ($err in @($summary.errors | Select-Object -First 10)) {
    Write-Host "    $err" -ForegroundColor Red
}

if ($Pretty) {
    $summary | ConvertTo-Json -Depth 8
} else {
    foreach ($d in @($summary.details | Select-Object -First 20)) {
        Write-Host ("  {0} | {1} | emails={2} states={3} db={4}" -f $d.folder, $d.stem, $d.emailChanges.Count, $d.stateChanges.Count, $d.dbRowsUpdated) -ForegroundColor Gray
    }
    if ($summary.details.Count -gt 20) {
        Write-Host "  ... $($summary.details.Count - 20) more groups with changes (use -Pretty for full JSON)" -ForegroundColor Gray
    }
}
