<#
.SYNOPSIS
Reconciles designer/reviewer/checker emails and workflow state across DGN, sheet PDF, and QC PDF.

.DESCRIPTION
Reads all rows from sheet_index, groups by folder + sheet stem, then applies:

  - Email source of truth: DGN (EM_Designer_Email / EM_Reviewer_Email / EM_Checker_Email)
    -> sheet PDF and lane QC PDFs (*-prod/-chk/-rev.pdf) are updated to match DGN when they differ.

  - State source of truth: lane QC PDF workflow state
    -> DGN and sheet PDF are updated to match the lane QC PDF when they differ.

ProjectWise documents are located by document_guid from the index (and qc_pdf_guid
when the QC PDF is linked but not indexed as its own row). sheet_index is updated
after successful PW reconciliation.

.EXAMPLE
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -DryRun
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -ConfirmWrites
PS> .\scripts\Reconcile-QCSheetOwnership.ps1 -ConfirmWrites -FolderPathFilter 'Caltrans\CAFWY2200-I-15_ELPSE\CADD\Sheets\Seg_1'

Do not combine -DryRun with -ConfirmWrites. Use -DryRun to preview only; use -ConfirmWrites alone to apply PW and database changes.
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
Import-Module (Join-Path $repoRoot 'modules\QC.ProcessType.psm1') -Force

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

function _RSO-NormalizeFolderPath {
    param([AllowNull()][string]$FolderPath)
    $p = ($FolderPath -as [string]).Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    if ($p -match '^(?i)Documents\\') { return $p }
    return ('Documents\' + $p)
}

function _RSO-SheetStem([string]$DocumentName) {
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        return ([string](Get-PWSheetStemFromDocumentName -DocumentName $DocumentName)).ToLowerInvariant()
    }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ($stem -match '(?i)-(prod|chk|rev)$') { $stem = $stem -replace '(?i)-(prod|chk|rev)$', '' }
    return $stem.ToLowerInvariant()
}

function _RSO-IsQcPdfName([string]$DocumentName) {
    if (Get-Command -Name 'Test-PWQcPdfLaneSuffix' -ErrorAction SilentlyContinue) {
        return [bool](Test-PWQcPdfLaneSuffix -DocumentName $DocumentName)
    }
    return [bool]([string]$DocumentName -match '(?i)-(prod|chk|rev)\.pdf$')
}

function _RSO-GetEmailColumns {
    param([hashtable]$Config)
    $designer = 'EM_Designer_Email'
    $reviewer = 'EM_Reviewer_Email'
    $checker = 'EM_Checker_Email'
    try {
        $na = $Config['notifications']['attributes']
        if ($na) {
            if ($na['designerEmailField']) { $designer = [string]$na['designerEmailField'] }
            if ($na['reviewerEmailField']) { $reviewer = [string]$na['reviewerEmailField'] }
            if ($na['checkerEmailField']) { $checker = [string]$na['checkerEmailField'] }
        }
        $pw = $Config.projectWise
        if ($pw -and $pw.environmentEmailAttributes) {
            $ea = $pw.environmentEmailAttributes
            if ($ea.default) {
                if ($ea.default.designerEmailColumn) { $designer = [string]$ea.default.designerEmailColumn }
                if ($ea.default.reviewerEmailColumn) { $reviewer = [string]$ea.default.reviewerEmailColumn }
                if ($ea.default.checkerEmailColumn) { $checker = [string]$ea.default.checkerEmailColumn }
            }
        }
    } catch { }
    return @{ designer = $designer; reviewer = $reviewer; checker = $checker }
}

function _RSO-BuildFolderEmailMap {
    param(
        [string]$FolderPath,
        [string]$DesignerColumn,
        [string]$ReviewerColumn,
        [string]$CheckerColumn
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
        $returnCols = @($DesignerColumn, $ReviewerColumn, $CheckerColumn) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
            $params['ColumnsToReturn'] = $returnCols
        } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
            $params['ReturnColumns'] = $returnCols
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
            $checker = if ($attrs.ContainsKey($CheckerColumn)) { [string]$attrs[$CheckerColumn] } else { '' }
            $key = $name.ToLowerInvariant()
            $map[$key] = @{
                documentGuid  = $guid
                designerEmail = $designer.Trim()
                reviewerEmail = $reviewer.Trim()
                checkerEmail  = $checker.Trim()
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
    if ([string]::IsNullOrWhiteSpace($StateName)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_EMPTY_STATE_GUARDED' `
                -Message 'Skipped reconcile workflow state write because StateName was empty.' -Data @{
                callSite = '_RSO-SetPwDocumentState.StateName'
                auditEventId = $null
                documentName = ''
                folderPath = ''
                sourceVariableName = 'StateName'
                sourceValue = $StateName
                livePwState = ''
                changedByUsername = ''
            } | Out-Null
        }
        return
    }
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

function _RSO-ResolvePwDocument {
    param(
        [hashtable]$DocByGuid,
        [hashtable]$EmailMap,
        [string]$FolderPath,
        [string]$DocumentName,
        [string]$DocumentGuid
    )

    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $gk = $DocumentGuid.ToLowerInvariant()
        if ($DocByGuid.ContainsKey($gk)) { return $DocByGuid[$gk] }
    }
    if ($DocumentName -and $EmailMap.ContainsKey($DocumentName.ToLowerInvariant())) {
        $hit = $EmailMap[$DocumentName.ToLowerInvariant()]
        if ($hit.document) { return $hit.document }
    }
    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
        if ($searchCmd) {
            $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
            if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }
            try {
                $params = @{
                    FolderPath     = $apiPath
                    JustThisFolder = $true
                    DocumentName   = $DocumentName
                    ErrorAction    = 'Stop'
                }
                if ($searchCmd.Parameters.ContainsKey('PopulatePath')) { $params['PopulatePath'] = $true }
                return (& $searchCmd @params | Select-Object -First 1)
            } catch { }
        }
    }
    return $null
}

function _RSO-LoadPwDocumentsByGuid {
    param(
        [string[]]$DocumentGuids,
        [System.Collections.Generic.List[string]]$InvalidGuids,
        [System.Collections.Generic.List[string]]$Errors
    )

    $docByGuid = @{}
    $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
    if (-not $guidCmd) { return $docByGuid }

    $valid = @($DocumentGuids | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    foreach ($raw in @($DocumentGuids | Select-Object -Unique)) {
        if (-not (Test-PWValidDocumentGuid -DocumentGuid $raw)) {
            if (-not [string]::IsNullOrWhiteSpace($raw)) { $InvalidGuids.Add([string]$raw) | Out-Null }
        }
    }
    if ($valid.Count -eq 0) { return $docByGuid }

    $chunkSize = 200
    for ($i = 0; $i -lt $valid.Count; $i += $chunkSize) {
        $chunk = @($valid[$i..[Math]::Min($i + $chunkSize - 1, $valid.Count - 1)])
        try {
            foreach ($doc in @(& $guidCmd -DocumentGUIDs $chunk -ErrorAction Stop)) {
                $dg = ''
                try { $dg = [string]$doc.DocumentGUID } catch { }
                if ($dg) { $docByGuid[$dg.ToLowerInvariant()] = $doc }
            }
        } catch {
            foreach ($oneGuid in $chunk) {
                try {
                    $doc = & $guidCmd -DocumentGUIDs @($oneGuid) -ErrorAction Stop | Select-Object -First 1
                    if (-not $doc) { continue }
                    $dg = [string]$doc.DocumentGUID
                    if ($dg) { $docByGuid[$dg.ToLowerInvariant()] = $doc }
                } catch {
                    $Errors.Add("GUID read failed for $oneGuid : $($_.Exception.Message)") | Out-Null
                }
            }
        }
    }
    return $docByGuid
}

$cfgRes = Read-QCAppSettings -Path $AppSettingsPath
if (-not $cfgRes.IsSuccess) { throw $cfgRes.Message }
$config = [hashtable]$cfgRes.Data.config

if (-not (Test-QCDatabaseEnabled -Config $config)) {
    throw 'database.enabled must be true for sheet_index reconciliation.'
}

if ($DryRun.IsPresent -and $ConfirmWrites.IsPresent) {
    throw 'Use -DryRun (preview only) OR -ConfirmWrites (apply changes), not both.'
}

$doWrites = $ConfirmWrites.IsPresent
if (-not $DryRun.IsPresent -and -not $ConfirmWrites.IsPresent) {
    Write-Host 'Refusing PW/DB writes: pass -ConfirmWrites to apply changes, or -DryRun to preview only.' -ForegroundColor Yellow
    $DryRun = $true
}

if ($doWrites) {
    Write-Host 'WRITE MODE: ProjectWise and sheet_index will be updated.' -ForegroundColor Green
} else {
    Write-Host 'PREVIEW MODE: no ProjectWise or database changes will be made.' -ForegroundColor Yellow
}

Initialize-QCDatabaseSchema -Config $config | Out-Null

$emailCols = _RSO-GetEmailColumns -Config $config
$designerCol = $emailCols.designer
$reviewerCol = $emailCols.reviewer
$checkerCol = $emailCols.checker

$sql = @"
SELECT document_guid, document_name, folder_path, extension, source_type,
       qc_pdf_guid, qc_pdf_name, designer_email, reviewer_email, checker_email, pw_state_name
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
        folderPath     = _RSO-NormalizeFolderPath -FolderPath ([string]$r.folder_path)
        extension      = if ($r.extension -is [DBNull]) { '' } else { [string]$r.extension }
        sourceType     = if ($r.source_type -is [DBNull]) { '' } else { [string]$r.source_type }
        qcPdfGuid      = if ($r.qc_pdf_guid -is [DBNull]) { '' } else { if (Test-PWValidDocumentGuid -DocumentGuid ([string]$r.qc_pdf_guid)) { [string]$r.qc_pdf_guid } else { '' } }
        qcPdfGuidRaw   = if ($r.qc_pdf_guid -is [DBNull]) { '' } else { [string]$r.qc_pdf_guid }
        qcPdfName      = if ($r.qc_pdf_name -is [DBNull]) { '' } else { [string]$r.qc_pdf_name }
        designerEmail  = if ($r.designer_email -is [DBNull]) { '' } else { [string]$r.designer_email }
        reviewerEmail  = if ($r.reviewer_email -is [DBNull]) { '' } else { [string]$r.reviewer_email }
        checkerEmail   = if ($r.Table.Columns.Contains('checker_email') -and -not ($r.checker_email -is [DBNull])) { [string]$r.checker_email } else { '' }
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
            qcPdfGuidRaw = ''
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
    if ($row.qcPdfGuidRaw -and -not $g.qcPdfGuidRaw) { $g.qcPdfGuidRaw = $row.qcPdfGuidRaw }
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
    invalidGuids      = @()
    dryRun            = -not $doWrites
    writeMode         = $doWrites
    details           = @()
}

try {
    $folderEmailCache = @{}
    $invalidGuids = [System.Collections.Generic.List[string]]::new()
    $guidLoadErrors = [System.Collections.Generic.List[string]]::new()

    $allGuids = [System.Collections.Generic.List[string]]::new()
    foreach ($g in $groupList) {
        foreach ($part in @($g.dgn, $g.pdf, $g.qcIndexed)) {
            if ($part -and (Test-PWValidDocumentGuid -DocumentGuid $part.documentGuid)) {
                $allGuids.Add($part.documentGuid) | Out-Null
            }
        }
        if (Test-PWValidDocumentGuid -DocumentGuid $g.qcPdfGuid) { $allGuids.Add($g.qcPdfGuid) | Out-Null }
    }
    $uniqueGuids = @($allGuids | Select-Object -Unique)
    $docByGuid = _RSO-LoadPwDocumentsByGuid -DocumentGuids $uniqueGuids -InvalidGuids $invalidGuids -Errors $guidLoadErrors
    if ($invalidGuids.Count -gt 0) {
        $summary.invalidGuids = @($invalidGuids | Select-Object -Unique)
    }
    foreach ($ge in @($guidLoadErrors)) { $summary.errors += $ge }

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
                -DesignerColumn $designerCol -ReviewerColumn $reviewerCol -CheckerColumn $checkerCol
        }
        $emailMap = $folderEmailCache[$folderKey]

        $dgnNameKey = $g.dgn.documentName.ToLowerInvariant()
        $dgnEmails = if ($emailMap.ContainsKey($dgnNameKey)) { $emailMap[$dgnNameKey] } else { $null }
        $canonicalDesigner = if ($dgnEmails) { [string]$dgnEmails.designerEmail } else { $g.dgn.designerEmail }
        $canonicalReviewer = if ($dgnEmails) { [string]$dgnEmails.reviewerEmail } else { $g.dgn.reviewerEmail }
        $canonicalChecker = if ($dgnEmails) { [string]$dgnEmails.checkerEmail } else { $g.dgn.checkerEmail }

        $qcGuid = ''
        $qcName = ''
        if ($g.qcIndexed) {
            $qcGuid = $g.qcIndexed.documentGuid
            $qcName = $g.qcIndexed.documentName
        } elseif ($g.qcPdfGuid) {
            $qcGuid = $g.qcPdfGuid
            $qcName = $g.qcPdfName
        } elseif ($g.qcPdfName) {
            $qcName = $g.qcPdfName
        }
        $canonicalState = ''
        if (Test-PWValidDocumentGuid -DocumentGuid $qcGuid) {
            $qk = $qcGuid.ToLowerInvariant()
            if ($stateByGuid.ContainsKey($qk)) { $canonicalState = [string]$stateByGuid[$qk] }
        }
        if (-not $canonicalState -and $qcName) {
            try {
                $canonicalState = Get-PWDocumentWorkflowStateName -FolderPath $g.folderPath -DocumentName $qcName -DocumentGuid $qcGuid
            } catch { }
        }

        $detail = [ordered]@{
            folder            = $g.folderPath
            stem              = $g.stem
            dgn               = $g.dgn.documentName
            pdf               = if ($g.pdf) { $g.pdf.documentName } else { $null }
            qcPdf             = if ($g.qcPdfName) { $g.qcPdfName } elseif ($g.qcIndexed) { $g.qcIndexed.documentName } else { $null }
            canonicalDesigner = $canonicalDesigner
            canonicalReviewer = $canonicalReviewer
            canonicalChecker  = $canonicalChecker
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
            $currentChecker = ''
            if ($target.row) {
                $nk = $target.nameKey
                if ($emailMap.ContainsKey($nk)) {
                    $currentDesigner = [string]$emailMap[$nk].designerEmail
                    $currentReviewer = [string]$emailMap[$nk].reviewerEmail
                    $currentChecker = [string]$emailMap[$nk].checkerEmail
                } else {
                    $currentDesigner = [string]$target.row.designerEmail
                    $currentReviewer = [string]$target.row.reviewerEmail
                    $currentChecker = [string]$target.row.checkerEmail
                }
            } elseif ($target.nameKey -and $emailMap.ContainsKey($target.nameKey)) {
                $currentDesigner = [string]$emailMap[$target.nameKey].designerEmail
                $currentReviewer = [string]$emailMap[$target.nameKey].reviewerEmail
                $currentChecker = [string]$emailMap[$target.nameKey].checkerEmail
            }

            $needsDesigner = (_RSO-Norm $currentDesigner) -ne (_RSO-Norm $canonicalDesigner)
            $needsReviewer = (_RSO-Norm $currentReviewer) -ne (_RSO-Norm $canonicalReviewer)
            $needsChecker = (_RSO-Norm $currentChecker) -ne (_RSO-Norm $canonicalChecker)
            if (-not $needsDesigner -and -not $needsReviewer -and -not $needsChecker) { continue }

            $targetDocumentName = if ($target.row) { $target.row.documentName } else { $target.name }
            $docObj = _RSO-ResolvePwDocument -DocByGuid $docByGuid -EmailMap $emailMap `
                -FolderPath $g.folderPath -DocumentName $targetDocumentName `
                -DocumentGuid $(if ($target.row) { $target.row.documentGuid } else { $target.guid })

            $change = [ordered]@{
                role     = $target.role
                document = $targetDocumentName
                from     = @{ designer = $currentDesigner; reviewer = $currentReviewer; checker = $currentChecker }
                to       = @{ designer = $canonicalDesigner; reviewer = $canonicalReviewer; checker = $canonicalChecker }
                applied  = $false
            }

            if ($docObj -and $doWrites) {
                $toWrite = @{}
                if ($needsDesigner) { $toWrite[$designerCol] = $canonicalDesigner }
                if ($needsReviewer) { $toWrite[$reviewerCol] = $canonicalReviewer }
                if ($needsChecker) { $toWrite[$checkerCol] = $canonicalChecker }
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

                $docObj = _RSO-ResolvePwDocument -DocByGuid $docByGuid -EmailMap $emailMap `
                    -FolderPath $g.folderPath -DocumentName $target.row.documentName -DocumentGuid $target.row.documentGuid
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
                        if ([string]::IsNullOrWhiteSpace($canonicalState)) {
                            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_EMPTY_STATE_GUARDED' `
                                    -Message 'Skipped reconcile state write because canonical state was empty.' -Data @{
                                    callSite = 'Reconcile-QCSheetOwnership.canonicalState'
                                    auditEventId = $null
                                    documentName = $target.row.documentName
                                    folderPath = $g.folderPath
                                    sourceVariableName = 'canonicalState'
                                    sourceValue = $canonicalState
                                    livePwState = [string]$currentState
                                    changedByUsername = ''
                                } | Out-Null
                            }
                            $change.skipped = $true
                            $change.reason = 'empty canonical state'
                        } else {
                            try {
                                _RSO-SetPwDocumentState -Document $docObj -StateName $canonicalState
                                $change.applied = $true
                                $summary.stateUpdates++
                            } catch {
                                $summary.errors += "State update failed for $($target.row.documentName): $($_.Exception.Message)"
                            }
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
                -or (_RSO-Norm $dbRow.reviewerEmail) -ne (_RSO-Norm $canonicalReviewer) `
                -or (_RSO-Norm $dbRow.checkerEmail) -ne (_RSO-Norm $canonicalChecker)
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
                    -CheckerEmail $canonicalChecker `
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
Write-Host "  Email updates:    $($summary.emailUpdates)$(if (-not $doWrites) { ' (planned)' })"
Write-Host "  State updates:    $($summary.stateUpdates)$(if (-not $doWrites) { ' (planned)' })"
Write-Host "  DB row updates:   $($summary.dbUpdates)$(if (-not $doWrites) { ' (planned)' })"
Write-Host "  Errors:           $($summary.errors.Count)"
if ($summary.invalidGuids -and $summary.invalidGuids.Count -gt 0) {
    Write-Host "  Invalid GUIDs:      $($summary.invalidGuids.Count) (skipped; fix qc_pdf_guid in sheet_index)" -ForegroundColor Yellow
    foreach ($ig in @($summary.invalidGuids | Select-Object -First 5)) {
        Write-Host "    $ig" -ForegroundColor DarkYellow
    }
}
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
