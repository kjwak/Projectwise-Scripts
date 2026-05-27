<#
.SYNOPSIS
Compare ProjectWise APIs for reading workflow state on sheet PDFs and *-qc.pdf documents.

.DESCRIPTION
Read-only unless -ConfirmStateWrite is passed (uses Test-QCWorkflowWriteback patterns).
Reports which cmdlets/columns return StateName vs WorkflowState for a target document.
#>
[CmdletBinding()]
param(
    [string]$AppSettingsPath = '',
    [string]$FolderPath = '',
    [string]$DocumentName = '',
    [switch]$UseQcPdfFromDatabase,
    [switch]$AlsoTestSourcePdf,
    [switch]$ConfirmStateWrite,
    [string]$TargetState = 'Corrections In Progress',
    [switch]$Pretty
)

$ErrorActionPreference = 'Continue'

$scriptDir = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}

foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'Core.Database.psm1', 'PW.Connection.psm1', 'PW.Discovery.psm1')) {
    Import-Module (Join-Path $repoRoot "modules\$mod") -Force -ErrorAction Stop
}
foreach ($moduleName in @('pwps', 'pwps_dab')) {
    if (-not (Get-Module -Name $moduleName -ErrorAction SilentlyContinue)) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($available) { Import-Module $available.Name -ErrorAction SilentlyContinue | Out-Null }
    }
}

function _DeepHashtable ($obj) {
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = _DeepHashtable $p.Value }
        return $h
    }
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        return @($obj | ForEach-Object { _DeepHashtable $_ })
    }
    return $obj
}

function Get-PropValues {
    param([object]$Obj, [string[]]$Names)
    $out = [ordered]@{}
    if (-not $Obj) { return $out }
    foreach ($n in $Names) {
        try {
            if ($Obj.PSObject.Properties[$n]) { $out[$n] = [string]$Obj.$n }
        } catch { $out[$n] = $null }
    }
    return $out
}

function Test-ReadMethod {
    param(
        [string]$Label,
        [scriptblock]$Reader
    )
    $row = [ordered]@{ method = $Label; ok = $false; error = $null; properties = $null }
    try {
        $doc = & $Reader
        if ($doc) {
            $row.ok = $true
            $row.properties = Get-PropValues -Obj $doc -Names @(
                'Name', 'DocumentName', 'FileName', 'DocumentGUID', 'DocumentID', 'ProjectID',
                'Workflow', 'WorkflowName', 'WorkflowState', 'WorkflowStateName',
                'State', 'StateName', 'DocumentState', 'CurrentState'
            )
        } else {
            $row.error = 'no document returned'
        }
    } catch {
        $row.error = $_.Exception.Message
    }
    return [pscustomobject]$row
}

$config = _DeepHashtable ((Get-Content -LiteralPath $AppSettingsPath -Raw) | ConvertFrom-Json)

# Optional: pick a qc.pdf from sheet_index
if ($UseQcPdfFromDatabase -and (Test-QCDatabaseEnabled -Config $config)) {
    $pick = Invoke-QCDatabaseQuery -Config $config -Sql @"
SELECT TOP 1 folder_path, qc_pdf_name
FROM sheet_index
WHERE qc_pdf_name IS NOT NULL AND qc_pdf_name LIKE '%-qc.pdf'
ORDER BY last_updated_at DESC
"@
    if ($pick.IsSuccess -and $pick.Data.rowCount -gt 0) {
        $r = $pick.Data.table.Rows[0]
        if (-not $FolderPath) { $FolderPath = [string]$r.folder_path }
        if (-not $DocumentName) { $DocumentName = [string]$r.qc_pdf_name }
    }
}

if ([string]::IsNullOrWhiteSpace($FolderPath) -or [string]::IsNullOrWhiteSpace($DocumentName)) {
    Write-Error 'Provide -FolderPath and -DocumentName, or -UseQcPdfFromDatabase with a populated sheet_index.'
    return
}

$apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }

$report = [ordered]@{
    folderPath = $FolderPath
    apiFolderPath = $apiPath
    documentName = $DocumentName
    databaseEnabled = (Test-QCDatabaseEnabled -Config $config)
    pwConnected = $false
    getPWDocumentEmailContacts = $null
    readMethods = @()
    setPWDocumentStateParameters = $null
}

# DB row for this doc (by name in folder)
if ($report.databaseEnabled) {
    $db = Invoke-QCDatabaseQuery -Config $config -Sql @"
SELECT document_guid, document_name, pw_state_name, qc_pdf_name, qc_pdf_guid, source_type
FROM sheet_index
WHERE folder_path = @fp AND (document_name = @dn OR qc_pdf_name = @dn)
"@ -Parameters @{ fp = $FolderPath; dn = $DocumentName }
    if ($db.IsSuccess -and $db.Data.rowCount -gt 0) {
        $report.sheetIndexRows = @($db.Data.table.Rows | ForEach-Object {
            @{
                document_guid = [string]$_.document_guid
                document_name = [string]$_.document_name
                pw_state_name = if ($_.pw_state_name -is [DBNull]) { $null } else { [string]$_.pw_state_name }
                qc_pdf_name = if ($_.qc_pdf_name -is [DBNull]) { $null } else { [string]$_.qc_pdf_name }
                source_type = if ($_.source_type -is [DBNull]) { $null } else { [string]$_.source_type }
            }
        })
    }
}

try {
    $pw = $config.projectWise
    $ds = if ($pw.datasourceName) { [string]$pw.datasourceName } elseif ($pw.datasource) { [string]$pw.datasource } else { $null }
    $credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }
    $credRes = Get-PWCredentialFromFile -CredentialPath $credPath
    if (-not $credRes.IsSuccess) { throw $credRes.Message }
    $connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
    if (-not $connRes.IsSuccess) { throw $connRes.Message }
    $report.pwConnected = $true

    $report.getPWDocumentEmailContacts = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $DocumentName
    $report.getPWDocumentWorkflowStateName = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName

    $stateCols = @('Name', 'StateName', 'WorkflowState', 'WorkflowName', 'Workflow', 'DocumentID', 'DocumentGUID')
    $emailCols = @('EM_Designer_Email', 'EM_Reviewer_Email')

    $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsBySearch -PopulatePath' -Reader {
        Get-PWDocumentsBySearch -FolderPath $apiPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction Stop | Select-Object -First 1
    }
    $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsBySearchWithReturnColumns (StateName only)' -Reader {
        Get-PWDocumentsBySearchWithReturnColumns -FolderPath $apiPath -JustThisFolder -DocumentName $DocumentName -ColumnsToReturn @('StateName') -ErrorAction Stop | Select-Object -First 1
    }
    $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsBySearchWithReturnColumns (emails only — current indexer)' -Reader {
        Get-PWDocumentsBySearchWithReturnColumns -FolderPath $apiPath -JustThisFolder -ColumnsToReturn $emailCols -ErrorAction Stop | Select-Object -First 1
    }
    $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsBySearchWithReturnColumns (emails + StateName)' -Reader {
        Get-PWDocumentsBySearchWithReturnColumns -FolderPath $apiPath -JustThisFolder -DocumentName $DocumentName -ColumnsToReturn ($emailCols + 'StateName') -ErrorAction Stop | Select-Object -First 1
    }
    $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsBySearchWithReturnColumns (full stateCols)' -Reader {
        Get-PWDocumentsBySearchWithReturnColumns -FolderPath $apiPath -JustThisFolder -DocumentName $DocumentName -ColumnsToReturn $stateCols -ErrorAction Stop | Select-Object -First 1
    }

    $doc = $report.readMethods | Where-Object { $_.ok -and $_.properties.StateName } | Select-Object -First 1
    if (-not $doc) {
        $doc = $report.readMethods | Where-Object { $_.ok } | Select-Object -First 1
    }
    if ($doc -and $doc.properties.DocumentGUID) {
        $guid = [string]$doc.properties.DocumentGUID
        $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentsByGUIDs' -Reader {
            Get-PWDocumentsByGUIDs -DocumentGUIDs @($guid) -ErrorAction Stop | Select-Object -First 1
        }
        $docId = 0; $projId = 0
        [void][int]::TryParse([string]$doc.properties.DocumentID, [ref]$docId)
        [void][int]::TryParse([string]$doc.properties.ProjectID, [ref]$projId)
        if ($docId -gt 0 -and $projId -gt 0) {
            $report.readMethods += Test-ReadMethod -Label 'Get-PWDocumentEAttributes (same IDs)' -Reader {
                Get-PWDocumentEAttributes -DocumentID $docId -ProjectID $projId -ErrorAction Stop
            }
        }
    }

    $setCmd = Get-Command Set-PWDocumentState -ErrorAction SilentlyContinue
    if ($setCmd) {
        $report.setPWDocumentStateParameters = @($setCmd.Parameters.Keys | Sort-Object)
    }

    if ($AlsoTestSourcePdf -and $DocumentName -match '(?i)-qc\.pdf$') {
        $sourceName = $DocumentName -replace '(?i)-qc\.pdf$', '.pdf'
        $report.sourcePdfRead = Test-ReadMethod -Label "source pdf: $sourceName" -Reader {
            Get-PWDocumentsBySearchWithReturnColumns -FolderPath $apiPath -JustThisFolder -DocumentName $sourceName -ColumnsToReturn @('StateName', 'WorkflowState') -ErrorAction Stop | Select-Object -First 1
        }
    }

    Disconnect-PW | Out-Null
} catch {
    $report.pwError = $_.Exception.Message
}

if ($ConfirmStateWrite -and $report.pwConnected) {
    $testPath = Join-Path $FolderPath $DocumentName
    Write-Host "`n--- Controlled write test (StateOnly) ---" -ForegroundColor Yellow
    $writeScript = Join-Path $repoRoot 'tools\discovery\Test-QCWorkflowWriteback.ps1'
    & $writeScript -ConfirmWrites -TestDocumentPath $testPath -Mode StateOnly -TargetState $TargetState -Pretty:$Pretty
}

if ($Pretty) { $report | ConvertTo-Json -Depth 8 } else { $report | ConvertTo-Json -Depth 8 -Compress }
