# QC.Reporting.psm1
# Responsibility: Read-only QC attribute-first reporting aggregation and JSON snapshots.

Import-Module (Join-Path $PSScriptRoot 'Core.Results.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

function _QCR-ToHashtable([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [hashtable]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) { return $Value }
    if ($Value -is [string] -or $Value -is [System.ValueType]) { return @{ value = $Value } }
    if ($Value.PSObject -and $Value.PSObject.Properties) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    }
    return $null
}

function _QCR-IsBlank([object]$Value) {
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function _QCR-GetProp([object]$Object, [string[]]$Names) {
    foreach ($n in @($Names)) {
        try { if ($Object -and $Object.PSObject.Properties[$n] -and $null -ne $Object.$n) { return $Object.$n } } catch { }
    }
    return $null
}

function _QCR-SafeName([string]$Name) {
    if (_QCR-IsBlank $Name) { return 'project' }
    return (($Name -replace '[\\/:*?"<>|]+','_') -replace '\s+','_').Trim('_')
}

function _QCR-ToBool([object]$Value) {
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return [bool]$Value }
    $s = ([string]$Value).Trim()
    return ($s -match '(?i)^(true|1|yes|y|active|open)$')
}

function _QCR-ToDate([object]$Value) {
    if ($null -eq $Value) { return $null }
    try { return ([datetime]$Value) } catch { return $null }
}

function _QCR-GetAttributeValue([object]$Document, [string]$AttributeName) {
    if (_QCR-IsBlank $AttributeName) { return $null }
    $containers = @()
    foreach ($prop in @('qcAttributes','attributes','Attributes','CustomAttributes','EnvironmentAttributes')) {
        try { if ($Document -and $Document.PSObject.Properties[$prop] -and $Document.$prop) { $containers += $Document.$prop } } catch { }
    }
    foreach ($c in @($containers)) {
        $h = _QCR-ToHashtable $c
        if ($h -and $h.ContainsKey($AttributeName)) { return $h[$AttributeName] }
    }
    try { if ($Document -and $Document.PSObject.Properties[$AttributeName]) { return $Document.$AttributeName } } catch { }
    return $null
}

function Get-QCReportingSettings {
    [CmdletBinding()]
    param([hashtable]$Config)

    $raw = @{}
    if ($Config -and $Config.ContainsKey('qcReporting') -and $Config.qcReporting) { $raw = _QCR-ToHashtable $Config.qcReporting }
    $wf = @{}
    if ($Config -and $Config.ContainsKey('qcWorkflow') -and $Config.qcWorkflow) { $wf = _QCR-ToHashtable $Config.qcWorkflow }
    $attributeMap = if ($wf -and $wf.ContainsKey('attributeMap')) { _QCR-ToHashtable $wf.attributeMap } else { @{} }

    $settings = @{
        enabled = $false
        snapshotRoot = 'metrics\qc'
        staleDays = 14
        includeWorkflowStateCounts = $true
        scanSubtree = $true
        scheduledIntervalMinutes = 60
        attributeMap = $attributeMap
    }
    if ($raw) { foreach ($k in $raw.Keys) { $settings[$k] = $raw[$k] } }
    if (-not $settings.attributeMap -or $settings.attributeMap.Keys.Count -eq 0) {
        $settings.attributeMap = @{
            qcActive = 'QC_Active'
            reviewType = 'QC_Review_Type'
            cycleId = 'QC_Cycle_ID'
            status = 'QC_Status'
            lastActionDate = 'QC_Last_Action_Date'
            automationLastRun = 'QC_Automation_Last_Run'
            automationResult = 'QC_Automation_Result'
            automationError = 'QC_Automation_Error'
        }
    }
    try { $settings.staleDays = [int]$settings.staleDays } catch { $settings.staleDays = 14 }
    return $settings
}

function ConvertTo-QCReportingDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Document,
        [Parameter(Mandatory)]
        [hashtable]$AttributeMap
    )

    $attrs = @{}
    foreach ($key in @($AttributeMap.Keys)) { $attrs[$key] = _QCR-GetAttributeValue -Document $Document -AttributeName ([string]$AttributeMap[$key]) }
    $state = _QCR-GetProp -Object $Document -Names @('StateName','WorkflowState','State','DocumentState')
    $workflow = _QCR-GetProp -Object $Document -Names @('WorkflowName','Workflow')
    $name = _QCR-GetProp -Object $Document -Names @('Name','DocumentName','FileName')
    $path = _QCR-GetProp -Object $Document -Names @('FullPath','Path','FolderPath')
    $lastAction = _QCR-ToDate $attrs.lastActionDate
    $cycleStart = _QCR-ToDate (_QCR-GetAttributeValue -Document $Document -AttributeName 'QC_Cycle_Start_Date')

    return [pscustomobject]@{
        name = $name
        path = $path
        workflowName = $workflow
        stateName = $state
        qcActive = (_QCR-ToBool $attrs.qcActive)
        cycleId = $attrs.cycleId
        reviewType = $attrs.reviewType
        status = $attrs.status
        lastActionDate = $lastAction
        cycleStartDate = $cycleStart
        automationResult = $attrs.automationResult
        automationError = $attrs.automationError
        attributes = $attrs
    }
}

function Get-QCReportingDocuments {
    [CmdletBinding()]
    param(
        [hashtable]$Job,
        [hashtable]$Settings
    )

    if ($Job -and $Job.ContainsKey('documents') -and $Job.documents) { return @($Job.documents) }

    $folder = $null
    if ($Job -and $Job.ContainsKey('sourceFolder')) { $folder = [string]$Job.sourceFolder }
    elseif ($Job -and $Job.ContainsKey('folderPath')) { $folder = [string]$Job.folderPath }
    if (_QCR-IsBlank $folder) { return @() }

    $docs = @()
    $cmd = Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue
    if ($cmd) {
        $cols = @('Name','DocumentID','ProjectID','WorkflowName','StateName') + @($Settings.attributeMap.Values)
        $cols = @($cols | Where-Object { -not (_QCR-IsBlank $_) } | Select-Object -Unique)
        $param = if ($cmd.Parameters.ContainsKey('ReturnColumns')) { 'ReturnColumns' } elseif ($cmd.Parameters.ContainsKey('ColumnsToReturn')) { 'ColumnsToReturn' } else { $null }
        try {
            $args = @{ FolderPath = $folder; ErrorAction = 'Stop' }
            if ($param) { $args[$param] = $cols }
            if ($cmd.Parameters.ContainsKey('PopulatePath')) { $args['PopulatePath'] = $true }
            if (-not [bool]$Settings.scanSubtree -and $cmd.Parameters.ContainsKey('JustThisFolder')) { $args['JustThisFolder'] = $true }
            $docs = @(& $cmd @args)
        } catch { $docs = @() }
    }
    if ($docs.Count -eq 0 -and (Get-Command -Name Get-PWDocumentsBySearchExtended -ErrorAction SilentlyContinue)) {
        try { $docs = @(Get-PWDocumentsBySearchExtended -FolderPath $folder -ErrorAction Stop) } catch { $docs = @() }
    }
    return @($docs)
}

function Get-QCPackageReportingRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$FolderPath = ''
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return @() }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return @() }

    $sql = 'SELECT * FROM v_sheet_package_status'
    $params = @{}
    if (-not (_QCR-IsBlank $FolderPath)) {
        $sql += ' WHERE folder_path = @folderPath'
        $params['folderPath'] = ([string]$FolderPath).Trim()
    }
    try {
        $res = Invoke-QCDatabaseQuery -Config $Config -Sql $sql -Parameters $params
        if (-not $res.IsSuccess -or -not $res.Data.table) { return @() }
        return @($res.Data.table.Rows)
    } catch {
        return @()
    }
}

function ConvertTo-QCPackageReportingRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Row)

    $get = {
        param([string]$Name)
        try {
            if ($Row -is [System.Data.DataRow]) {
                if ($Row.Table.Columns.Contains($Name) -and $Row[$Name] -isnot [DBNull]) { return $Row[$Name] }
            } elseif ($Row.PSObject.Properties[$Name]) { return $Row.$Name }
        } catch { }
        return $null
    }

    return [pscustomobject]@{
        sheetPackageId = & $get 'sheet_package_id'
        sheetStem = & $get 'sheet_stem'
        folderPath = & $get 'folder_path'
        pwStateName = & $get 'pw_state_name'
        qcReviewType = & $get 'qc_review_type'
        qcAssignedTo = & $get 'qc_assigned_to'
        productionQcCompletedCount = & $get 'production_qc_completed_count'
        peerReviewCompletedCount = & $get 'peer_review_completed_count'
        independentCheckCompletedCount = & $get 'independent_check_completed_count'
        productionQcLastCompletedAt = & $get 'production_qc_last_completed_at'
        peerReviewLastCompletedAt = & $get 'peer_review_last_completed_at'
        independentCheckLastCompletedAt = & $get 'independent_check_last_completed_at'
        dgnGuid = & $get 'dgn_guid'
        sheetPdfGuid = & $get 'sheet_pdf_guid'
        qcPdfGuid = & $get 'qc_pdf_guid'
    }
}

function New-QCPackageReportingMetrics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Packages
    )

    $records = @($Packages | ForEach-Object { ConvertTo-QCPackageReportingRecord -Row $_ })
    $countByState = @{}
    foreach ($pkg in @($records)) {
        $state = if (_QCR-IsBlank $pkg.pwStateName) { '<none>' } else { [string]$pkg.pwStateName }
        if (-not $countByState.ContainsKey($state)) { $countByState[$state] = 0 }
        $countByState[$state]++
    }

    return [pscustomobject]@{
        packageCount = @($records).Count
        inProductionCount = @($records | Where-Object { ([string]$_.pwStateName) -ieq 'In Production' }).Count
        readyForQcCount = @($records | Where-Object { ([string]$_.pwStateName) -ieq 'Ready for QC' -or ([string]$_.pwStateName) -ieq 'QC Received' }).Count
        redlinesReceivedCount = @($records | Where-Object { ([string]$_.pwStateName) -ieq 'Redlines Received' -or ([string]$_.pwStateName) -ieq 'Redlines Issued' }).Count
        correctionsReceivedCount = @($records | Where-Object { ([string]$_.pwStateName) -match '(?i)Corrections (Received|In Progress)|Verification In Progress|Backcheck In Progress' }).Count
        qcFinalizingCount = @($records | Where-Object { ([string]$_.pwStateName) -ieq 'QC Finalizing' }).Count
        qcCompleteCount = @($records | Where-Object { ([string]$_.pwStateName) -ieq 'QC Complete' -or ([string]$_.pwStateName) -ieq 'Verified Closed' }).Count
        productionQcCompletedTotal = ($records | ForEach-Object { [int]$_.productionQcCompletedCount } | Measure-Object -Sum).Sum
        peerReviewCompletedTotal = ($records | ForEach-Object { [int]$_.peerReviewCompletedCount } | Measure-Object -Sum).Sum
        independentCheckCompletedTotal = ($records | ForEach-Object { [int]$_.independentCheckCompletedCount } | Measure-Object -Sum).Sum
        workflowStateCounts = [pscustomobject]$countByState
    }
}

function New-QCReportingSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Documents,
        [Parameter(Mandatory)]
        [hashtable]$Settings,
        [string]$Project = 'project',
        [object[]]$PackageRows = @()
    )

    $normalized = @($Documents | ForEach-Object { ConvertTo-QCReportingDocument -Document $_ -AttributeMap $Settings.attributeMap })
    $now = Get-Date
    $active = @($normalized | Where-Object { $_.qcActive -or -not (_QCR-IsBlank $_.status) -or -not (_QCR-IsBlank $_.reviewType) })
    $errors = @($active | Where-Object { -not (_QCR-IsBlank $_.automationError) -or ([string]$_.automationResult) -match '(?i)fail|error' })
    $completeStateNames = @('QC Complete', 'Verified Closed')
    $stale = @($active | Where-Object {
        $_.lastActionDate -and (($now - $_.lastActionDate).TotalDays -gt [int]$Settings.staleDays) -and
        (($completeStateNames | Where-Object { $_ -ieq [string]$_.stateName }).Count -eq 0)
    })
    $closed = @($active | Where-Object { ($completeStateNames | Where-Object { $_ -ieq [string]$_.stateName }).Count -gt 0 })
    $cycleDays = @($closed | ForEach-Object { if ($_.cycleStartDate -and $_.lastActionDate) { ($_.lastActionDate - $_.cycleStartDate).TotalDays } } | Where-Object { $null -ne $_ })
    $avg = if ($cycleDays.Count -gt 0) { [math]::Round((($cycleDays | Measure-Object -Average).Average),2) } else { $null }

    $inProduction = @($normalized | Where-Object { ([string]$_.stateName) -ieq 'In Production' })
    $readyForQc = @($active | Where-Object { ([string]$_.stateName) -ieq 'Ready for QC' -or ([string]$_.stateName) -ieq 'QC Received' })
    $redlinesReceived = @($active | Where-Object { ([string]$_.stateName) -ieq 'Redlines Received' -or ([string]$_.stateName) -ieq 'Redlines Issued' })
    $correctionsReceived = @($active | Where-Object { ([string]$_.stateName) -ieq 'Corrections Received' -or ([string]$_.stateName) -ieq 'Corrections In Progress' -or ([string]$_.stateName) -ieq 'Verification In Progress' -or ([string]$_.stateName) -ieq 'Backcheck In Progress' })
    $qcFinalizing = @($active | Where-Object { ([string]$_.stateName) -ieq 'QC Finalizing' })
    $qcComplete = @($active | Where-Object { ($completeStateNames | Where-Object { $_ -ieq [string]$_.stateName }).Count -gt 0 })
    $errorNeedsAttention = @($active | Where-Object { ([string]$_.stateName) -ieq 'Error Needs Attention' }) + @($errors)
    $errorNeedsAttention = @($errorNeedsAttention | Select-Object -Unique)

    $stateCounts = @{}
    if ([bool]$Settings.includeWorkflowStateCounts) {
        foreach ($d in @($normalized)) {
            $k = if (_QCR-IsBlank $d.stateName) { '<none>' } else { [string]$d.stateName }
            if (-not $stateCounts.ContainsKey($k)) { $stateCounts[$k] = 0 }
            $stateCounts[$k]++
        }
    }

    $packageRecords = @()
    $packageMetrics = $null
    if (@($PackageRows).Count -gt 0) {
        $packageRecords = @($PackageRows | ForEach-Object { ConvertTo-QCPackageReportingRecord -Row $_ })
        $packageMetrics = New-QCPackageReportingMetrics -Packages $PackageRows
    }

    return [pscustomobject]@{
        project = $Project
        generatedUtc = Get-QCTimestamp
        source = 'QC_REPORTING_SCAN'
        primaryEntity = if (@($packageRecords).Count -gt 0) { 'sheet_package_id' } else { 'document_guid' }
        metrics = [pscustomobject]@{
            documentCount = @($normalized).Count
            qcActiveCount = @($active).Count
            qcClosedCount = @($closed).Count
            qcErrorCount = @($errors).Count
            staleQcCount = @($stale).Count
            inProductionCount = @($inProduction).Count
            readyForQcCount = @($readyForQc).Count
            redlinesReceivedCount = @($redlinesReceived).Count
            correctionsReceivedCount = @($correctionsReceived).Count
            qcFinalizingCount = @($qcFinalizing).Count
            reviewInProgressCount = 0
            redlinesIssuedCount = @($redlinesReceived).Count
            correctionsInProgressCount = @($correctionsReceived).Count
            verificationInProgressCount = @($correctionsReceived).Count
            qcCompleteCount = @($qcComplete).Count
            errorNeedsAttentionCount = @($errorNeedsAttention).Count
            staleOpenQcCount = @($stale).Count
            avgQcCycleDays = $avg
        }
        workflowStateCounts = [pscustomobject]$stateCounts
        documents = @($normalized)
        packageMetrics = $packageMetrics
        packages = @($packageRecords)
    }
}

function Write-QCReportingSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Snapshot,
        [Parameter(Mandatory)]
        [hashtable]$Settings
    )

    $stamp = Get-QCTimestampShort
    $projectName = _QCR-SafeName ([string]$Snapshot.project)
    $root = [string]$Settings.snapshotRoot
    $dir = Join-Path $root $stamp
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $path = Join-Path $dir ($projectName + '.json')
    $Snapshot | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath $path -Encoding utf8
    return $path
}

function Invoke-QCReportingScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Job,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $settings = Get-QCReportingSettings -Config $Config
    $docs = @(Get-QCReportingDocuments -Job $Job -Settings $settings)
    $project = if ($Job.ContainsKey('project')) { [string]$Job.project } elseif ($Job.ContainsKey('sourceFolder')) { [string]$Job.sourceFolder } else { 'project' }
    $folderPath = if ($Job.ContainsKey('sourceFolder')) { [string]$Job.sourceFolder } else { '' }
    $packageRows = @(Get-QCPackageReportingRows -Config $Config -FolderPath $folderPath)
    $snapshot = New-QCReportingSnapshot -Documents $docs -Settings $settings -Project $project -PackageRows $packageRows
    $path = Write-QCReportingSnapshot -Snapshot $snapshot -Settings $settings
    return New-QCSuccessResult -Code 'QC_REPORTING_SCAN_OK' -Message 'QC reporting snapshot written.' -Data @{
        snapshotPath = $path; snapshot = $snapshot; documentCount = $docs.Count; packageCount = @($packageRows).Count
    }
}

function New-QCReportingScanJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Project,
        [Parameter(Mandatory)]
        [string]$SourceFolder
    )
    $bucket = (Get-QCWallClockNow).ToString('yyyyMMddHH')
    $safe = _QCR-SafeName $Project
    return @{
        id = ('qc_reporting_' + $safe + '_' + $bucket)
        type = 'QC_REPORTING_SCAN'
        sourceFolder = $SourceFolder
        project = $Project
        dedupeKey = ('dq_qc_reporting|' + $SourceFolder + '|' + $bucket).ToLowerInvariant()
        createdUtc = Get-QCTimestamp
    }
}

Export-ModuleMember -Function Get-QCReportingSettings,Get-QCReportingDocuments,ConvertTo-QCReportingDocument,Get-QCPackageReportingRows,ConvertTo-QCPackageReportingRecord,New-QCPackageReportingMetrics,New-QCReportingSnapshot,Write-QCReportingSnapshot,Invoke-QCReportingScan,New-QCReportingScanJob
