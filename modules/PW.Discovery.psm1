# PW.Discovery.psm1
# Responsibility: Read-only ProjectWise watch-path resolution and candidate discovery.

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue

function Ensure-PWDiscoveryModuleLoaded {
    <#
    .SYNOPSIS
    Ensures PW.Discovery and watcher session exports are present after nested Import-Module -Force.
    #>
    [CmdletBinding()]
    param()

    $required = @(
        'Find-PWSheetsFoldersUnderRoot'
        'Get-PWDocumentDescriptionForFolder'
        'Sync-PWAssociatedSheetWorkflowState'
        'Revert-PWAssociatedSheetWorkflowStates'
        'Sync-PWAssociatedSheetMembersToWorkflowState'
        'Sync-PWAssociatedSheetReviewTypeAttributes'
        'Sync-PWSheetIndexOwnership'
        'Get-PWDocumentsInFolder'
        'ConvertTo-PWCmdletFolderPath'
        'Test-PWSheetPdfHasMatchingPair'
        'Test-PWFolderResolvable'
        'Get-PWDocumentWorkflowStateName'
        'Get-PWDocName'
        'Get-PWDocDescription'
        'Get-PWDocLastModifiedUtc'
        'Get-PWDocumentWorkflowStateMapByGuid'
        'New-QCSuccessResult'
        'Get-StatusSetPWFolderState'
        'Invoke-AuditTrailScan'
    )

    $restoreOrder = @(
        'Core.Results.psm1'
        'Core.Runtime.psm1'
        'Core.Hashing.psm1'
        'Core.Database.psm1'
        'QC.StatusSet.psm1'
        'PW.Connection.psm1'
        'PW.AuditPoller.psm1'
        'PW.Discovery.psm1'
    )

    for ($pass = 0; $pass -lt 3; $pass++) {
        $missing = @($required | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
        if ($missing.Count -eq 0) { return $true }

        foreach ($modFile in $restoreOrder) {
            $modPath = Join-Path $PSScriptRoot $modFile
            if (Test-Path -LiteralPath $modPath) {
                Import-Module $modPath -Force -WarningAction SilentlyContinue | Out-Null
            }
        }
    }

    if (Get-Command -Name 'Ensure-QCJsonLogAvailable' -ErrorAction SilentlyContinue) {
        [void](Ensure-QCJsonLogAvailable -ModulesRoot $PSScriptRoot)
    } elseif (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force -WarningAction SilentlyContinue | Out-Null
    }

    $stillMissing = @($required | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
    return ($stillMissing.Count -eq 0)
}

function Resolve-WatchPaths {
    <#
    .SYNOPSIS
    Resolves effective watch paths from configuration.
    .DESCRIPTION
    Expands configured roots/folders into a de-duplicated watch path list.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only discovery operations.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}

function Get-PWTriggerCandidates {
    <#
    .SYNOPSIS
    Retrieves trigger candidate files/events from watch paths.
    .DESCRIPTION
    Performs read-only lookup of candidate documents relevant for trigger evaluation.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .PARAMETER WatchPaths
    Effective watch paths to query.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only ProjectWise/API calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string[]]$WatchPaths
    )
}

function Get-PWCandidateMetadata {
    <#
    .SYNOPSIS
    Enriches a trigger candidate with metadata.
    .DESCRIPTION
    Reads additional metadata fields required for filtering and trigger classification.
    .PARAMETER Candidate
    Candidate object from discovery.
    .PARAMETER Config
    Loaded app configuration hashtable.
    .OUTPUTS
    PSCustomObject with shape: IsSuccess [bool], Code [string], Message [string], Data [object].
    .NOTES
    Side effects: read-only ProjectWise/API calls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Candidate,
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
}


function Get-PWObjectPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Object,
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        if ($null -eq $Object) { return $null }
        if ($Object.PSObject -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    } catch { }
    return $null
}

function Get-PWDocName {
    [CmdletBinding()]
    param([AllowNull()][object]$Doc)

    foreach ($n in @('Name', 'DocumentName')) {
        $v = Get-PWObjectPropertyValue -Object $Doc -Name $n
        if ($v) { return [string]$v }
    }
    return ''
}

function Resolve-PWAuditDocumentName {
    <#
    .SYNOPSIS
    Resolves a ProjectWise document file name for audit trigger processing when pw_itemname is empty.
    .DESCRIPTION
    Bulk workflow attach often leaves o_itemname blank on DOCUMENT_STATE rows. Prefer sheet_index
    (populated by status-set / reconciliation scans), then Get-PWDocumentsByGUIDs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentGuid,
        [string]$DocumentName = '',
        [hashtable]$Config = @{},
        [string]$FolderPath = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($DocumentName)) { return [string]$DocumentName.Trim() }
    $guid = ([string]$DocumentGuid).Trim()
    if ([string]::IsNullOrWhiteSpace($guid)) { return '' }

    if ($Config -and (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) {
        if (Test-QCDatabaseEnabled -Config $Config) {
            try {
                $siRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT TOP 1 document_name, folder_path
FROM sheet_index
WHERE document_guid = @docGuid OR qc_pdf_guid = @docGuid
ORDER BY CASE WHEN document_guid = @docGuid THEN 0 ELSE 1 END
"@ -Parameters @{ docGuid = $guid }
                if ($siRes.IsSuccess -and $siRes.Data.table -and $siRes.Data.table.Rows.Count -gt 0) {
                    $r = $siRes.Data.table.Rows[0]
                    if (-not ($r.document_name -is [DBNull]) -and -not [string]::IsNullOrWhiteSpace([string]$r.document_name)) {
                        return [string]$r.document_name
                    }
                }
            } catch { }
        }
    }

    try {
        $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs @($guid) -ErrorAction SilentlyContinue)
        if ($docs.Count -gt 0) {
            $fromPw = Get-PWDocName -Doc $docs[0]
            if (-not [string]::IsNullOrWhiteSpace($fromPw)) { return [string]$fromPw }
        }
    } catch { }

    return ''
}

function Get-PWDocDescription {
    [CmdletBinding()]
    param([AllowNull()][object]$Doc)

    foreach ($n in @('Description', 'DocumentDescription')) {
        $v = Get-PWObjectPropertyValue -Object $Doc -Name $n
        if ($null -ne $v) { return [string]$v }
    }
    return ''
}

function Get-PWDocumentDescriptionForFolder {
    <#
    .SYNOPSIS
    Reads document Description reliably for audit-trail processing (GUID lookup, then folder search).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = ''
    )

    if ($DocumentGuid) {
        try {
            $byGuid = @(Get-PWDocumentsByGUIDs -DocumentGUIDs @($DocumentGuid) -ErrorAction SilentlyContinue)
            if ($byGuid -and $byGuid.Count -gt 0) {
                $dd = Get-PWDocDescription -Doc $byGuid[0]
                if (-not [string]::IsNullOrWhiteSpace($dd)) { return $dd }
            }
        } catch { }
    }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }

    if (-not (Test-PWFolderResolvable -FolderPath $FolderPath)) { return '' }

    $cmd = Get-Command -Name 'Get-PWDocumentsBySearchWithReturnColumns' -ErrorAction SilentlyContinue
    if ($cmd) {
        try {
            $descParams = @{
                FolderPath     = $apiPath
                JustThisFolder = $true
                DocumentName   = $DocumentName
                ErrorAction    = 'SilentlyContinue'
            }
            if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
                $descParams['ColumnsToReturn'] = @('Description', 'Name', 'DocumentName')
            } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
                $descParams['ReturnColumns'] = @('Description', 'Name', 'DocumentName')
            }
            $row = & $cmd @descParams | Select-Object -First 1
            if ($row) {
                $dd = Get-PWDocDescription -Doc $row
                if (-not [string]::IsNullOrWhiteSpace($dd)) { return $dd }
            }
        } catch { }
    }

    try {
        $doc = Get-PWDocumentsBySearch -FolderPath $apiPath -JustThisFolder -DocumentName $DocumentName -PopulatePath -ErrorAction SilentlyContinue
        if ($doc) {
            $dd = Get-PWDocDescription -Doc $doc
            if (-not [string]::IsNullOrWhiteSpace($dd)) { return $dd }
        }
    } catch { }

    return ''
}

function Get-PWDocLastModifiedUtc {
    [CmdletBinding()]
    param([AllowNull()][object]$Doc)

    foreach ($n in @('FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date')) {
        $v = Get-PWObjectPropertyValue -Object $Doc -Name $n
        if ($v) {
            try {
                if ($v -is [DateTime]) { return ConvertTo-QCTimestamp $v }
                if ($v -is [DateTimeOffset]) { return ConvertTo-QCTimestamp $v.UtcDateTime }
                $dt = [DateTime]::Parse([string]$v, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
                return ConvertTo-QCTimestamp $dt
            } catch {
                try { return ([string]$v).Trim() } catch { }
            }
        }
    }
    return ''
}

function ConvertTo-PWCmdletFolderPath {
    [CmdletBinding()]
    param([AllowNull()][string]$InternalFolderPath)

    $s = ($InternalFolderPath -as [string]).Trim().TrimEnd('\')
    while ($s -match '^(?i)Documents\\') { $s = $s -replace '^(?i)Documents\\', '' }
    return $s
}

function Test-PWFolderResolvable {
    <#
    .SYNOPSIS
    True when Get-PWFolders can resolve the folder (avoids spurious search/state cmdlet errors on bad paths).
    #>
    [CmdletBinding()]
    param([AllowNull()][string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return $false }
    $folderCmd = Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue
    if (-not $folderCmd) { return $true }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }
    try {
        $f = Get-PWFolders -FolderPath $apiPath -JustOne -ErrorAction SilentlyContinue
        return ($null -ne $f)
    } catch {
        return $false
    }
}

function ConvertTo-PWCanonicalDocumentsFolderPath {
    [CmdletBinding()]
    param([AllowNull()][string]$FolderPathProperty)

    $t = ($FolderPathProperty -as [string]).Trim().TrimStart('\').TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    if ($t -match '^(?i)documents\\') { return $t }
    return ('Documents\' + $t)
}

function Get-PWDocumentsInFolder {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FolderPath)

    try {
        $allDocs = @()

        $view = $null
        try {
            $cmd = Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue
            if ($cmd -and $cmd.Parameters.ContainsKey('InputFolder')) {
                $f = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction Stop
                if ($f) { $view = $f | Get-PWFolderView -ErrorAction Stop }
            } else {
                $view = Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop
            }
        } catch { }
        if ($view -and $view.Documents) {
            $allDocs = @($view.Documents)
        } elseif ($view -and $view.Children) {
            $allDocs = @($view.Children | Where-Object { $_.DocumentID -or $_.Name })
        }

        if ($allDocs.Count -eq 0) {
            try {
                $allDocs = @(Get-PWDocumentsBySearch -FolderPath $FolderPath -JustThisFolder -PopulatePath -ErrorAction Stop)
            } catch {
                $cmd = Get-Command -Name Get-PWDocumentsBySearchWithReturnColumns -ErrorAction SilentlyContinue
                if ($cmd) {
                    $returnColsParam = if ($cmd.Parameters.ContainsKey('ReturnColumns')) { 'ReturnColumns' } elseif ($cmd.Parameters.ContainsKey('ColumnsToReturn')) { 'ColumnsToReturn' } else { $null }
                    if ($returnColsParam) {
                        $cols = @('Description', 'Name', 'DocumentID')
                        if ($returnColsParam -eq 'ReturnColumns') {
                            $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -ReturnColumns $cols -ErrorAction SilentlyContinue
                        } else {
                            $withCols = Get-PWDocumentsBySearchWithReturnColumns -FolderPath $FolderPath -JustThisFolder -ColumnsToReturn $cols -ErrorAction SilentlyContinue
                        }
                        if ($withCols) { $allDocs = @($withCols) }
                    }
                }
            }
        }

        return @($allDocs)
    } catch {
        try {
            $folder = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction SilentlyContinue
            if ($folder) {
                try {
                    $cmd = Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue
                    if ($cmd -and $cmd.Parameters.ContainsKey('InputFolder')) {
                        $f2 = Get-PWFolders -FolderPath $FolderPath -JustOne -ErrorAction SilentlyContinue
                        if ($f2) { $view = $f2 | Get-PWFolderView -ErrorAction SilentlyContinue }
                    } else {
                        $view = Get-PWFolderView -FolderPath $FolderPath -ErrorAction SilentlyContinue
                    }
                } catch { }
                if ($view -and $view.Documents) { return @($view.Documents) }
            }
        } catch { }
    }
    return @()
}

function Find-PWSheetsFoldersUnderRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootPath,
        [string]$SheetsSuffix,
        [string]$DatasourceName,
        [int]$ProjectDepth = 1
    )

    $rootPathRaw = ($RootPath -as [string]).Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($rootPathRaw)) { return @() }

    $suffix = ($SheetsSuffix -as [string]).Trim().TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($suffix)) { $suffix = 'CADD\Sheets' }

    $depth = $ProjectDepth
    if ($depth -lt 1) { $depth = 1 }
    if ($depth -gt 5) { $depth = 5 }

    $rootDocs = $rootPathRaw
    if ($rootDocs -notmatch '^(?i)Documents\\') { $rootDocs = ('Documents\' + $rootDocs.TrimStart('\')) }

    function _TestPwFolderExists([string]$DocsFolderPath) {
        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $DocsFolderPath
        if ([string]::IsNullOrWhiteSpace($apiPath)) { return $false }
        foreach ($p in @($apiPath, ('Documents\' + $apiPath))) {
            try {
                $cmd = Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue
                if ($cmd -and $cmd.Parameters.ContainsKey('InputFolder')) {
                    $ff = Get-PWFolders -FolderPath $p -JustOne -WarningAction SilentlyContinue -ErrorAction Stop
                    if (-not $ff) { throw 'missing' }
                    $null = $ff | Get-PWFolderView -WarningAction SilentlyContinue -ErrorAction Stop
                } else {
                    $null = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                }
                return $true
            } catch {
                try {
                    $f = Get-PWFolders -FolderPath $p -JustOne -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    if ($f) { return $true }
                } catch { }
            }
        }
        return $false
    }

    function _GetChildNames([string]$DocsFolderPath) {
        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $DocsFolderPath
        if ([string]::IsNullOrWhiteSpace($apiPath)) { return @() }
        $names = @()

        foreach ($p in @($apiPath, ('Documents\' + $apiPath))) {
            # Prefer ImmediateChildren because Get-PWFolderView can succeed but return empty on some servers/paths.
            try {
                $children = Get-PWFoldersImmediateChildren -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                foreach ($c in @($children)) {
                    if (-not $c) { continue }
                    $name = Get-PWObjectPropertyValue -Object $c -Name 'Name'
                    if (-not $name) {
                        $fp = Get-PWObjectPropertyValue -Object $c -Name 'FolderPath'
                        if ($fp) { $name = [System.IO.Path]::GetFileName(([string]$fp).TrimEnd('\')) }
                    }
                    if ($name) { $names += [string]$name }
                }
            } catch { }
            if ($names.Count -gt 0) { break }

            try {
                $cmd = Get-Command -Name Get-PWFolderView -ErrorAction SilentlyContinue
                if ($cmd -and $cmd.Parameters.ContainsKey('InputFolder')) {
                    $ff2 = Get-PWFolders -FolderPath $p -JustOne -WarningAction SilentlyContinue -ErrorAction Stop
                    if (-not $ff2) { throw 'missing' }
                    $view = $ff2 | Get-PWFolderView -WarningAction SilentlyContinue -ErrorAction Stop
                } else {
                    $view = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                }
                if ($view.Children) {
                    foreach ($c in $view.Children) {
                        $name = Get-PWObjectPropertyValue -Object $c -Name 'Name'
                        if (-not $name) {
                            $fp = Get-PWObjectPropertyValue -Object $c -Name 'FolderPath'
                            if ($fp) { $name = [System.IO.Path]::GetFileName(([string]$fp).TrimEnd('\')) }
                        }
                        if ($name) { $names += [string]$name }
                    }
                }
                if ($names.Count -eq 0 -and $view.Folders) {
                    foreach ($f in $view.Folders) {
                        $name = Get-PWObjectPropertyValue -Object $f -Name 'Name'
                        if ($name) { $names += [string]$name }
                    }
                }
            } catch {
                try {
                    $children = Get-PWFoldersImmediateChildren -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
                    foreach ($c in @($children)) {
                        $name = Get-PWObjectPropertyValue -Object $c -Name 'Name'
                        if (-not $name) {
                            $fp = Get-PWObjectPropertyValue -Object $c -Name 'FolderPath'
                            if ($fp) { $name = [System.IO.Path]::GetFileName(([string]$fp).TrimEnd('\')) }
                        }
                        if ($name) { $names += [string]$name }
                    }
                } catch { }
            }

            if ($names.Count -gt 0) { break }
        }

        return @($names | Where-Object { $_ } | Select-Object -Unique)
    }

    $rootSheets = (($rootDocs.TrimEnd('\') + '\' + $suffix).TrimEnd('\'))
    if (_TestPwFolderExists -DocsFolderPath $rootSheets) {
        $walkDepth = ($depth - 1)
        if ($walkDepth -lt 0) { $walkDepth = 0 }

        $all = @($rootSheets)
        $current = @($rootSheets)
        for ($i = 1; $i -le $walkDepth; $i++) {
            $next = @()
            foreach ($p in @($current)) {
                foreach ($n in @(_GetChildNames -DocsFolderPath $p)) {
                    $next += (($p.TrimEnd('\') + '\' + $n).TrimEnd('\'))
                }
            }
            $current = @($next | Select-Object -Unique)
            if ($current.Count -eq 0) { break }
            $all += $current
        }

        $list = @()
        foreach ($sPath in @($all | Select-Object -Unique)) {
            $list += @{ DatasourceName = $DatasourceName; FolderPath = $sPath; OneLevelDeep = $true }
        }
        return $list
    }

    $levelProjects = @($rootDocs)
    for ($i = 1; $i -le $depth; $i++) {
        $next = @()
        foreach ($p in @($levelProjects)) {
            foreach ($n in @(_GetChildNames -DocsFolderPath $p)) {
                $next += (($p.TrimEnd('\') + '\' + $n).TrimEnd('\'))
            }
        }
        $levelProjects = @($next | Select-Object -Unique)
        if ($levelProjects.Count -eq 0) { break }
    }

    $list = @()
    foreach ($projPath in @($levelProjects)) {
        $folderPath = (($projPath.TrimEnd('\') + '\' + $suffix).TrimEnd('\'))
        $list += @{ DatasourceName = $DatasourceName; FolderPath = $folderPath; OneLevelDeep = $true }
    }
    return $list
}

function _PWD-MergePwAttributeSourceIntoMap {
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [AllowNull()][object]$Source
    )

    if (-not $Source) { return }
    if ($Source -is [System.Collections.IDictionary]) {
        foreach ($k in $Source.Keys) {
            $Map[[string]$k] = [string]$Source[$k]
        }
        return
    }
    foreach ($bag in @($Source)) {
        if ($bag -is [System.Collections.IDictionary]) {
            foreach ($k in $bag.Keys) {
                $Map[[string]$k] = [string]$bag[$k]
            }
        }
    }
}

function Get-PWDocumentAttributeMap {
    <#
    .SYNOPSIS
    Parses attribute bags from a Get-PWDocumentsBySearchWithReturnColumns row into a hashtable.
    Merges .Attributes (sorted-list bags or dictionary), .CustomAttributes, and .EnvironmentAttributes.
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$DocRow)

    $map = @{}
    if (-not $DocRow) { return $map }
    _PWD-MergePwAttributeSourceIntoMap -Map $map -Source $DocRow.Attributes
    try { _PWD-MergePwAttributeSourceIntoMap -Map $map -Source $DocRow.CustomAttributes } catch { }
    try { _PWD-MergePwAttributeSourceIntoMap -Map $map -Source $DocRow.EnvironmentAttributes } catch { }
    return $map
}

function Test-PWValidDocumentGuid {
    [CmdletBinding()]
    param([AllowNull()][string]$DocumentGuid)

    if ([string]::IsNullOrWhiteSpace($DocumentGuid)) { return $false }
    try {
        $parsed = [guid]::Parse([string]$DocumentGuid)
        return $parsed -ne [guid]::Empty
    } catch {
        return $false
    }
}

function _PWD-GetWorkflowStateFromDocumentRow {
    param([AllowNull()][object]$DocRow)

    if (-not $DocRow) { return '' }
    foreach ($prop in @('WorkflowState', 'StateName', 'State', 'DocumentState', 'CurrentState', 'WorkflowStateName')) {
        try {
            if ($DocRow.PSObject.Properties[$prop]) {
                $v = [string]$DocRow.$prop
                if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
            }
        } catch { }
    }
    return ''
}

function Get-PWDocumentWorkflowStateName {
    <#
    .SYNOPSIS
    Reads workflow state for one document. WithReturnColumns often omits state in this environment;
    Get-PWDocumentsByGUIDs and Get-PWDocumentsBySearch -PopulatePath return WorkflowState reliably.
    #>
    [CmdletBinding()]
    param(
        [string]$FolderPath,
        [string]$DocumentName,
        [string]$DocumentGuid
    )

    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
        if ($guidCmd) {
            try {
                $doc = & $guidCmd -DocumentGUIDs @([string]$DocumentGuid) -ErrorAction Stop | Select-Object -First 1
                $state = _PWD-GetWorkflowStateFromDocumentRow -DocRow $doc
                if ($state) { return $state }
            } catch { }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        if (-not (Test-PWFolderResolvable -FolderPath $FolderPath)) { return '' }
        $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
        if ($searchCmd) {
            $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
            if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }
            try {
                $searchParams = @{
                    FolderPath     = $apiPath
                    JustThisFolder = $true
                    DocumentName   = $DocumentName
                    ErrorAction    = 'SilentlyContinue'
                }
                if ($searchCmd.Parameters.ContainsKey('PopulatePath')) { $searchParams['PopulatePath'] = $true }
                $doc = & $searchCmd @searchParams | Select-Object -First 1
                $state = _PWD-GetWorkflowStateFromDocumentRow -DocRow $doc
                if ($state) { return $state }
            } catch { }
        }
    }

    return ''
}

function Get-PWDocumentWorkflowStateMapByGuid {
    <#
    .SYNOPSIS
    Batch workflow-state lookup keyed by document GUID (lowercase).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$DocumentGuids
    )

    $map = @{}
    $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
    if (-not $guidCmd) { return $map }

    $unique = @($DocumentGuids | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    if ($unique.Count -eq 0) { return $map }

    $chunkSize = 200
    for ($i = 0; $i -lt $unique.Count; $i += $chunkSize) {
        $chunk = @($unique[$i..[Math]::Min($i + $chunkSize - 1, $unique.Count - 1)])
        try {
            $docs = @(& $guidCmd -DocumentGUIDs $chunk -ErrorAction Stop)
            foreach ($doc in $docs) {
                $guid = $null
                try { $guid = [string]$doc.DocumentGUID } catch { }
                if ([string]::IsNullOrWhiteSpace($guid)) { continue }
                $state = _PWD-GetWorkflowStateFromDocumentRow -DocRow $doc
                if ($state) { $map[$guid.ToLowerInvariant()] = $state }
            }
        } catch {
            foreach ($oneGuid in $chunk) {
                try {
                    $doc = & $guidCmd -DocumentGUIDs @($oneGuid) -ErrorAction Stop | Select-Object -First 1
                    if (-not $doc) { continue }
                    $guid = [string]$doc.DocumentGUID
                    if ([string]::IsNullOrWhiteSpace($guid)) { continue }
                    $state = _PWD-GetWorkflowStateFromDocumentRow -DocRow $doc
                    if ($state) { $map[$guid.ToLowerInvariant()] = $state }
                } catch { }
            }
        }
    }
    return $map
}

function Get-PWSheetIndexSyncColumnNames {
    <#
    .SYNOPSIS
    PW attribute column names to re-read after DOCUMENT_ATTR audit events (EM_* plus qcWorkflow.attributeMap).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    [void]$set.Add('EM_Designer_Email')
    [void]$set.Add('EM_Reviewer_Email')
    try {
        $na = $Config['notifications']['attributes']
        if ($na) {
            if ($na['designerEmailField']) { [void]$set.Add([string]$na['designerEmailField']) }
            if ($na['reviewerEmailField']) { [void]$set.Add([string]$na['reviewerEmailField']) }
            if ($na['checkerEmailField']) { [void]$set.Add([string]$na['checkerEmailField']) }
        }
    } catch { }
    if (-not ($set | Where-Object { $_ -ieq 'EM_Checker_Email' })) { [void]$set.Add('EM_Checker_Email') }
    try {
        $am = $Config['qcWorkflow']['attributeMap']
        if ($am) {
            foreach ($v in $am.Values) {
                if (-not [string]::IsNullOrWhiteSpace([string]$v)) { [void]$set.Add([string]$v) }
            }
        }
    } catch { }
    return @($set)
}

function _PWD-GetPwAttributeValue {
    param(
        [Parameter(Mandatory)][hashtable]$PwAttributes,
        [AllowNull()][string]$ColumnName
    )
    if ([string]::IsNullOrWhiteSpace($ColumnName)) { return '' }
    if ($PwAttributes.ContainsKey($ColumnName)) { return ([string]$PwAttributes[$ColumnName]).Trim() }
    foreach ($key in $PwAttributes.Keys) {
        if ([string]$key -ieq $ColumnName) { return ([string]$PwAttributes[$key]).Trim() }
    }
    return ''
}

function ConvertTo-SheetIndexFieldValues {
    <#
    .SYNOPSIS
    Maps ProjectWise attribute bags to sheet_index column values (EM_* preferred for role emails).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$PwAttributes,
        [string]$PwStateName = ''
    )

    $wfMap = @{}
    try {
        $am = $Config['qcWorkflow']['attributeMap']
        if ($am) {
            foreach ($k in $am.Keys) { $wfMap[[string]$k] = [string]$am[$k] }
        }
    } catch { }

    $emDesignerCol = 'EM_Designer_Email'
    $emReviewerCol = 'EM_Reviewer_Email'
    $emCheckerCol = 'EM_Checker_Email'
    try {
        $na = $Config['notifications']['attributes']
        if ($na) {
            if ($na['designerEmailField']) { $emDesignerCol = [string]$na['designerEmailField'] }
            if ($na['reviewerEmailField']) { $emReviewerCol = [string]$na['reviewerEmailField'] }
            if ($na['checkerEmailField']) { $emCheckerCol = [string]$na['checkerEmailField'] }
        }
    } catch { }

    $qcDesignerCol = if ($wfMap.ContainsKey('designerEmail')) { $wfMap['designerEmail'] } else { 'QC_Designer_Email' }
    $qcReviewerCol = if ($wfMap.ContainsKey('reviewerEmail')) { $wfMap['reviewerEmail'] } else { 'QC_Reviewer_Email' }
    $qcCheckerCol = if ($wfMap.ContainsKey('checkerEmail')) { $wfMap['checkerEmail'] } else { 'QC_Checker_Email' }
    $qcReviewTypeCol = if ($wfMap.ContainsKey('reviewType')) { $wfMap['reviewType'] } else { 'QC_Review_Type' }
    $qcAssignedCol = if ($wfMap.ContainsKey('assignedTo')) { $wfMap['assignedTo'] } else { 'QC_Assigned_To' }
    $qcStatusCol = if ($wfMap.ContainsKey('status')) { $wfMap['status'] } else { 'QC_Status' }

    $emDesigner = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $emDesignerCol
    $emReviewer = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $emReviewerCol
    $emChecker = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $emCheckerCol
    $qcDesigner = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcDesignerCol
    $qcReviewer = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcReviewerCol
    $qcChecker = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcCheckerCol

    $designer = if ($emDesigner) { $emDesigner } elseif ($qcDesigner) { $qcDesigner } else { '' }
    $reviewer = if ($emReviewer) { $emReviewer } elseif ($qcReviewer) { $qcReviewer } else { '' }
    $checker = if ($emChecker) { $emChecker } elseif ($qcChecker) { $qcChecker } else { '' }

    return @{
        designerEmail  = $designer
        reviewerEmail  = $reviewer
        checkerEmail   = $checker
        qcReviewType   = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcReviewTypeCol
        qcAssignedTo   = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcAssignedCol
        qcStatus       = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcStatusCol
        pwStateName    = if ($PwStateName) { $PwStateName.Trim() } else { '' }
    }
}

function _PWD-EnrichSheetIndexReviewType {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Fields,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Fields.qcReviewType)) { return $Fields }

    $sheetPdfName = [string]$DocumentName
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        if ($sheetPdfName -match '(?i)(?:-qc\.pdf|\.dgn)$') {
            $stem = Get-PWSheetStemFromDocumentName -DocumentName $sheetPdfName
            if (-not [string]::IsNullOrWhiteSpace($stem)) { $sheetPdfName = $stem + '.pdf' }
        }
    }

    if (Get-Command -Name 'Get-PWQcPrependRoleFieldsFromSourcePdf' -ErrorAction SilentlyContinue) {
        $pw = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $sheetPdfName -Config $Config
        if ($pw.found -and -not [string]::IsNullOrWhiteSpace([string]$pw.qcReviewType)) {
            $Fields.qcReviewType = [string]$pw.qcReviewType
        }
    }
    return $Fields
}

function _PWD-ResolveSheetIndexQcReviewType {
    <#
    .SYNOPSIS
    Resolves QC_Review_Type for sheet_index sync: direct PW read, source-PDF fallback, then column-only re-read.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$FieldsFromPwRead,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [bool]$EnrichFromSourcePdf = $true
    )

    $resolved = [string]$FieldsFromPwRead.qcReviewType
    if ($EnrichFromSourcePdf -and [string]::IsNullOrWhiteSpace($resolved)) {
        $enriched = _PWD-EnrichSheetIndexReviewType -Config $Config -Fields $FieldsFromPwRead `
            -FolderPath $FolderPath -DocumentName $DocumentName
        $resolved = [string]$enriched.qcReviewType
    }
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config
        $solo = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $DocumentName `
            -ColumnsToReturn @($reviewCol)
        if ($solo.found) {
            $resolved = _PWD-GetPwAttributeValue -PwAttributes $solo.attributes -ColumnName $reviewCol
        }
    }
    return ([string]$resolved).Trim()
}

function Get-PWDocumentAttributesByColumns {
    <#
    .SYNOPSIS
    Reads configured PW document attributes via search-with-columns API.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string[]]$ColumnsToReturn
    )

    $cmd = Get-Command -Name 'Get-PWDocumentsBySearchWithReturnColumns' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return @{ attributes = @{}; pwStateName = ''; found = $false; error = 'Get-PWDocumentsBySearchWithReturnColumns not available' }
    }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }

    $cols = @($ColumnsToReturn | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($cols.Count -eq 0) {
        return @{ attributes = @{}; pwStateName = ''; found = $false; error = 'No columns requested' }
    }

    $row = $null
    try {
        $params = @{
            FolderPath     = $apiPath
            JustThisFolder = $true
            DocumentName   = $DocumentName
            ErrorAction    = 'Stop'
        }
        if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
            $params['ColumnsToReturn'] = $cols
        } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
            $params['ReturnColumns'] = $cols
        }
        $row = & $cmd @params | Select-Object -First 1
    } catch {
        return @{ attributes = @{}; pwStateName = ''; found = $false; error = $_.Exception.Message }
    }

    if (-not $row) {
        return @{ attributes = @{}; pwStateName = ''; found = $false; error = 'Document not found' }
    }

    $attrs = Get-PWDocumentAttributeMap -DocRow $row
    $pwState = _PWD-GetWorkflowStateFromDocumentRow -DocRow $row
    if ([string]::IsNullOrWhiteSpace($pwState)) {
        $docGuid = ''
        try { $docGuid = [string]$row.DocumentGUID } catch { }
        $pwState = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $docGuid
    }
    return @{
        attributes  = $attrs
        pwStateName = if ($pwState) { $pwState.Trim() } else { '' }
        found       = $true
        document    = $row
        error       = $null
    }
}

function Get-PWDocumentEmailContacts {
    <#
    .SYNOPSIS
    Reads EM_Designer_Email and EM_Reviewer_Email from a document in a folder via search-with-columns API.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DesignerEmailColumn = 'EM_Designer_Email',
        [string]$ReviewerEmailColumn = 'EM_Reviewer_Email'
    )

    $cols = @($DesignerEmailColumn, $ReviewerEmailColumn) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $read = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $DocumentName -ColumnsToReturn $cols
    if (-not $read.found) {
        return @{ designerEmail = ''; reviewerEmail = ''; found = $false; error = $read.error }
    }

    $designer = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $DesignerEmailColumn
    $reviewer = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $ReviewerEmailColumn
    return @{
        designerEmail = $designer
        reviewerEmail = $reviewer
        pwStateName   = [string]$read.pwStateName
        found         = $true
        document      = $read.document
        error         = $null
    }
}

function _PWD-NormalizeSheetIndexValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function _PWD-GetSheetMemberDocRole {
    param([string]$DocumentName)
    $dn = [string]$DocumentName
    if ($dn -match '(?i)-qc\.pdf$') { return 'qcPdf' }
    if ($dn -match '(?i)\.dgn$') { return 'dgn' }
    if ($dn -match '(?i)\.pdf$') { return 'pdf' }
    return 'other'
}

function _PWD-WriteDocumentStateLiveVerificationLog {
    param(
        [Nullable[long]]$AuditEventId = $null,
        [string]$SourceDocumentGuid = '',
        [string]$SourceDocumentName = '',
        [string]$FolderPath = '',
        [string]$CanonicalState = '',
        [string]$CanonicalStateSource = 'liveProjectWise',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [array]$Members = @(),
        [hashtable]$StateByGuid = @{},
        [hashtable]$Config = $null
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }

    $byRole = @{ dgn = $null; pdf = $null; qcPdf = $null }
    $associatedStates = [System.Collections.Generic.List[object]]::new()
    foreach ($member in $Members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dn) { continue }
        $role = _PWD-GetSheetMemberDocRole -DocumentName $dn
        $liveState = ''
        $stateSource = 'liveProjectWise'
        $key = $dg.ToLowerInvariant()
        if ($dg -and $StateByGuid.ContainsKey($key)) {
            $liveState = [string]$StateByGuid[$key]
        } else {
            $liveState = _PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document
            if (-not $liveState) { $stateSource = 'memberDocumentRow' }
        }
        $sheetIndexState = ''
        if ($dg -and $Config -and (Get-Command -Name '_PWD-GetSheetIndexPwStateName' -ErrorAction SilentlyContinue)) {
            $sheetIndexState = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg
        }
        $entry = @{
            role           = $role
            documentGuid   = $dg
            documentName   = $dn
            livePwState    = $liveState
            liveStateSource = $stateSource
            sheetIndexState = $sheetIndexState
        }
        [void]$associatedStates.Add($entry)
        if ($byRole.ContainsKey($role)) { $byRole[$role] = $entry }
    }

    Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_DOCUMENT_STATE_LIVE_VERIFY' `
        -Message 'Live associated document workflow states before DOCUMENT_STATE sibling sync.' -Data @{
        auditEventId = $AuditEventId
        sourceDocumentGuid = $SourceDocumentGuid
        sourceDocumentName = $SourceDocumentName
        folderPath = $FolderPath
        pwStateName = $CanonicalState
        pwStateNameSource = $CanonicalStateSource
        changedByUser = $ChangedByUser
        changedByUsername = $ChangedByUsername
        associatedStates = @($associatedStates)
    } | Out-Null

    $dgnState = if ($byRole.dgn) { [string]$byRole.dgn.livePwState } else { '' }
    $pdfState = if ($byRole.pdf) { [string]$byRole.pdf.livePwState } else { '' }
    $qcState = if ($byRole.qcPdf) { [string]$byRole.qcPdf.livePwState } else { '' }
    $dgnName = if ($byRole.dgn) { [string]$byRole.dgn.documentName } else { '' }
    $pdfName = if ($byRole.pdf) { [string]$byRole.pdf.documentName } else { '' }
    $qcName = if ($byRole.qcPdf) { [string]$byRole.qcPdf.documentName } else { '' }

    $divergences = [System.Collections.Generic.List[string]]::new()
    if ($dgnState -and $pdfState -and ((_PWD-NormalizeSheetIndexValue $dgnState) -ne (_PWD-NormalizeSheetIndexValue $pdfState))) {
        [void]$divergences.Add('dgn_vs_pdf')
    }
    if ($pdfState -and $qcState -and ((_PWD-NormalizeSheetIndexValue $pdfState) -ne (_PWD-NormalizeSheetIndexValue $qcState))) {
        [void]$divergences.Add('pdf_vs_qcPdf')
    }
    if ($divergences.Count -gt 0) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_STATE_DIVERGENCE' `
            -Message 'Associated sheet workflow states differ before sibling sync (expected when QC PDF or a member changes first).' -Data @{
            auditEventId = $AuditEventId
            sourceDocumentGuid = $SourceDocumentGuid
            sourceDocumentName = $SourceDocumentName
            folderPath = $FolderPath
            changedByUser = $ChangedByUser
            changedByUsername = $ChangedByUsername
            divergences = @($divergences)
            dgnDocumentName = $dgnName
            dgnState = $dgnState
            pdfDocumentName = $pdfName
            pdfState = $pdfState
            qcPdfDocumentName = $qcName
            qcPdfState = $qcState
            canonicalState = $CanonicalState
            canonicalStateSource = $CanonicalStateSource
        } | Out-Null
    }
}

function _PWD-GetSheetIndexDocumentNameFromRow {
    param([Parameter(Mandatory)][object]$DocRow)
    $name = $null
    try { $name = [string]$DocRow.Name } catch { }
    if (-not $name) { try { $name = [string]$DocRow.DocumentName } catch { } }
    if (-not $name) { try { $name = [string]$DocRow.FileName } catch { } }
    return $name
}

function Build-PWSheetIndexRowsForPairedSheets {
    <#
    .SYNOPSIS
    Builds sheet_index row hashtables for paired PDF/DGN sheets using EM_* and QC_* PW attributes.
    Used during full-folder reconciliation scans (reconcileEveryNCycles).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$WatchRoot = '',
        [Parameter(Mandatory)][object[]]$PairedSheets,
        [hashtable]$StateByGuid = @{}
    )

    if (-not (Test-QCSheetIndexFolderPath -FolderPath $FolderPath)) { return @() }

    $nameToFields = @{}
    try {
        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
        if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }
        $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
        $searchRows = @(Get-PWDocumentsBySearchWithReturnColumns `
            -FolderPath $apiPath -JustThisFolder `
            -ColumnsToReturn $cols `
            -ErrorAction SilentlyContinue)
        foreach ($sr in $searchRows) {
            $srName = _PWD-GetSheetIndexDocumentNameFromRow -DocRow $sr
            if (-not $srName) { continue }
            $attrs = Get-PWDocumentAttributeMap -DocRow $sr
            $pwState = _PWD-GetWorkflowStateFromDocumentRow -DocRow $sr
            $fields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $attrs -PwStateName $pwState
            $nameToFields[$srName.ToLowerInvariant()] = $fields
        }
    } catch { }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($ps in @($PairedSheets)) {
        try {
            $pdfName = if ($ps.pdf -and $ps.pdf.name) { [string]$ps.pdf.name } else { $null }
            $dgnName = if ($ps.dgn -and $ps.dgn.name) { [string]$ps.dgn.name } else { $null }
            $pdfGuid = if ($ps.pdf -and $ps.pdf.documentGuid) { [string]$ps.pdf.documentGuid } else { $null }
            $dgnGuid = if ($ps.dgn -and $ps.dgn.documentGuid) { [string]$ps.dgn.documentGuid } else { $null }

            foreach ($pair in @(
                @{ name = $pdfName; guid = $pdfGuid; sourceType = 'pdf' }
                @{ name = $dgnName; guid = $dgnGuid; sourceType = 'dgn' }
            )) {
                if (-not $pair.name -or -not $pair.guid) { continue }
                if (-not (Test-PWValidDocumentGuid -DocumentGuid $pair.guid)) { continue }
                $f = if ($nameToFields.ContainsKey($pair.name.ToLowerInvariant())) {
                    $nameToFields[$pair.name.ToLowerInvariant()]
                } else {
                    @{ designerEmail = ''; reviewerEmail = ''; checkerEmail = ''; qcReviewType = ''; qcAssignedTo = ''; qcStatus = ''; pwStateName = '' }
                }
                $state = if ($StateByGuid.ContainsKey($pair.guid.ToLowerInvariant())) {
                    [string]$StateByGuid[$pair.guid.ToLowerInvariant()]
                } elseif ($f.pwStateName) {
                    [string]$f.pwStateName
                } else { '' }
                $rows.Add(@{
                    documentGuid   = $pair.guid
                    documentName   = $pair.name
                    folderPath     = $FolderPath
                    watchRoot      = $WatchRoot
                    sourceType     = $pair.sourceType
                    designerEmail  = [string]$f.designerEmail
                    reviewerEmail  = [string]$f.reviewerEmail
                    checkerEmail   = [string]$f.checkerEmail
                    qcReviewType   = [string]$f.qcReviewType
                    qcAssignedTo   = [string]$f.qcAssignedTo
                    qcStatus       = [string]$f.qcStatus
                    pwStateName    = $state
                }) | Out-Null
            }
        } catch { }
    }
    return @($rows)
}

function Get-PWSheetStemFromDocumentName {
    <#
    .SYNOPSIS
    Normalized sheet stem from a DGN, sheet PDF, or *-qc.pdf filename.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentName
    )
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) { return '' }
    if ($stem -match '(?i)-qc$') { $stem = $stem -replace '(?i)-qc$', '' }
    return $stem
}

function Test-PWSheetPdfHasMatchingPair {
    <#
    .SYNOPSIS
    True when a sheet PDF filename has a matching DGN in the same PW folder (same stem).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [hashtable]$PairCache = $null
    )

    if ($DocumentName -notmatch '(?i)\.pdf$' -or $DocumentName -match '(?i)-qc\.pdf$') { return $false }
    if ($DocumentName -match '(?i)_statusset\.pdf$') { return $false }

    $stem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    if ([string]::IsNullOrWhiteSpace($stem)) { return $false }

    $cacheKey = ($FolderPath.ToLowerInvariant() + '|' + $stem.ToLowerInvariant())
    if ($PairCache -and $PairCache.ContainsKey($cacheKey)) { return [bool]$PairCache[$cacheKey] }

    $dgnName = $stem + '.dgn'
    $found = $false
    $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
    if ($searchCmd) {
        $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
        if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }
        try {
            $params = @{
                FolderPath     = $apiPath
                JustThisFolder = $true
                DocumentName   = $dgnName
                ErrorAction    = 'Stop'
            }
            if ($searchCmd.Parameters.ContainsKey('PopulatePath')) { $params['PopulatePath'] = $true }
            $doc = & $searchCmd @params | Select-Object -First 1
            $found = ($null -ne $doc)
        } catch { }
    }

    if ($PairCache) { $PairCache[$cacheKey] = $found }
    return $found
}

function Get-PWAssociatedSheetDocumentNames {
    <#
    .SYNOPSIS
    Expected sibling filenames (sheet PDF, DGN, QC PDF) for one sheet stem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SheetStem
    )
    if ([string]::IsNullOrWhiteSpace($SheetStem)) { return @() }
    return @(
        ($SheetStem + '.pdf')
        ($SheetStem + '.dgn')
        ($SheetStem + '-qc.pdf')
    )
}

function _PWD-InvokeSetPwDocumentState {
    param(
        [Parameter(Mandatory)][object]$Document,
        [string]$StateName = '',
        [hashtable]$GuardContext = @{}
    )
    if ([string]::IsNullOrWhiteSpace($StateName)) {
        if (-not $GuardContext) { $GuardContext = @{} }
        $guardCallSite = '_PWD-InvokeSetPwDocumentState'
        if ($GuardContext.callSite) { $guardCallSite = [string]$GuardContext.callSite }
        _PWD-WriteEmptyStateGuardLog -CallSite $guardCallSite `
            -AuditEventId $GuardContext.auditEventId -DocumentName ([string]$GuardContext.documentName) `
            -FolderPath ([string]$GuardContext.folderPath) -SourceVariableName ([string]$GuardContext.sourceVariableName) `
            -SourceValue $StateName -LivePwState ([string]$GuardContext.livePwState) `
            -ChangedByUsername ([string]$GuardContext.changedByUsername)
        return $false
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
    $result = $null
    if ($docParam -and $stateParam) { $result = & $cmd @args -ErrorAction Stop }
    elseif ($stateParam) { $result = & $cmd $Document @args -ErrorAction Stop }
    else { $result = & $cmd $Document $StateName -ErrorAction Stop }
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) {
        try {
            if ($result -eq $false) {
                throw "Set-PWDocumentState returned false for state '$StateName'."
            }
        } catch {
            if ($_.Exception.Message -match 'returned false') { throw }
        }
    }
}

function _PWD-LoadPwDocumentsByGuid {
    param([Parameter(Mandatory)][string[]]$DocumentGuids)
    $docByGuid = @{}
    $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
    if (-not $guidCmd) { return $docByGuid }
    $valid = @($DocumentGuids | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
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
                } catch { }
            }
        }
    }
    return $docByGuid
}

function _PWD-ResolvePwDocumentInFolder {
    param(
        [hashtable]$DocByGuid,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = ''
    )
    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $gk = $DocumentGuid.ToLowerInvariant()
        if ($DocByGuid.ContainsKey($gk)) { return $DocByGuid[$gk] }
    }
    $searchCmd = Get-Command -Name 'Get-PWDocumentsBySearch' -ErrorAction SilentlyContinue
    if (-not $searchCmd -or [string]::IsNullOrWhiteSpace($DocumentName)) { return $null }
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
    } catch { return $null }
}

function Get-PWAssociatedSheetMembers {
    <#
    .SYNOPSIS
    Resolves DGN, sheet PDF, and QC PDF siblings for one sheet stem in a folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = ''
    )

    $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    if ([string]::IsNullOrWhiteSpace($sheetStem)) { return @() }

    $expectedNames = @(Get-PWAssociatedSheetDocumentNames -SheetStem $sheetStem)
    $members = @{}
    $guidsToLoad = [System.Collections.Generic.List[string]]::new()
    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $guidsToLoad.Add([string]$DocumentGuid) | Out-Null
    }

    if (Get-Command -Name 'Invoke-QCDatabaseQuery' -ErrorAction SilentlyContinue) {
        if (Test-QCDatabaseEnabled -Config $Config) {
            try {
                $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT document_guid, document_name, source_type, qc_pdf_guid, qc_pdf_name
FROM sheet_index
WHERE folder_path = @folderPath
  AND (
    LOWER(document_name) = LOWER(@pdfName)
    OR LOWER(document_name) = LOWER(@dgnName)
    OR LOWER(document_name) = LOWER(@qcName)
  )
"@ -Parameters @{
                    folderPath = $FolderPath
                    pdfName    = $expectedNames[0]
                    dgnName    = $expectedNames[1]
                    qcName     = $expectedNames[2]
                }
                if ($dbRes.IsSuccess -and $dbRes.Data.table) {
                    foreach ($row in @($dbRes.Data.table.Rows)) {
                        $dg = if ($row.document_guid -is [DBNull]) { '' } else { [string]$row.document_guid }
                        $dn = if ($row.document_name -is [DBNull]) { '' } else { [string]$row.document_name }
                        if (-not $dn) { continue }
                        $key = $dn.ToLowerInvariant()
                        if (-not $members.ContainsKey($key)) {
                            $members[$key] = @{
                                documentGuid = $dg
                                documentName = $dn
                                sourceType   = if ($row.source_type -is [DBNull]) { '' } else { [string]$row.source_type }
                            }
                        }
                        if ($dg) { $guidsToLoad.Add($dg) | Out-Null }
                        if ($row.Table.Columns.Contains('qc_pdf_guid') -and -not ($row.qc_pdf_guid -is [DBNull])) {
                            $qcg = [string]$row.qc_pdf_guid
                            if ($qcg) { $guidsToLoad.Add($qcg) | Out-Null }
                            $qcn = if ($row.qc_pdf_name -is [DBNull]) { '' } else { [string]$row.qc_pdf_name }
                            if ($qcn -and -not $members.ContainsKey($qcn.ToLowerInvariant())) {
                                $members[$qcn.ToLowerInvariant()] = @{
                                    documentGuid = $qcg
                                    documentName = $qcn
                                    sourceType   = 'pdf'
                                }
                            }
                        }
                    }
                }
            } catch { }
        }
    }

    foreach ($name in $expectedNames) {
        $key = $name.ToLowerInvariant()
        if (-not $members.ContainsKey($key)) {
            $members[$key] = @{
                documentGuid = ''
                documentName = $name
                sourceType   = if ($name -match '(?i)\.dgn$') { 'dgn' } elseif ($name -match '(?i)-qc\.pdf$') { 'pdf' } else { 'pdf' }
            }
        }
    }

    $docByGuid = _PWD-LoadPwDocumentsByGuid -DocumentGuids @($guidsToLoad)
    $resolved = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $members.Values) {
        $dn = [string]$entry.documentName
        $dg = [string]$entry.documentGuid
        $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid $docByGuid -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        if (-not $doc) { continue }
        try {
            if (-not $dg) { $dg = [string]$doc.DocumentGUID }
        } catch { }
        $resolved.Add(@{
            documentGuid = $dg
            documentName = $dn
            document     = $doc
            sourceType   = [string]$entry.sourceType
        }) | Out-Null
    }
    return @($resolved)
}

function Sync-PWAssociatedSheetMembersToWorkflowState {
    <#
    .SYNOPSIS
    Sets the same workflow state on all associated DGN, sheet PDF, and QC PDF siblings in a folder.
  .DESCRIPTION
    Used after QC_PREPEND writeback so every sheet member reaches the configured post-prepend state
    (e.g. Ready for QC). Does not enqueue jobs or fire audit notifications on siblings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [Parameter(Mandatory)][string]$TargetStateName,
        [bool]$DryRun = $false,
        [string]$TriggerSource = 'prepend_writeback'
    )

    $target = ([string]$TargetStateName).Trim()
    if ([string]::IsNullOrWhiteSpace($target)) { return @{ updates = @(); memberCount = 0 } }

    $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    if ($members.Count -eq 0) { return @{ updates = @(); memberCount = 0 } }

    $guids = @($members | ForEach-Object { [string]$_.documentGuid } | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ })
    $stateByGuid = @{}
    if ($guids.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids } catch { }
    }

    $stateUpdates = [System.Collections.Generic.List[object]]::new()
    $dbEnabled = $false
    if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
        $dbEnabled = Test-QCDatabaseEnabled -Config $Config
    }

    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dg -and -not $member.document) { continue }

        $currentState = if ($dg -and $stateByGuid.ContainsKey($dg.ToLowerInvariant())) {
            [string]$stateByGuid[$dg.ToLowerInvariant()]
        } else {
            _PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document
        }
        if ([string]::IsNullOrWhiteSpace($currentState) -and $dg) {
            $currentState = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromState    = [string]$currentState
            toState      = $target
            applied      = $false
            planned      = $false
            verified     = $false
        }

        if ((_PWD-NormalizeSheetIndexValue $currentState) -eq (_PWD-NormalizeSheetIndexValue $target)) {
            $change.skipped = 'already_at_target'
            if ($dbEnabled -and $dg) {
                try { [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $target) } catch { }
            }
            $stateUpdates.Add($change) | Out-Null
            continue
        }

        if ($DryRun) {
            $change.planned = $true
            $stateUpdates.Add($change) | Out-Null
            continue
        }

        if (-not (_PWD-TestStateNameNotEmpty -StateName $target -CallSite 'Sync-PWAssociatedSheetReviewTypeAttributes.target' `
                -DocumentName $dn -FolderPath $FolderPath -SourceVariableName 'target' -LivePwState ([string]$currentState) `
                -ChangedByUsername '')) {
            $change.skipped = 'empty_target_state'
            $stateUpdates.Add($change) | Out-Null
            continue
        }
        try {
            _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $target -GuardContext @{
                callSite = 'Sync-PWAssociatedSheetReviewTypeAttributes.target'; documentName = $dn; folderPath = $FolderPath
                sourceVariableName = 'target'; livePwState = [string]$currentState
            }
            $change.applied = $true
            $verifiedState = if ($dg) {
                Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
            } else {
                Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $dn -DocumentGuid ''
            }
            if (-not [string]::IsNullOrWhiteSpace($verifiedState)) {
                $change.verified = (_PWD-NormalizeSheetIndexValue $verifiedState) -eq (_PWD-NormalizeSheetIndexValue $target)
                $change.stateAfter = [string]$verifiedState
            }
            if ($change.applied -and $change.verified -eq $false -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_SHEET_STATE_VERIFY_MISMATCH' `
                    -Message 'Set-PWDocumentState completed but PW state does not match target after read-back.' -Data @{
                    documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                    targetState = $target; stateAfter = [string]$change.stateAfter; triggerSource = $TriggerSource
                } | Out-Null
            }
            if ($dbEnabled -and $dg) {
                try { [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $target) } catch { }
            }
        } catch {
            $change.error = [string]$_.Exception.Message
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_SHEET_STATE_SYNC_FAILED' -Message 'Failed to set associated sheet workflow state.' -Data @{
                    documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                    fromState = [string]$currentState; toState = $target; error = [string]$_.Exception.Message
                    triggerSource = $TriggerSource
                } | Out-Null
            }
        }
        $stateUpdates.Add($change) | Out-Null
    }

    $result = @{ updates = @($stateUpdates); memberCount = $members.Count; targetState = $target; triggerSource = $TriggerSource }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_SHEET_STATE_SYNC' -Message 'Associated sheet workflow states aligned to target.' -Data @{
            folderPath          = $FolderPath
            triggerDocumentName = $DocumentName
            triggerDocumentGuid = $DocumentGuid
            targetState         = $target
            memberCount         = $members.Count
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
            triggerSource       = $TriggerSource
        } | Out-Null
    }
    return $result
}

function _PWD-GetSheetIndexQcReviewType {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid
    )
    if ([string]::IsNullOrWhiteSpace($DocumentGuid)) { return '' }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return '' }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return '' }
    try {
        $siRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT qc_review_type FROM sheet_index WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if ($siRes.IsSuccess -and $siRes.Data.table -and $siRes.Data.table.Rows.Count -gt 0) {
            $r = $siRes.Data.table.Rows[0]
            if (-not ($r.qc_review_type -is [DBNull])) { return [string]$r.qc_review_type }
        }
    } catch { }
    return ''
}

function Sync-PWAssociatedSheetReviewTypeAttributes {
    <#
    .SYNOPSIS
    Propagates QC_Review_Type from a DOCUMENT_ATTR change to associated DGN, sheet PDF, and QC PDF siblings.
    .DESCRIPTION
    Updates ProjectWise environment attributes and sheet_index.qc_review_type for every associated member
    when the canonical review type differs. Mirrors sibling workflow-state sync for user-owned QC attributes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$CanonicalReviewType,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [bool]$DryRun = $false
    )

    if ([string]::IsNullOrWhiteSpace($CanonicalReviewType)) { return }

    $canonical = ([string]$CanonicalReviewType).Trim()
    if ([string]::IsNullOrWhiteSpace($canonical)) { return }

    $pwWritesEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath
    $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config
    $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    if ($members.Count -eq 0) { return }

    $stateUpdates = @()
    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        $doc = $member.document
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }

        $prevDb = _PWD-GetSheetIndexQcReviewType -Config $Config -DocumentGuid $dg
        $currentPw = ''
        $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $dn -ColumnsToReturn @($reviewCol)
        if ($attrRead.found) {
            $currentPw = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $reviewCol
            if (-not $doc -and $attrRead.document) { $doc = $attrRead.document }
        }
        if (-not $doc) {
            $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }
        if (-not $doc) { continue }

        $pwNeedsWrite = (_PWD-NormalizeSheetIndexValue $currentPw) -ne (_PWD-NormalizeSheetIndexValue $canonical)
        $indexNeedsWrite = (_PWD-NormalizeSheetIndexValue $prevDb) -ne (_PWD-NormalizeSheetIndexValue $canonical)
        if (-not $pwNeedsWrite -and -not $indexNeedsWrite) { continue }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromValue    = [string]$currentPw
            toValue      = $canonical
            applied      = $false
            planned      = $false
            indexUpdated = $false
        }

        if ($pwNeedsWrite) {
            if (-not $pwWritesEnabled) {
                $change.pwWriteSkipped = 'qc_review_type_not_enabled_for_environment'
            } elseif ($DryRun) {
                $change.planned = $true
            } else {
                try {
                    [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes @{ $reviewCol = $canonical })
                    $change.applied = $true
                } catch {
                    $change.error = [string]$_.Exception.Message
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_REVIEW_TYPE_SYNC_FAILED' `
                            -Message 'Failed to align associated sheet QC_Review_Type.' -Data @{
                            documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                            fromValue = [string]$currentPw; toValue = $canonical; error = [string]$_.Exception.Message
                        }
                    }
                    continue
                }
            }
        }

        if ($indexNeedsWrite -and (Get-Command -Name 'Write-QCSheetIndex' -ErrorAction SilentlyContinue)) {
            if (-not $DryRun) {
                try {
                    $ext = [System.IO.Path]::GetExtension($dn)
                    if ($ext) { $ext = $ext.ToLowerInvariant() }
                    $sourceType = $null
                    if ($ext -eq '.pdf') { $sourceType = 'pdf' }
                    elseif ($ext -eq '.dgn') { $sourceType = 'dgn' }
                    Write-QCSheetIndex -Config $Config -DocumentGuid $dg -DocumentName $dn -FolderPath $FolderPath `
                        -WatchRoot $WatchRoot -Extension $ext -SourceType $sourceType -QcReviewType $canonical `
                        -LastAuditEventAt $LastAuditEventAt -SetOwnershipFromProjectWise | Out-Null
                    $change.indexUpdated = $true
                } catch { }
            }
        }

        if (Get-Command -Name 'Invoke-QCAuditWorkflowAttributeChangeTriggers' -ErrorAction SilentlyContinue) {
            if ($pwNeedsWrite -or $indexNeedsWrite) {
                $prevAudit = if ($currentPw) { [string]$currentPw } else { [string]$prevDb }
                Invoke-QCAuditWorkflowAttributeChangeTriggers -Config $Config -DocumentGuid $dg -DocumentName $dn `
                    -FolderPath $FolderPath -FieldChanges @{
                    qc_review_type = @{ oldValue = $prevAudit; newValue = $canonical }
                } | Out-Null
            }
        }

        $stateUpdates += $change
    }

    if ($stateUpdates.Count -gt 0 -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_REVIEW_TYPE_SYNC' `
            -Message 'Associated sheet QC_Review_Type aligned from DOCUMENT_ATTR audit event.' -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalReviewType = $canonical
            reviewTypeColumn    = $reviewCol
            pwWritesEnabled     = $pwWritesEnabled
            memberCount         = $members.Count
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
        }
    }
}



function _PWD-WriteEmptyStateGuardLog {
    param(
        [string]$CallSite = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$SourceVariableName = '',
        [AllowNull()][object]$SourceValue = $null,
        [string]$LivePwState = '',
        [string]$ChangedByUsername = ''
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_EMPTY_STATE_GUARDED' `
        -Message 'Skipped workflow state call because StateName source value was empty.' -Data @{
        callSite = $CallSite
        auditEventId = $AuditEventId
        documentName = $DocumentName
        folderPath = $FolderPath
        sourceVariableName = $SourceVariableName
        sourceValue = if ($null -eq $SourceValue) { $null } else { [string]$SourceValue }
        livePwState = $LivePwState
        changedByUsername = $ChangedByUsername
    } | Out-Null
}

function _PWD-TestStateNameNotEmpty {
    param(
        [string]$StateName = '',
        [string]$CallSite = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$SourceVariableName = '',
        [string]$LivePwState = '',
        [string]$ChangedByUsername = ''
    )
    if (-not [string]::IsNullOrWhiteSpace($StateName)) { return $true }
    _PWD-WriteEmptyStateGuardLog -CallSite $CallSite -AuditEventId $AuditEventId `
        -DocumentName $DocumentName -FolderPath $FolderPath -SourceVariableName $SourceVariableName `
        -SourceValue $StateName -LivePwState $LivePwState -ChangedByUsername $ChangedByUsername
    return $false
}

function _PWD-GetConfiguredWorkflowStateNames {
    param([hashtable]$Config)
    $states = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    $add = {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return }
        $n = ([string]$Name).Trim()
        $k = $n.ToLowerInvariant()
        if ($seen.ContainsKey($k)) { return }
        $seen[$k] = $true
        [void]$states.Add($n)
    }
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try {
            $wf = Get-QCWorkflowSettings -Config $Config
            if ($wf -and (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
                foreach ($key in @('production','qcInitiated','qcReceived','readyForQc','redlinesReceived','correctionsReceived','correctionsInProgress','qcFinalizing','complete','error')) {
                    try { & $add (Get-QCWorkflowStateName -Settings $wf -StateKey $key) } catch { }
                }
            }
        } catch { }
    }
    foreach ($fallback in @('In Production','QC Initiated','Ready for QC','Redlines Received','Corrections Received','Corrections In Progress','QC Finalizing','QC Complete','Error Needs Attention')) {
        & $add $fallback
    }
    return @($states)
}

function _PWD-ResolveAuditWorkflowTargetStateName {
    param(
        [hashtable]$Config,
        [string]$AuditTargetStateName = '',
        [string]$AuditRawItemDesc = '',
        [string]$AuditRawTextParam = ''
    )
    $knownStates = @(_PWD-GetConfiguredWorkflowStateNames -Config $Config)
    foreach ($raw in @($AuditTargetStateName, $AuditRawTextParam, $AuditRawItemDesc)) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $value = ([string]$raw).Trim()
        foreach ($state in $knownStates) {
            if ([string]::IsNullOrWhiteSpace($state)) { continue }
            if ($value.Equals([string]$state, [System.StringComparison]::OrdinalIgnoreCase)) { return [string]$state }
        }
        foreach ($state in $knownStates) {
            if ([string]::IsNullOrWhiteSpace($state)) { continue }
            $escaped = [regex]::Escape(([string]$state).Trim())
            if ($escaped -and $value -match ('(?i)(^|[^A-Za-z0-9])' + $escaped + '([^A-Za-z0-9]|$)')) { return [string]$state }
        }
    }
    return ''
}

function _PWD-GetSheetIndexStateSnapshot {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid
    )
    $snapshot = @{ pwStateName = ''; lastAuditEventAt = $null }
    if ([string]::IsNullOrWhiteSpace($DocumentGuid)) { return $snapshot }
    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return $snapshot }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $snapshot }
    try {
        $siRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT pw_state_name, last_audit_event_at FROM sheet_index WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if ($siRes.IsSuccess -and $siRes.Data.table -and $siRes.Data.table.Rows.Count -gt 0) {
            $r = $siRes.Data.table.Rows[0]
            if (-not ($r.pw_state_name -is [DBNull])) { $snapshot.pwStateName = [string]$r.pw_state_name }
            if (-not ($r.last_audit_event_at -is [DBNull])) { $snapshot.lastAuditEventAt = $r.last_audit_event_at }
        }
    } catch { }
    return $snapshot
}

function _PWD-GetSheetIndexPwStateName {
    param(
        [hashtable]$Config,
        [string]$DocumentGuid
    )
    $snapshot = _PWD-GetSheetIndexStateSnapshot -Config $Config -DocumentGuid $DocumentGuid
    return [string]$snapshot.pwStateName
}

function _PWD-TryParseAuditDateTime {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try { return [DateTimeOffset]::Parse([string]$Value) } catch { }
    try { return [DateTimeOffset]::ParseExact(([string]$Value).Trim().TrimEnd('Z'), 'yyyy-MM-dd HH:mm:ss', $null, [Globalization.DateTimeStyles]::AssumeUniversal) } catch { }
    return $null
}

function _PWD-TestShouldBlockStaleRestartOverwrite {
    param(
        [hashtable]$Config,
        [string]$CurrentState = '',
        [string]$TargetState = '',
        [hashtable]$SheetIndexSnapshot = $null,
        [string]$SourceAuditEventAt = ''
    )
    if (-not (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue)) { return $false }
    if ([string]::IsNullOrWhiteSpace($TargetState)) { return $false }
    if (Test-QCWorkflowStateIsQcInitiated -StateName $TargetState -Config $Config) { return $false }
    if (-not $SheetIndexSnapshot) { return $false }
    $snapshotState = [string]$SheetIndexSnapshot.pwStateName
    if ([string]::IsNullOrWhiteSpace($snapshotState)) { return $false }
    if (-not (Test-QCWorkflowStateIsQcInitiated -StateName $snapshotState -Config $Config)) { return $false }
    if (Get-Command -Name 'Test-QCWorkflowStateIsRestartIntakeTransition' -ErrorAction SilentlyContinue) {
        if (-not (Test-QCWorkflowStateIsRestartIntakeTransition -Config $Config -PreviousState $TargetState -CurrentState $snapshotState)) { return $false }
    }
    $sourceAt = _PWD-TryParseAuditDateTime -Value $SourceAuditEventAt
    $targetAt = $null
    try {
        if ($SheetIndexSnapshot.lastAuditEventAt) { $targetAt = [DateTimeOffset]$SheetIndexSnapshot.lastAuditEventAt }
    } catch { $targetAt = _PWD-TryParseAuditDateTime -Value ([string]$SheetIndexSnapshot.lastAuditEventAt) }
    if ($sourceAt -and $targetAt) { return ($sourceAt -le $targetAt) }
    return $true
}

function _PWD-InvokeStaleSheetIndexAuditStateTriggers {
    <#
    .SYNOPSIS
    Records transition_events / qc_workflow_events when PW already shows the new state (member loop skips aligned rows).
    .DESCRIPTION
    Audit DOCUMENT_STATE means the trigger document is already at the target state in PW. Sibling sync only
    invokes audit triggers for members that still need Set-PWDocumentState. When the user moves *-qc.pdf (or PW
    pre-aligns all members), aligned rows are skipped and telemetry never runs. Compare sheet_index pw_state_name
    to the canonical state and invoke audit triggers for every stale member before the sync loop updates the index.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Members,
        [Parameter(Mandatory)][hashtable]$StateByGuid,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$CanonicalState,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null
    )

    if (-not (Get-Command -Name 'Get-QCAuditWorkflowTriggerSettings' -ErrorAction SilentlyContinue)) { return }
    if (-not (Get-Command -Name 'Invoke-QCAuditWorkflowStateChangeTriggers' -ErrorAction SilentlyContinue)) { return }

    $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$wt.enabled) { return }
    if (-not [bool]$wt.recordStateHistory -and -not [bool]$wt.recordTransitions -and -not [bool]$wt.notifyOnStateChange) { return }

    $canonical = _PWD-NormalizeSheetIndexValue $CanonicalState
    if ([string]::IsNullOrWhiteSpace($canonical)) { return }

    foreach ($member in $Members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dg) { continue }

        $prevDb = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg
        if ((_PWD-NormalizeSheetIndexValue $prevDb) -eq $canonical) { continue }

        $pwCurrent = ''
        $key = $dg.ToLowerInvariant()
        if ($StateByGuid.ContainsKey($key)) {
            $pwCurrent = [string]$StateByGuid[$key]
        } else {
            $pwCurrent = _PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document
        }
        if ((_PWD-NormalizeSheetIndexValue $pwCurrent) -ne $canonical) { continue }

        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STALE_INDEX_STATE_TRIGGER' `
                -Message 'Firing workflow triggers for member already at canonical PW state but stale sheet_index.' -Data @{
                auditEventId = $AuditEventId
                documentGuid = $dg
                documentName = $dn
                folderPath = $FolderPath
                previousState = $prevDb
                previousStateSource = 'sheet_index'
                currentState = $CanonicalState
                currentStateSource = 'liveProjectWise'
                changedByUser = $ChangedByUser
                changedByUsername = $ChangedByUsername
            } | Out-Null
        }
        Invoke-QCAuditWorkflowStateChangeTriggers -Config $Config -DocumentGuid $dg -DocumentName $dn `
            -FolderPath $FolderPath -PreviousState $prevDb -CurrentState $CanonicalState -Document $member.document `
            -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId `
            -PreviousStateSource 'sheet_index' -CurrentStateSource 'liveProjectWise' `
            -StaleCheckMembers $Members -StaleCheckStateByGuid $StateByGuid -StaleCheckCanonicalState $CanonicalState | Out-Null
    }
}

function _PWD-GetPrependEnqueueStateTransitionKey {
    param(
        [Nullable[long]]$AuditEventId = $null,
        [string]$LastAuditEventAt = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$TriggerDocumentGuid = '',
        [string]$SheetStem = '',
        [string]$PreviousSheetState = '',
        [string]$TargetStateName = '',
        [string]$PrependTrigger = ''
    )
    if (Get-Command -Name 'Get-QCPrependStateTransitionDedupeKey' -ErrorAction SilentlyContinue) {
        return Get-QCPrependStateTransitionDedupeKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
            -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid `
            -SheetStem $SheetStem -PreviousSheetState $PreviousSheetState -TargetStateName $TargetStateName `
            -PrependTrigger $PrependTrigger
    }
    if (Get-Command -Name '_QCP-GetSheetPrependStateTransitionKey' -ErrorAction SilentlyContinue) {
        return _QCP-GetSheetPrependStateTransitionKey -SheetStem $SheetStem -FromState $PreviousSheetState -ToState $TargetStateName
    }
    return $null
}

function _PWD-EnqueuePrependJobsFromAssociatedQcPdfState {
    <#
    .SYNOPSIS
    Enqueues QC_PREPEND once per sheet workflow transition after sibling state sync.
    .DESCRIPTION
    QC Initiated uses the canonical synced workflow state (first prepend often runs before *-qc.pdf exists).
    QC Finalizing prefers *-qc.pdf workflow state when that file exists. Dedupe uses the audit event id when
    available (see Get-QCPrependStateTransitionDedupeKey). Sync-PWAssociatedSheetWorkflowState only calls this
    when at least one sibling state was applied (echo audits with empty updates are skipped).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][array]$Members,
        [Parameter(Mandatory)][string]$CanonicalState,
        [string]$PreviousSheetState = '',
        [string]$TriggerDocumentGuid = '',
        [string]$TriggerDocumentName = '',
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [hashtable]$StateByGuid = $null
    )

    if (-not (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) { return }
    $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $TriggerDocumentName
    if ([string]::IsNullOrWhiteSpace($sheetStem)) { return }

    $sheetPdfName = $sheetStem + '.pdf'
    $sheetPdfGuid = ''
    $qcPdfGuid = ''
    $qcPdfName = ''
    foreach ($member in $Members) {
        $dn = [string]$member.documentName
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$')) {
            $sheetPdfGuid = [string]$member.documentGuid
        } elseif ($dn -match '(?i)-qc\.pdf$') {
            $qcPdfGuid = [string]$member.documentGuid
            $qcPdfName = $dn
        }
    }

    $qcState = ''
    if (-not [string]::IsNullOrWhiteSpace($qcPdfName)) {
        if ($StateByGuid -and $qcPdfGuid -and $StateByGuid.ContainsKey($qcPdfGuid.ToLowerInvariant())) {
            $qcState = [string]$StateByGuid[$qcPdfGuid.ToLowerInvariant()]
        }
        if ([string]::IsNullOrWhiteSpace($qcState) -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
            try {
                $qcState = [string](Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $qcPdfName -DocumentGuid $qcPdfGuid)
            } catch { }
        }
    }
    $hasQcPdf = -not [string]::IsNullOrWhiteSpace($qcPdfName)
    $canonical = ([string]$CanonicalState).Trim()
    if ($canonical.Length -eq 0) { return }

    # Finalizing: *-qc.pdf is the history target; use its state when the file exists.
    $finalizingState = $canonical
    if ($hasQcPdf) {
        if ([string]::IsNullOrWhiteSpace($qcState)) { $qcState = $canonical }
        $finalizingState = ([string]$qcState).Trim()
        if ($finalizingState.Length -eq 0) { $finalizingState = $canonical }
    }

    $prependGuid = if (-not [string]::IsNullOrWhiteSpace($qcPdfGuid)) { $qcPdfGuid } else { $TriggerDocumentGuid }
    if ([string]::IsNullOrWhiteSpace($prependGuid)) { $prependGuid = $sheetPdfGuid }

    if (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue) {
        if (Test-QCWorkflowStateIsQcInitiated -StateName $canonical -Config $Config) {
            if (-not (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue)) {
                try {
                    $notifPath = Join-Path $PSScriptRoot 'QC.Notifications.psm1'
                    Import-Module $notifPath -ErrorAction SilentlyContinue
                    if (-not (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue)) {
                        Import-Module $notifPath -Force -ErrorAction SilentlyContinue
                    }
                } catch { }
            }
            if (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue) {
                $emailGate = Test-QCPrependBlockedByMissingEmailAttributes -Config $Config -FolderPath $FolderPath `
                    -SheetPdfName $sheetPdfName -DocumentGuid $sheetPdfGuid
                if ($emailGate -and [bool]$emailGate.blocked) {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_PREPEND_SKIPPED_MISSING_EMAIL' `
                            -Message 'QC_PREPEND skipped after sheet sync: required notification email attributes are missing.' -Data @{
                            folderPath = $FolderPath; sheetPdf = $sheetPdfName; sheetStem = $sheetStem
                            missingFields = @($emailGate.missingFields); postPrependState = [string]$emailGate.postPrependState
                            canonicalState = $canonical
                        } | Out-Null
                    }
                    return
                }
            }
            $stKey = _PWD-GetPrependEnqueueStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
                -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid `
                -SheetStem $sheetStem -PreviousSheetState $PreviousSheetState -TargetStateName $canonical `
                -PrependTrigger 'initialQcPdf'
            if (Get-Command -Name 'Add-QCPrependJobForQcInitiatedStateChange' -ErrorAction SilentlyContinue) {
                try {
                    Add-QCPrependJobForQcInitiatedStateChange -Config $Config `
                        -TriggerDocumentGuid $prependGuid -TriggerDocumentName $sheetPdfName -FolderPath $FolderPath `
                        -CurrentStateName $canonical -DryRun:$DryRun -ChangedByUser $ChangedByUser `
                        -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt `
                        -AuditEventId $AuditEventId -StateTransitionKey $stKey | Out-Null
                } catch {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_PREPEND_STATE_ENQUEUE_ERROR' -Message $_.Exception.Message -Data @{
                            folderPath = $FolderPath; sheetPdf = $sheetPdfName; qcPdf = $qcPdfName
                            canonicalState = $canonical; hasQcPdf = $hasQcPdf; stateTransitionKey = $stKey
                        } | Out-Null
                    }
                }
            }
            return
        }
    }

    if (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue) {
        if (Test-QCWorkflowStateIsQcFinalizing -StateName $finalizingState -Config $Config) {
            $stKeyFinal = _PWD-GetPrependEnqueueStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
                -ChangedByUser $ChangedByUser -TriggerDocumentGuid $TriggerDocumentGuid `
                -SheetStem $sheetStem -PreviousSheetState $PreviousSheetState -TargetStateName $finalizingState `
                -PrependTrigger 'finalQcComplete'
            if (Get-Command -Name 'Add-QCPrependJobForQcFinalizingStateChange' -ErrorAction SilentlyContinue) {
                try {
                    Add-QCPrependJobForQcFinalizingStateChange -Config $Config `
                        -TriggerDocumentGuid $prependGuid -TriggerDocumentName $sheetPdfName -FolderPath $FolderPath `
                        -CurrentStateName $finalizingState -DryRun:$DryRun -ChangedByUser $ChangedByUser `
                        -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt `
                        -AuditEventId $AuditEventId -StateTransitionKey $stKeyFinal | Out-Null
                } catch {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_PREPEND_FINAL_ENQUEUE_ERROR' -Message $_.Exception.Message -Data @{
                            folderPath = $FolderPath; sheetPdf = $sheetPdfName; qcPdf = $qcPdfName
                            finalizingState = $finalizingState; hasQcPdf = $hasQcPdf; stateTransitionKey = $stKeyFinal
                        } | Out-Null
                    }
                }
            }
        }
    }
}

function Revert-PWAssociatedSheetWorkflowStates {
    <#
    .SYNOPSIS
    Restores prior workflow states for sheet members that were advanced to a blocked target state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Members,
        [hashtable]$StateByGuid = @{},
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$TargetStateName,
        [string]$FallbackPreviousState = '',
        [string[]]$AdditionalStatesToRevert = @(),
        [bool]$DryRun = $false
    )

    $target = ([string]$TargetStateName).Trim()
    $reverts = [System.Collections.Generic.List[object]]::new()
    $dbEnabled = $false
    if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
        $dbEnabled = Test-QCDatabaseEnabled -Config $Config
    }

    $statesToRevert = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not [string]::IsNullOrWhiteSpace($target)) { [void]$statesToRevert.Add($target) }
    foreach ($extra in @($AdditionalStatesToRevert)) {
        if (-not [string]::IsNullOrWhiteSpace($extra)) { [void]$statesToRevert.Add([string]$extra.Trim()) }
    }

    $fallbackPrevious = ([string]$FallbackPreviousState).Trim()
    if ([string]::IsNullOrWhiteSpace($fallbackPrevious) -and (Get-Command -Name 'Resolve-QCWorkflowRollbackPreviousState' -ErrorAction SilentlyContinue)) {
        try {
            $fallbackPrevious = [string](Resolve-QCWorkflowRollbackPreviousState -Config $Config -TargetStateName $target -Members $Members)
        } catch { }
    }

    foreach ($member in $Members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dg) { continue }

        $previous = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg
        $previousNorm = _PWD-NormalizeSheetIndexValue $previous
        $fallbackNorm = _PWD-NormalizeSheetIndexValue $fallbackPrevious
        if ([string]::IsNullOrWhiteSpace($previousNorm) -or $statesToRevert.Contains($previous)) {
            $previous = $fallbackPrevious
            $previousNorm = $fallbackNorm
        }
        if ([string]::IsNullOrWhiteSpace($previousNorm)) { continue }

        $current = if ($StateByGuid.ContainsKey($dg.ToLowerInvariant())) {
            [string]$StateByGuid[$dg.ToLowerInvariant()]
        } else {
            _PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document
        }
        if ([string]::IsNullOrWhiteSpace($current)) {
            $current = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }
        $currentNorm = _PWD-NormalizeSheetIndexValue $current
        if ([string]::IsNullOrWhiteSpace($currentNorm)) { continue }

        $shouldRevert = $false
        foreach ($revertState in @($statesToRevert)) {
            if (_PWD-NormalizeSheetIndexValue $revertState -eq $currentNorm) {
                $shouldRevert = $true
                break
            }
        }
        if (-not $shouldRevert) { continue }
        if ($currentNorm -eq $previousNorm) { continue }

        $doc = $member.document
        if (-not $doc) {
            try {
                $loaded = _PWD-LoadPwDocumentsByGuid -DocumentGuids @($dg)
                if ($loaded -and $loaded.ContainsKey($dg.ToLowerInvariant())) {
                    $doc = $loaded[$dg.ToLowerInvariant()]
                }
            } catch { }
        }
        if (-not $doc) {
            try { $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg } catch { }
        }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromState    = [string]$current
            toState      = [string]$previous
            applied      = $false
            planned      = $false
        }

        if ($DryRun) {
            $change.planned = $true
        } elseif (-not $doc) {
            $change.error = 'ProjectWise document object unavailable for state rollback.'
        } elseif (-not (_PWD-TestStateNameNotEmpty -StateName $previous -CallSite 'Invoke-QCWorkflowStateEmailAttributeGate.rollback.previous' `
                -AuditEventId $AuditEventId -DocumentName $dn -FolderPath $FolderPath -SourceVariableName 'previous' `
                -LivePwState ([string]$current) -ChangedByUsername $ChangedByUsername)) {
            $change.skipped = 'empty_previous_state'
        } else {
            try {
                _PWD-InvokeSetPwDocumentState -Document $doc -StateName $previous -GuardContext @{
                    callSite = 'Invoke-QCWorkflowStateEmailAttributeGate.rollback.previous'; auditEventId = $AuditEventId
                    documentName = $dn; folderPath = $FolderPath; sourceVariableName = 'previous'
                    livePwState = [string]$current; changedByUsername = $ChangedByUsername
                }
                $change.applied = $true
                if ($dbEnabled) {
                    try { [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $previous) } catch { }
                }
            } catch {
                $change.error = [string]$_.Exception.Message
            }
        }
        $reverts.Add($change) | Out-Null
    }

    return @{
        targetState = $target
        fallbackPreviousState = $fallbackPrevious
        statesToRevert = @($statesToRevert)
        reverts     = @($reverts)
        dryRun      = [bool]$DryRun
    }
}

function Sync-PWAssociatedSheetWorkflowState {
    <#
    .SYNOPSIS
    Propagates a manual DOCUMENT_STATE change to associated DGN, sheet PDF, and QC PDF siblings.
    .DESCRIPTION
    The audit event document is the source of truth for the new workflow state. Associated files
    in the same folder (same sheet stem) are updated via Set-PWDocumentState when they differ.
    sheet_index pw_state_name is updated for every member that was aligned.
    QC_PREPEND is enqueued once after sync (Initiated from canonical state; Finalizing from *-qc.pdf when present).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [string]$AuditTargetStateName = '',
        [string]$AuditRawItemDesc = '',
        [string]$AuditRawTextParam = ''
    )

    if (Get-Command -Name 'Test-QCShouldSuppressAuditSheetStateSync' -ErrorAction SilentlyContinue) {
        if (Test-QCShouldSuppressAuditSheetStateSync -Config $Config -DocumentName $DocumentName `
                -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_STATE_SYNC_SKIPPED_AUTOMATION' `
                    -Message 'Skipped sibling sheet state sync for automation-originated DOCUMENT_STATE echo.' -Data @{
                    documentGuid = $DocumentGuid; documentName = $DocumentName; folderPath = $FolderPath
                    changedByUser = $ChangedByUser; changedByUsername = $ChangedByUsername
                } | Out-Null
            }
            return
        }
    }

    $sourcePwState = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid
    if ([string]::IsNullOrWhiteSpace($sourcePwState)) { return }
    $auditTargetState = _PWD-ResolveAuditWorkflowTargetStateName -Config $Config `
        -AuditTargetStateName $AuditTargetStateName -AuditRawItemDesc $AuditRawItemDesc -AuditRawTextParam $AuditRawTextParam
    # DOCUMENT_STATE audit rows are sparse: the parsed audit state is only a diagnostic hint.
    # Live ProjectWise state remains authoritative for sibling sync and workflow actions.
    $canonicalState = $sourcePwState
    $actorIsAutomation = $false
    if (Get-Command -Name 'Test-QCIsAutomationPwActor' -ErrorAction SilentlyContinue) {
        try { $actorIsAutomation = Test-QCIsAutomationPwActor -Config $Config -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername } catch { $actorIsAutomation = $false }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        if ([string]::IsNullOrWhiteSpace($auditTargetState)) {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATE_TARGET_UNAVAILABLE' `
                -Message 'DOCUMENT_STATE audit row did not include a usable target state; using live ProjectWise state.' -Data @{
                auditEventId = $AuditEventId
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
                changedByUser = $ChangedByUser
                changedByUsername = $ChangedByUsername
                auditTimestamp = $LastAuditEventAt
                rawItemDesc = $AuditRawItemDesc
                rawTextParam = $AuditRawTextParam
                liveWorkflowState = $sourcePwState
            } | Out-Null
        } else {
            Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATE_TARGET_HINT' `
                -Message 'DOCUMENT_STATE audit row included a target-state hint; live ProjectWise state remains authoritative.' -Data @{
                auditEventId = $AuditEventId
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
                changedByUser = $ChangedByUser
                changedByUsername = $ChangedByUsername
                auditTimestamp = $LastAuditEventAt
                rawItemDesc = $AuditRawItemDesc
                rawTextParam = $AuditRawTextParam
                auditTargetStateHint = $auditTargetState
                liveWorkflowState = $sourcePwState
            } | Out-Null
        }
    }

    $triggerSheetIndexState = ''
    if (Get-Command -Name '_PWD-GetSheetIndexPwStateName' -ErrorAction SilentlyContinue) {
        $triggerSheetIndexState = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $DocumentGuid
    }
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATE_SYNC_SOURCE' `
            -Message 'DOCUMENT_STATE sibling sync source state resolved from live ProjectWise state.' -Data @{
            auditEventId = $AuditEventId
            sourceDocumentGuid = $DocumentGuid
            sourceDocumentName = $DocumentName
            folderPath = $FolderPath
            pwStateName = $sourcePwState
            pwStateNameSource = 'liveProjectWise'
            sourceCurrentPwState = $sourcePwState
            sheetIndexState = $triggerSheetIndexState
            sheetIndexStateSource = 'sheet_index'
            auditTargetState = $auditTargetState
            canonicalState = $canonicalState
            canonicalStateSource = 'liveProjectWise'
            sourceAuditActionTime = $LastAuditEventAt
            changedByUser = $ChangedByUser
            changedByUsername = $ChangedByUsername
            actorIsAutomation = $actorIsAutomation
            rawItemDesc = $AuditRawItemDesc
            rawTextParam = $AuditRawTextParam
        } | Out-Null
    }

    if (-not (Get-Command -Name 'Add-QCRenditionJobForReadyForQcStateChange' -ErrorAction SilentlyContinue)) {
        try {
            $rendPath = Join-Path $PSScriptRoot 'QC.Rendition.psm1'
            Import-Module $rendPath -ErrorAction SilentlyContinue
            if (-not (Get-Command -Name 'Add-QCRenditionJobForReadyForQcStateChange' -ErrorAction SilentlyContinue)) {
                Import-Module $rendPath -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        [void](Ensure-PWDiscoveryModuleLoaded)
    }

    if (-not (Get-Command -Name 'Add-QCPrependJobForQcInitiatedStateChange' -ErrorAction SilentlyContinue)) {
        try {
            $procPath = Join-Path $PSScriptRoot 'QC.Processors.psm1'
            Import-Module $procPath -Force -ErrorAction SilentlyContinue
        } catch { }
    }

    if (([string]$DocumentName -match '(?i)\.dgn$') -and (Get-Command -Name 'Add-QCRenditionJobForReadyForQcStateChange' -ErrorAction SilentlyContinue)) {
        try {
            Add-QCRenditionJobForReadyForQcStateChange -Config $Config `
                -TriggerDocumentGuid $DocumentGuid `
                -TriggerDocumentName $DocumentName `
                -FolderPath $FolderPath `
                -CurrentStateName $canonicalState `
                -DryRun:$DryRun `
                -ChangedByUser $ChangedByUser `
                -ChangedByUsername $ChangedByUsername | Out-Null
        } catch {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_RENDITION_STATE_ENQUEUE_ERROR' -Message $_.Exception.Message -Data @{
                    documentGuid = $DocumentGuid; documentName = $DocumentName; folderPath = $FolderPath
                    changedByUser = $ChangedByUser; changedByUsername = $ChangedByUsername
                    currentState = $canonicalState
                } | Out-Null
            }
        }
    }

    $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    if ($members.Count -eq 0) { return }

    $previousSheetState = ''
    $sheetStemForPrepend = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStemForPrepend = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }
    if (-not [string]::IsNullOrWhiteSpace($sheetStemForPrepend)) {
        foreach ($member in $members) {
            $dn = [string]$member.documentName
            if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-qc\.pdf$')) {
                $dg = [string]$member.documentGuid
                if ($dg) {
                    $previousSheetState = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg
                }
                break
            }
        }
    }

    $guids = @($members | ForEach-Object { [string]$_.documentGuid } | Where-Object { Test-PWValidDocumentGuid -DocumentGuid $_ })
    $stateByGuid = @{}
    if ($guids.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids } catch { }
    }

    _PWD-WriteDocumentStateLiveVerificationLog -AuditEventId $AuditEventId `
        -SourceDocumentGuid $DocumentGuid -SourceDocumentName $DocumentName -FolderPath $FolderPath `
        -CanonicalState $canonicalState -CanonicalStateSource 'liveProjectWise' `
        -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
        -Members $members -StateByGuid $stateByGuid -Config $Config

    if ($guids.Count -gt 0) {
        try {
            $liveRefresh = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids
            if ($liveRefresh) {
                foreach ($k in @($liveRefresh.Keys)) { $stateByGuid[$k] = $liveRefresh[$k] }
            }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Level 'Information' -Code 'WATCH_SHEET_STATE_LIVE_REFRESH' `
                    -Message 'Re-read live ProjectWise workflow states before DOCUMENT_STATE recency evaluation.' -Data @{
                    auditEventId = $AuditEventId
                    sourceDocumentGuid = $DocumentGuid
                    sourceDocumentName = $DocumentName
                    folderPath = $FolderPath
                    memberCount = $members.Count
                    stateCount = if ($liveRefresh) { @($liveRefresh.Keys).Count } else { 0 }
                    sourceAuditActionTime = $LastAuditEventAt
                    changedByUser = $ChangedByUser
                    changedByUsername = $ChangedByUsername
                } | Out-Null
            }
        } catch {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_STATE_LIVE_REFRESH_FAILED' `
                    -Message 'Failed to re-read live ProjectWise workflow states before DOCUMENT_STATE recency evaluation.' -Data @{
                    auditEventId = $AuditEventId
                    sourceDocumentGuid = $DocumentGuid
                    sourceDocumentName = $DocumentName
                    folderPath = $FolderPath
                    error = [string]$_.Exception.Message
                } | Out-Null
            }
        }
    }

    if (Get-Command -Name 'Test-QCDocumentStateAuditEventIsStale' -ErrorAction SilentlyContinue) {
        $staleDecision = Test-QCDocumentStateAuditEventIsStale -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -CanonicalState $canonicalState `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -Members $members -StateByGuid $stateByGuid -SheetStem $sheetStemForPrepend
        if ($staleDecision -and [bool]$staleDecision.isStale) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                $staleLog = @{
                    decision = 'skipped'
                    sync = 'skipped'
                    notify = 'skipped'
                    canonicalState = $canonicalState
                    canonicalStateSource = 'liveProjectWise'
                }
                foreach ($k in @($staleDecision.Keys)) { $staleLog[$k] = $staleDecision[$k] }
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_STATE_SYNC_SKIPPED_STALE_EVENT' `
                    -Message 'Skipped DOCUMENT_STATE sibling sync for superseded audit event.' -Data $staleLog | Out-Null
            }
            return
        }
    }

    if (-not (Get-Command -Name 'Invoke-QCWorkflowStateEmailAttributeGate' -ErrorAction SilentlyContinue)) {
        try {
            $notifPath = Join-Path $PSScriptRoot 'QC.Notifications.psm1'
            Import-Module $notifPath -ErrorAction SilentlyContinue
            if (-not (Get-Command -Name 'Invoke-QCWorkflowStateEmailAttributeGate' -ErrorAction SilentlyContinue)) {
                Import-Module $notifPath -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
    if (Get-Command -Name 'Invoke-QCWorkflowStateEmailAttributeGate' -ErrorAction SilentlyContinue) {
        $gate = Invoke-QCWorkflowStateEmailAttributeGate -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -TargetStateName $canonicalState `
            -Members $members -StateByGuid $stateByGuid -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -DryRun:$DryRun
        if ($gate -and $gate.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_STATE_SYNC_BLOCKED' `
                    -Message 'Sheet state sync stopped: required email attributes missing; states reverted.' -Data @{
                    triggerDocumentGuid = $DocumentGuid
                    triggerDocumentName = $DocumentName
                    folderPath          = $FolderPath
                    canonicalState      = $canonicalState
                    missingFields       = @($gate.missingFields)
                    rollback            = $gate.rollback
                } | Out-Null
            }
            return
        }
    }

    $stateUpdates = @()
    $dbEnabled = $false
    if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
        $dbEnabled = Test-QCDatabaseEnabled -Config $Config
    }

    _PWD-InvokeStaleSheetIndexAuditStateTriggers -Config $Config -Members $members -StateByGuid $stateByGuid `
        -FolderPath $FolderPath -CanonicalState $canonicalState -DryRun:$DryRun -ChangedByUser $ChangedByUser `
        -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId

    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (-not $dg) { continue }

        $currentState = if ($stateByGuid.ContainsKey($dg.ToLowerInvariant())) {
            [string]$stateByGuid[$dg.ToLowerInvariant()]
        } else {
            _PWD-GetWorkflowStateFromDocumentRow -DocRow $member.document
        }
        $memberSheetIndexSnapshot = _PWD-GetSheetIndexStateSnapshot -Config $Config -DocumentGuid $dg
        $regressionBlocked = _PWD-TestShouldBlockStaleRestartOverwrite -Config $Config -CurrentState ([string]$currentState) `
            -TargetState ([string]$canonicalState) -SheetIndexSnapshot $memberSheetIndexSnapshot `
            -SourceAuditEventAt $LastAuditEventAt
        if ($regressionBlocked) {
            $restartStateToKeep = [string]$memberSheetIndexSnapshot.pwStateName
            $targetLastAuditEventAtForLog = $null
            if ($memberSheetIndexSnapshot.lastAuditEventAt) { $targetLastAuditEventAtForLog = [string]$memberSheetIndexSnapshot.lastAuditEventAt }
            $restored = $false
            $restoreError = ''
            if ((-not $DryRun) -and ((_PWD-NormalizeSheetIndexValue $currentState) -ne (_PWD-NormalizeSheetIndexValue $restartStateToKeep))) {
                if (-not (_PWD-TestStateNameNotEmpty -StateName $restartStateToKeep -CallSite 'Sync-PWAssociatedSheetWorkflowState.restartStateToKeep' `
                        -AuditEventId $AuditEventId -DocumentName $dn -FolderPath $FolderPath -SourceVariableName 'restartStateToKeep' `
                        -LivePwState ([string]$currentState) -ChangedByUsername $ChangedByUsername)) {
                    $restoreError = 'Restart state was empty; restore skipped.'
                } else {
                try {
                    _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $restartStateToKeep -GuardContext @{
                        callSite = 'Sync-PWAssociatedSheetWorkflowState.restartStateToKeep'; auditEventId = $AuditEventId
                        documentName = $dn; folderPath = $FolderPath; sourceVariableName = 'restartStateToKeep'
                        livePwState = [string]$currentState; changedByUsername = $ChangedByUsername
                    }
                    $restored = $true
                } catch {
                    $restoreError = [string]$_.Exception.Message
                }
                }
            }
            if ($dbEnabled -and (-not $DryRun)) {
                try { [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $restartStateToKeep) } catch { }
            }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                $level = if ($restoreError) { 'Warning' } else { 'Warning' }
                Write-QCJsonLog -Flush -Level $level -Code 'WATCH_AUDIT_STATE_REGRESSION_BLOCKED' `
                    -Message 'Blocked stale DOCUMENT_STATE sync from regressing a newer manual QC Initiated restart.' -Data @{
                    sourceAuditId = $AuditEventId
                    sourceDocumentGuid = $DocumentGuid
                    sourceDocumentName = $DocumentName
                    sourceCurrentPwState = $sourcePwState
                    sourceAuditActionTime = $LastAuditEventAt
                    sourceActorUserNo = $ChangedByUser
                    sourceActorUsername = $ChangedByUsername
                    sourceActorIsAutomation = $actorIsAutomation
                    targetDocumentGuid = $dg
                    targetDocumentName = $dn
                    targetPreviousPwState = [string]$currentState
                    targetNewPwState = $restartStateToKeep
                    targetLastAuditEventAt = $targetLastAuditEventAtForLog
                    targetChangeIsAutomationOriginated = $actorIsAutomation
                    reason = 'newer_manual_qc_initiated_restart'
                    restored = $restored
                    dryRun = [bool]$DryRun
                    error = $restoreError
                } | Out-Null
            }
            continue
        }
        if ((_PWD-NormalizeSheetIndexValue $currentState) -eq (_PWD-NormalizeSheetIndexValue $canonicalState)) {
            if ($dbEnabled) {
                try {
                    [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $canonicalState)
                    if ($LastAuditEventAt) {
                        Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index SET last_audit_event_at = @lastAudit, last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $dg; lastAudit = $LastAuditEventAt } | Out-Null
                    }
                } catch { }
            }
            continue
        }

        if (Get-Command -Name 'Invoke-QCAuditWorkflowStateChangeTriggers' -ErrorAction SilentlyContinue) {
            Invoke-QCAuditWorkflowStateChangeTriggers -Config $Config -DocumentGuid $dg -DocumentName $dn `
                -FolderPath $FolderPath -PreviousState ([string]$currentState) -CurrentState ([string]$canonicalState) `
                -Document $member.document -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' `
                -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
                -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId `
                -PreviousStateSource 'liveProjectWise' -CurrentStateSource 'canonicalLiveTrigger' `
                -StaleCheckMembers $members -StaleCheckStateByGuid $stateByGuid -StaleCheckSheetStem $sheetStemForPrepend `
                -StaleCheckCanonicalState $canonicalState | Out-Null
        }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromState    = [string]$currentState
            toState      = [string]$canonicalState
            applied      = $false
            planned      = $false
        }

        if ($DryRun) {
            $change.planned = $true
        } elseif (-not (_PWD-TestStateNameNotEmpty -StateName $canonicalState -CallSite 'Sync-PWAssociatedSheetWorkflowState.canonicalState' `
                -AuditEventId $AuditEventId -DocumentName $dn -FolderPath $FolderPath -SourceVariableName 'canonicalState' `
                -LivePwState ([string]$currentState) -ChangedByUsername $ChangedByUsername)) {
            $change.skipped = 'empty_canonical_state'
            $stateUpdates += $change
            continue
        } else {
            try {
                _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $canonicalState -GuardContext @{
                    callSite = 'Sync-PWAssociatedSheetWorkflowState.canonicalState'; auditEventId = $AuditEventId
                    documentName = $dn; folderPath = $FolderPath; sourceVariableName = 'canonicalState'
                    livePwState = [string]$currentState; changedByUsername = $ChangedByUsername
                }
                $change.applied = $true
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_STATE_SYNC_CHANGE' `
                        -Message 'Sibling workflow state changed from DOCUMENT_STATE audit event.' -Data @{
                        sourceAuditId = $AuditEventId
                        sourceDocument = $DocumentName
                        sourceDocumentGuid = $DocumentGuid
                        sourceDocumentName = $DocumentName
                        sourceState = [string]$canonicalState
                        sourceCurrentPwState = $sourcePwState
                        sourceAuditActionTime = $LastAuditEventAt
                        sourceActorUserNo = $ChangedByUser
                        sourceActorUsername = $ChangedByUsername
                        actor = $ChangedByUsername
                        sourceActorIsAutomation = $actorIsAutomation
                        targetDocument = $dn
                        targetDocumentGuid = $dg
                        targetDocumentName = $dn
                        targetState = [string]$canonicalState
                        previousTargetState = [string]$currentState
                        targetPreviousPwState = [string]$currentState
                        newTargetState = [string]$canonicalState
                        targetNewPwState = [string]$canonicalState
                        changedByUser = $ChangedByUser
                        changedByUsername = $ChangedByUsername
                        targetChangeIsAutomationOriginated = $actorIsAutomation
                        reason = 'document_state_sibling_sync'
                        dryRun = [bool]$DryRun
                    } | Out-Null
                }
                if (Get-Command -Name 'Write-QCStateChangeJobTelemetry' -ErrorAction SilentlyContinue) {
                    $recordJobs = $true
                    if (Get-Command -Name 'Get-QCAuditWorkflowTriggerSettings' -ErrorAction SilentlyContinue) {
                        try {
                            $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
                            $recordJobs = [bool]$wt.recordProcessingJobs
                        } catch { }
                    }
                    if ($recordJobs) {
                        Write-QCStateChangeJobTelemetry -Config $Config `
                            -PreviousState ([string]$currentState) -CurrentState ([string]$canonicalState) `
                            -DocumentGuid $dg -DocumentName $dn -SourceFolder $FolderPath `
                            -TriggerSource 'audit_sheet_sync' -Operation 'sheet_state_sync' | Out-Null
                    }
                }
            } catch {
                $change.error = [string]$_.Exception.Message
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_STATE_SYNC_FAILED' -Message 'Failed to align associated sheet workflow state.' -Data @{
                        documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                        fromState = [string]$currentState; toState = [string]$canonicalState; error = [string]$_.Exception.Message
                    }
                }
                continue
            }
        }

        if ($dbEnabled -and (-not $DryRun)) {
            try {
                [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $canonicalState)
                if ($LastAuditEventAt) {
                    Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index SET last_audit_event_at = @lastAudit, last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $dg; lastAudit = $LastAuditEventAt } | Out-Null
                }
            } catch { }
        }

        $stateUpdates += $change
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_STATE_SYNC' -Message 'Associated sheet workflow states aligned from DOCUMENT_STATE audit event.' -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalState      = $canonicalState
            memberCount         = $members.Count
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
        }
    }

    $qcInitiated = $false
    if ((-not [string]::IsNullOrWhiteSpace($canonicalState)) -and (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue)) {
        $qcInitiated = Test-QCWorkflowStateIsQcInitiated -StateName $canonicalState -Config $Config
    } elseif ([string]::IsNullOrWhiteSpace($canonicalState)) {
        _PWD-WriteEmptyStateGuardLog -CallSite 'Sync-PWAssociatedSheetWorkflowState.qcInitiatedTest' `
            -AuditEventId $AuditEventId -DocumentName $DocumentName -FolderPath $FolderPath `
            -SourceVariableName 'canonicalState' -SourceValue $canonicalState -LivePwState $sourcePwState `
            -ChangedByUsername $ChangedByUsername
    }
    if ($qcInitiated -and $dbEnabled -and -not $DryRun -and (Get-Command -Name 'Sync-PWSheetIndexOwnership' -ErrorAction SilentlyContinue)) {
        foreach ($member in $members) {
            $dg = [string]$member.documentGuid
            $dn = [string]$member.documentName
            if (-not $dg -or -not $dn) { continue }
            try {
                Sync-PWSheetIndexOwnership -Config $Config -DocumentGuid $dg -DocumentName $dn `
                    -FolderPath $FolderPath -IsSheetsFolder:$true -WatchRoot $WatchRoot `
                    -LastAuditEventAt $LastAuditEventAt -AuditActionName 'DOCUMENT_ATTR' `
                    -ChangedByUser $ChangedByUser -SkipQcInitiatedFallback
            } catch {
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_INDEX_QC_ATTR_SYNC_ERROR' `
                        -Message $_.Exception.Message -Data @{
                        documentGuid = $dg; documentName = $dn; folderPath = $FolderPath; canonicalState = $canonicalState
                    } | Out-Null
                }
            }
        }
    }

    if (-not $DryRun) {
        try {
            $refreshed = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids
            if ($refreshed) {
                foreach ($k in @($refreshed.Keys)) { $stateByGuid[$k] = $refreshed[$k] }
            }
        } catch { }
    }

    $appliedStateUpdates = @($stateUpdates | Where-Object { $_ -and $_.applied -eq $true })
    if ($appliedStateUpdates.Count -eq 0) {
        $allowPrependForRestartEcho = $false
        if ($qcInitiated -and (Get-Command -Name 'Test-QCWorkflowStateIsRestartIntakeTransition' -ErrorAction SilentlyContinue)) {
            $allowPrependForRestartEcho = Test-QCWorkflowStateIsRestartIntakeTransition -Config $Config `
                -PreviousState $previousSheetState -CurrentState $canonicalState
        }
        if ($allowPrependForRestartEcho) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PREPEND_RESTART_NO_STATE_CHANGES' `
                    -Message 'QC_PREPEND evaluation allowed for QC Initiated restart even though sibling states were already aligned.' -Data @{
                    triggerDocumentGuid = $DocumentGuid
                    triggerDocumentName = $DocumentName
                    folderPath          = $FolderPath
                    previousSheetState  = $previousSheetState
                    canonicalState      = $canonicalState
                    memberCount         = $members.Count
                    auditEventId        = $AuditEventId
                    reason              = 'qc_initiated_restart_intake'
                } | Out-Null
            }
        } else {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PREPEND_SKIPPED_NO_SHEET_STATE_CHANGES' `
                    -Message 'QC_PREPEND skipped (QC Initiated or QC Finalizing): sibling sync made no state changes (echo audit).' -Data @{
                    triggerDocumentGuid = $DocumentGuid
                    triggerDocumentName = $DocumentName
                    folderPath          = $FolderPath
                    canonicalState      = $canonicalState
                    memberCount         = $members.Count
                    auditEventId        = $AuditEventId
                } | Out-Null
            }
            return
        }
    }

    _PWD-EnqueuePrependJobsFromAssociatedQcPdfState -Config $Config -FolderPath $FolderPath -Members $members `
        -CanonicalState $canonicalState -PreviousSheetState $previousSheetState `
        -TriggerDocumentGuid $DocumentGuid -TriggerDocumentName $DocumentName -DryRun:$DryRun `
        -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
        -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId -StateByGuid $stateByGuid
}

function _PWD-TryTriggerQcInitiatedFromAssociatedSheetPdf {
    <#
    .SYNOPSIS
    When DOCUMENT_STATE is backlog-starved, sibling DOCUMENT_ATTR events may still run.
    If the associated sheet PDF is already QC Initiated in PW, run the normal state-sync + prepend path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [bool]$IsSheetsFolder = $false,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    if (-not $IsSheetsFolder) { return }
    if (-not (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue)) { return }
    if (-not (Get-Command -Name 'Sync-PWAssociatedSheetWorkflowState' -ErrorAction SilentlyContinue)) { return }
    if (-not (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) { return }

    $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    if ([string]::IsNullOrWhiteSpace($sheetStem)) { return }
    $sheetPdfName = $sheetStem + '.pdf'

    $sheetGuid = ''
    if (($DocumentName -match '(?i)\.pdf$') -and ($DocumentName -notmatch '(?i)-qc\.pdf$')) {
        $sheetGuid = [string]$DocumentGuid
    } else {
        try {
            $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
                -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
            foreach ($member in $members) {
                $mn = [string]$member.documentName
                if (($mn -match '(?i)\.pdf$') -and ($mn -notmatch '(?i)-qc\.pdf$')) {
                    $sheetGuid = [string]$member.documentGuid
                    break
                }
            }
        } catch { return }
    }

    if (-not (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) { return }
    $sheetState = ''
    try {
        $sheetState = [string](Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $sheetPdfName -DocumentGuid $sheetGuid)
    } catch { return }
    if ([string]::IsNullOrWhiteSpace($sheetState)) {
        _PWD-WriteEmptyStateGuardLog -CallSite '_PWD-TryTriggerQcInitiatedFromAssociatedSheetPdf.sheetState' `
            -AuditEventId $null -DocumentName $sheetPdfName -FolderPath $FolderPath `
            -SourceVariableName 'sheetState' -SourceValue $sheetState -LivePwState $sheetState `
            -ChangedByUsername $ChangedByUsername
        return
    }
    if (-not (Test-QCWorkflowStateIsQcInitiated -StateName $sheetState -Config $Config)) { return }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_QC_INITIATED_SHEET_PDF_FALLBACK' `
            -Message 'Associated sheet PDF is QC Initiated; running state sync/prepend fallback from sibling audit event.' -Data @{
            triggerDocumentGuid = $DocumentGuid; triggerDocumentName = $DocumentName; folderPath = $FolderPath
            sheetPdfName = $sheetPdfName; sheetPdfGuid = $sheetGuid; sheetPdfState = $sheetState
        } | Out-Null
    }

    try {
        Sync-PWAssociatedSheetWorkflowState -Config $Config -DocumentGuid $sheetGuid -DocumentName $sheetPdfName `
            -FolderPath $FolderPath -WatchRoot $WatchRoot -LastAuditEventAt $LastAuditEventAt `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername
    } catch {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_QC_INITIATED_SHEET_PDF_FALLBACK_ERROR' `
                -Message $_.Exception.Message -Data @{
                triggerDocumentGuid = $DocumentGuid; triggerDocumentName = $DocumentName; folderPath = $FolderPath
                sheetPdfName = $sheetPdfName; sheetPdfGuid = $sheetGuid
            } | Out-Null
        }
    }
}

function Sync-PWSheetIndexOwnership {
    <#
    .SYNOPSIS
    Re-reads ProjectWise attributes into sheet_index for audit events on watchlist documents.
    .DESCRIPTION
    DOCUMENT_ATTR: always refreshes EM_* and QC_* fields from PW (audit does not include old/new values).
    DOCUMENT_STATE: updates workflow state when it differs.
    Inserts new rows for Sheets-folder paths, or on DOCUMENT_ATTR when the document is not yet indexed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [bool]$IsSheetsFolder = $false,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [string]$AuditActionName = '',
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [switch]$SkipQcInitiatedFallback
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }
    if (-not (Test-QCSheetIndexFolderPath -FolderPath $FolderPath)) { return }

    $isDocumentAttr = ([string]$AuditActionName).Trim() -eq 'DOCUMENT_ATTR'
    $isDocumentState = ([string]$AuditActionName).Trim() -eq 'DOCUMENT_STATE'

    $dbDesigner = ''
    $dbReviewer = ''
    $dbChecker = ''
    $dbReviewType = ''
    $dbAssignedTo = ''
    $dbQcStatus = ''
    $dbState = ''
    $rowExists = $false
    try {
        $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, checker_email, qc_review_type, qc_assigned_to, qc_status, pw_state_name
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if ($dbRes.IsSuccess -and $dbRes.Data.table -and $dbRes.Data.table.Rows.Count -gt 0) {
            $rowExists = $true
            $r = $dbRes.Data.table.Rows[0]
            if (-not ($r.designer_email -is [DBNull])) { $dbDesigner = [string]$r.designer_email }
            if (-not ($r.reviewer_email -is [DBNull])) { $dbReviewer = [string]$r.reviewer_email }
            if ($r.Table.Columns.Contains('checker_email') -and -not ($r.checker_email -is [DBNull])) { $dbChecker = [string]$r.checker_email }
            if ($r.Table.Columns.Contains('qc_review_type') -and -not ($r.qc_review_type -is [DBNull])) { $dbReviewType = [string]$r.qc_review_type }
            if ($r.Table.Columns.Contains('qc_assigned_to') -and -not ($r.qc_assigned_to -is [DBNull])) { $dbAssignedTo = [string]$r.qc_assigned_to }
            if (-not ($r.qc_status -is [DBNull])) { $dbQcStatus = [string]$r.qc_status }
            if (-not ($r.pw_state_name -is [DBNull])) { $dbState = [string]$r.pw_state_name }
        }
    } catch { return }

    if (-not $rowExists -and -not $IsSheetsFolder -and -not $isDocumentAttr) { return }

    if ($isDocumentState -and -not $isDocumentAttr) {
        return
    }

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $read = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $DocumentName -ColumnsToReturn $cols
    if ($isDocumentAttr -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_ATTR_LIVE_READ' `
            -Message 'DOCUMENT_ATTR audit signal resolved by live ProjectWise attribute read.' -Data @{
            auditEventId = $AuditEventId
            documentGuid = $DocumentGuid
            documentName = $DocumentName
            folderPath = $FolderPath
            changedByUser = $ChangedByUser
            changedByUsername = $ChangedByUsername
            auditTimestamp = $LastAuditEventAt
            liveAttributeReadStatus = if ($read.found) { 'found' } else { 'not_found' }
            liveWorkflowState = if ($read.pwStateName) { [string]$read.pwStateName } else { '' }
            error = if ($read.error) { [string]$read.error } else { '' }
        } | Out-Null
    }
    if (-not $read.found) { return }

    $fieldsRaw = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $read.attributes -PwStateName ([string]$read.pwStateName)
    $rawReviewType = [string]$fieldsRaw.qcReviewType
    $fields = @{}
    foreach ($k in @($fieldsRaw.Keys)) { $fields[$k] = $fieldsRaw[$k] }
    if (-not $isDocumentAttr) {
        $fields = _PWD-EnrichSheetIndexReviewType -Config $Config -Fields $fields -FolderPath $FolderPath -DocumentName $DocumentName
    }
    $pwDesigner = [string]$fields.designerEmail
    $pwReviewer = [string]$fields.reviewerEmail
    $pwChecker = [string]$fields.checkerEmail
    $pwReviewType = _PWD-ResolveSheetIndexQcReviewType -Config $Config -FieldsFromPwRead $fieldsRaw `
        -FolderPath $FolderPath -DocumentName $DocumentName -EnrichFromSourcePdf:$isDocumentAttr
    if ([string]::IsNullOrWhiteSpace($pwReviewType) -and -not $isDocumentAttr) {
        $pwReviewType = [string]$fields.qcReviewType
    }
    $pwAssignedTo = [string]$fields.qcAssignedTo
    $pwQcStatus = [string]$fields.qcStatus
    $pwState = [string]$fields.pwStateName

    $emailsDiffer = (_PWD-NormalizeSheetIndexValue $pwDesigner) -ne (_PWD-NormalizeSheetIndexValue $dbDesigner) `
        -or (_PWD-NormalizeSheetIndexValue $pwReviewer) -ne (_PWD-NormalizeSheetIndexValue $dbReviewer)
    $reviewTypeDiffer = (_PWD-NormalizeSheetIndexValue $pwReviewType) -ne (_PWD-NormalizeSheetIndexValue $dbReviewType)
    $qcFieldsDiffer = (_PWD-NormalizeSheetIndexValue $pwChecker) -ne (_PWD-NormalizeSheetIndexValue $dbChecker) `
        -or $reviewTypeDiffer `
        -or (_PWD-NormalizeSheetIndexValue $pwAssignedTo) -ne (_PWD-NormalizeSheetIndexValue $dbAssignedTo) `
        -or (_PWD-NormalizeSheetIndexValue $pwQcStatus) -ne (_PWD-NormalizeSheetIndexValue $dbQcStatus)
    $stateDiffers = (_PWD-NormalizeSheetIndexValue $pwState) -ne (_PWD-NormalizeSheetIndexValue $dbState)

    if ($rowExists -and -not $isDocumentAttr -and -not $emailsDiffer -and -not $qcFieldsDiffer -and -not $stateDiffers) {
        if ($LastAuditEventAt) {
            try {
                Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index
SET last_audit_event_at = @lastAudit, last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid; lastAudit = $LastAuditEventAt } | Out-Null
            } catch { }
        }
        return
    }

    $ext = $null
    if ($DocumentName) {
        $ext = [System.IO.Path]::GetExtension($DocumentName)
        if ($ext) { $ext = $ext.ToLowerInvariant() }
    }
    $sourceType = $null
    if ($ext -eq '.pdf') { $sourceType = 'pdf' }
    elseif ($ext -eq '.dgn') { $sourceType = 'dgn' }

    Write-QCSheetIndex -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
        -FolderPath $FolderPath -WatchRoot $WatchRoot -Extension $ext -SourceType $sourceType `
        -DesignerEmail $pwDesigner -ReviewerEmail $pwReviewer -CheckerEmail $pwChecker `
        -QcReviewType $pwReviewType -QcAssignedTo $pwAssignedTo -QcStatus $pwQcStatus `
        -PwStateName $pwState -LastAuditEventAt $LastAuditEventAt -SetOwnershipFromProjectWise

    if (Get-Command -Name 'Invoke-QCAuditWorkflowStateChangeTriggers' -ErrorAction SilentlyContinue) {
        if ($stateDiffers) {
            Invoke-QCAuditWorkflowStateChangeTriggers -Config $Config -DocumentGuid $DocumentGuid `
                -DocumentName $DocumentName -FolderPath $FolderPath -PreviousState $dbState -CurrentState $pwState `
                -PwAttributes $read.attributes -AuditActionName $AuditActionName -ChangedByUser $ChangedByUser `
                -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId | Out-Null
        }
    }

    if (Get-Command -Name 'Invoke-QCAuditWorkflowAttributeChangeTriggers' -ErrorAction SilentlyContinue) {
        $fieldChanges = @{}
        if ($emailsDiffer) {
            $fieldChanges['designer_email'] = @{ oldValue = $dbDesigner; newValue = $pwDesigner }
            $fieldChanges['reviewer_email'] = @{ oldValue = $dbReviewer; newValue = $pwReviewer }
        }
        if ($qcFieldsDiffer) {
            $fieldChanges['checker_email'] = @{ oldValue = $dbChecker; newValue = $pwChecker }
            $fieldChanges['qc_review_type'] = @{ oldValue = $dbReviewType; newValue = $pwReviewType }
            $fieldChanges['qc_assigned_to'] = @{ oldValue = $dbAssignedTo; newValue = $pwAssignedTo }
            $fieldChanges['qc_status'] = @{ oldValue = $dbQcStatus; newValue = $pwQcStatus }
        }
        if ($fieldChanges.Count -gt 0) {
            Invoke-QCAuditWorkflowAttributeChangeTriggers -Config $Config -DocumentGuid $DocumentGuid `
                -DocumentName $DocumentName -FolderPath $FolderPath -FieldChanges $fieldChanges `
                -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername | Out-Null
        }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_INDEX_SYNC' -Message 'sheet_index synced from ProjectWise.' -Data @{
            documentGuid    = $DocumentGuid
            documentName    = $DocumentName
            folderPath      = $FolderPath
            auditActionName = $AuditActionName
            designerEmail   = $pwDesigner
            reviewerEmail   = $pwReviewer
            checkerEmail    = $pwChecker
            qcReviewType    = $pwReviewType
            rawReviewType   = $rawReviewType
            qcAssignedTo    = $pwAssignedTo
            qcStatus        = $pwQcStatus
            pwStateName     = $pwState
            wasInsert       = (-not $rowExists)
            forceAttrSync   = $isDocumentAttr
            emailsChanged   = $emailsDiffer
            reviewTypeChanged = $reviewTypeDiffer
            qcFieldsChanged = $qcFieldsDiffer
            stateChanged    = $stateDiffers
        }
    }

    if ($isDocumentAttr -and $reviewTypeDiffer -and -not [string]::IsNullOrWhiteSpace($pwReviewType)) {
        Sync-PWAssociatedSheetReviewTypeAttributes -Config $Config -DocumentGuid $DocumentGuid `
            -DocumentName $DocumentName -FolderPath $FolderPath -CanonicalReviewType $pwReviewType `
            -WatchRoot $WatchRoot -LastAuditEventAt $LastAuditEventAt -DryRun:$false
    } elseif ($isDocumentAttr -and $reviewTypeDiffer -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_REVIEW_TYPE_SYNC_SKIPPED' `
            -Message 'DOCUMENT_ATTR review type change detected but canonical QC_Review_Type could not be read from ProjectWise.' -Data @{
            documentGuid = $DocumentGuid; documentName = $DocumentName; folderPath = $FolderPath
            rawReviewType = $rawReviewType; dbReviewType = $dbReviewType; resolvedReviewType = $pwReviewType
        }
    }

    if ($IsSheetsFolder -and $isDocumentAttr -and -not $SkipQcInitiatedFallback) {
        _PWD-TryTriggerQcInitiatedFromAssociatedSheetPdf -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -IsSheetsFolder:$IsSheetsFolder `
            -WatchRoot $WatchRoot -LastAuditEventAt $LastAuditEventAt -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername
    }
}

function _PWD-InvokeUpdatePWDocumentAttributes {
    param(
        [Parameter(Mandatory)][object]$Document,
        [Parameter(Mandatory)][hashtable]$Attributes
    )

    $cmd = Get-Command -Name 'Update-PWDocumentAttributes' -ErrorAction SilentlyContinue
    if (-not $cmd) { throw 'Update-PWDocumentAttributes is not available.' }
    if (-not $Attributes -or $Attributes.Keys.Count -eq 0) { return $false }

    $args = @{}
    $docParam = if ($cmd.Parameters.ContainsKey('InputDocuments')) { 'InputDocuments' }
        elseif ($cmd.Parameters.ContainsKey('InputDocument')) { 'InputDocument' }
        elseif ($cmd.Parameters.ContainsKey('Document')) { 'Document' }
        else { $null }
    $attrsParam = if ($cmd.Parameters.ContainsKey('Attributes')) { 'Attributes' } else { $null }
    if ($docParam) { $args[$docParam] = @($Document) }
    if ($attrsParam) { $args[$attrsParam] = $Attributes }
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) { $args['ReturnBoolean'] = $true }

    if ($docParam -and $attrsParam) {
        $r = & $cmd @args -ErrorAction Stop
        if ($null -ne $r -and $r -is [bool]) { return $r }
        return $true
    }
    if ($attrsParam) {
        & $cmd $Document @args -ErrorAction Stop | Out-Null
        return $true
    }
    & $cmd $Document $Attributes -ErrorAction Stop | Out-Null
    return $true
}

function Sync-PWQcPdfEmailAttributesFromSourcePdf {
    <#
    .SYNOPSIS
    Copies EM_Designer_Email and EM_Reviewer_Email from the source sheet PDF to the matching *-qc.pdf document.
    .DESCRIPTION
    The source .pdf is the source of truth. When the QC PDF is missing these values or they differ,
    updates the QC PDF environment attributes to match the source.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][string]$QcDocumentName,
        [string]$DesignerEmailColumn = 'EM_Designer_Email',
        [string]$ReviewerEmailColumn = 'EM_Reviewer_Email',
        [switch]$PassThru
    )

    $result = @{
        updated = $false
        skipped = $true
        reason = ''
        sourceDesignerEmail = ''
        sourceReviewerEmail = ''
        qcDesignerEmailBefore = ''
        qcReviewerEmailBefore = ''
        attributesWritten = @()
    }

    $source = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $SourceDocumentName `
        -DesignerEmailColumn $DesignerEmailColumn -ReviewerEmailColumn $ReviewerEmailColumn
    if (-not $source.found) {
        $result.reason = if ($source.error) { "Source PDF not found or unreadable: $($source.error)" } else { 'Source PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $result.sourceDesignerEmail = [string]$source.designerEmail
    $result.sourceReviewerEmail = [string]$source.reviewerEmail

    $qc = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $QcDocumentName `
        -DesignerEmailColumn $DesignerEmailColumn -ReviewerEmailColumn $ReviewerEmailColumn
    if (-not $qc.found) {
        $result.reason = if ($qc.error) { "QC PDF not found or unreadable: $($qc.error)" } else { 'QC PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $qcDoc = $qc.document
    $result.qcDesignerEmailBefore = [string]$qc.designerEmail
    $result.qcReviewerEmailBefore = [string]$qc.reviewerEmail

    $toWrite = @{}
    $sourceDesigner = [string]$source.designerEmail
    $sourceReviewer = [string]$source.reviewerEmail
    $qcDesigner = [string]$qc.designerEmail
    $qcReviewer = [string]$qc.reviewerEmail

    if ($sourceDesigner -ne $qcDesigner) {
        $toWrite[$DesignerEmailColumn] = $sourceDesigner
        $result.attributesWritten += $DesignerEmailColumn
    }
    if ($sourceReviewer -ne $qcReviewer) {
        $toWrite[$ReviewerEmailColumn] = $sourceReviewer
        $result.attributesWritten += $ReviewerEmailColumn
    }

    if ($toWrite.Keys.Count -eq 0) {
        $result.skipped = $true
        $result.reason = 'QC PDF email attributes already match source PDF.'
        if ($PassThru) { return $result }
        return
    }

    $target = "$FolderPath\$QcDocumentName"
    if ($PSCmdlet.ShouldProcess($target, "Sync email attributes from $SourceDocumentName")) {
        try {
            [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $qcDoc -Attributes $toWrite)
            $result.updated = $true
            $result.skipped = $false
            $result.reason = 'Updated QC PDF email attributes from source PDF.'
        } catch {
            $result.reason = "Failed to update QC PDF attributes: $($_.Exception.Message)"
        }
    } else {
        $result.skipped = $true
        $result.reason = 'WhatIf: would update QC PDF email attributes from source PDF.'
        $result.attributesWritten = @($toWrite.Keys)
    }

    if ($PassThru) { return $result }
}

function Get-PWQcReviewTypeEnabledEnvironments {
    <#
    .SYNOPSIS
    PW environment names where QC_Review_Type (and related QC workflow attributes) are defined.
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $enabled = @('Caltrans')
    if (-not $Config) { return @($enabled) }
    try {
        $rta = $null
        if ($Config.ContainsKey('projectWise') -and $Config.projectWise) {
            $pw = $Config.projectWise
            if ($pw -is [hashtable] -and $pw.ContainsKey('qcReviewTypeAttributes')) { $rta = $pw['qcReviewTypeAttributes'] }
            elseif ($pw.qcReviewTypeAttributes) { $rta = $pw.qcReviewTypeAttributes }
        }
        if ($rta) {
            $list = $null
            if ($rta -is [hashtable] -and $rta.ContainsKey('enabledEnvironments')) { $list = @($rta['enabledEnvironments']) }
            elseif ($rta.enabledEnvironments) { $list = @($rta.enabledEnvironments) }
            if ($list -and $list.Count -gt 0) {
                $enabled = @($list | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            }
        }
    } catch { }
    return @($enabled)
}

function _PWD-InferPwEnvironmentFromFolderPath {
    param([AllowNull()][string]$FolderPath)

    $p = ([string]$FolderPath).Trim().Trim('\').Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    if ($p -match '(?i)(?:^|/)documents/(Caltrans)(?:/|$)') { return 'Caltrans' }
    if ($p -match '(?i)(?:^|/)Caltrans(?:/|$)') { return 'Caltrans' }
    if ($p -match '(?i)(?:^|/)documents/AZDOT(?:\s+2024)?(?:/|$)') { return 'ADOT' }
    if ($p -match '(?i)(?:^|/)AZDOT(?:\s+2024)?(?:/|$)') { return 'ADOT' }
    return ''
}

function _PWD-NormalizePwEnvironmentForQcReviewType {
    <#
    .SYNOPSIS
    Maps PW folder environment labels (e.g. AZDOT 2024) to config keys in qcReviewTypeAttributes.enabledEnvironments.
    #>
    param([AllowNull()][string]$EnvName)

    $e = ([string]$EnvName).Trim()
    if ([string]::IsNullOrWhiteSpace($e)) { return '' }
    if ($e -match '(?i)^AZDOT') { return 'ADOT' }
    if ($e -match '(?i)^ADOT') { return 'ADOT' }
    if ($e -match '(?i)^Caltrans') { return 'Caltrans' }
    return $e
}

function Get-PWFolderEnvironmentName {
    <#
    .SYNOPSIS
    Resolves the ProjectWise environment name for a folder (Get-PWFolders.Environment), with path fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return '' }

    $apiPath = $FolderPath
    if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
        try {
            $converted = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
            if (-not [string]::IsNullOrWhiteSpace($converted)) { $apiPath = $converted }
        } catch { }
    }

    $folderCmd = Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue
    if ($folderCmd) {
        foreach ($tryPath in @($apiPath, $FolderPath)) {
            if ([string]::IsNullOrWhiteSpace($tryPath)) { continue }
            try {
                $folder = Get-PWFolders -FolderPath $tryPath -JustOne -ErrorAction Stop
                if ($folder) {
                    $envName = ''
                    try { $envName = [string]$folder.Environment } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($envName)) { return $envName.Trim() }
                }
            } catch { }
        }
    }

    return _PWD-InferPwEnvironmentFromFolderPath -FolderPath $FolderPath
}

function Test-PWQcReviewTypeAttributesEnabled {
    <#
    .SYNOPSIS
    True when QC_Review_Type automation is enabled for the folder's PW environment (e.g. Caltrans only until ADOT is integrated).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $enabledEnvs = @(Get-PWQcReviewTypeEnabledEnvironments -Config $Config)
    if ($enabledEnvs.Count -eq 0) { return $false }

    $envName = _PWD-NormalizePwEnvironmentForQcReviewType -EnvName (Get-PWFolderEnvironmentName -FolderPath $FolderPath)
    if ([string]::IsNullOrWhiteSpace($envName)) { return $false }

    foreach ($allowed in $enabledEnvs) {
        $normAllowed = _PWD-NormalizePwEnvironmentForQcReviewType -EnvName ([string]$allowed)
        if (-not [string]::IsNullOrWhiteSpace($normAllowed) -and $normAllowed -ieq $envName) { return $true }
    }
    return $false
}

function Get-PWQcReviewTypeAttributeName {
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $col = 'QC_Review_Type'
    try {
        $am = $Config['qcWorkflow']['attributeMap']
        if ($am -and $am['reviewType']) { $col = [string]$am['reviewType'] }
    } catch { }
    return $col
}

function Get-PWQcDefaultReviewType {
    <#
    .SYNOPSIS
    Configured default QC_Review_Type (qcWorkflow.defaultReviewType), usually Production QC.
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $default = 'Production QC'
    if (-not $Config) { return $default }
    try {
        $wf = $null
        if ($Config.ContainsKey('qcWorkflow')) { $wf = $Config.qcWorkflow }
        if ($wf -is [hashtable] -and $wf.ContainsKey('defaultReviewType') -and -not [string]::IsNullOrWhiteSpace([string]$wf.defaultReviewType)) {
            return [string]$wf.defaultReviewType
        }
        if ($wf -and $wf.PSObject.Properties['defaultReviewType'] -and -not [string]::IsNullOrWhiteSpace([string]$wf.defaultReviewType)) {
            return [string]$wf.defaultReviewType
        }
    } catch { }
    return $default
}

function Ensure-PWQcReviewTypeOnAssociatedSheet {
    <#
    .SYNOPSIS
    Sets QC_Review_Type on associated DGN, sheet PDF, and QC PDF when the source sheet PDF has no review type.
    .DESCRIPTION
    Used during QC_Prepend when QC_Review_Type is unset: writes qcWorkflow.defaultReviewType (Production QC)
    to every associated sibling in the folder that is missing the attribute, then downstream logic treats the sheet as Production QC.
    Skipped when the folder PW environment is not listed in projectWise.qcReviewTypeAttributes.enabledEnvironments (ADOT until integrated).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [string]$DocumentGuid = '',
        [switch]$PassThru
    )

    $pwEnv = Get-PWFolderEnvironmentName -FolderPath $FolderPath
    if (-not (Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath)) {
        $result = @{
            applied = $false
            skipped = $true
            reason = if ($pwEnv) {
                "QC_Review_Type is not enabled for PW environment '$pwEnv' (Caltrans only until ADOT integration)."
            } else {
                'QC_Review_Type is not enabled for this folder (unknown PW environment).'
            }
            defaultReviewType = Get-PWQcDefaultReviewType -Config $Config
            reviewTypeColumn = Get-PWQcReviewTypeAttributeName -Config $Config
            pwEnvironment = $pwEnv
            sourceReviewTypeBefore = ''
            membersChecked = 0
            membersUpdated = @()
            attributesWritten = @()
        }
        if ($PassThru) { return $result }
        return
    }

    $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config
    $defaultReviewType = Get-PWQcDefaultReviewType -Config $Config
    $result = @{
        applied = $false
        skipped = $true
        reason = ''
        defaultReviewType = $defaultReviewType
        reviewTypeColumn = $reviewCol
        sourceReviewTypeBefore = ''
        membersChecked = 0
        membersUpdated = @()
        attributesWritten = @()
    }

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $read = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $SourceDocumentName -ColumnsToReturn $cols
    if (-not $read.found) {
        $result.reason = if ($read.error) { "Source PDF not found or unreadable: $($read.error)" } else { 'Source PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $fields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $read.attributes -PwStateName ([string]$read.pwStateName)
    $sourceRt = [string]$fields.qcReviewType
    $result.sourceReviewTypeBefore = $sourceRt
    if (-not [string]::IsNullOrWhiteSpace($sourceRt)) {
        $result.reason = 'Source sheet PDF already has QC_Review_Type; no default applied.'
        if ($PassThru) { return $result }
        return
    }

    $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $SourceDocumentName -DocumentGuid $DocumentGuid)
    if ($members.Count -eq 0) {
        $members = @(@{
            documentName = $SourceDocumentName
            document     = $read.document
            documentGuid = try { [string]$read.document.DocumentGUID } catch { '' }
        })
    }

    $toWrite = @{ $reviewCol = $defaultReviewType }
    foreach ($member in $members) {
        $dn = [string]$member.documentName
        $doc = $member.document
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        $result.membersChecked++

        if (-not $doc) {
            $dg = [string]$member.documentGuid
            $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }
        if (-not $doc) { continue }

        $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $dn -ColumnsToReturn @($reviewCol)
        if (-not $attrRead.found) { continue }
        $current = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $reviewCol
        if (-not [string]::IsNullOrWhiteSpace([string]$current)) { continue }

        $target = "$FolderPath\$dn"
        if ($PSCmdlet.ShouldProcess($target, "Set $reviewCol to $defaultReviewType (QC_Prepend default)")) {
            try {
                [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes $toWrite)
                $result.applied = $true
                $result.skipped = $false
                $result.membersUpdated += $dn
                if ($result.attributesWritten -notcontains $reviewCol) {
                    $result.attributesWritten += $reviewCol
                }
            } catch {
                $result.reason = "Failed to set $reviewCol on ${dn}: $($_.Exception.Message)"
            }
        } else {
            $result.skipped = $true
            $result.reason = "WhatIf: would set $reviewCol to $defaultReviewType on associated sheet files."
            $result.membersUpdated += $dn
        }
    }

    if ($result.applied) {
        $result.reason = "Set $reviewCol to $defaultReviewType on $($result.membersUpdated.Count) associated file(s)."
    } elseif ([string]::IsNullOrWhiteSpace($result.reason)) {
        if ($result.membersUpdated.Count -gt 0) {
            $result.reason = "WhatIf: would set $reviewCol to $defaultReviewType on associated sheet files."
        } else {
            $result.reason = 'No associated sheet files needed QC_Review_Type default.'
            $result.skipped = $true
        }
    }

    if ($PassThru) { return $result }
}

function Get-PWQcPrependRoleFieldsFromSourcePdf {
    <#
    .SYNOPSIS
    Reads QC role emails and QC_Review_Type from the source sheet PDF in ProjectWise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $folderCandidates = [System.Collections.Generic.List[string]]::new()
    foreach ($fp in @($FolderPath)) {
        if ([string]::IsNullOrWhiteSpace($fp)) { continue }
        $t = $fp.Trim().TrimEnd('\')
        if (-not $folderCandidates.Contains($t)) { $folderCandidates.Add($t) | Out-Null }
        if ($t -notmatch '^(?i)Documents\\') {
            $withDocs = 'Documents\' + $t
            if (-not $folderCandidates.Contains($withDocs)) { $folderCandidates.Add($withDocs) | Out-Null }
        }
    }

    $read = @{ found = $false; error = 'Document not found'; attributes = @{}; pwStateName = '' }
    $resolvedFolder = $FolderPath
    foreach ($fp in @($folderCandidates)) {
        $tryRead = Get-PWDocumentAttributesByColumns -FolderPath $fp -DocumentName $SourceDocumentName -ColumnsToReturn $cols
        if ($tryRead.found) {
            $read = $tryRead
            $resolvedFolder = $fp
            break
        }
        if (-not $read.found -and $tryRead.error) { $read.error = [string]$tryRead.error }
    }
    if (-not $read.found) {
        return @{ found = $false; error = [string]$read.error; designerEmail = ''; reviewerEmail = ''; checkerEmail = ''; qcReviewType = '' }
    }

    $fields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $read.attributes -PwStateName ([string]$read.pwStateName)
    $designerEmail = [string]$fields.designerEmail
    $reviewerEmail = [string]$fields.reviewerEmail
    $checkerEmail = [string]$fields.checkerEmail

    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$SourceDocumentName)
    if (-not [string]::IsNullOrWhiteSpace($stem)) {
        $dgnName = $stem + '.dgn'
        if ($dgnName -ne [string]$SourceDocumentName) {
            $needDgn = [string]::IsNullOrWhiteSpace($designerEmail) -or [string]::IsNullOrWhiteSpace($reviewerEmail)
            if ($needDgn) {
                $dgnRead = Get-PWDocumentAttributesByColumns -FolderPath $resolvedFolder -DocumentName $dgnName -ColumnsToReturn $cols
                if ($dgnRead.found) {
                    $dgnFields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $dgnRead.attributes -PwStateName ([string]$dgnRead.pwStateName)
                    if ([string]::IsNullOrWhiteSpace($designerEmail)) { $designerEmail = [string]$dgnFields.designerEmail }
                    if ([string]::IsNullOrWhiteSpace($reviewerEmail)) { $reviewerEmail = [string]$dgnFields.reviewerEmail }
                    if ([string]::IsNullOrWhiteSpace($checkerEmail)) { $checkerEmail = [string]$dgnFields.checkerEmail }
                    if ([string]::IsNullOrWhiteSpace($fields.qcReviewType) -and $dgnFields.qcReviewType) {
                        $fields.qcReviewType = [string]$dgnFields.qcReviewType
                    }
                }
            }
        }
    }

    $qcReviewType = [string]$fields.qcReviewType
    if ([string]::IsNullOrWhiteSpace($qcReviewType) -and (Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath)) {
        $qcReviewType = Get-PWQcDefaultReviewType -Config $Config
    }
    return @{
        found          = $true
        error          = ''
        designerEmail  = $designerEmail
        reviewerEmail  = $reviewerEmail
        checkerEmail   = $checkerEmail
        qcReviewType   = $qcReviewType
    }
}

function Sync-PWQcPdfReviewTypeFromSourcePdf {
    <#
    .SYNOPSIS
    Copies QC_Review_Type from the source sheet PDF to the matching *-qc.pdf document in PW.
    When the source has no review type, applies the configured default to DGN, sheet PDF, and QC PDF first.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][string]$QcDocumentName,
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$PassThru
    )

    $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config
    $result = @{
        updated = $false
        skipped = $true
        reason = ''
        reviewTypeColumn = $reviewCol
        sourceReviewType = ''
        qcReviewTypeBefore = ''
        attributesWritten = @()
    }

    if (-not (Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath)) {
        $pwEnv = Get-PWFolderEnvironmentName -FolderPath $FolderPath
        $result.reason = if ($pwEnv) {
            "QC_Review_Type sync skipped for PW environment '$pwEnv' (not integrated)."
        } else {
            'QC_Review_Type sync skipped (PW environment unknown or not enabled).'
        }
        if ($PassThru) { return $result }
        return
    }

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $readSource = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $SourceDocumentName -ColumnsToReturn $cols
    if (-not $readSource.found) {
        $result.reason = if ($readSource.error) { "Source PDF not found or unreadable: $($readSource.error)" } else { 'Source PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $fields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $readSource.attributes -PwStateName ([string]$readSource.pwStateName)
    $rawSourceRt = [string]$fields.qcReviewType
    if ([string]::IsNullOrWhiteSpace($rawSourceRt)) {
        $ensure = Ensure-PWQcReviewTypeOnAssociatedSheet -Config $Config -FolderPath $FolderPath `
            -SourceDocumentName $SourceDocumentName -PassThru
        if ($ensure) {
            $result.sourceReviewType = [string]$ensure.defaultReviewType
            if ($ensure.applied) {
                $result.updated = $true
                $result.skipped = $false
                $result.reason = [string]$ensure.reason
                $result.attributesWritten = @($reviewCol)
                if ($PassThru) { return $result }
                return
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ensure.reason)) {
                $result.reason = [string]$ensure.reason
            }
        }
        $result.sourceReviewType = Get-PWQcDefaultReviewType -Config $Config
    } else {
        $result.sourceReviewType = $rawSourceRt
    }

    $read = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $QcDocumentName -ColumnsToReturn @($reviewCol)
    if (-not $read.found) {
        $result.reason = if ($read.error) { "QC PDF not found or unreadable: $($read.error)" } else { 'QC PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $qcDoc = $read.document
    if (-not $qcDoc) {
        $result.reason = 'QC PDF row missing from attribute read.'
        if ($PassThru) { return $result }
        return
    }

    $before = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $reviewCol
    $result.qcReviewTypeBefore = $before
    if ($before -eq $result.sourceReviewType) {
        $result.reason = 'QC PDF QC_Review_Type already matches source PDF.'
        if ($PassThru) { return $result }
        return
    }

    $toWrite = @{ $reviewCol = $result.sourceReviewType }
    $target = "$FolderPath\$QcDocumentName"
    if ($PSCmdlet.ShouldProcess($target, "Sync $reviewCol from $SourceDocumentName")) {
        try {
            [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $qcDoc -Attributes $toWrite)
            $result.updated = $true
            $result.skipped = $false
            $result.reason = "Updated QC PDF $reviewCol from source PDF."
            $result.attributesWritten = @($reviewCol)
        } catch {
            $result.reason = "Failed to update QC PDF attributes: $($_.Exception.Message)"
        }
    } else {
        $result.skipped = $true
        $result.reason = 'WhatIf: would update QC PDF QC_Review_Type from source PDF.'
        $result.attributesWritten = @($reviewCol)
    }

    if ($PassThru) { return $result }
}

function Sync-PrependQcPdfAttributesFromSource {
    <#
    .SYNOPSIS
    Syncs email and QC_Review_Type from source sheet PDF to *-qc.pdf (PW environment attributes).
    When QC_Review_Type is unset on the source sheet PDF, sets Production QC on DGN, sheet PDF, and QC PDF.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][string]$QcDocumentName,
        [Parameter(Mandatory)][hashtable]$Config
    )

    if (Get-Command -Name 'Sync-PWQcPdfEmailAttributesFromSourcePdf' -ErrorAction SilentlyContinue) {
        Sync-PWQcPdfEmailAttributesFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName `
            -QcDocumentName $QcDocumentName | Out-Null
    }
    Sync-PWQcPdfReviewTypeFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName `
        -QcDocumentName $QcDocumentName -Config $Config | Out-Null
}

Export-ModuleMember -Function *
