# PW.Discovery.psm1
# Responsibility: Read-only ProjectWise watch-path resolution and candidate discovery.

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'QC.AuditTriggers.psm1') -Force -ErrorAction SilentlyContinue
Import-Module (Join-Path $PSScriptRoot 'QC.ProcessType.psm1') -Force -ErrorAction SilentlyContinue

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
        'Sync-PWAssociatedSheetEmailAttributes'
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
    $qcProcessTypeCol = if ($wfMap.ContainsKey('processType')) { $wfMap['processType'] } else { 'QC_Process_Type' }
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
        qcProcessType  = _PWD-GetPwAttributeValue -PwAttributes $PwAttributes -ColumnName $qcProcessTypeCol
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
        if ((Test-QCIsQcPdfDocumentName -DocumentName $sheetPdfName) -or ($sheetPdfName -match '(?i)\.dgn$')) {
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
    Resolves process/review type for sheet_index sync: QC_Process_Type first, QC_Review_Type read-only fallback.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$FieldsFromPwRead,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [bool]$EnrichFromSourcePdf = $true
    )

    $resolved = [string]$FieldsFromPwRead.qcProcessType
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $resolved = [string]$FieldsFromPwRead.qcReviewType
    }
    if ($EnrichFromSourcePdf -and [string]::IsNullOrWhiteSpace($resolved)) {
        $enriched = _PWD-EnrichSheetIndexReviewType -Config $Config -Fields $FieldsFromPwRead `
            -FolderPath $FolderPath -DocumentName $DocumentName
        if (-not [string]::IsNullOrWhiteSpace([string]$enriched.qcProcessType)) {
            $resolved = [string]$enriched.qcProcessType
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$enriched.qcReviewType)) {
            $resolved = [string]$enriched.qcReviewType
        }
    }
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $processCol = Get-PWQcProcessTypeAttributeName -Config $Config
        $reviewCol = Get-PWQcReviewTypeAttributeName -Config $Config
        $solo = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $DocumentName `
            -ColumnsToReturn @($processCol, $reviewCol)
        if ($solo.found) {
            $resolved = _PWD-GetPwAttributeValue -PwAttributes $solo.attributes -ColumnName $processCol
            if ([string]::IsNullOrWhiteSpace($resolved)) {
                $resolved = _PWD-GetPwAttributeValue -PwAttributes $solo.attributes -ColumnName $reviewCol
            }
        }
    }
    return ([string]$resolved).Trim()
}

function _PWD-ResolveCanonicalProcessTypeForSync {
    param(
        [string]$RawValue,
        [hashtable]$Config = $null
    )
    if ([string]::IsNullOrWhiteSpace($RawValue)) { return $null }
    if (-not (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue)) { return $null }
    return Normalize-QCProcessType -ProcessType ([string]$RawValue).Trim()
}

function _PWD-LogProcessTypeSyncSkipped {
    param(
        [Parameter(Mandatory)][string]$Reason,
        [hashtable]$Config = $null,
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [hashtable]$ExtraData = $null
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    $data = @{
        reason = [string]$Reason
        documentGuid = [string]$DocumentGuid
        documentName = [string]$DocumentName
        folderPath = [string]$FolderPath
    }
    if ($ExtraData) {
        foreach ($k in $ExtraData.Keys) { $data[$k] = $ExtraData[$k] }
    }
    Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PROCESS_TYPE_SYNC_SKIPPED' `
        -Message 'Skipped QC process type attribute sync.' -Data $data | Out-Null
}

function _PWD-LogProcessTypeUnknown {
    param(
        [hashtable]$Config = $null,
        [string]$DocumentGuid = '',
        [string]$DocumentName = '',
        [string]$FolderPath = '',
        [string]$RawValue = '',
        [string]$Source = ''
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_PROCESS_TYPE_UNKNOWN' `
        -Message 'Could not resolve canonical QC process type; no ProjectWise attribute write performed.' -Data @{
        documentGuid = [string]$DocumentGuid
        documentName = [string]$DocumentName
        folderPath = [string]$FolderPath
        rawValue = [string]$RawValue
        source = [string]$Source
    } | Out-Null
}

function _PWD-SyncReferenceSheetProcessTypeAttributes {
    <#
    .SYNOPSIS
    Writes QC_Process_Type on stem PDF and optionally DGN (prepend reset path). Never touches lane QC PDFs or QC_Review_Type.
    When -ControlDocumentOnly is set, only the stem PDF is updated (not DGN or siblings).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$CanonicalProcessType,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [bool]$DryRun = $false,
        [bool]$ControlDocumentOnly = $false
    )

    $canonicalNorm = _PWD-ResolveCanonicalProcessTypeForSync -RawValue $CanonicalProcessType -Config $Config
    if (-not $canonicalNorm) {
        _PWD-LogProcessTypeUnknown -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -RawValue $CanonicalProcessType -Source 'prepend_stem_reset'
        _PWD-LogProcessTypeSyncSkipped -Reason 'null_canonical_process_type' -Config $Config `
            -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
            -ExtraData @{ source = 'prepend_stem_reset'; rawValue = [string]$CanonicalProcessType }
        return @{ skipped = $true; reason = 'null_canonical_process_type' }
    }

    $canonicalDisplay = $canonicalNorm
    if (Get-Command -Name 'Format-QCProcessTypeAttributeValue' -ErrorAction SilentlyContinue) {
        $canonicalDisplay = Format-QCProcessTypeAttributeValue -ProcessType $canonicalNorm
    }

    $sheetStem = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }
    if ([string]::IsNullOrWhiteSpace($sheetStem)) {
        $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    }
    if ([string]::IsNullOrWhiteSpace($sheetStem)) {
        return @{ skipped = $true; reason = 'missing_sheet_stem' }
    }

    $pwWritesEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath
    $processCol = Get-PWQcProcessTypeAttributeName -Config $Config
    # QC_Process_Type on DGN is user-owned; only the stem sheet PDF may be updated by automation.
    $targetNames = @($sheetStem + '.pdf')
    $updates = @()

    foreach ($dn in @($targetNames)) {
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid ''
        if (-not $doc) { continue }
        $dg = try { [string]$doc.DocumentGUID } catch { '' }

        $currentPw = ''
        $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $dn -ColumnsToReturn @($processCol)
        if ($attrRead.found) {
            $currentPw = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $processCol
            if (-not $doc -and $attrRead.document) { $doc = $attrRead.document }
        }
        if ((_PWD-NormalizeSheetIndexValue $currentPw) -eq (_PWD-NormalizeSheetIndexValue $canonicalNorm)) { continue }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromValue = [string]$currentPw
            toValue = $canonicalDisplay
            applied = $false
            planned = $false
        }
        if (-not $pwWritesEnabled) {
            $change.pwWriteSkipped = 'qc_process_type_not_enabled_for_environment'
        } elseif ($DryRun) {
            $change.planned = $true
        } else {
            try {
                [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes @{ $processCol = $canonicalNorm } -Config $Config)
                $change.applied = $true
                if ($dg -and (Get-Command -Name 'Write-QCSheetIndex' -ErrorAction SilentlyContinue)) {
                    try {
                        $ext = [System.IO.Path]::GetExtension($dn)
                        if ($ext) { $ext = $ext.ToLowerInvariant() }
                        $sourceType = if ($ext -eq '.pdf') { 'pdf' } elseif ($ext -eq '.dgn') { 'dgn' } else { $null }
                        Write-QCSheetIndex -Config $Config -DocumentGuid $dg -DocumentName $dn -FolderPath $FolderPath `
                            -WatchRoot $WatchRoot -Extension $ext -SourceType $sourceType -QcReviewType $canonicalDisplay `
                            -LastAuditEventAt $LastAuditEventAt -SetOwnershipFromProjectWise | Out-Null
                    } catch { }
                }
            } catch {
                $change.error = [string]$_.Exception.Message
            }
        }
        $updates += $change
    }

    if ($updates.Count -gt 0 -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        $logCode = if ($ControlDocumentOnly) { 'QC_PROCESS_TYPE_RESET_CONTROL_DOCUMENT_ONLY' } else { 'QC_PREPEND_STEM_PROCESS_TYPE_RESET' }
        $logMsg = if ($ControlDocumentOnly) {
            'Reset control sheet PDF QC_Process_Type after lane prepend.'
        } else {
            'Reset stem/DGN QC_Process_Type after Initiate Origination prepend.'
        }
        Write-QCJsonLog -Level 'Information' -Code $logCode `
            -Message $logMsg -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath = $FolderPath
            canonicalProcessType = $canonicalNorm
            controlDocumentOnly = [bool]$ControlDocumentOnly
            updateCount = $updates.Count
            updates = @($updates)
            dryRun = [bool]$DryRun
        } | Out-Null
    }

    return @{ updates = @($updates); canonicalProcessType = $canonicalNorm; targetState = $canonicalDisplay; controlDocumentOnly = [bool]$ControlDocumentOnly }
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

function _PWD-GetSheetRoleEmailColumnNames {
    param([hashtable]$Config = @{})

    $designer = 'EM_Designer_Email'
    $reviewer = 'EM_Reviewer_Email'
    $checker = 'EM_Checker_Email'
    try {
        if ($Config) {
            $na = $Config['notifications']['attributes']
            if ($na) {
                if ($na['designerEmailField']) { $designer = [string]$na['designerEmailField'] }
                if ($na['reviewerEmailField']) { $reviewer = [string]$na['reviewerEmailField'] }
                if ($na['checkerEmailField']) { $checker = [string]$na['checkerEmailField'] }
            }
            $pw = $Config['projectWise']
            if ($pw -and $pw['environmentEmailAttributes']) {
                $ea = $pw['environmentEmailAttributes']
                if ($ea['default']) {
                    if ($ea['default']['designerEmailColumn']) { $designer = [string]$ea['default']['designerEmailColumn'] }
                    if ($ea['default']['reviewerEmailColumn']) { $reviewer = [string]$ea['default']['reviewerEmailColumn'] }
                    if ($ea['default']['checkerEmailColumn']) { $checker = [string]$ea['default']['checkerEmailColumn'] }
                }
            }
        }
    } catch { }
    return @{ designer = $designer; reviewer = $reviewer; checker = $checker }
}

function Get-PWDocumentEmailContacts {
    <#
    .SYNOPSIS
    Reads EM_Designer_Email, EM_Reviewer_Email, and EM_Checker_Email from a document in a folder via search-with-columns API.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DesignerEmailColumn = 'EM_Designer_Email',
        [string]$ReviewerEmailColumn = 'EM_Reviewer_Email',
        [string]$CheckerEmailColumn = 'EM_Checker_Email'
    )

    $cols = @($DesignerEmailColumn, $ReviewerEmailColumn, $CheckerEmailColumn) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $read = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $DocumentName -ColumnsToReturn $cols
    if (-not $read.found) {
        return @{ designerEmail = ''; reviewerEmail = ''; checkerEmail = ''; found = $false; error = $read.error }
    }

    $designer = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $DesignerEmailColumn
    $reviewer = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $ReviewerEmailColumn
    $checker = _PWD-GetPwAttributeValue -PwAttributes $read.attributes -ColumnName $CheckerEmailColumn
    return @{
        designerEmail = $designer
        reviewerEmail = $reviewer
        checkerEmail  = $checker
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
    if (Test-QCIsQcPdfDocumentName -DocumentName $dn) { return 'qcPdf' }
    if ($dn -match '(?i)\.dgn$') { return 'dgn' }
    if (Test-QCIsSheetPdfDocumentName -DocumentName $dn) { return 'pdf' }
    return 'other'
}

function _PWD-TestAutomationDgnProtectedDocument {
    param([string]$DocumentName)
    return ([string]$DocumentName -match '(?i)\.dgn$')
}

function _PWD-TestTriggerIsLaneQcPdf {
    param([string]$DocumentName)
    if (Get-Command -Name 'Test-PWQcPdfLaneSuffix' -ErrorAction SilentlyContinue) {
        return [bool](Test-PWQcPdfLaneSuffix -DocumentName $DocumentName)
    }
    return ([string]$DocumentName -match '(?i)-(prod|chk|rev)\.pdf$')
}

function _PWD-TestAutomationStemPdfAllowedTargetState {
    param(
        [hashtable]$Config = $null,
        [string]$TargetState = ''
    )
    if ([string]::IsNullOrWhiteSpace($TargetState)) { return $false }
    $allowed = @()
    if (Get-Command -Name 'Get-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        foreach ($key in @('production', 'qcInitiated')) {
            try {
                $name = [string](Get-QCWorkflowStateName -Settings $(if ($Config) { Get-QCWorkflowSettings -Config $Config } else { $null }) -StateKey $key)
                if (-not [string]::IsNullOrWhiteSpace($name)) { $allowed += $name }
            } catch { }
        }
    }
    if ($allowed.Count -eq 0) {
        $allowed = @('In Development', 'Initiate Origination')
    }
    foreach ($candidate in @($allowed)) {
        if ((_PWD-NormalizeSheetIndexValue $TargetState) -eq (_PWD-NormalizeSheetIndexValue $candidate)) {
            return $true
        }
    }
    return $false
}

function _PWD-TestAutomationStemPdfStateWriteBlocked {
    <#
    Stem PDF automation may only write In Development (post-prepend) or Initiate Origination intake paths.
    Lane QC states must never be copied onto the stem sheet PDF.
    #>
    param(
        [string]$DocumentName = '',
        [string]$TargetState = '',
        [string]$WriteScope = 'workflow',
        [hashtable]$Config = $null
    )
    if (-not (Get-Command -Name 'Test-QCIsSheetPdfDocumentName' -ErrorAction SilentlyContinue)) { return $false }
    if (-not (Test-QCIsSheetPdfDocumentName -DocumentName $DocumentName)) { return $false }
    if ($WriteScope -eq 'stem') { return $false }
    return -not (_PWD-TestAutomationStemPdfAllowedTargetState -Config $Config -TargetState $TargetState)
}

function _PWD-LogAutomationStemPdfWriteBlocked {
    param(
        [string]$Operation = '',
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$CallSite = '',
        [string]$TargetState = '',
        [string]$TriggerDocumentName = ''
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    Write-QCJsonLog -Level 'Information' -Code 'QC_STEM_PDF_AUTOMATION_WRITE_BLOCKED' `
        -Message 'Automation blocked from copying lane workflow state onto stem sheet PDF.' -Data @{
        operation = [string]$Operation
        documentName = [string]$DocumentName
        documentGuid = [string]$DocumentGuid
        folderPath = [string]$FolderPath
        callSite = [string]$CallSite
        targetState = [string]$TargetState
        triggerDocumentName = [string]$TriggerDocumentName
    } | Out-Null
}

function _PWD-TestAutomationWritesQcProcessTypeOnAttributes {
    param(
        [hashtable]$Attributes,
        [hashtable]$Config = $null
    )
    if (-not $Attributes -or $Attributes.Keys.Count -eq 0) { return $false }
    $processCol = 'QC_Process_Type'
    if ($Config -and (Get-Command -Name 'Get-PWQcProcessTypeAttributeName' -ErrorAction SilentlyContinue)) {
        try { $processCol = Get-PWQcProcessTypeAttributeName -Config $Config } catch { }
    }
    foreach ($key in @($Attributes.Keys)) {
        if ((_PWD-NormalizeSheetIndexValue ([string]$key)) -eq (_PWD-NormalizeSheetIndexValue $processCol)) {
            return $true
        }
    }
    return $false
}

function _PWD-LogAutomationDgnWriteBlocked {
    param(
        [string]$Operation,
        [string]$DocumentName = '',
        [string]$DocumentGuid = '',
        [string]$FolderPath = '',
        [string]$CallSite = ''
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    Write-QCJsonLog -Level 'Information' -Code 'QC_DGN_AUTOMATION_WRITE_BLOCKED' `
        -Message 'Automation blocked from writing DGN workflow state or QC_Process_Type.' -Data @{
        operation = [string]$Operation
        documentName = [string]$DocumentName
        documentGuid = [string]$DocumentGuid
        folderPath = [string]$FolderPath
        callSite = [string]$CallSite
    } | Out-Null
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
    Normalized sheet stem from a DGN, sheet PDF, or lane PDF (*-prod/-chk/-rev.pdf) filename.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DocumentName
    )
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) { return '' }
    if ($stem -match '(?i)-(prod|chk|rev)$') { $stem = $stem -replace '(?i)-(prod|chk|rev)$', '' }
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

    if ($DocumentName -notmatch '(?i)\.pdf$') { return $false }
    if (Get-Command -Name 'Test-PWQcPdfLaneSuffix' -ErrorAction SilentlyContinue) {
        if (Test-PWQcPdfLaneSuffix -DocumentName $DocumentName) { return $false }
    } elseif ($DocumentName -match '(?i)-(prod|chk|rev)\.pdf$') { return $false }
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

function Get-PWAssociatedSheetSyncDocumentNames {
    <#
    .SYNOPSIS
    Sibling filenames that participate in workflow state sync (sheet PDF and production QC PDF only; DGN excluded).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SheetStem
    )
    if ([string]::IsNullOrWhiteSpace($SheetStem)) { return @() }
    # DGN workflow state is user-owned; automation never participates in sibling state sync.
    return @(
        ($SheetStem + '.pdf')
        ($SheetStem + '-prod.pdf')
    )
}

function Get-PWAssociatedSheetDocumentNames {
    <#
    .SYNOPSIS
    Expected sibling filenames for discovery (sheet PDF, DGN, all lane QC PDFs).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SheetStem
    )
    if ([string]::IsNullOrWhiteSpace($SheetStem)) { return @() }
    return @(
        ($SheetStem + '.pdf')
        ($SheetStem + '.dgn')
        ($SheetStem + '-prod.pdf')
        ($SheetStem + '-chk.pdf')
        ($SheetStem + '-rev.pdf')
    )
}

function _PWD-TestLegacySiblingStateSyncEnabled {
    param([hashtable]$Config)
    if (Get-Command -Name 'Test-QCLegacySiblingStateSyncEnabled' -ErrorAction SilentlyContinue) {
        try { return (Test-QCLegacySiblingStateSyncEnabled -Config $Config) } catch { }
    }
    return $false
}

function _PWD-LogLaneStateIndependentTelemetry {
    param(
        [array]$AllMembers,
        [string]$FolderPath = '',
        [string]$TriggerSource = '',
        [string]$TriggerDocumentName = '',
        [string]$TriggerDocumentGuid = ''
    )
    if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
    $sheetStem = ''
    if ($TriggerDocumentName -and (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue)) {
        $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $TriggerDocumentName
    }
    Write-QCJsonLog -Level 'Information' -Code 'QC_LANE_STATE_INDEPENDENT' `
        -Message 'QC process lanes maintain independent workflow state; sibling sync disabled by design.' -Data @{
        folderPath = $FolderPath
        triggerSource = $TriggerSource
        triggerDocumentName = $TriggerDocumentName
        triggerDocumentGuid = $TriggerDocumentGuid
        sheetStem = $sheetStem
        memberCount = @($AllMembers).Count
    } | Out-Null
    if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
        foreach ($m in @($AllMembers)) {
            $dn = [string]$m.documentName
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }
            $lane = Get-PWQcPdfLaneFromDocumentName -DocumentName $dn
            if (-not $lane) { continue }
            Write-QCJsonLog -Level 'Information' -Code 'QC_SIBLING_STATE_SYNC_DISABLED' `
                -Message 'Sibling workflow state sync skipped; lane state is independent.' -Data @{
                documentName = $dn
                documentGuid = [string]$m.documentGuid
                qcProcessType = $lane
                folderPath = $FolderPath
                triggerSource = $TriggerSource
            } | Out-Null
        }
    }
}

function _PWD-SyncSheetPackageLaneQcPdfs {
    param(
        [hashtable]$Config,
        [string]$FolderPath,
        [string]$SheetStem,
        [array]$Members,
        [hashtable]$StateByGuid = @{},
        [string]$ActiveQcProcessType = '',
        [switch]$RequireActiveLane
    )
    if (-not (Get-Command -Name 'Sync-SheetPackageLaneQcPdfsFromMembers' -ErrorAction SilentlyContinue)) {
        try {
            Import-Module (Join-Path $PSScriptRoot 'Core.Database.psm1') -Force -ErrorAction SilentlyContinue
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($SheetStem)) {
        if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
            foreach ($m in @($Members)) {
                $dn = [string]$m.documentName
                if (-not [string]::IsNullOrWhiteSpace($dn)) {
                    $SheetStem = Get-PWSheetStemFromDocumentName -DocumentName $dn
                    if (-not [string]::IsNullOrWhiteSpace($SheetStem)) { break }
                }
            }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($SheetStem) `
            -and (Get-Command -Name 'Sync-QCLaneQcPdfGuidFromProjectWise' -ErrorAction SilentlyContinue)) {
        $laneSuffixMap = @{ production = '-prod.pdf'; check = '-chk.pdf'; review = '-rev.pdf' }
        foreach ($lane in @('production', 'check', 'review')) {
            $laneName = $SheetStem + $laneSuffixMap[$lane]
            $pwState = ''
            $laneMember = @($Members | Where-Object {
                $n = [string]$_.documentName
                $n -and ($n.ToLowerInvariant() -eq $laneName.ToLowerInvariant())
            } | Select-Object -First 1)
            if ($laneMember -and $laneMember.documentGuid -and $StateByGuid) {
                $lk = ([string]$laneMember.documentGuid).ToLowerInvariant()
                if ($StateByGuid.ContainsKey($lk)) { $pwState = [string]$StateByGuid[$lk] }
            }
            try {
                Sync-QCLaneQcPdfGuidFromProjectWise -Config $Config -FolderPath $FolderPath `
                    -QcPdfName $laneName -QcProcessType $lane -CurrentPwState $pwState | Out-Null
            } catch { }
        }
    }
    if (Get-Command -Name 'Sync-SheetPackageLaneQcPdfsFromMembers' -ErrorAction SilentlyContinue) {
        try {
            Sync-SheetPackageLaneQcPdfsFromMembers -Config $Config -FolderPath $FolderPath -SheetStem $SheetStem `
                -Members $Members -StateByGuid $StateByGuid -ActiveQcProcessType $ActiveQcProcessType `
                -RequireActiveLane:$RequireActiveLane | Out-Null
        } catch { }
    }
}

function Get-PWAssociatedSheetSyncMembers {
    <#
    .SYNOPSIS
    Resolves DGN, sheet PDF, and production QC PDF siblings for workflow state sync.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [string]$TriggerSource = ''
    )
    $all = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    if (_PWD-TestTriggerIsLaneQcPdf -DocumentName $DocumentName) {
        _PWD-LogLaneStateIndependentTelemetry -AllMembers $all -FolderPath $FolderPath -TriggerSource $TriggerSource `
            -TriggerDocumentName $DocumentName -TriggerDocumentGuid $DocumentGuid
        return @()
    }
    $syncNames = @{}
    $stem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    foreach ($n in @(Get-PWAssociatedSheetSyncDocumentNames -SheetStem $stem)) {
        $syncNames[$n.ToLowerInvariant()] = $true
    }
    $sync = @($all | Where-Object {
        $dn = [string]$_.documentName
        $dn -and $syncNames.ContainsKey($dn.ToLowerInvariant()) -and -not (_PWD-TestAutomationDgnProtectedDocument -DocumentName $dn)
    })
    if (-not (_PWD-TestLegacySiblingStateSyncEnabled -Config $Config)) {
        _PWD-LogLaneStateIndependentTelemetry -AllMembers $all -FolderPath $FolderPath -TriggerSource $TriggerSource `
            -TriggerDocumentName $DocumentName -TriggerDocumentGuid $DocumentGuid
        return @()
    }
    return $sync
}

function _PWD-GetLaneIndependentAuditMembers {
    param(
        [array]$AllMembers,
        [string]$TriggerDocumentGuid = '',
        [string]$TriggerDocumentName = ''
    )
    $triggerKey = ([string]$TriggerDocumentName).ToLowerInvariant()
    $triggerGuidKey = ([string]$TriggerDocumentGuid).ToLowerInvariant()
    $seen = @{}
    $match = [System.Collections.Generic.List[object]]::new()
    foreach ($m in @($AllMembers)) {
        $dn = [string]$m.documentName
        $dg = ([string]$m.documentGuid).ToLowerInvariant()
        $hit = $false
        if ($triggerGuidKey -and $dg -eq $triggerGuidKey) { $hit = $true }
        elseif ($triggerKey -and $dn -and ($dn.ToLowerInvariant() -eq $triggerKey)) { $hit = $true }
        if (-not $hit) { continue }
        $dedupeKey = if ($dg) { $dg } else { $dn.ToLowerInvariant() }
        if ($seen.ContainsKey($dedupeKey)) { continue }
        $seen[$dedupeKey] = $true
        $entry = $m
        if ($triggerGuidKey -and $dn -and ($dn.ToLowerInvariant() -eq $triggerKey) -and $dg -ne $triggerGuidKey) {
            $entry = @{
                documentGuid = [string]$TriggerDocumentGuid
                documentName = $dn
            }
            if ($null -ne $m.document) { $entry['document'] = $m.document }
            if ($m.sourceType) { $entry['sourceType'] = [string]$m.sourceType }
        }
        $match.Add($entry) | Out-Null
    }
    return @($match)
}

function _PWD-GetWorkflowStateFromPwDocument {
    param(
        [object]$Document,
        [string]$FolderPath = '',
        [string]$DocumentName = '',
        [string]$DocumentGuid = ''
    )
    if ($Document) {
        foreach ($name in @('WorkflowState', 'StateName', 'State', 'WorkflowStateName', 'CurrentState')) {
            try {
                if ($Document.PSObject.Properties[$name]) {
                    $v = [string]$Document.$name
                    if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
                }
            } catch { }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        try {
            $pw = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid
            if (-not [string]::IsNullOrWhiteSpace($pw)) { return [string]$pw }
        } catch { }
    }
    return ''
}

function _PWD-GetCleanPwDocumentForStateChange {
    <#
    .SYNOPSIS
    Reloads a ProjectWise document without attribute bags for Set-PWDocumentState.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = ''
    )
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
            if ($searchCmd.Parameters.ContainsKey('GetAttributes')) { $params['GetAttributes'] = $false }
            $doc = & $searchCmd @params | Select-Object -First 1
            if ($doc) { return $doc }
        } catch { }
    }
    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $guidCmd = Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue
        if ($guidCmd) {
            try {
                $guidParams = @{ DocumentGUIDs = @($DocumentGuid); ErrorAction = 'Stop' }
                if ($guidCmd.Parameters.ContainsKey('GetAttributes')) { $guidParams['GetAttributes'] = $false }
                $byGuid = & $guidCmd @guidParams | Select-Object -First 1
                if ($byGuid) { return $byGuid }
            } catch { }
        }
    }
    return $null
}

function Set-PWDocumentWorkflowStateVerified {
    <#
    .SYNOPSIS
    Writes ProjectWise workflow state using clean document reload, -Force (default), and read-back verification.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [Parameter(Mandatory)][string]$TargetState,
        [string]$QcProcessType = '',
        [int]$ReadBackDelayMs = 2000,
        [bool]$DryRun = $false,
        [string]$WorkflowName = '',
        [bool]$IsLaneAuthority = $true,
        [ValidateSet('lane', 'stem', 'workflow')]
        [string]$WriteScope = ''
    )

    if ([string]::IsNullOrWhiteSpace($WriteScope)) {
        $WriteScope = if ($IsLaneAuthority) { 'lane' } else { 'workflow' }
    }

    $targetFormatted = ([string]$TargetState).Trim()
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try { $targetFormatted = Format-QCWorkflowStateName -StateName $TargetState -Config $Config } catch { }
    }

    $useForce = $true
    if (Get-Command -Name 'Test-QCWorkflowStateWritebackUseForce' -ErrorAction SilentlyContinue) {
        try { $useForce = Test-QCWorkflowStateWritebackUseForce -Config $Config } catch { }
    }

    $commandShapeBase = "Set-PWDocumentState -InputDocuments @(`$cleanDoc) -State '$targetFormatted'"
    if ($useForce) { $commandShapeBase += ' -Force' }

    $result = @{
        documentGuid   = ''
        documentName   = $DocumentName
        fromState      = ''
        targetState    = $targetFormatted
        readBackState  = ''
        applied        = $false
        verified       = $false
        usedForce      = [bool]$useForce
        commandShape   = $commandShapeBase
        error          = ''
        planned        = $false
        qcProcessType  = [string]$QcProcessType
        attemptedCommand = 'Set-PWDocumentState'
        isLaneAuthority = [bool]$IsLaneAuthority
    }

    $telemetryPrefix = switch ($WriteScope) {
        'lane' { 'QC_LANE' }
        'stem' { 'QC_STEM_REFERENCE' }
        default { 'QC_WORKFLOW' }
    }

    $cleanDoc = _PWD-GetCleanPwDocumentForStateChange -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid
    if (-not $cleanDoc) {
        $result.error = 'clean_document_not_found'
        return $result
    }

    try { $result.documentGuid = [string]$cleanDoc.DocumentGUID } catch { }
    if ([string]::IsNullOrWhiteSpace($result.documentGuid) -and $DocumentGuid) {
        $result.documentGuid = [string]$DocumentGuid
    }

    $fromState = _PWD-GetWorkflowStateFromPwDocument -Document $cleanDoc -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $result.documentGuid
    $result.fromState = [string]$fromState

    function _PWD-WriteStateWriteTelemetry {
        param([string]$Code, [string]$Level, [string]$Message)
        if (-not (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) { return }
        Write-QCJsonLog -Level $Level -Code $Code -Message $Message -Data @{
            documentGuid   = $result.documentGuid
            documentName   = $DocumentName
            qcProcessType  = [string]$QcProcessType
            fromState      = [string]$result.fromState
            targetState    = $targetFormatted
            commandShape   = $commandShapeBase
            usedForce      = [bool]$useForce
            readBackState  = [string]$result.readBackState
            verified       = [bool]$result.verified
            workflowName   = $WorkflowName
            attemptedCommand = 'Set-PWDocumentState'
            error          = [string]$result.error
            isLaneAuthority = [bool]$IsLaneAuthority
            writeScope = [string]$WriteScope
        } | Out-Null
    }

    _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_WORKFLOW_WRITEBACK_TARGET_RESOLVED') -Level 'Information' `
        -Message $(if ($IsLaneAuthority) {
            'Lane workflow state write target resolved with clean document.'
        } else {
            'Stem PDF reference workflow state write target resolved with clean document.'
        })

    if ((_PWD-NormalizeSheetIndexValue $fromState) -eq (_PWD-NormalizeSheetIndexValue $targetFormatted)) {
        $result.readBackState = [string]$fromState
        $result.verified = $true
        _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_STATE_WRITE_VERIFIED') -Level 'Information' `
            -Message $(if ($IsLaneAuthority) {
                'Lane PDF workflow state already at target.'
            } else {
                'Stem PDF workflow state already at target.'
            })
        return $result
    }

    if ($DryRun) {
        $result.planned = $true
        return $result
    }

    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $result.error = 'Set-PWDocumentState unavailable.'
        _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_STATE_WRITE_UNVERIFIED') -Level 'Warning' `
            -Message 'Workflow state write failed before read-back.'
        _PWD-WriteStateWriteTelemetry -Code 'QC_WORKFLOW_STATE_WRITE_UNVERIFIED' -Level 'Warning' `
            -Message 'Workflow state write failed before read-back.'
        return $result
    }

    try {
        $invokeParams = @{ ErrorAction = 'Stop' }
        $docParam = if ($cmd.Parameters.ContainsKey('InputDocuments')) { 'InputDocuments' }
            elseif ($cmd.Parameters.ContainsKey('InputDocument')) { 'InputDocument' }
            elseif ($cmd.Parameters.ContainsKey('Document')) { 'Document' }
            else { $null }
        $stateParam = if ($cmd.Parameters.ContainsKey('State')) { 'State' }
            elseif ($cmd.Parameters.ContainsKey('StateName')) { 'StateName' }
            else { $null }
        if (-not $docParam -or -not $stateParam) {
            throw 'Set-PWDocumentState missing InputDocuments/State parameters.'
        }
        $invokeParams[$docParam] = @($cleanDoc)
        $invokeParams[$stateParam] = $targetFormatted
        if ($useForce -and $cmd.Parameters.ContainsKey('Force')) { $invokeParams['Force'] = $true }
        if ($cmd.Parameters.ContainsKey('ReturnBoolean')) { $invokeParams['ReturnBoolean'] = $true }
        $writeResult = & $cmd @invokeParams
        if ($cmd.Parameters.ContainsKey('ReturnBoolean')) {
            try {
                if ($writeResult -eq $false) {
                    throw "Set-PWDocumentState returned false for state '$targetFormatted'."
                }
            } catch {
                if ($_.Exception.Message -match 'returned false') { throw }
            }
        }
        $result.applied = $true
    } catch {
        $result.error = [string]$_.Exception.Message
        _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_STATE_WRITE_UNVERIFIED') -Level 'Warning' `
            -Message 'Workflow state write threw or returned false.'
        _PWD-WriteStateWriteTelemetry -Code 'QC_WORKFLOW_STATE_WRITE_UNVERIFIED' -Level 'Warning' `
            -Message 'Workflow state write threw or returned false.'
        return $result
    }

    if ($ReadBackDelayMs -gt 0) { Start-Sleep -Milliseconds $ReadBackDelayMs }

    $readDoc = _PWD-GetCleanPwDocumentForStateChange -FolderPath $FolderPath -DocumentName $DocumentName `
        -DocumentGuid $result.documentGuid
    $readBack = _PWD-GetWorkflowStateFromPwDocument -Document $readDoc -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $result.documentGuid
    $result.readBackState = [string]$readBack
    $result.verified = ((_PWD-NormalizeSheetIndexValue $readBack) -eq (_PWD-NormalizeSheetIndexValue $targetFormatted))

    if ($result.verified) {
        _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_STATE_WRITE_VERIFIED') -Level 'Information' `
            -Message $(if ($IsLaneAuthority) {
                'Lane PDF workflow state verified after write.'
            } else {
                'Stem PDF workflow state verified after write.'
            })
    } else {
        _PWD-WriteStateWriteTelemetry -Code ($telemetryPrefix + '_STATE_WRITE_UNVERIFIED') -Level 'Warning' `
            -Message 'Workflow state did not read back at target after write.'
        _PWD-WriteStateWriteTelemetry -Code 'QC_WORKFLOW_STATE_WRITE_UNVERIFIED' -Level 'Warning' `
            -Message 'Workflow state did not read back at target after write.'
    }

    return $result
}

function _PWD-InvokeSetPwDocumentState {
    <#
    .SYNOPSIS
    Routes ProjectWise workflow state writes through Set-PWDocumentWorkflowStateVerified by default.
    Legacy non-verified writes require qcWorkflow.stateWriteback.useVerified=false and emit QC_LEGACY_UNVERIFIED_STATE_WRITE.
    #>
    param(
        [object]$Document = $null,
        [string]$StateName = '',
        [hashtable]$GuardContext = @{}
    )
    if (-not $GuardContext) { $GuardContext = @{} }
    $cfg = $null
    if ($GuardContext.ContainsKey('config') -and $GuardContext.config) { $cfg = $GuardContext.config }

    if (-not [string]::IsNullOrWhiteSpace($StateName) -and (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue)) {
        try { $StateName = Format-QCWorkflowStateName -StateName $StateName -Config $cfg } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($StateName)) {
        $guardCallSite = '_PWD-InvokeSetPwDocumentState'
        if ($GuardContext.callSite) { $guardCallSite = [string]$GuardContext.callSite }
        _PWD-WriteEmptyStateGuardLog -CallSite $guardCallSite `
            -AuditEventId $GuardContext.auditEventId -DocumentName ([string]$GuardContext.documentName) `
            -FolderPath ([string]$GuardContext.folderPath) -SourceVariableName ([string]$GuardContext.sourceVariableName) `
            -SourceValue $StateName -LivePwState ([string]$GuardContext.livePwState) `
            -ChangedByUsername ([string]$GuardContext.changedByUsername)
        return @{ applied = $false; verified = $false; error = 'empty_target_state' }
    }

    $folderPath = [string]$GuardContext.folderPath
    $documentName = [string]$GuardContext.documentName
    $documentGuid = [string]$GuardContext.documentGuid
    if ($Document) {
        if ([string]::IsNullOrWhiteSpace($documentName)) {
            try { $documentName = [string]$Document.Name } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($documentGuid)) {
            try { $documentGuid = [string]$Document.DocumentGUID } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($folderPath)) {
            try { $folderPath = [string]$Document.FolderPath } catch { }
        }
    }
    $writeScope = if ($GuardContext.writeScope) { [string]$GuardContext.writeScope } else { 'workflow' }

    if (_PWD-TestAutomationDgnProtectedDocument -DocumentName $documentName) {
        _PWD-LogAutomationDgnWriteBlocked -Operation 'workflow_state' -DocumentName $documentName `
            -DocumentGuid $documentGuid -FolderPath $folderPath -CallSite $(if ($GuardContext.callSite) { [string]$GuardContext.callSite } else { '_PWD-InvokeSetPwDocumentState' })
        return @{
            applied = $false
            verified = $false
            skipped = $true
            skipReason = 'dgn_automation_write_blocked'
            documentName = $documentName
            documentGuid = $documentGuid
            targetState = $StateName
        }
    }
    if (_PWD-TestAutomationStemPdfStateWriteBlocked -DocumentName $documentName -TargetState $StateName `
            -WriteScope $writeScope -Config $cfg) {
        $triggerName = if ($GuardContext.triggerDocumentName) { [string]$GuardContext.triggerDocumentName } else { '' }
        _PWD-LogAutomationStemPdfWriteBlocked -Operation 'workflow_state' -DocumentName $documentName `
            -DocumentGuid $documentGuid -FolderPath $folderPath `
            -CallSite $(if ($GuardContext.callSite) { [string]$GuardContext.callSite } else { '_PWD-InvokeSetPwDocumentState' }) `
            -TargetState $StateName -TriggerDocumentName $triggerName
        return @{
            applied = $false
            verified = $false
            skipped = $true
            skipReason = 'stem_pdf_lane_state_write_blocked'
            documentName = $documentName
            documentGuid = $documentGuid
            targetState = $StateName
        }
    }
    $useVerified = $true
    if (Get-Command -Name 'Test-QCWorkflowStateWritebackUseVerified' -ErrorAction SilentlyContinue) {
        try { $useVerified = Test-QCWorkflowStateWritebackUseVerified -Config $cfg } catch { }
    }
    if ($GuardContext.legacyUnverified -eq $true) { $useVerified = $false }

    $dryRun = $false
    if ($GuardContext.ContainsKey('dryRun')) {
        try { $dryRun = [bool]$GuardContext.dryRun } catch { }
    }
    $workflowName = if ($GuardContext.workflowName) { [string]$GuardContext.workflowName } else { '' }
    $qcProcessType = if ($GuardContext.qcProcessType) { [string]$GuardContext.qcProcessType } else { '' }
    $callSite = if ($GuardContext.callSite) { [string]$GuardContext.callSite } else { '_PWD-InvokeSetPwDocumentState' }

    if ($useVerified -and -not [string]::IsNullOrWhiteSpace($folderPath) -and -not [string]::IsNullOrWhiteSpace($documentName)) {
        $isLaneAuthority = ($writeScope -eq 'lane')
        $writeResult = Set-PWDocumentWorkflowStateVerified -Config $(if ($cfg) { $cfg } else { @{} }) `
            -FolderPath $folderPath -DocumentName $documentName -DocumentGuid $documentGuid `
            -TargetState $StateName -QcProcessType $qcProcessType -DryRun:$dryRun `
            -WorkflowName $workflowName -IsLaneAuthority:$isLaneAuthority -WriteScope $writeScope
        if (-not $writeResult.verified) {
            $msg = if (-not [string]::IsNullOrWhiteSpace([string]$writeResult.error)) {
                [string]$writeResult.error
            } else {
                "Workflow state write unverified for '$documentName' (read-back: '$($writeResult.readBackState)', target: '$StateName')."
            }
            throw $msg
        }
        return $writeResult
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Warning' -Code 'QC_LEGACY_UNVERIFIED_STATE_WRITE' `
            -Message 'Using legacy non-verified Set-PWDocumentState path.' -Data @{
            callSite = $callSite
            documentName = $documentName
            documentGuid = $documentGuid
            folderPath = $folderPath
            targetState = $StateName
            useVerified = $false
            missingFolderPath = [string]::IsNullOrWhiteSpace($folderPath)
            missingDocumentName = [string]::IsNullOrWhiteSpace($documentName)
        } | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($folderPath) -and -not [string]::IsNullOrWhiteSpace($documentName)) {
        $cleanDoc = _PWD-GetCleanPwDocumentForStateChange -FolderPath $folderPath -DocumentName $documentName -DocumentGuid $documentGuid
        if ($cleanDoc) { $Document = $cleanDoc }
    }

    $cmd = Get-Command -Name 'Set-PWDocumentState' -ErrorAction SilentlyContinue
    if (-not $cmd -or -not $Document) { throw 'Set-PWDocumentState or document unavailable.' }

    $useForce = $true
    if (Get-Command -Name 'Test-QCWorkflowStateWritebackUseForce' -ErrorAction SilentlyContinue) {
        try { $useForce = Test-QCWorkflowStateWritebackUseForce -Config $cfg } catch { }
    }

    $args = @{ ErrorAction = 'Stop' }
    $docParam = if ($cmd.Parameters.ContainsKey('InputDocuments')) { 'InputDocuments' }
        elseif ($cmd.Parameters.ContainsKey('InputDocument')) { 'InputDocument' }
        elseif ($cmd.Parameters.ContainsKey('Document')) { 'Document' }
        else { $null }
    $stateParam = if ($cmd.Parameters.ContainsKey('State')) { 'State' }
        elseif ($cmd.Parameters.ContainsKey('StateName')) { 'StateName' }
        else { $null }
    if (-not $docParam -or -not $stateParam) { throw 'Set-PWDocumentState missing document/state parameters.' }
    $args[$docParam] = @($Document)
    $args[$stateParam] = $StateName
    if ($useForce -and $cmd.Parameters.ContainsKey('Force')) { $args['Force'] = $true }
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) { $args['ReturnBoolean'] = $true }
    $result = & $cmd @args
    if ($cmd.Parameters.ContainsKey('ReturnBoolean')) {
        try {
            if ($result -eq $false) {
                throw "Set-PWDocumentState returned false for state '$StateName'."
            }
        } catch {
            if ($_.Exception.Message -match 'returned false') { throw }
        }
    }

    return @{
        applied = $true
        verified = $false
        legacyUnverified = $true
        documentName = $documentName
        documentGuid = $documentGuid
        targetState = $StateName
        usedForce = [bool]$useForce
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
    $nameCandidates = [System.Collections.Generic.List[string]]::new()
    $seenNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    function Add-NameCandidate([string]$Candidate) {
        if ([string]::IsNullOrWhiteSpace($Candidate)) { return }
        $trimmed = $Candidate.Trim()
        if ($seenNames.Add($trimmed)) { [void]$nameCandidates.Add($trimmed) }
    }
    Add-NameCandidate $DocumentName
    if ($DocumentName -match '(?i)^(.+)-(prod|chk|rev)\.pdf$') {
        Add-NameCandidate ([string]$Matches[1] + '-' + [string]$Matches[2].ToLowerInvariant() + '.pdf')
    }
    foreach ($searchName in @($nameCandidates)) {
        try {
            $params = @{
                FolderPath     = $apiPath
                JustThisFolder = $true
                DocumentName   = $searchName
                ErrorAction    = 'Stop'
            }
            if ($searchCmd.Parameters.ContainsKey('PopulatePath')) { $params['PopulatePath'] = $true }
            $found = (& $searchCmd @params | Select-Object -First 1)
            if ($found) { return $found }
        } catch { }
    }
    if (Get-Command -Name 'Get-PWDocumentsInFolder' -ErrorAction SilentlyContinue) {
        try {
            $all = @(Get-PWDocumentsInFolder -FolderPath $apiPath -ErrorAction SilentlyContinue)
            foreach ($doc in $all) {
                $actualName = ''
                try { $actualName = [string]$doc.Name } catch { }
                if ([string]::IsNullOrWhiteSpace($actualName)) { continue }
                foreach ($searchName in @($nameCandidates)) {
                    if ($actualName.Equals($searchName, [StringComparison]::OrdinalIgnoreCase)) {
                        return $doc
                    }
                }
            }
        } catch { }
    }
    return $null
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
    OR LOWER(document_name) = LOWER(@prodName)
    OR LOWER(document_name) = LOWER(@chkName)
    OR LOWER(document_name) = LOWER(@revName)
  )
ORDER BY last_updated_at DESC
"@ -Parameters @{
                    folderPath = $FolderPath
                    pdfName    = $expectedNames[0]
                    dgnName    = $expectedNames[1]
                    prodName   = $expectedNames[2]
                    chkName    = if ($expectedNames.Count -gt 3) { $expectedNames[3] } else { '' }
                    revName    = if ($expectedNames.Count -gt 4) { $expectedNames[4] } else { '' }
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
                sourceType   = if ($name -match '(?i)\.dgn$') { 'dgn' } elseif ($name -match '(?i)-(prod|chk|rev)\.pdf$') { 'pdf' } else { 'pdf' }
            }
        }
    }

    $docByGuid = _PWD-LoadPwDocumentsByGuid -DocumentGuids @($guidsToLoad)
    $resolved = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $members.Values) {
        $dn = [string]$entry.documentName
        $dg = [string]$entry.documentGuid
        if ($dn -match '(?i)-(prod|chk|rev)\.pdf$') {
            if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
                if ($dn.Equals($DocumentName, [StringComparison]::OrdinalIgnoreCase)) {
                    $dg = [string]$DocumentGuid
                }
            }
            if (Get-Command -Name 'Resolve-QCSheetQcPdfGuid' -ErrorAction SilentlyContinue) {
                try {
                    $liveLaneGuid = Resolve-QCSheetQcPdfGuid -Config $Config -FolderPath $FolderPath `
                        -QcPdfName $dn -SourceDocumentGuid $DocumentGuid
                    if (-not [string]::IsNullOrWhiteSpace($liveLaneGuid)) { $dg = $liveLaneGuid.Trim() }
                } catch { }
            }
        }
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

    $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    if ($resolved.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($sheetStem)) {
        $stateByGuid = @{}
        foreach ($r in $resolved) {
            $dg = [string]$r.documentGuid
            if (-not $dg) { continue }
            try {
                $st = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName ([string]$r.documentName) -DocumentGuid $dg
                if ($st) { $stateByGuid[$dg.ToLowerInvariant()] = $st }
            } catch { }
        }
        _PWD-SyncSheetPackageLaneQcPdfs -Config $Config -FolderPath $FolderPath -SheetStem $sheetStem `
            -Members @($resolved) -StateByGuid $stateByGuid
    }
    return @($resolved)
}

function _PWD-EnsureLaneQcPdfProcessTypeAttribute {
    <#
    .SYNOPSIS
    Sets QC_Process_Type on a lane QC PDF from the triggering lane type when unset or incorrect.
    Skips only when the current value already matches the expected lane type.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$LanePdfName,
        [Parameter(Mandatory)][string]$QcProcessType,
        [bool]$DryRun = $false
    )

    $laneType = [string]$QcProcessType
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType $laneType -AllowNullOnEmpty
        if ($norm) { $laneType = $norm }
    }
    if ([string]::IsNullOrWhiteSpace($laneType) -or [string]::IsNullOrWhiteSpace($LanePdfName)) {
        return @{ ensured = $false; skipped = 'missing_input' }
    }

    $targetDisplay = $laneType
    if (Get-Command -Name 'Format-QCProcessTypeAttributeValue' -ErrorAction SilentlyContinue) {
        $targetDisplay = Format-QCProcessTypeAttributeValue -ProcessType $laneType
    }

    $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $LanePdfName -DocumentGuid ''
    if (-not $doc) {
        return @{ ensured = $false; skipped = 'lane_doc_not_found'; lanePdfName = $LanePdfName }
    }

    $processCol = Get-PWQcProcessTypeAttributeName -Config $Config
    $currentPw = ''
    $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $LanePdfName -ColumnsToReturn @($processCol)
    if ($attrRead.found) {
        $currentPw = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $processCol
    }
    $currentCanonical = _PWD-ResolveCanonicalProcessTypeForSync -RawValue $currentPw -Config $Config
    $expectedCanonical = _PWD-ResolveCanonicalProcessTypeForSync -RawValue $laneType -Config $Config
    if (-not [string]::IsNullOrWhiteSpace($currentCanonical) -and -not [string]::IsNullOrWhiteSpace($expectedCanonical) `
        -and $currentCanonical -eq $expectedCanonical) {
        return @{
            ensured = $false
            skipped = 'lane_process_type_already_set'
            lanePdfName = $LanePdfName
            currentValue = [string]$currentPw
        }
    }
    $correctingWrongValue = -not [string]::IsNullOrWhiteSpace($currentPw)

    $pwWritesEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath
    if (-not $pwWritesEnabled) {
        return @{ ensured = $false; skipped = 'qc_process_type_not_enabled_for_environment'; lanePdfName = $LanePdfName }
    }
    if ($DryRun) {
        return @{ ensured = $false; planned = $true; lanePdfName = $LanePdfName; toValue = $targetDisplay }
    }

    try {
        $attrs = @{ $processCol = $targetDisplay }
        [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes $attrs -Config $Config)
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_LANE_PROCESS_TYPE_ENSURED' `
                -Message $(if ($correctingWrongValue) {
                    'Corrected lane QC PDF QC_Process_Type from triggering process type.'
                } else {
                    'Set lane QC PDF QC_Process_Type from triggering process type.'
                }) -Data @{
                folderPath = $FolderPath
                lanePdfName = $LanePdfName
                qcProcessType = $laneType
                fromValue = [string]$currentPw
                toValue = $targetDisplay
                corrected = [bool]$correctingWrongValue
            } | Out-Null
        }
        return @{
            ensured = $true
            lanePdfName = $LanePdfName
            toValue = $targetDisplay
            qcProcessType = $laneType
            fromValue = [string]$currentPw
            corrected = [bool]$correctingWrongValue
        }
    } catch {
        return @{ ensured = $false; error = [string]$_.Exception.Message; lanePdfName = $LanePdfName }
    }
}

function Sync-PWPostInitialPrependLaneStates {
    <#
    .SYNOPSIS
    After initial QC prepend (Initiate Origination), sets the active lane QC PDF to Originated and returns the stem PDF to In Development.
    .DESCRIPTION
    The lane PDF is the process record. The stem sheet PDF is returned to the reference state (In Development)
    using the same verified Set-PWDocumentState pattern. DGN and other lane PDFs are not written.
    Success notification requires verified lane PDF state.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$DocumentName,
        [string]$DocumentGuid = '',
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$QcProcessType,
        [Parameter(Mandatory)][string]$LaneTargetState,
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$ReferenceState,
        [string]$ExpectedLanePdfName = '',
        [bool]$WriteStemPdfReferenceState = $true,
        [bool]$DryRun = $false
    )

    $laneType = [string]$QcProcessType
    if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
        $norm = Normalize-QCProcessType -ProcessType $laneType -AllowNullOnEmpty
        if ($norm) { $laneType = $norm }
    }
    if ([string]::IsNullOrWhiteSpace($laneType)) {
        return @{ updates = @(); skipped = $true; skipReason = 'empty_qc_process_type' }
    }

    $laneTarget = ([string]$LaneTargetState).Trim()
    $referenceTarget = ([string]$ReferenceState).Trim()
    if (Get-Command -Name 'Format-QCWorkflowStateName' -ErrorAction SilentlyContinue) {
        try {
            $laneTarget = Format-QCWorkflowStateName -StateName $laneTarget -Config $Config
            $referenceTarget = Format-QCWorkflowStateName -StateName $referenceTarget -Config $Config
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($laneTarget) -or ($WriteStemPdfReferenceState -and [string]::IsNullOrWhiteSpace($referenceTarget))) {
        return @{ updates = @(); skipped = $true; skipReason = 'empty_target_state' }
    }

    # Lane PDF is workflow authority; stem PDF may return to reference state on initial prepend only.
    $writeStemPdfReferenceState = [bool]$WriteStemPdfReferenceState

    $sheetStem = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }
    if ([string]::IsNullOrWhiteSpace($sheetStem)) {
        $sheetStem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
    }

    $lanePdfName = ([string]$ExpectedLanePdfName).Trim()
    if ([string]::IsNullOrWhiteSpace($lanePdfName) -and (Get-Command -Name 'Get-QCLaneQcPdfExpectedName' -ErrorAction SilentlyContinue)) {
        $lanePdfName = Get-QCLaneQcPdfExpectedName -SheetBaseName $sheetStem -ProcessType $laneType -Config $Config
    }
    if ([string]::IsNullOrWhiteSpace($lanePdfName)) {
        $suffix = 'prod'
        if (Get-Command -Name 'Get-QCProcessTypePdfSuffix' -ErrorAction SilentlyContinue) {
            $resolvedSuffix = Get-QCProcessTypePdfSuffix -ProcessType $laneType -Config $Config
            if (-not [string]::IsNullOrWhiteSpace($resolvedSuffix)) { $suffix = [string]$resolvedSuffix }
        }
        $lanePdfName = ($sheetStem + '-' + $suffix + '.pdf')
    }

    $stemPdfName = ($sheetStem + '.pdf')
    $stateUpdates = [System.Collections.Generic.List[object]]::new()
    $workflowName = ''
    if (Get-Command -Name 'Get-QCWorkflowSettings' -ErrorAction SilentlyContinue) {
        try { $workflowName = [string](Get-QCWorkflowSettings -Config $Config).expectedWorkflowName } catch { }
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_LANE_WORKFLOW_WRITEBACK_TARGET_RESOLVED' `
            -Message 'Resolved lane workflow writeback target for initial prepend.' -Data @{
            folderPath = $FolderPath
            sheetStem = $sheetStem
            qcProcessType = $laneType
            lanePdfName = $lanePdfName
            stemPdfName = $stemPdfName
            laneTargetState = $laneTarget
            referenceState = $referenceTarget
            writeStemPdfReferenceState = [bool]$writeStemPdfReferenceState
            workflowName = $workflowName
        } | Out-Null
    }

    function _PWD-ApplyPostPrependState {
        param(
            [string]$TargetName,
            [string]$TargetGuid,
            [string]$TargetState,
            [bool]$IsLaneAuthority = $false
        )
        if ([string]::IsNullOrWhiteSpace($TargetName)) { return }
        $writeResult = Set-PWDocumentWorkflowStateVerified -Config $Config -FolderPath $FolderPath `
            -DocumentName $TargetName -DocumentGuid $TargetGuid -TargetState $TargetState `
            -QcProcessType $laneType -DryRun:$DryRun -WorkflowName $workflowName `
            -IsLaneAuthority:$IsLaneAuthority -WriteScope $(if ($IsLaneAuthority) { 'lane' } else { 'stem' })
        $change = @{
            documentGuid = [string]$writeResult.documentGuid
            documentName = $TargetName
            fromState = [string]$writeResult.fromState
            toState = $TargetState
            applied = [bool]$writeResult.applied
            planned = [bool]$writeResult.planned
            verified = [bool]$writeResult.verified
            readBackState = [string]$writeResult.readBackState
            usedForce = [bool]$writeResult.usedForce
            commandShape = [string]$writeResult.commandShape
            isLaneAuthority = [bool]$IsLaneAuthority
            isStemReference = -not [bool]$IsLaneAuthority
            qcProcessType = $laneType
            attemptedCommand = 'Set-PWDocumentState'
            workflowName = $workflowName
            error = [string]$writeResult.error
        }
        if ($change.verified -and -not $DryRun -and [string]::IsNullOrWhiteSpace([string]$writeResult.error)) {
            $dg = [string]$change.documentGuid
            if ($dg -and (Get-Command -Name 'Update-QCSheetIndexPwStateName' -ErrorAction SilentlyContinue)) {
                try { [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $dg -PwStateName $TargetState) } catch { }
            }
            if ($IsLaneAuthority -and $dg -and (Get-Command -Name 'Update-SheetPackageQcPdfLaneState' -ErrorAction SilentlyContinue)) {
                $laneForRow = Get-PWQcPdfLaneFromDocumentName -DocumentName $TargetName
                if ($laneForRow) {
                    try {
                        [void](Update-SheetPackageQcPdfLaneState -Config $Config -DocumentGuid $dg `
                            -CurrentPwState $TargetState -QcProcessType $laneForRow)
                    } catch { }
                }
            }
        }
        $stateUpdates.Add($change) | Out-Null
    }

    if (-not $DryRun) {
        $processTypeReset = @{ skipped = $true; reason = 'sync_unavailable' }
        if (Get-Command -Name '_PWD-SyncReferenceSheetProcessTypeAttributes' -ErrorAction SilentlyContinue) {
            $resetGuid = if ($DocumentGuid) { [string]$DocumentGuid } else { '' }
            $resetName = if ($DocumentName -match '(?i)\.pdf$' -and $DocumentName -notmatch '(?i)-(prod|chk|rev)\.pdf$') {
                [string]$DocumentName
            } else {
                $sheetStem + '.pdf'
            }
            try {
                $processTypeReset = _PWD-SyncReferenceSheetProcessTypeAttributes -Config $Config -DocumentGuid $resetGuid `
                    -DocumentName $resetName -FolderPath $FolderPath -CanonicalProcessType 'production' `
                    -DryRun:$DryRun -ControlDocumentOnly
            } catch {
                $processTypeReset = @{ skipped = $true; error = [string]$_.Exception.Message }
            }
        }
        $laneProcessType = _PWD-EnsureLaneQcPdfProcessTypeAttribute -Config $Config -FolderPath $FolderPath `
            -LanePdfName $lanePdfName -QcProcessType $laneType -DryRun:$DryRun
    } else {
        $processTypeReset = @{ planned = $true }
        $laneProcessType = @{ planned = $true }
    }

    $laneGuid = ''
    if (Get-Command -Name 'Sync-QCLaneQcPdfGuidFromProjectWise' -ErrorAction SilentlyContinue) {
        try {
            $reg = Sync-QCLaneQcPdfGuidFromProjectWise -Config $Config -FolderPath $FolderPath -QcPdfName $lanePdfName `
                -QcProcessType $laneType -SourceDocumentGuid $DocumentGuid -CurrentPwState $laneTarget
            if ($reg.IsSuccess -and $reg.Data -and $reg.Data.documentGuid) {
                $laneGuid = [string]$reg.Data.documentGuid
            }
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($laneGuid)) {
        $laneDoc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $lanePdfName -DocumentGuid ''
        if ($laneDoc) {
            try { $laneGuid = [string]$laneDoc.DocumentGUID } catch { }
        }
    }

    _PWD-ApplyPostPrependState -TargetName $lanePdfName -TargetGuid $laneGuid -TargetState $laneTarget -IsLaneAuthority:$true

    if ($writeStemPdfReferenceState) {
        $stemGuid = ''
        if ($DocumentName -match '(?i)\.pdf$' -and $DocumentName -notmatch '(?i)-(prod|chk|rev)\.pdf$') {
            $stemGuid = [string]$DocumentGuid
        }
        _PWD-ApplyPostPrependState -TargetName $stemPdfName -TargetGuid $stemGuid -TargetState $referenceTarget -IsLaneAuthority:$false
    }

    $laneStateVerified = $false
    $laneUpdateFound = $false
    $unverifiedUpdates = [System.Collections.Generic.List[object]]::new()
    foreach ($upd in @($stateUpdates)) {
        $u = $upd
        if ($u -isnot [hashtable]) {
            try { $u = @{}; foreach ($p in $upd.PSObject.Properties) { $u[$p.Name] = $p.Value } } catch { continue }
        }
        if ([string]$u.documentName -ne $lanePdfName) { continue }
        $laneUpdateFound = $true
        if ($u.verified -eq $true -or [string]$u.skipped -eq 'already_at_target') {
            $laneStateVerified = $true
        } elseif ($u.applied -eq $true -and $u.verified -eq $false) {
            $unverifiedUpdates.Add($u) | Out-Null
        }
    }
    $allVerified = $laneUpdateFound -and ($unverifiedUpdates.Count -eq 0) -and $laneStateVerified

    $stemStateVerified = $false
    foreach ($upd in @($stateUpdates)) {
        $u = $upd
        if ($u -isnot [hashtable]) {
            try { $u = @{}; foreach ($p in $upd.PSObject.Properties) { $u[$p.Name] = $p.Value } } catch { continue }
        }
        if ([string]$u.documentName -ne $stemPdfName) { continue }
        if ($u.verified -eq $true) { $stemStateVerified = $true }
        break
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Level 'Information' -Code 'QC_POST_PREPEND_LANE_SPLIT' `
            -Message 'Applied lane-independent post-prepend workflow writeback.' -Data @{
            folderPath = $FolderPath
            sheetStem = $sheetStem
            qcProcessType = $laneType
            lanePdfName = $lanePdfName
            stemPdfName = $stemPdfName
            laneTargetState = $laneTarget
            referenceState = $referenceTarget
            writeStemPdfReferenceState = [bool]$writeStemPdfReferenceState
            updateCount = $stateUpdates.Count
            laneStateVerified = [bool]$laneStateVerified
            stemStateVerified = [bool]$stemStateVerified
            allVerified = [bool]$allVerified
        } | Out-Null
    }

    return @{
        updates = @($stateUpdates)
        lanePdfName = $lanePdfName
        stemPdfName = $stemPdfName
        qcProcessType = $laneType
        laneTargetState = $laneTarget
        referenceState = $referenceTarget
        writeStemPdfReferenceState = [bool]$writeStemPdfReferenceState
        writeReferenceStates = [bool]$writeStemPdfReferenceState
        laneProcessTypeEnsure = $laneProcessType
        processTypeReset = $processTypeReset
        laneStateVerified = [bool]$laneStateVerified
        stemStateVerified = [bool]$stemStateVerified
        allVerified = [bool]$allVerified
        unverifiedUpdates = @($unverifiedUpdates)
    }
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

    if (-not (_PWD-TestLegacySiblingStateSyncEnabled -Config $Config)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Level 'Information' -Code 'QC_SIBLING_STATE_SYNC_DISABLED' `
                -Message 'Post-prepend sibling workflow state sync skipped; lanes are independent.' -Data @{
                folderPath = $FolderPath
                triggerDocumentName = $DocumentName
                triggerDocumentGuid = $DocumentGuid
                targetState = $target
                triggerSource = $TriggerSource
            } | Out-Null
        }
        if (Get-Command -Name 'Update-SheetPackageQcPdfLaneState' -ErrorAction SilentlyContinue) {
            $lane = 'production'
            if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
                $resolvedLane = Get-PWQcPdfLaneFromDocumentName -DocumentName $DocumentName
                if ($resolvedLane) { $lane = $resolvedLane }
            }
            [void](Update-SheetPackageQcPdfLaneState -Config $Config -DocumentGuid $DocumentGuid `
                -CurrentPwState $target -QcProcessType $lane)
        }
        return @{ updates = @(); memberCount = 0; siblingSyncDisabled = $true; targetState = $target }
    }

    $members = @(Get-PWAssociatedSheetSyncMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -TriggerSource $TriggerSource)
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
        if (_PWD-TestAutomationDgnProtectedDocument -DocumentName $dn) { continue }
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
            $writeResult = _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $target -GuardContext @{
                callSite = 'Sync-PWAssociatedSheetMembersToWorkflowState'
                documentName = $dn
                folderPath = $FolderPath
                documentGuid = $dg
                config = $Config
                sourceVariableName = 'target'
                livePwState = [string]$currentState
                writeScope = 'workflow'
            }
            $change.applied = [bool]$writeResult.applied
            $change.verified = [bool]$writeResult.verified
            $change.readBackState = [string]$writeResult.readBackState
            $change.stateAfter = [string]$writeResult.readBackState
            if ($change.applied -and $change.verified -eq $false -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_SHEET_STATE_VERIFY_MISMATCH' `
                    -Message 'Workflow state write did not verify after read-back.' -Data @{
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
    Legacy DOCUMENT_ATTR sibling QC process type sync (disabled by default).
    .DESCRIPTION
    When EnableLegacyReviewTypeAttributeSync is true, propagates a resolved QC_Process_Type to associated stem
    members. Lane QC PDFs are excluded. QC_Review_Type is read-only and is never written.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$CanonicalReviewType,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [bool]$DryRun = $false
    )

    if (-not (Get-Command -Name 'Test-QCLegacyReviewTypeAttributeSyncEnabled' -ErrorAction SilentlyContinue) `
        -or -not (Test-QCLegacyReviewTypeAttributeSyncEnabled -Config $Config)) {
        _PWD-LogProcessTypeSyncSkipped -Reason 'legacy_sync_disabled' -Config $Config `
            -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
            -ExtraData @{ source = 'legacy_review_type_attribute_sync' }
        return
    }

    $canonicalNorm = _PWD-ResolveCanonicalProcessTypeForSync -RawValue $CanonicalReviewType -Config $Config
    if (-not $canonicalNorm) {
        _PWD-LogProcessTypeUnknown -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -RawValue $CanonicalReviewType -Source 'legacy_review_type_attribute_sync'
        _PWD-LogProcessTypeSyncSkipped -Reason 'null_canonical_process_type' -Config $Config `
            -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
            -ExtraData @{ source = 'legacy_review_type_attribute_sync'; rawValue = [string]$CanonicalReviewType }
        return
    }

    $canonicalDisplay = $canonicalNorm
    if (Get-Command -Name 'Format-QCProcessTypeAttributeValue' -ErrorAction SilentlyContinue) {
        $canonicalDisplay = Format-QCProcessTypeAttributeValue -ProcessType $canonicalNorm
    }

    $pwWritesEnabled = Test-PWQcReviewTypeAttributesEnabled -Config $Config -FolderPath $FolderPath
    $processCol = Get-PWQcProcessTypeAttributeName -Config $Config
    $members = @(Get-PWAssociatedSheetSyncMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid -TriggerSource 'process_type_sync')
    if ($members.Count -eq 0) { return }

    $stateUpdates = @()
    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        $doc = $member.document
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (_PWD-TestAutomationDgnProtectedDocument -DocumentName $dn) { continue }

        if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
            $memberLane = Get-PWQcPdfLaneFromDocumentName -DocumentName $dn
            if ($memberLane) { continue }
        }

        $prevDb = _PWD-GetSheetIndexQcReviewType -Config $Config -DocumentGuid $dg
        $currentPw = ''
        $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $dn -ColumnsToReturn @($processCol)
        if ($attrRead.found) {
            $currentPw = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $processCol
            if (-not $doc -and $attrRead.document) { $doc = $attrRead.document }
        }
        if (-not $doc) {
            $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }
        if (-not $doc) { continue }

        $pwNeedsWrite = (_PWD-NormalizeSheetIndexValue $currentPw) -ne (_PWD-NormalizeSheetIndexValue $canonicalNorm)
        $indexNeedsWrite = (_PWD-NormalizeSheetIndexValue $prevDb) -ne (_PWD-NormalizeSheetIndexValue $canonicalNorm)
        if (-not $pwNeedsWrite -and -not $indexNeedsWrite) { continue }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromValue    = [string]$currentPw
            toValue      = $canonicalDisplay
            applied      = $false
            planned      = $false
            indexUpdated = $false
        }

        if ($pwNeedsWrite) {
            if (-not $pwWritesEnabled) {
                $change.pwWriteSkipped = 'qc_process_type_not_enabled_for_environment'
            } elseif ($DryRun) {
                $change.planned = $true
            } else {
                try {
                    [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes @{ $processCol = $canonicalNorm } -Config $Config)
                    $change.applied = $true
                } catch {
                    $change.error = [string]$_.Exception.Message
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_REVIEW_TYPE_SYNC_FAILED' `
                            -Message 'Failed to align associated sheet QC_Process_Type (legacy sync).' -Data @{
                            documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                            fromValue = [string]$currentPw; toValue = $canonicalDisplay; error = [string]$_.Exception.Message
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
                        -WatchRoot $WatchRoot -Extension $ext -SourceType $sourceType -QcReviewType $canonicalDisplay `
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
                    qc_review_type = @{ oldValue = $prevAudit; newValue = $canonicalDisplay }
                } | Out-Null
            }
        }

        $stateUpdates += $change
    }

    if ($stateUpdates.Count -gt 0 -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_REVIEW_TYPE_SYNC' `
            -Message 'Associated sheet QC_Process_Type aligned from DOCUMENT_ATTR audit event (legacy sync enabled).' -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalProcessType = $canonicalNorm
            canonicalReviewType = $canonicalNorm
            processTypeColumn   = $processCol
            pwWritesEnabled     = $pwWritesEnabled
            memberCount         = $members.Count
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
        } | Out-Null
    }
}

function Sync-PWAssociatedSheetEmailAttributes {
    <#
    .SYNOPSIS
    Propagates role emails from the DGN-first canonical source to associated DGN, sheet PDF, and QC PDF siblings.
    .DESCRIPTION
    Updates ProjectWise environment attributes and sheet_index role email columns for every associated member
    when canonical designer, reviewer, or checker emails differ. Mirrors sibling QC_Review_Type sync.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$WatchRoot = '',
        [string]$LastAuditEventAt = '',
        [bool]$DryRun = $false
    )

    $emailCols = _PWD-GetSheetRoleEmailColumnNames -Config $Config
    $designerCol = [string]$emailCols.designer
    $reviewerCol = [string]$emailCols.reviewer
    $checkerCol = [string]$emailCols.checker

    $canonical = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $DocumentName -Config $Config
    if (-not $canonical.found) { return }

    $canonicalDesigner = [string]$canonical.designerEmail
    $canonicalReviewer = [string]$canonical.reviewerEmail
    $canonicalChecker = [string]$canonical.checkerEmail

    $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    if ($members.Count -eq 0) { return }

    $readCols = @($designerCol, $reviewerCol, $checkerCol) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $stateUpdates = @()
    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        $doc = $member.document
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }

        $prevDbDesigner = ''
        $prevDbReviewer = ''
        $prevDbChecker = ''
        if ($dg -and (Get-Command -Name 'Invoke-QCDatabaseQuery' -ErrorAction SilentlyContinue) -and (Test-QCDatabaseEnabled -Config $Config)) {
            try {
                $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, checker_email
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $dg }
                if ($dbRes.IsSuccess -and $dbRes.Data.table -and $dbRes.Data.table.Rows.Count -gt 0) {
                    $r = $dbRes.Data.table.Rows[0]
                    if (-not ($r.designer_email -is [DBNull])) { $prevDbDesigner = [string]$r.designer_email }
                    if (-not ($r.reviewer_email -is [DBNull])) { $prevDbReviewer = [string]$r.reviewer_email }
                    if ($r.Table.Columns.Contains('checker_email') -and -not ($r.checker_email -is [DBNull])) { $prevDbChecker = [string]$r.checker_email }
                }
            } catch { }
        }

        $currentDesigner = ''
        $currentReviewer = ''
        $currentChecker = ''
        $attrRead = Get-PWDocumentAttributesByColumns -FolderPath $FolderPath -DocumentName $dn -ColumnsToReturn $readCols
        if ($attrRead.found) {
            $currentDesigner = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $designerCol
            $currentReviewer = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $reviewerCol
            $currentChecker = _PWD-GetPwAttributeValue -PwAttributes $attrRead.attributes -ColumnName $checkerCol
            if (-not $doc -and $attrRead.document) { $doc = $attrRead.document }
        }
        if (-not $doc) {
            $doc = _PWD-ResolvePwDocumentInFolder -DocByGuid @{} -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg
        }
        if (-not $doc) { continue }

        $toWrite = @{}
        if ((_PWD-NormalizeSheetIndexValue $currentDesigner) -ne (_PWD-NormalizeSheetIndexValue $canonicalDesigner)) {
            $toWrite[$designerCol] = $canonicalDesigner
        }
        if ((_PWD-NormalizeSheetIndexValue $currentReviewer) -ne (_PWD-NormalizeSheetIndexValue $canonicalReviewer)) {
            $toWrite[$reviewerCol] = $canonicalReviewer
        }
        if ((_PWD-NormalizeSheetIndexValue $currentChecker) -ne (_PWD-NormalizeSheetIndexValue $canonicalChecker)) {
            $toWrite[$checkerCol] = $canonicalChecker
        }

        $indexNeedsWrite = (_PWD-NormalizeSheetIndexValue $prevDbDesigner) -ne (_PWD-NormalizeSheetIndexValue $canonicalDesigner) `
            -or (_PWD-NormalizeSheetIndexValue $prevDbReviewer) -ne (_PWD-NormalizeSheetIndexValue $canonicalReviewer) `
            -or (_PWD-NormalizeSheetIndexValue $prevDbChecker) -ne (_PWD-NormalizeSheetIndexValue $canonicalChecker)
        if ($toWrite.Keys.Count -eq 0 -and -not $indexNeedsWrite) { continue }

        $change = @{
            documentGuid = $dg
            documentName = $dn
            fromValue    = @{ designer = $currentDesigner; reviewer = $currentReviewer; checker = $currentChecker }
            toValue      = @{ designer = $canonicalDesigner; reviewer = $canonicalReviewer; checker = $canonicalChecker }
            applied      = $false
            planned      = $false
            indexUpdated = $false
            attributesWritten = @($toWrite.Keys)
        }

        if ($toWrite.Keys.Count -gt 0) {
            if ($DryRun) {
                $change.planned = $true
            } else {
                try {
                    [void](_PWD-InvokeUpdatePWDocumentAttributes -Document $doc -Attributes $toWrite)
                    $change.applied = $true
                } catch {
                    $change.error = [string]$_.Exception.Message
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_SHEET_EMAIL_SYNC_FAILED' `
                            -Message 'Failed to align associated sheet role emails.' -Data @{
                            documentGuid = $dg; documentName = $dn; folderPath = $FolderPath
                            toValue = $change.toValue; error = [string]$_.Exception.Message
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
                        -WatchRoot $WatchRoot -Extension $ext -SourceType $sourceType `
                        -DesignerEmail $canonicalDesigner -ReviewerEmail $canonicalReviewer -CheckerEmail $canonicalChecker `
                        -LastAuditEventAt $LastAuditEventAt -SetOwnershipFromProjectWise | Out-Null
                    $change.indexUpdated = $true
                } catch { }
            }
        }

        if (Get-Command -Name 'Invoke-QCAuditWorkflowAttributeChangeTriggers' -ErrorAction SilentlyContinue) {
            $fieldChanges = @{}
            if ((_PWD-NormalizeSheetIndexValue $currentDesigner) -ne (_PWD-NormalizeSheetIndexValue $canonicalDesigner)) {
                $fieldChanges['designer_email'] = @{ oldValue = [string]$currentDesigner; newValue = $canonicalDesigner }
            }
            if ((_PWD-NormalizeSheetIndexValue $currentReviewer) -ne (_PWD-NormalizeSheetIndexValue $canonicalReviewer)) {
                $fieldChanges['reviewer_email'] = @{ oldValue = [string]$currentReviewer; newValue = $canonicalReviewer }
            }
            if ((_PWD-NormalizeSheetIndexValue $currentChecker) -ne (_PWD-NormalizeSheetIndexValue $canonicalChecker)) {
                $fieldChanges['checker_email'] = @{ oldValue = [string]$currentChecker; newValue = $canonicalChecker }
            }
            if ($fieldChanges.Count -gt 0) {
                Invoke-QCAuditWorkflowAttributeChangeTriggers -Config $Config -DocumentGuid $dg -DocumentName $dn `
                    -FolderPath $FolderPath -FieldChanges $fieldChanges | Out-Null
            }
        }

        $stateUpdates += $change
    }

    if ($stateUpdates.Count -gt 0 -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_EMAIL_SYNC' `
            -Message 'Associated sheet role emails aligned from DOCUMENT_ATTR audit event.' -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalDesigner   = $canonicalDesigner
            canonicalReviewer   = $canonicalReviewer
            canonicalChecker    = $canonicalChecker
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
                foreach ($key in @('production','qcInitiated','qcReceived','readyForQc','redlinesReceived','correctionsInProgress','qcFinalizing','readyForVerification','complete','error')) {
                    try { & $add (Get-QCWorkflowStateName -Settings $wf -StateKey $key) } catch { }
                }
            }
        } catch { }
    }
    foreach ($fallback in @('In Production','QC Initiated','Ready for QC','Redlines Received','Corrections In Progress','Initiate Verification','Ready for Verification','QC Complete','Error Needs Attention')) {
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

function _PWD-EnsureAuditWorkflowExports {
    <#
    .SYNOPSIS
    Ensures QC.AuditTriggers exports survive nested Import-Module -Force in the watcher session.
    #>
    [CmdletBinding()]
    param()

    $required = @(
        'Get-QCAuditWorkflowTriggerSettings'
        'Invoke-QCAuditWorkflowStateChangeTriggers'
        'Invoke-QCSheetGroupWorkflowTransition'
        'Test-QCDocumentStateAuditEventIsStale'
        'Test-QCIsQcPdfDocumentName'
    )
    $restoreOrder = @(
        'QC.ProcessType.psm1'
        'Core.Database.psm1'
        'QC.AuditTriggers.psm1'
        'QC.Notifications.psm1'
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
        if (Get-Command -Name 'Ensure-PWDiscoveryModuleLoaded' -ErrorAction SilentlyContinue) {
            [void](Ensure-PWDiscoveryModuleLoaded)
        }
    }
    return (@($required | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }).Count -eq 0)
}

function _PWD-ResolveSheetPackageNotificationDocumentGuid {
    param([array]$Members)

    foreach ($m in @($Members)) {
        $dn = [string]$m.documentName
        if ([string]::IsNullOrWhiteSpace($dn)) { continue }
        if (Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) {
            if (Test-QCIsQcPdfDocumentName -DocumentName $dn) { return [string]$m.documentGuid }
        }
    }
    foreach ($m in @($Members)) {
        $dn = [string]$m.documentName
        if (Test-QCIsSheetPdfDocumentName -DocumentName $dn) { return [string]$m.documentGuid }
    }
    if (@($Members).Count -gt 0) { return [string]$Members[0].documentGuid }
    return ''
}

function _PWD-TryEnqueueLaneQcPrependFromState {
    <#
    .SYNOPSIS
    Enqueues QC_PREPEND for lane QC PDF automation intake states (Initiate Origination / Initiate Verification).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$CanonicalState,
        [string]$PreviousLaneState = '',
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = ''
    )

    $isInitiated = $false
    $isFinalizing = $false
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue) {
        $isInitiated = Test-QCWorkflowStateIsQcInitiated -StateName $CanonicalState -Config $Config
    }
    if (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue) {
        $isFinalizing = Test-QCWorkflowStateIsQcFinalizing -StateName $CanonicalState -Config $Config
    }
    if (-not $isInitiated -and -not $isFinalizing) { return }

    if (-not (Get-Command -Name 'Add-QCPrependJobForQcInitiatedStateChange' -ErrorAction SilentlyContinue) `
        -or -not (Get-Command -Name 'Add-QCPrependJobForQcFinalizingStateChange' -ErrorAction SilentlyContinue)) {
        try {
            $procPath = Join-Path $PSScriptRoot 'QC.Processors.psm1'
            Import-Module $procPath -Force -ErrorAction SilentlyContinue
        } catch { }
    }

    $sheetStem = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStem = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }
    if ([string]::IsNullOrWhiteSpace($sheetStem)) { return }
    $sheetPdfName = $sheetStem + '.pdf'

    $prependTrigger = if ($isInitiated) { 'initialQcPdf' } else { 'finalQcComplete' }

    if (-not (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue)) {
        try { Import-Module (Join-Path $PSScriptRoot 'QC.Processors.psm1') -Force -ErrorAction SilentlyContinue } catch { }
    }
    if (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue) {
        $intendedLane = ''
        if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
            $intendedLane = Get-PWQcPdfLaneFromDocumentName -DocumentName $DocumentName
        }
        $sheetBlock = Test-QCPrependEnqueueBlockedForSheet -Config $Config -FolderPath $FolderPath `
            -SheetPdfName $sheetPdfName -PrependTrigger $prependTrigger -QcProcessType $intendedLane
        if ($sheetBlock -and [bool]$sheetBlock.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PREPEND_SKIPPED_SHEET_ACTIVE' `
                    -Message 'QC_PREPEND skipped for lane intake: prepend already pending, running, or recently succeeded for this sheet lane.' -Data @{
                    folderPath = $FolderPath; sheetPdf = $sheetPdfName; sheetStem = $sheetStem
                    triggerDocumentName = $DocumentName; qcProcessType = $intendedLane
                    reason = [string]$sheetBlock.reason; prependTrigger = $prependTrigger
                } | Out-Null
            }
            return
        }
    }

    if (-not (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue)) {
        try { Import-Module (Join-Path $PSScriptRoot 'QC.Notifications.psm1') -Force -ErrorAction SilentlyContinue } catch { }
    }
    if (Get-Command -Name 'Test-QCPrependBlockedByMissingEmailAttributes' -ErrorAction SilentlyContinue) {
        $emailGate = Test-QCPrependBlockedByMissingEmailAttributes -Config $Config -FolderPath $FolderPath `
            -SheetPdfName $sheetPdfName -DocumentGuid $DocumentGuid
        if ($emailGate -and [bool]$emailGate.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'QC_PREPEND_SKIPPED_MISSING_EMAIL' `
                    -Message 'QC_PREPEND skipped for lane intake: required notification email attributes are missing.' -Data @{
                    folderPath = $FolderPath; sheetPdf = $sheetPdfName; sheetStem = $sheetStem
                    triggerDocumentName = $DocumentName; missingFields = @($emailGate.missingFields)
                    postPrependState = [string]$emailGate.postPrependState; prependTrigger = $prependTrigger
                } | Out-Null
            }
            return
        }
    }

    $stKey = _PWD-GetPrependEnqueueStateTransitionKey -AuditEventId $AuditEventId -LastAuditEventAt $LastAuditEventAt `
        -ChangedByUser $ChangedByUser -TriggerDocumentGuid $DocumentGuid `
        -SheetStem $sheetStem -PreviousSheetState $PreviousLaneState -TargetStateName $CanonicalState `
        -PrependTrigger $prependTrigger

    try {
        if ($isInitiated) {
            Add-QCPrependJobForQcInitiatedStateChange -Config $Config `
                -TriggerDocumentGuid $DocumentGuid -TriggerDocumentName $DocumentName -FolderPath $FolderPath `
                -CurrentStateName $CanonicalState -DryRun:$DryRun -ChangedByUser $ChangedByUser `
                -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt `
                -AuditEventId $AuditEventId -StateTransitionKey $stKey | Out-Null
        } else {
            Add-QCPrependJobForQcFinalizingStateChange -Config $Config `
                -TriggerDocumentGuid $DocumentGuid -TriggerDocumentName $DocumentName -FolderPath $FolderPath `
                -CurrentStateName $CanonicalState -DryRun:$DryRun -ChangedByUser $ChangedByUser `
                -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt `
                -AuditEventId $AuditEventId -StateTransitionKey $stKey | Out-Null
        }
    } catch {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            $code = if ($isInitiated) { 'QC_PREPEND_STATE_ENQUEUE_ERROR' } else { 'QC_PREPEND_FINAL_ENQUEUE_ERROR' }
            Write-QCJsonLog -Flush -Level 'Warning' -Code $code -Message $_.Exception.Message -Data @{
                folderPath = $FolderPath; triggerDocumentName = $DocumentName; prependTrigger = $prependTrigger
            } | Out-Null
        }
    }
}

function _PWD-InvokeLaneQcPdfDocumentStateWorkflow {
    <#
    .SYNOPSIS
    Records lane QC PDF workflow telemetry and notifications without sheet-group sibling sync.
    .DESCRIPTION
    rev/prod/chk PDFs are independent process lanes. Stem PDF and DGN are excluded from this path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$DocumentGuid,
        [Parameter(Mandatory)][string]$DocumentName,
        [Parameter(Mandatory)][string]$FolderPath,
        [string]$WatchRoot = '',
        [Parameter(Mandatory)][string]$CanonicalState,
        [string]$LastAuditEventAt = '',
        [Nullable[long]]$AuditEventId = $null,
        [bool]$DryRun = $false,
        [Nullable[int]]$ChangedByUser = $null,
        [string]$ChangedByUsername = '',
        [bool]$ActorIsAutomation = $false
    )

    [void](_PWD-EnsureAuditWorkflowExports)

    $members = @(@{
        documentGuid = $DocumentGuid
        documentName = $DocumentName
        document     = $null
    })

    $guids = @()
    if (Test-PWValidDocumentGuid -DocumentGuid $DocumentGuid) {
        $guids = @([string]$DocumentGuid)
    }

    $stateByGuid = @{}
    if ($guids.Count -gt 0) {
        try { $stateByGuid = Get-PWDocumentWorkflowStateMapByGuid -DocumentGuids $guids } catch { }
    }

    $triggerGuidKey = ([string]$DocumentGuid).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($triggerGuidKey) -and $stateByGuid.ContainsKey($triggerGuidKey)) {
        $batchLiveState = [string]$stateByGuid[$triggerGuidKey]
        if (-not [string]::IsNullOrWhiteSpace($batchLiveState)) {
            $normalizedBatch = _PWD-NormalizeSheetIndexValue $batchLiveState
            $normalizedCanonical = _PWD-NormalizeSheetIndexValue $CanonicalState
            if ([string]::IsNullOrWhiteSpace($CanonicalState) -or ($normalizedBatch -ne $normalizedCanonical)) {
                $CanonicalState = $batchLiveState
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($CanonicalState)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_LANE_STATE_NO_SOURCE' `
                -Message 'Lane QC PDF DOCUMENT_STATE skipped: live PW state unavailable after batch read.' -Data @{
                auditEventId = $AuditEventId
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
            } | Out-Null
        }
        return
    }

    $previousStateByGuid = @{}
    $prevIndex = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $DocumentGuid
    $previousStateByGuid[$triggerGuidKey] = [string]$prevIndex

    _PWD-WriteDocumentStateLiveVerificationLog -AuditEventId $AuditEventId `
        -SourceDocumentGuid $DocumentGuid -SourceDocumentName $DocumentName -FolderPath $FolderPath `
        -CanonicalState $CanonicalState -CanonicalStateSource 'liveProjectWise' `
        -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
        -Members $members -StateByGuid $stateByGuid -Config $Config

    $sheetStemForPrepend = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStemForPrepend = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }

    if (Get-Command -Name 'Test-QCDocumentStateAuditEventIsStale' -ErrorAction SilentlyContinue) {
        $staleDecision = Test-QCDocumentStateAuditEventIsStale -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -AuditEventId $AuditEventId `
            -LastAuditEventAt $LastAuditEventAt -CanonicalState $CanonicalState `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -Members $members -StateByGuid $stateByGuid -SheetStem $sheetStemForPrepend
        if ($staleDecision -and [bool]$staleDecision.isStale) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                $staleLog = @{
                    decision = 'skipped'
                    sync = 'skipped'
                    notify = 'skipped'
                    canonicalState = $CanonicalState
                    canonicalStateSource = 'liveProjectWise'
                    laneIndependent = $true
                }
                foreach ($k in @($staleDecision.Keys)) { $staleLog[$k] = $staleDecision[$k] }
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_STATE_SYNC_SKIPPED_STALE_EVENT' `
                    -Message 'Skipped lane QC PDF DOCUMENT_STATE workflow for superseded audit event.' -Data $staleLog | Out-Null
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
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -TargetStateName $CanonicalState `
            -Members $members -StateByGuid $stateByGuid -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -DryRun:$DryRun
        if ($gate -and $gate.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_LANE_STATE_SYNC_BLOCKED' `
                    -Message 'Lane QC PDF state workflow stopped: required email attributes missing.' -Data @{
                    triggerDocumentGuid = $DocumentGuid
                    triggerDocumentName = $DocumentName
                    folderPath          = $FolderPath
                    canonicalState      = $CanonicalState
                    missingFields       = @($gate.missingFields)
                } | Out-Null
            }
            if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
                Invoke-QCSheetGroupWorkflowTransition -Config $Config -TriggerDocumentGuid $DocumentGuid `
                    -TriggerDocumentName $DocumentName -FolderPath $FolderPath -SourceState $prevIndex `
                    -TargetState $CanonicalState -TransitionSource 'user_audit' -Members $members `
                    -StateByGuid $stateByGuid -PreviousStateByGuid $previousStateByGuid `
                    -AuditEventId $AuditEventId -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
                    -LastAuditEventAt $LastAuditEventAt -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' `
                    -LaneIndependentMode | Out-Null
            }
            return
        }
    }

    if (Get-Command -Name '_PWD-InvokeStaleSheetIndexAuditStateTriggers' -ErrorAction SilentlyContinue) {
        _PWD-InvokeStaleSheetIndexAuditStateTriggers -Config $Config -Members $members -StateByGuid $stateByGuid `
            -FolderPath $FolderPath -CanonicalState $CanonicalState -DryRun:$DryRun `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId | Out-Null
    }

    $dbEnabled = $false
    if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
        $dbEnabled = Test-QCDatabaseEnabled -Config $Config
    }

    $currentState = if ($stateByGuid.ContainsKey($triggerGuidKey)) {
        [string]$stateByGuid[$triggerGuidKey]
    } else {
        ''
    }
    if ([string]::IsNullOrWhiteSpace($currentState)) {
        try {
            $currentState = [string](Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
        } catch { }
    }

    $stateUpdates = @()
    if ((_PWD-NormalizeSheetIndexValue $currentState) -eq (_PWD-NormalizeSheetIndexValue $CanonicalState)) {
        if ($dbEnabled -and (-not $DryRun)) {
            try {
                [void](Update-QCSheetIndexPwStateName -Config $Config -DocumentGuid $DocumentGuid -PwStateName $CanonicalState)
                if ($LastAuditEventAt) {
                    Invoke-QCDatabaseNonQuery -Config $Config -Sql @"
UPDATE sheet_index SET last_audit_event_at = @lastAudit, last_updated_at = SYSDATETIMEOFFSET()
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid; lastAudit = $LastAuditEventAt } | Out-Null
                }
                if (Get-Command -Name 'Update-SheetPackageQcPdfLaneState' -ErrorAction SilentlyContinue) {
                    $laneType = $null
                    if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
                        $laneType = Get-PWQcPdfLaneFromDocumentName -DocumentName $DocumentName
                    }
                    if ($laneType) {
                        [void](Update-SheetPackageQcPdfLaneState -Config $Config -DocumentGuid $DocumentGuid `
                            -CurrentPwState $CanonicalState -PreviousPwState ([string]$prevIndex) -QcProcessType $laneType)
                    }
                }
            } catch { }
        }
    }

    if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
        Invoke-QCSheetGroupWorkflowTransition -Config $Config -TriggerDocumentGuid $DocumentGuid `
            -TriggerDocumentName $DocumentName -FolderPath $FolderPath -SourceState $prevIndex `
            -TargetState $CanonicalState -TransitionSource 'user_audit' -Members $members `
            -StateByGuid $stateByGuid -PreviousStateByGuid $previousStateByGuid `
            -AuditEventId $AuditEventId -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' `
            -LaneIndependentMode | Out-Null
    } elseif (Get-Command -Name 'Invoke-QCAuditWorkflowStateChangeTriggers' -ErrorAction SilentlyContinue) {
        Invoke-QCAuditWorkflowStateChangeTriggers -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -PreviousState $prevIndex -CurrentState $CanonicalState `
            -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId `
            -PreviousStateSource 'sheet_index' -CurrentStateSource 'liveProjectWise' `
            -StaleCheckMembers $members -StaleCheckStateByGuid $stateByGuid -StaleCheckCanonicalState $CanonicalState | Out-Null
    }

    _PWD-TryEnqueueLaneQcPrependFromState -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
        -FolderPath $FolderPath -CanonicalState $CanonicalState -PreviousLaneState $prevIndex `
        -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId -DryRun:$DryRun `
        -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_LANE_STATE_INDEPENDENT' `
            -Message 'Lane workflow state recorded from DOCUMENT_STATE audit event without sibling sync.' -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalState      = $CanonicalState
            previousLaneState   = $prevIndex
            memberCount         = 1
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
            actorIsAutomation   = [bool]$ActorIsAutomation
        } | Out-Null
    }
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
        [Nullable[long]]$AuditEventId = $null,
        [switch]$DeferNotification
    )

    if (-not (Get-Command -Name 'Get-QCAuditWorkflowTriggerSettings' -ErrorAction SilentlyContinue)) { return }
    if (-not (Get-Command -Name 'Invoke-QCAuditWorkflowStateChangeTriggers' -ErrorAction SilentlyContinue)) { return }

    $wt = Get-QCAuditWorkflowTriggerSettings -Config $Config
    if (-not [bool]$wt.enabled) { return }
    if (-not [bool]$wt.recordStateHistory -and -not [bool]$wt.recordTransitions -and -not [bool]$wt.notifyOnStateChange) { return }

    $canonical = _PWD-NormalizeSheetIndexValue $CanonicalState
    if ([string]::IsNullOrWhiteSpace($canonical)) { return }

    $packageNotifyGuid = _PWD-ResolveSheetPackageNotificationDocumentGuid -Members $Members

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
        if ([string]::IsNullOrWhiteSpace($pwCurrent) -and (Get-Command -Name 'Get-PWDocumentWorkflowStateName' -ErrorAction SilentlyContinue)) {
            try {
                $pwCurrent = [string](Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $dn -DocumentGuid $dg)
            } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($pwCurrent) -and -not [string]::IsNullOrWhiteSpace($canonical)) {
            $pwCurrent = $CanonicalState
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
        $laneOnlyMember = (@($Members).Count -eq 1) -and (Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) `
            -and (Test-QCIsQcPdfDocumentName -DocumentName $dn)
        $suppressNotify = $false
        if ($DeferNotification.IsPresent -and -not $laneOnlyMember) {
            $suppressNotify = $true
        }
        if (-not $suppressNotify -and -not $laneOnlyMember) {
            $suppressNotify = ((-not [string]::IsNullOrWhiteSpace($packageNotifyGuid)) -and ($dg -ne $packageNotifyGuid))
        }
        Invoke-QCAuditWorkflowStateChangeTriggers -Config $Config -DocumentGuid $dg -DocumentName $dn `
            -FolderPath $FolderPath -PreviousState $prevDb -CurrentState $CanonicalState -Document $member.document `
            -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' -ChangedByUser $ChangedByUser `
            -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId `
            -PreviousStateSource 'sheet_index' -CurrentStateSource 'liveProjectWise' `
            -StaleCheckMembers $Members -StaleCheckStateByGuid $StateByGuid -StaleCheckCanonicalState $CanonicalState `
            -SuppressNotification:$suppressNotify | Out-Null
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
    QC Finalizing prefers the active lane QC PDF workflow state. Dedupe uses the audit event id when
    available (see Get-QCPrependStateTransitionDedupeKey). Initiate Origination and Initiate Verification
    intake always enqueue prepend even when lane-independent sync applies no PW state writes.
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
        if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-(prod|chk|rev)\.pdf$')) {
            $sheetPdfGuid = [string]$member.documentGuid
        } elseif ($dn -match '(?i)-(prod|chk|rev)\.pdf$') {
            if ([string]::IsNullOrWhiteSpace($qcPdfName)) {
                $qcPdfGuid = [string]$member.documentGuid
                $qcPdfName = $dn
            }
        }
    }

    if ((Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) `
        -and (Test-QCIsQcPdfDocumentName -DocumentName $TriggerDocumentName)) {
        $qcPdfName = [string]$TriggerDocumentName
        if (-not [string]::IsNullOrWhiteSpace($TriggerDocumentGuid)) {
            $qcPdfGuid = [string]$TriggerDocumentGuid
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

    $prependTriggerName = $sheetPdfName
    if ((Get-Command -Name 'Test-QCIsQcPdfDocumentName' -ErrorAction SilentlyContinue) `
        -and (Test-QCIsQcPdfDocumentName -DocumentName $TriggerDocumentName)) {
        $prependTriggerName = [string]$TriggerDocumentName
    }

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
            if (-not (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue)) {
                try {
                    $procPath = Join-Path $PSScriptRoot 'QC.Processors.psm1'
                    Import-Module $procPath -Force -ErrorAction SilentlyContinue
                } catch { }
            }
            if (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue) {
                $intendedLane = ''
                if (Get-Command -Name '_QCP-ResolveIntendedPrependProcessType' -ErrorAction SilentlyContinue) {
                    $intendedLane = _QCP-ResolveIntendedPrependProcessType -Config $Config -FolderPath $FolderPath `
                        -SheetPdfName $sheetPdfName -SheetPdfGuid $sheetPdfGuid
                }
                $sheetBlock = Test-QCPrependEnqueueBlockedForSheet -Config $Config -FolderPath $FolderPath `
                    -SheetPdfName $sheetPdfName -PrependTrigger 'initialQcPdf' -QcProcessType $intendedLane
                if ($sheetBlock -and [bool]$sheetBlock.blocked) {
                    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                        Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PREPEND_SKIPPED_SHEET_ACTIVE' `
                            -Message 'QC_PREPEND skipped after sheet sync: prepend already pending, running, or recently succeeded for this sheet lane.' -Data @{
                            folderPath = $FolderPath; sheetPdf = $sheetPdfName; sheetStem = $sheetStem
                            qcProcessType = $intendedLane; reason = [string]$sheetBlock.reason
                            matches = @($sheetBlock.matches); canonicalState = $canonical
                        } | Out-Null
                    }
                    return
                }
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
                        -TriggerDocumentGuid $prependGuid -TriggerDocumentName $prependTriggerName -FolderPath $FolderPath `
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
                        -TriggerDocumentGuid $prependGuid -TriggerDocumentName $prependTriggerName -FolderPath $FolderPath `
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
                $writeResult = _PWD-InvokeSetPwDocumentState -Document $doc -StateName $previous -GuardContext @{
                    callSite = 'Invoke-QCWorkflowStateEmailAttributeGate.rollback.previous'
                    auditEventId = $AuditEventId
                    documentName = $dn
                    folderPath = $FolderPath
                    documentGuid = $dg
                    config = $Config
                    sourceVariableName = 'previous'
                    livePwState = [string]$current
                    changedByUsername = $ChangedByUsername
                    writeScope = 'workflow'
                }
                $change.applied = [bool]$writeResult.applied
                $change.verified = [bool]$writeResult.verified
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

    $auditTargetState = _PWD-ResolveAuditWorkflowTargetStateName -Config $Config `
        -AuditTargetStateName $AuditTargetStateName -AuditRawItemDesc $AuditRawItemDesc -AuditRawTextParam $AuditRawTextParam
    $sourcePwState = Get-PWDocumentWorkflowStateName -FolderPath $FolderPath -DocumentName $DocumentName -DocumentGuid $DocumentGuid
    $sourcePwStateSource = 'liveProjectWise'
    if ([string]::IsNullOrWhiteSpace($sourcePwState) -and -not [string]::IsNullOrWhiteSpace($auditTargetState)) {
        $sourcePwState = $auditTargetState
        $sourcePwStateSource = 'auditTargetHint'
    }
    # Do not seed canonical state from sheet_index here: stale index rows poison lane-PDF
    # notifications when the single-doc PW read fails. Batch GUID reads reconcile below.
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
            pwStateNameSource = $sourcePwStateSource
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

    if (_PWD-TestTriggerIsLaneQcPdf -DocumentName $DocumentName) {
        _PWD-InvokeLaneQcPdfDocumentStateWorkflow -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
            -FolderPath $FolderPath -WatchRoot $WatchRoot -CanonicalState $canonicalState -DryRun:$DryRun `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername -LastAuditEventAt $LastAuditEventAt `
            -AuditEventId $AuditEventId -ActorIsAutomation $actorIsAutomation
        return
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
        return
    }

    if (([string]$DocumentName -match '(?i)\.dgn$')) {
        return
    }

    $allMembers = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
        -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
    $legacySiblingSync = _PWD-TestLegacySiblingStateSyncEnabled -Config $Config
    if ($legacySiblingSync) {
        $members = @(Get-PWAssociatedSheetSyncMembers -Config $Config -FolderPath $FolderPath `
            -DocumentName $DocumentName -DocumentGuid $DocumentGuid -TriggerSource 'user_audit')
    } else {
        _PWD-LogLaneStateIndependentTelemetry -AllMembers $allMembers -FolderPath $FolderPath `
            -TriggerSource 'user_audit' -TriggerDocumentName $DocumentName -TriggerDocumentGuid $DocumentGuid
        $members = @(_PWD-GetLaneIndependentAuditMembers -AllMembers $allMembers `
            -TriggerDocumentGuid $DocumentGuid -TriggerDocumentName $DocumentName)
    }
    if ($members.Count -eq 0) {
        if (-not $legacySiblingSync) {
            $members = @(@{
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                document     = $null
                sourceType   = 'pdf'
            })
        } else {
            return
        }
    }

    $previousSheetState = ''
    $sheetStemForPrepend = ''
    if (Get-Command -Name 'Get-PWSheetStemFromDocumentName' -ErrorAction SilentlyContinue) {
        $sheetStemForPrepend = Get-PWSheetStemFromDocumentName -DocumentName $DocumentName
    }
    if (-not [string]::IsNullOrWhiteSpace($sheetStemForPrepend)) {
        foreach ($member in $allMembers) {
            $dn = [string]$member.documentName
            if (($dn -match '(?i)\.pdf$') -and ($dn -notmatch '(?i)-(prod|chk|rev)\.pdf$')) {
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

    $previousStateByGuid = @{}
    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        if (-not $dg) { continue }
        $prevIndex = _PWD-GetSheetIndexPwStateName -Config $Config -DocumentGuid $dg
        $previousStateByGuid[$dg.ToLowerInvariant()] = [string]$prevIndex
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

    $triggerGuidKey = ([string]$DocumentGuid).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($triggerGuidKey) -and $stateByGuid.ContainsKey($triggerGuidKey)) {
        $batchLiveState = [string]$stateByGuid[$triggerGuidKey]
        if (-not [string]::IsNullOrWhiteSpace($batchLiveState)) {
            $normalizedBatch = _PWD-NormalizeSheetIndexValue $batchLiveState
            $normalizedCanonical = _PWD-NormalizeSheetIndexValue $canonicalState
            if ([string]::IsNullOrWhiteSpace($canonicalState) -or ($normalizedBatch -ne $normalizedCanonical)) {
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Level 'Information' -Code 'WATCH_AUDIT_STATE_SYNC_CANONICAL_RECONCILED' `
                        -Message 'Reconciled DOCUMENT_STATE canonical state from batch live ProjectWise read.' -Data @{
                        auditEventId = $AuditEventId
                        documentGuid = $DocumentGuid
                        documentName = $DocumentName
                        folderPath = $FolderPath
                        previousCanonicalState = $canonicalState
                        previousCanonicalStateSource = $sourcePwStateSource
                        reconciledCanonicalState = $batchLiveState
                        reconciledCanonicalStateSource = 'liveProjectWiseBatch'
                    } | Out-Null
                }
                $canonicalState = $batchLiveState
                $sourcePwState = $batchLiveState
                $sourcePwStateSource = 'liveProjectWiseBatch'
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($canonicalState)) {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'WATCH_AUDIT_STATE_SYNC_NO_SOURCE_STATE' `
                -Message 'DOCUMENT_STATE sync skipped: live PW state and audit hint unavailable after batch read.' -Data @{
                auditEventId = $AuditEventId
                documentGuid = $DocumentGuid
                documentName = $DocumentName
                folderPath = $FolderPath
                auditTargetStateHint = $auditTargetState
                auditTargetStateName = $AuditTargetStateName
            } | Out-Null
        }
        return
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
            if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
                Invoke-QCSheetGroupWorkflowTransition -Config $Config -TriggerDocumentGuid $DocumentGuid `
                    -TriggerDocumentName $DocumentName -FolderPath $FolderPath -SourceState $previousSheetState `
                    -TargetState $canonicalState -TransitionSource 'user_audit' -Members $members `
                    -StateByGuid $stateByGuid -PreviousStateByGuid $previousStateByGuid `
                    -AuditEventId $AuditEventId -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
                    -LastAuditEventAt $LastAuditEventAt -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' | Out-Null
            }
            return
        }
    }

    if (Get-Command -Name '_PWD-InvokeStaleSheetIndexAuditStateTriggers' -ErrorAction SilentlyContinue) {
        _PWD-InvokeStaleSheetIndexAuditStateTriggers -Config $Config -Members $members -StateByGuid $stateByGuid `
            -FolderPath $FolderPath -CanonicalState $canonicalState -DryRun:$DryRun `
            -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -AuditEventId $AuditEventId -DeferNotification | Out-Null
    }

    $stateUpdates = @()
    $dbEnabled = $false
    if (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue) {
        $dbEnabled = Test-QCDatabaseEnabled -Config $Config
    }
    $triggerIsLanePdf = _PWD-TestTriggerIsLaneQcPdf -DocumentName $DocumentName

    foreach ($member in $members) {
        $dg = [string]$member.documentGuid
        $dn = [string]$member.documentName
        if (_PWD-TestAutomationDgnProtectedDocument -DocumentName $dn) { continue }
        if ($triggerIsLanePdf -and (Get-Command -Name 'Test-QCIsSheetPdfDocumentName' -ErrorAction SilentlyContinue) `
                -and (Test-QCIsSheetPdfDocumentName -DocumentName $dn)) {
            _PWD-LogAutomationStemPdfWriteBlocked -Operation 'audit_state_sync' -DocumentName $dn -DocumentGuid $dg `
                -FolderPath $FolderPath -CallSite 'Sync-PWAssociatedSheetWorkflowState' `
                -TargetState $canonicalState -TriggerDocumentName $DocumentName
            continue
        }
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
                    $writeResult = _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $restartStateToKeep -GuardContext @{
                        callSite = 'Sync-PWAssociatedSheetWorkflowState.restartStateToKeep'
                        auditEventId = $AuditEventId
                        documentName = $dn
                        folderPath = $FolderPath
                        documentGuid = $dg
                        config = $Config
                        sourceVariableName = 'restartStateToKeep'
                        livePwState = [string]$currentState
                        changedByUsername = $ChangedByUsername
                        writeScope = 'workflow'
                    }
                    $restored = [bool]$writeResult.verified
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
                $writeResult = _PWD-InvokeSetPwDocumentState -Document $member.document -StateName $canonicalState -GuardContext @{
                    callSite = 'Sync-PWAssociatedSheetWorkflowState.canonicalState'
                    auditEventId = $AuditEventId
                    documentName = $dn
                    folderPath = $FolderPath
                    documentGuid = $dg
                    config = $Config
                    sourceVariableName = 'canonicalState'
                    livePwState = [string]$currentState
                    changedByUsername = $ChangedByUsername
                    writeScope = 'workflow'
                    triggerDocumentName = $DocumentName
                }
                $change.applied = [bool]$writeResult.applied
                $change.verified = [bool]$writeResult.verified
                if (-not $change.verified) {
                    throw "Workflow state write unverified for '$dn' (read-back: '$($writeResult.readBackState)', target: '$canonicalState')."
                }
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
            if (-not $legacySiblingSync -and (Get-Command -Name 'Update-SheetPackageQcPdfLaneState' -ErrorAction SilentlyContinue)) {
                $laneType = $null
                if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
                    $laneType = Get-PWQcPdfLaneFromDocumentName -DocumentName $dn
                }
                if (-not $laneType) {
                    if (Get-Command -Name 'Get-PWQcPdfLaneFromDocumentName' -ErrorAction SilentlyContinue) {
                        $laneType = Get-PWQcPdfLaneFromDocumentName -DocumentName $dn
                    }
                }
                if ($laneType) {
                    [void](Update-SheetPackageQcPdfLaneState -Config $Config -DocumentGuid $dg `
                        -CurrentPwState $canonicalState -PreviousPwState ([string]$currentState) -QcProcessType $laneType)
                }
            }
        }

        $stateUpdates += $change
    }

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        $syncCode = if ($legacySiblingSync) { 'WATCH_SHEET_STATE_SYNC' } else { 'QC_LANE_STATE_INDEPENDENT' }
        $syncMsg = if ($legacySiblingSync) {
            'Associated sheet workflow states aligned from DOCUMENT_STATE audit event.'
        } else {
            'Lane workflow state recorded from DOCUMENT_STATE audit event without sibling sync.'
        }
        Write-QCJsonLog -Flush -Level 'Information' -Code $syncCode -Message $syncMsg -Data @{
            triggerDocumentGuid = $DocumentGuid
            triggerDocumentName = $DocumentName
            folderPath          = $FolderPath
            canonicalState      = $canonicalState
            memberCount         = $members.Count
            updates             = @($stateUpdates)
            dryRun              = [bool]$DryRun
        }
    }

    if (Get-Command -Name 'Invoke-QCSheetGroupWorkflowTransition' -ErrorAction SilentlyContinue) {
        Invoke-QCSheetGroupWorkflowTransition -Config $Config -TriggerDocumentGuid $DocumentGuid `
            -TriggerDocumentName $DocumentName -FolderPath $FolderPath -SourceState $previousSheetState `
            -TargetState $canonicalState -TransitionSource 'user_audit' -Members $members `
            -StateByGuid $stateByGuid -PreviousStateByGuid $previousStateByGuid `
            -AuditEventId $AuditEventId -ChangedByUser $ChangedByUser -ChangedByUsername $ChangedByUsername `
            -LastAuditEventAt $LastAuditEventAt -DryRun:$DryRun -AuditActionName 'DOCUMENT_STATE' | Out-Null
    }

    $qcInitiated = $false
    $qcFinalizing = $false
    if ((-not [string]::IsNullOrWhiteSpace($canonicalState)) -and (Get-Command -Name 'Test-QCWorkflowStateIsQcInitiated' -ErrorAction SilentlyContinue)) {
        $qcInitiated = Test-QCWorkflowStateIsQcInitiated -StateName $canonicalState -Config $Config
    }
    if ((-not [string]::IsNullOrWhiteSpace($canonicalState)) -and (Get-Command -Name 'Test-QCWorkflowStateIsQcFinalizing' -ErrorAction SilentlyContinue)) {
        $qcFinalizing = Test-QCWorkflowStateIsQcFinalizing -StateName $canonicalState -Config $Config
    }
    if ([string]::IsNullOrWhiteSpace($canonicalState)) {
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
        if ($qcInitiated -or $qcFinalizing) {
            $intakeReason = if ($qcInitiated) { 'qc_initiated_intake' } else { 'qc_finalizing_intake' }
            $intakeCode = if ($qcInitiated) { 'QC_PREPEND_INTAKE_WITHOUT_SIBLING_SYNC' } else { 'QC_PREPEND_FINAL_INTAKE_WITHOUT_SIBLING_SYNC' }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code $intakeCode `
                    -Message 'QC_PREPEND allowed for automation intake even though lane-independent sync applied no PW state writes.' -Data @{
                    triggerDocumentGuid = $DocumentGuid
                    triggerDocumentName = $DocumentName
                    folderPath          = $FolderPath
                    previousSheetState  = $previousSheetState
                    canonicalState      = $canonicalState
                    memberCount         = $members.Count
                    auditEventId        = $AuditEventId
                    reason              = $intakeReason
                } | Out-Null
            }
        } else {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'QC_PREPEND_SKIPPED_NO_SHEET_STATE_CHANGES' `
                    -Message 'QC_PREPEND skipped: sibling sync made no state changes (echo audit).' -Data @{
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
    if (Test-QCIsSheetPdfDocumentName -DocumentName $DocumentName) {
        $sheetGuid = [string]$DocumentGuid
    } else {
        try {
            $members = @(Get-PWAssociatedSheetMembers -Config $Config -FolderPath $FolderPath `
                -DocumentName $DocumentName -DocumentGuid $DocumentGuid)
            foreach ($member in $members) {
                $mn = [string]$member.documentName
                if (Test-QCIsSheetPdfDocumentName -DocumentName $mn) {
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

    if (-not (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue)) {
        try {
            $procPath = Join-Path $PSScriptRoot 'QC.Processors.psm1'
            Import-Module $procPath -Force -ErrorAction SilentlyContinue
        } catch { }
    }
    if (Get-Command -Name 'Test-QCPrependEnqueueBlockedForSheet' -ErrorAction SilentlyContinue) {
        $intendedLane = ''
        if (Get-Command -Name '_QCP-ResolveIntendedPrependProcessType' -ErrorAction SilentlyContinue) {
            $intendedLane = _QCP-ResolveIntendedPrependProcessType -Config $Config -FolderPath $FolderPath `
                -SheetPdfName $sheetPdfName -SheetPdfGuid $sheetGuid
        }
        $sheetBlock = Test-QCPrependEnqueueBlockedForSheet -Config $Config -FolderPath $FolderPath `
            -SheetPdfName $sheetPdfName -PrependTrigger 'initialQcPdf' -QcProcessType $intendedLane
        if ($sheetBlock -and [bool]$sheetBlock.blocked) {
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_QC_INITIATED_FALLBACK_SKIPPED' `
                    -Message 'QC Initiated sheet-PDF fallback skipped: prepend already pending, running, or recently succeeded for this sheet lane.' -Data @{
                    triggerDocumentGuid = $DocumentGuid; triggerDocumentName = $DocumentName; folderPath = $FolderPath
                    sheetPdfName = $sheetPdfName; qcProcessType = $intendedLane; reason = [string]$sheetBlock.reason
                    matches = @($sheetBlock.matches)
                } | Out-Null
            }
            return
        }
    }

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

    $checkerEmailDiffer = (_PWD-NormalizeSheetIndexValue $pwChecker) -ne (_PWD-NormalizeSheetIndexValue $dbChecker)
    $roleEmailsNeedSiblingSync = $emailsDiffer -or $checkerEmailDiffer

    if ($isDocumentAttr -and $roleEmailsNeedSiblingSync) {
        Sync-PWAssociatedSheetEmailAttributes -Config $Config -DocumentGuid $DocumentGuid `
            -DocumentName $DocumentName -FolderPath $FolderPath -WatchRoot $WatchRoot `
            -LastAuditEventAt $LastAuditEventAt -DryRun:$false
    }

    if ($isDocumentAttr -and $reviewTypeDiffer) {
        $canonicalRaw = [string]$fieldsRaw.qcProcessType
        if ([string]::IsNullOrWhiteSpace($canonicalRaw)) {
            $canonicalRaw = [string]$fieldsRaw.qcReviewType
        }
        $canonicalProcess = _PWD-ResolveCanonicalProcessTypeForSync -RawValue $canonicalRaw -Config $Config
        if (-not (Get-Command -Name 'Test-QCLegacyReviewTypeAttributeSyncEnabled' -ErrorAction SilentlyContinue) `
            -or -not (Test-QCLegacyReviewTypeAttributeSyncEnabled -Config $Config)) {
            _PWD-LogProcessTypeSyncSkipped -Reason 'legacy_sync_disabled' -Config $Config `
                -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
                -ExtraData @{
                source = 'document_attr_audit'
                rawProcessType = [string]$fieldsRaw.qcProcessType
                rawReviewType = [string]$fieldsRaw.qcReviewType
                resolvedReviewType = [string]$pwReviewType
                dbReviewType = [string]$dbReviewType
            }
        } elseif (-not $canonicalProcess) {
            _PWD-LogProcessTypeUnknown -Config $Config -DocumentGuid $DocumentGuid -DocumentName $DocumentName `
                -FolderPath $FolderPath -RawValue $canonicalRaw -Source 'document_attr_audit'
            _PWD-LogProcessTypeSyncSkipped -Reason 'null_canonical_process_type' -Config $Config `
                -DocumentGuid $DocumentGuid -DocumentName $DocumentName -FolderPath $FolderPath `
                -ExtraData @{
                source = 'document_attr_audit'
                rawProcessType = [string]$fieldsRaw.qcProcessType
                rawReviewType = [string]$fieldsRaw.qcReviewType
                resolvedReviewType = [string]$pwReviewType
                dbReviewType = [string]$dbReviewType
            }
        } else {
            Sync-PWAssociatedSheetReviewTypeAttributes -Config $Config -DocumentGuid $DocumentGuid `
                -DocumentName $DocumentName -FolderPath $FolderPath -CanonicalReviewType $canonicalProcess `
                -WatchRoot $WatchRoot -LastAuditEventAt $LastAuditEventAt -DryRun:$false
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
        [Parameter(Mandatory)][hashtable]$Attributes,
        [hashtable]$Config = $null
    )

    $cmd = Get-Command -Name 'Update-PWDocumentAttributes' -ErrorAction SilentlyContinue
    if (-not $cmd) { throw 'Update-PWDocumentAttributes is not available.' }
    if (-not $Attributes -or $Attributes.Keys.Count -eq 0) { return $false }

    $documentName = ''
    try { $documentName = [string]$Document.Name } catch { }
    if (_PWD-TestAutomationDgnProtectedDocument -DocumentName $documentName `
        -and _PWD-TestAutomationWritesQcProcessTypeOnAttributes -Attributes $Attributes -Config $Config) {
        $documentGuid = ''
        try { $documentGuid = [string]$Document.DocumentGUID } catch { }
        _PWD-LogAutomationDgnWriteBlocked -Operation 'qc_process_type' -DocumentName $documentName `
            -DocumentGuid $documentGuid -CallSite '_PWD-InvokeUpdatePWDocumentAttributes'
        return $false
    }

    $attrsToWrite = @{}
    $processCol = ''
    if ($Config) {
        try { $processCol = Get-PWQcProcessTypeAttributeName -Config $Config } catch { $processCol = '' }
    }
    foreach ($key in @($Attributes.Keys)) {
        $val = [string]$Attributes[$key]
        if ($processCol -and ([string]$key -eq $processCol) -and (Get-Command -Name 'Format-QCProcessTypeAttributeValue' -ErrorAction SilentlyContinue)) {
            try { $val = Format-QCProcessTypeAttributeValue -ProcessType $val } catch { }
        }
        $attrsToWrite[[string]$key] = $val
    }
    $Attributes = $attrsToWrite

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
    Copies EM_Designer_Email, EM_Reviewer_Email, and EM_Checker_Email from the source sheet PDF to the matching *-qc.pdf document.
    .DESCRIPTION
    The source .pdf is the source of truth. When the QC PDF is missing these values or they differ,
    updates the QC PDF environment attributes to match the source.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][string]$QcDocumentName,
        [hashtable]$Config = @{},
        [string]$DesignerEmailColumn = 'EM_Designer_Email',
        [string]$ReviewerEmailColumn = 'EM_Reviewer_Email',
        [string]$CheckerEmailColumn = 'EM_Checker_Email',
        [switch]$PassThru
    )

    if ($Config -and $Config.Count -gt 0) {
        $emailCols = _PWD-GetSheetRoleEmailColumnNames -Config $Config
        if ($emailCols.designer) { $DesignerEmailColumn = [string]$emailCols.designer }
        if ($emailCols.reviewer) { $ReviewerEmailColumn = [string]$emailCols.reviewer }
        if ($emailCols.checker) { $CheckerEmailColumn = [string]$emailCols.checker }
    }

    $result = @{
        updated = $false
        skipped = $true
        reason = ''
        sourceDesignerEmail = ''
        sourceReviewerEmail = ''
        sourceCheckerEmail = ''
        qcDesignerEmailBefore = ''
        qcReviewerEmailBefore = ''
        qcCheckerEmailBefore = ''
        attributesWritten = @()
    }

    $sourceDesigner = ''
    $sourceReviewer = ''
    $sourceChecker = ''
    if ($Config -and $Config.Count -gt 0) {
        $roleFields = Get-PWQcPrependRoleFieldsFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName -Config $Config
        if ($roleFields.found) {
            $sourceDesigner = [string]$roleFields.designerEmail
            $sourceReviewer = [string]$roleFields.reviewerEmail
            $sourceChecker = [string]$roleFields.checkerEmail
        }
    }
    if (-not $sourceDesigner -and -not $sourceReviewer -and -not $sourceChecker) {
        $source = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $SourceDocumentName `
            -DesignerEmailColumn $DesignerEmailColumn -ReviewerEmailColumn $ReviewerEmailColumn `
            -CheckerEmailColumn $CheckerEmailColumn
        if (-not $source.found) {
            $result.reason = if ($source.error) { "Source PDF not found or unreadable: $($source.error)" } else { 'Source PDF not found.' }
            if ($PassThru) { return $result }
            return
        }
        $sourceDesigner = [string]$source.designerEmail
        $sourceReviewer = [string]$source.reviewerEmail
        $sourceChecker = [string]$source.checkerEmail
    }

    $result.sourceDesignerEmail = $sourceDesigner
    $result.sourceReviewerEmail = $sourceReviewer
    $result.sourceCheckerEmail = $sourceChecker

    $qc = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $QcDocumentName `
        -DesignerEmailColumn $DesignerEmailColumn -ReviewerEmailColumn $ReviewerEmailColumn `
        -CheckerEmailColumn $CheckerEmailColumn
    if (-not $qc.found) {
        $result.reason = if ($qc.error) { "QC PDF not found or unreadable: $($qc.error)" } else { 'QC PDF not found.' }
        if ($PassThru) { return $result }
        return
    }

    $qcDoc = $qc.document
    $result.qcDesignerEmailBefore = [string]$qc.designerEmail
    $result.qcReviewerEmailBefore = [string]$qc.reviewerEmail
    $result.qcCheckerEmailBefore = [string]$qc.checkerEmail

    $toWrite = @{}
    $qcDesigner = [string]$qc.designerEmail
    $qcReviewer = [string]$qc.reviewerEmail
    $qcChecker = [string]$qc.checkerEmail

    if ($sourceDesigner -ne $qcDesigner) {
        $toWrite[$DesignerEmailColumn] = $sourceDesigner
        $result.attributesWritten += $DesignerEmailColumn
    }
    if ($sourceReviewer -ne $qcReviewer) {
        $toWrite[$ReviewerEmailColumn] = $sourceReviewer
        $result.attributesWritten += $ReviewerEmailColumn
    }
    if ($sourceChecker -ne $qcChecker) {
        $toWrite[$CheckerEmailColumn] = $sourceChecker
        $result.attributesWritten += $CheckerEmailColumn
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
            if ($pw -is [hashtable] -and $pw.ContainsKey('qcProcessTypeAttributes')) { $rta = $pw['qcProcessTypeAttributes'] }
            elseif ($pw.qcProcessTypeAttributes) { $rta = $pw.qcProcessTypeAttributes }
            if (-not $rta -and $pw -is [hashtable] -and $pw.ContainsKey('qcReviewTypeAttributes')) { $rta = $pw['qcReviewTypeAttributes'] }
            elseif (-not $rta -and $pw.qcReviewTypeAttributes) { $rta = $pw.qcReviewTypeAttributes }
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

function Get-PWQcProcessTypeAttributeName {
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $col = 'QC_Process_Type'
    try {
        $am = $Config['qcWorkflow']['attributeMap']
        if ($am -and $am['processType']) { $col = [string]$am['processType'] }
    } catch { }
    return $col
}

function Get-PWQcReviewTypeAttributeName {
    <#
    .SYNOPSIS
    Legacy read column for QC_Review_Type (compatibility).
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $col = 'QC_Review_Type'
    try {
        $am = $Config['qcWorkflow']['attributeMap']
        if ($am -and $am['reviewType']) { $col = [string]$am['reviewType'] }
    } catch { }
    return $col
}

function Get-PWQcDefaultProcessType {
    <#
    .SYNOPSIS
    Configured default qc_process_type (normalized lowercase), usually production.
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $default = 'production'
    if (-not $Config) { return $default }
    try {
        $wf = $null
        if ($Config.ContainsKey('qcWorkflow')) { $wf = $Config.qcWorkflow }
        if ($wf -is [hashtable] -and $wf.ContainsKey('defaultProcessType') -and -not [string]::IsNullOrWhiteSpace([string]$wf.defaultProcessType)) {
            $norm = if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
                Normalize-QCProcessType -ProcessType ([string]$wf.defaultProcessType)
            } else { [string]$wf.defaultProcessType }
            if ($norm) { return $norm }
        }
        if ($wf -is [hashtable] -and $wf.ContainsKey('defaultReviewType') -and -not [string]::IsNullOrWhiteSpace([string]$wf.defaultReviewType)) {
            $norm = if (Get-Command -Name 'Normalize-QCProcessType' -ErrorAction SilentlyContinue) {
                Normalize-QCProcessType -ProcessType ([string]$wf.defaultReviewType)
            } else { $null }
            if ($norm) { return $norm }
        }
    } catch { }
    return $default
}

function Get-PWQcDefaultReviewType {
    <#
    .SYNOPSIS
    Display label for default process type (backward compatibility).
    #>
    [CmdletBinding()]
    param([hashtable]$Config = @{})

    $pt = Get-PWQcDefaultProcessType -Config $Config
    if (Get-Command -Name 'Get-QCProcessTypeDisplayLabel' -ErrorAction SilentlyContinue) {
        return Get-QCProcessTypeDisplayLabel -ProcessType $pt
    }
    return 'Production'
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

function _PWD-GetSheetRoleFolderCandidates {
    param([string]$FolderPath)

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
    return @($folderCandidates)
}

function _PWD-TryReadSheetRoleFieldsFromDocument {
    param(
        [string[]]$FolderCandidates,
        [string]$DocumentName,
        [string[]]$Columns,
        [hashtable]$Config
    )

    $read = @{ found = $false; error = 'Document not found'; attributes = @{}; pwStateName = '' }
    $resolvedFolder = ''
    foreach ($fp in @($FolderCandidates)) {
        $tryRead = Get-PWDocumentAttributesByColumns -FolderPath $fp -DocumentName $DocumentName -ColumnsToReturn $Columns
        if ($tryRead.found) {
            $read = $tryRead
            $resolvedFolder = $fp
            break
        }
        if (-not $read.found -and $tryRead.error) { $read.error = [string]$tryRead.error }
    }
    if (-not $read.found) {
        return @{ found = $false; error = [string]$read.error; fields = $null; resolvedFolder = '' }
    }
    $fields = ConvertTo-SheetIndexFieldValues -Config $Config -PwAttributes $read.attributes -PwStateName ([string]$read.pwStateName)
    return @{ found = $true; error = ''; fields = $fields; resolvedFolder = $resolvedFolder }
}

function _PWD-NormalizeQcReviewTypeRoleValue([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $v = ([string]$Value).Trim()
    # Civil/PW sometimes stores unset QC_Review_Type as literal "0"; treat as missing.
    if ($v -eq '0') { return '' }
    return $v
}

function _PWD-PickSheetRoleFieldValue {
    param(
        [hashtable]$DgnFields,
        [hashtable]$PdfFields,
        [string]$Name
    )

    if ($Name -eq 'qcReviewType') {
        $dgnRt = if ($DgnFields -and $DgnFields.ContainsKey($Name)) {
            _PWD-NormalizeQcReviewTypeRoleValue ([string]$DgnFields[$Name])
        } else { '' }
        if (-not [string]::IsNullOrWhiteSpace($dgnRt)) { return $dgnRt }
        if ($PdfFields -and $PdfFields.ContainsKey($Name)) {
            return _PWD-NormalizeQcReviewTypeRoleValue ([string]$PdfFields[$Name])
        }
        return ''
    }

    if ($DgnFields -and $DgnFields.ContainsKey($Name) -and -not [string]::IsNullOrWhiteSpace([string]$DgnFields[$Name])) {
        return [string]$DgnFields[$Name]
    }
    if ($PdfFields -and $PdfFields.ContainsKey($Name) -and -not [string]::IsNullOrWhiteSpace([string]$PdfFields[$Name])) {
        return [string]$PdfFields[$Name]
    }
    return ''
}

function _PWD-PickSheetProcessIntentFieldValue {
    param(
        [hashtable]$DgnFields,
        [hashtable]$PdfFields,
        [string]$Name
    )

    if ($PdfFields -and $PdfFields.ContainsKey($Name) -and -not [string]::IsNullOrWhiteSpace([string]$PdfFields[$Name])) {
        return [string]$PdfFields[$Name]
    }
    if ($DgnFields -and $DgnFields.ContainsKey($Name) -and -not [string]::IsNullOrWhiteSpace([string]$DgnFields[$Name])) {
        return [string]$DgnFields[$Name]
    }
    return ''
}

function Get-PWQcPrependProcessIntentFromSourcePdf {
    <#
    .SYNOPSIS
    Resolves QC_Process_Type for prepend lane selection from the stem sheet PDF first, then the paired DGN.
    .DESCRIPTION
    QC_Review_Type is deprecated and is not read. Role emails still prefer the DGN
    (see Get-PWQcPrependRoleFieldsFromSourcePdf). For lane selection, the control sheet PDF
    QC_Process_Type reflects the user's Initiate Origination intent and wins over stale DGN values.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $folderCandidates = _PWD-GetSheetRoleFolderCandidates -FolderPath $FolderPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$SourceDocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) {
        return @{ found = $false; error = 'Invalid source document name'; qcProcessType = '' }
    }

    $dgnName = if ([string]$SourceDocumentName -match '(?i)\.dgn$') { [string]$SourceDocumentName } else { $stem + '.dgn' }
    $pdfName = if ([string]$SourceDocumentName -match '(?i)\.pdf$' -and [string]$SourceDocumentName -notmatch '(?i)-(prod|chk|rev|qc)\.pdf$') {
        [string]$SourceDocumentName
    } else {
        $stem + '.pdf'
    }

    $dgnRead = _PWD-TryReadSheetRoleFieldsFromDocument -FolderCandidates $folderCandidates -DocumentName $dgnName -Columns $cols -Config $Config
    $pdfRead = $null
    if ($pdfName -ne $dgnName) {
        $pdfRead = _PWD-TryReadSheetRoleFieldsFromDocument -FolderCandidates $folderCandidates -DocumentName $pdfName -Columns $cols -Config $Config
    }
    if (-not $dgnRead.found -and (-not $pdfRead -or -not $pdfRead.found)) {
        $err = if ($dgnRead.error) { [string]$dgnRead.error } elseif ($pdfRead -and $pdfRead.error) { [string]$pdfRead.error } else { 'Document not found' }
        return @{ found = $false; error = $err; qcProcessType = '' }
    }

    $dgnFields = if ($dgnRead.found) { $dgnRead.fields } else { $null }
    $pdfFields = if ($pdfRead -and $pdfRead.found) { $pdfRead.fields } else { $null }
    return @{
        found         = $true
        error         = ''
        qcProcessType = _PWD-PickSheetProcessIntentFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'qcProcessType'
    }
}

function Get-PWQcPrependRoleFieldsFromSourcePdf {
    <#
    .SYNOPSIS
    Reads QC role emails and QC_Review_Type for a sheet package member.
    .DESCRIPTION
    The paired DGN is the source of truth for role emails and review type. The sheet PDF
    is read only to fill fields that are blank on the DGN.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$SourceDocumentName,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $cols = @(Get-PWSheetIndexSyncColumnNames -Config $Config)
    $folderCandidates = _PWD-GetSheetRoleFolderCandidates -FolderPath $FolderPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$SourceDocumentName)
    if ([string]::IsNullOrWhiteSpace($stem)) {
        return @{ found = $false; error = 'Invalid source document name'; designerEmail = ''; reviewerEmail = ''; checkerEmail = ''; qcReviewType = ''; qcProcessType = '' }
    }

    $dgnName = if ([string]$SourceDocumentName -match '(?i)\.dgn$') { [string]$SourceDocumentName } else { $stem + '.dgn' }
    $pdfName = if ([string]$SourceDocumentName -match '(?i)\.pdf$' -and [string]$SourceDocumentName -notmatch '(?i)-(prod|chk|rev|qc)\.pdf$') {
        [string]$SourceDocumentName
    } else {
        $stem + '.pdf'
    }

    $dgnRead = _PWD-TryReadSheetRoleFieldsFromDocument -FolderCandidates $folderCandidates -DocumentName $dgnName -Columns $cols -Config $Config
    $pdfRead = $null
    if ($pdfName -ne $dgnName) {
        $pdfRead = _PWD-TryReadSheetRoleFieldsFromDocument -FolderCandidates $folderCandidates -DocumentName $pdfName -Columns $cols -Config $Config
    }

    if (-not $dgnRead.found -and (-not $pdfRead -or -not $pdfRead.found)) {
        $err = if ($dgnRead.error) { [string]$dgnRead.error } elseif ($pdfRead -and $pdfRead.error) { [string]$pdfRead.error } else { 'Document not found' }
        return @{ found = $false; error = $err; designerEmail = ''; reviewerEmail = ''; checkerEmail = ''; qcReviewType = ''; qcProcessType = '' }
    }

    $dgnFields = if ($dgnRead.found) { $dgnRead.fields } else { $null }
    $pdfFields = if ($pdfRead -and $pdfRead.found) { $pdfRead.fields } else { $null }
    $designerEmail = _PWD-PickSheetRoleFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'designerEmail'
    $reviewerEmail = _PWD-PickSheetRoleFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'reviewerEmail'
    $checkerEmail = _PWD-PickSheetRoleFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'checkerEmail'
    $qcProcessType = _PWD-PickSheetProcessIntentFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'qcProcessType'
    $qcReviewType = _PWD-PickSheetRoleFieldValue -DgnFields $dgnFields -PdfFields $pdfFields -Name 'qcReviewType'
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
        qcProcessType  = $qcProcessType
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
    Syncs role emails (designer, reviewer, checker) and QC_Review_Type from source sheet PDF to *-qc.pdf (PW environment attributes).
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
            -QcDocumentName $QcDocumentName -Config $Config | Out-Null
    }
    Sync-PWQcPdfReviewTypeFromSourcePdf -FolderPath $FolderPath -SourceDocumentName $SourceDocumentName `
        -QcDocumentName $QcDocumentName -Config $Config | Out-Null
}

Export-ModuleMember -Function *
