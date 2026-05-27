# PW.Discovery.psm1
# Responsibility: Read-only ProjectWise watch-path resolution and candidate discovery.

Import-Module (Join-Path $PSScriptRoot 'Core.Runtime.psm1') -Force

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

    $cmd = Get-Command -Name 'Get-PWDocumentsBySearchWithReturnColumns' -ErrorAction SilentlyContinue
    if ($cmd) {
        try {
            $params = @{
                FolderPath     = $apiPath
                JustThisFolder = $true
                DocumentName   = $DocumentName
                ErrorAction    = 'SilentlyContinue'
            }
            if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
                $params['ColumnsToReturn'] = @('Description', 'Name', 'DocumentName')
            } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
                $params['ReturnColumns'] = @('Description', 'Name', 'DocumentName')
            }
            $row = & $cmd @params | Select-Object -First 1
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

function Get-PWDocumentAttributeMap {
    <#
    .SYNOPSIS
    Parses .Attributes sorted-list bags from a Get-PWDocumentsBySearchWithReturnColumns row into a hashtable.
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$DocRow)

    $map = @{}
    if (-not $DocRow -or -not $DocRow.Attributes) { return $map }
    foreach ($bag in @($DocRow.Attributes)) {
        if ($bag -is [System.Collections.IDictionary]) {
            foreach ($k in $bag.Keys) {
                $map[[string]$k] = [string]$bag[$k]
            }
        }
    }
    return $map
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

    $cmd = Get-Command -Name 'Get-PWDocumentsBySearchWithReturnColumns' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return @{ designerEmail = ''; reviewerEmail = ''; found = $false; error = 'Get-PWDocumentsBySearchWithReturnColumns not available' }
    }

    $apiPath = ConvertTo-PWCmdletFolderPath -InternalFolderPath $FolderPath
    if ([string]::IsNullOrWhiteSpace($apiPath)) { $apiPath = $FolderPath }

    $cols = @($DesignerEmailColumn, $ReviewerEmailColumn) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $row = $null
    try {
        $params = @{
            FolderPath = $apiPath
            JustThisFolder = $true
            DocumentName = $DocumentName
            ErrorAction = 'Stop'
        }
        if ($cmd.Parameters.ContainsKey('ColumnsToReturn')) {
            $params['ColumnsToReturn'] = $cols
        } elseif ($cmd.Parameters.ContainsKey('ReturnColumns')) {
            $params['ReturnColumns'] = $cols
        }
        $row = & $cmd @params | Select-Object -First 1
    } catch {
        return @{ designerEmail = ''; reviewerEmail = ''; found = $false; error = $_.Exception.Message }
    }

    if (-not $row) {
        return @{ designerEmail = ''; reviewerEmail = ''; found = $false; error = 'Document not found' }
    }

    $attrs = Get-PWDocumentAttributeMap -DocRow $row
    $designer = if ($attrs.ContainsKey($DesignerEmailColumn)) { [string]$attrs[$DesignerEmailColumn] } else { '' }
    $reviewer = if ($attrs.ContainsKey($ReviewerEmailColumn)) { [string]$attrs[$ReviewerEmailColumn] } else { '' }
    $pwState = $null
    try { $pwState = [string]$row.WorkflowState } catch { }
    if ([string]::IsNullOrWhiteSpace($pwState)) {
        try { $pwState = [string]$row.StateName } catch { }
    }
    return @{
        designerEmail = $designer.Trim()
        reviewerEmail = $reviewer.Trim()
        pwStateName   = if ($pwState) { $pwState.Trim() } else { '' }
        found = $true
        document = $row
        error = $null
    }
}

function _PWD-NormalizeSheetIndexValue {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Sync-PWSheetIndexOwnership {
    <#
    .SYNOPSIS
    Reads designer/reviewer emails and workflow state from ProjectWise and updates sheet_index when they differ from the database.
    .DESCRIPTION
    Intended for audit-trail DOCUMENT_ATTR (and DOCUMENT_STATE) events on watchlist documents.
    Inserts new sheet_index rows only for Sheets-folder paths; updates existing rows for any watchlist path.
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
        [string]$AuditActionName = '',
        [string]$DesignerEmailColumn = 'EM_Designer_Email',
        [string]$ReviewerEmailColumn = 'EM_Reviewer_Email'
    )

    if (-not (Get-Command -Name 'Test-QCDatabaseEnabled' -ErrorAction SilentlyContinue)) { return }
    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return }

    $pw = Get-PWDocumentEmailContacts -FolderPath $FolderPath -DocumentName $DocumentName `
        -DesignerEmailColumn $DesignerEmailColumn -ReviewerEmailColumn $ReviewerEmailColumn
    if (-not $pw.found) { return }

    $pwDesigner = [string]$pw.designerEmail
    $pwReviewer = [string]$pw.reviewerEmail
    $pwState    = [string]$pw.pwStateName

    $dbDesigner = ''
    $dbReviewer = ''
    $dbState    = ''
    $rowExists  = $false
    try {
        $dbRes = Invoke-QCDatabaseQuery -Config $Config -Sql @"
SELECT designer_email, reviewer_email, pw_state_name
FROM sheet_index
WHERE document_guid = @docGuid
"@ -Parameters @{ docGuid = $DocumentGuid }
        if ($dbRes.IsSuccess -and $dbRes.Data.table -and $dbRes.Data.table.Rows.Count -gt 0) {
            $rowExists = $true
            $r = $dbRes.Data.table.Rows[0]
            if (-not ($r.designer_email -is [DBNull])) { $dbDesigner = [string]$r.designer_email }
            if (-not ($r.reviewer_email -is [DBNull])) { $dbReviewer = [string]$r.reviewer_email }
            if (-not ($r.pw_state_name -is [DBNull])) { $dbState = [string]$r.pw_state_name }
        }
    } catch { return }

    if (-not $rowExists -and -not $IsSheetsFolder) { return }

    $emailsDiffer = (_PWD-NormalizeSheetIndexValue $pwDesigner) -ne (_PWD-NormalizeSheetIndexValue $dbDesigner) `
        -or (_PWD-NormalizeSheetIndexValue $pwReviewer) -ne (_PWD-NormalizeSheetIndexValue $dbReviewer)
    $stateDiffers = (_PWD-NormalizeSheetIndexValue $pwState) -ne (_PWD-NormalizeSheetIndexValue $dbState)

    if ($rowExists -and -not $emailsDiffer -and -not $stateDiffers) {
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
        -DesignerEmail $pwDesigner -ReviewerEmail $pwReviewer -PwStateName $pwState `
        -LastAuditEventAt $LastAuditEventAt -SetOwnershipFromProjectWise

    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_SHEET_INDEX_SYNC' -Message 'sheet_index ownership synced from ProjectWise.' -Data @{
            documentGuid    = $DocumentGuid
            documentName    = $DocumentName
            folderPath      = $FolderPath
            auditActionName = $AuditActionName
            designerEmail   = $pwDesigner
            reviewerEmail   = $pwReviewer
            pwStateName     = $pwState
            wasInsert       = (-not $rowExists)
            emailsChanged   = $emailsDiffer
            stateChanged    = $stateDiffers
        }
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

Export-ModuleMember -Function *
