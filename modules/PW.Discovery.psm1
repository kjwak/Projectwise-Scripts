# PW.Discovery.psm1
# Responsibility: Read-only ProjectWise watch-path resolution and candidate discovery.

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

function Get-PWDocLastModifiedUtc {
    [CmdletBinding()]
    param([AllowNull()][object]$Doc)

    foreach ($n in @('FileUpdatedDate','FileUpdateDate','DocumentUpdateDate','VersionModifiedDate','Version Modified Date')) {
        $v = Get-PWObjectPropertyValue -Object $Doc -Name $n
        if ($v) {
            try {
                if ($v -is [DateTime]) { return $v.ToUniversalTime().ToString('o') }
                if ($v -is [DateTimeOffset]) { return $v.UtcDateTime.ToString('o') }
                $dt = [DateTime]::Parse([string]$v, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
                return $dt.ToUniversalTime().ToString('o')
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
        try { $view = Get-PWFolderView -FolderPath $FolderPath -ErrorAction Stop } catch { }
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
                $view = $folder | Get-PWFolderView -ErrorAction SilentlyContinue
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
                $null = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
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
            try {
                $view = Get-PWFolderView -FolderPath $p -WarningAction SilentlyContinue -ErrorAction Stop
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

Export-ModuleMember -Function *
