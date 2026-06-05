# PW.AuditPoller.psm1
# Responsibility: Query ProjectWise audit trail (dms_audt), resolve folders,
# match against watch roots, and return candidate events for job creation.
# Extracted from the POC Test-AuditEventIngestion.ps1 for production use.

# Dependencies (Core.Results, Core.Runtime, Core.Database) must be imported by the
# caller before this module. Re-importing with -Force here would clobber their
# global-scope exports.

# Bump when trigger/candidate logic changes; appears in WATCH_AUDIT_SCAN_DONE logs.
$script:AuditPollerLogicVersion = '2026-06-05-watermark-pagination-safety-v8'

$script:QCRelevantActions = @{
    1001 = 'DOCUMENT_CREATE'
    1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR'
    1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN'
    1012 = 'DOCUMENT_STATE'
    1015 = 'DOCUMENT_VERSION'
    1020 = 'DOCUMENT_DELETE'
}

# Full PW audit trail action map (dms_audt o_action / pw_action). Only QCRelevantActions are written to audit_events;
# all other actions still advance the watermark on fetch. Source: Bentley ProjectWise audit action constants.
$script:AuditActionNames = @{
    # Folders
    1    = 'FOLDER_CREATE';           2    = 'FOLDER_MODIFY';           3    = 'FOLDER_WFLOW'
    4    = 'FOLDER_DELETE';           5    = 'FOLDER_STATE';            6    = 'FOLDER_ACL_ASSIGN'
    7    = 'FOLDER_ACL_MODIFY';       8    = 'FOLDER_ACL_REMOVE';       9    = 'FOLDER_CONNECT_ASSIGN'
    10   = 'FOLDER_CONNECT_CHANGE';   11   = 'FOLDER_CONNECT_REMOVE';   12   = 'FOLDER_RESTORE'
    100  = 'FOLDER_CUSTOM_FIRST';     999  = 'FOLDER_CUSTOM_LAST'
    # Documents
    1000 = 'DOCUMENT_UNKNOWN';        1001 = 'DOCUMENT_CREATE';        1002 = 'DOCUMENT_MODIFY'
    1003 = 'DOCUMENT_ATTR';           1004 = 'DOCUMENT_FILE_ADD';      1005 = 'DOCUMENT_FILE_REM'
    1006 = 'DOCUMENT_FILE_REP';       1007 = 'DOCUMENT_CIN';           1008 = 'DOCUMENT_VIEW'
    1009 = 'DOCUMENT_CHOUT';          1010 = 'DOCUMENT_CPOUT';         1011 = 'DOCUMENT_GOUT'
    1012 = 'DOCUMENT_STATE';          1013 = 'DOCUMENT_FINAL_S';       1014 = 'DOCUMENT_FINAL_R'
    1015 = 'DOCUMENT_VERSION';        1016 = 'DOCUMENT_MOVE';          1017 = 'DOCUMENT_COPY'
    1018 = 'DOCUMENT_SECUR';          1019 = 'DOCUMENT_REDLINE';       1020 = 'DOCUMENT_DELETE'
    1021 = 'DOCUMENT_EXPORT';         1022 = 'DOCUMENT_FREE';          1023 = 'DOCUMENT_EXTRACT'
    1024 = 'DOCUMENT_DISTRIBUTE';     1025 = 'DOCUMENT_SEND_TO';       1026 = 'DOCUMENT_COMMENT'
    1027 = 'DOCUMENT_IMPORT';         1028 = 'DOCUMENT_ACL_ASSIGN';    1029 = 'DOCUMENT_ACL_MODIFY'
    1030 = 'DOCUMENT_ACL_REMOVE';     1031 = 'DOCUMENT_REVIT';         1032 = 'DOCUMENT_PACK'
    1033 = 'DOCUMENT_UNPACK';         1034 = 'DOCUMENT_WRE_START';     1035 = 'DOCUMENT_WRE_END'
    1036 = 'DOCUMENT_WRE_FAILURE';    1037 = 'DOCUMENT_RESTORE'
    1100 = 'DOCUMENT_CUSTOM_FIRST';   1999 = 'DOCUMENT_CUSTOM_LAST'
    # Sets
    2001 = 'SET_CREATE';              2002 = 'SET_ADD';                2003 = 'SET_REMOVE'
    2100 = 'SET_CUSTOM_FIRST';        2999 = 'SET_CUSTOM_LAST'
    # Users
    3001 = 'USER_LOGIN';              3002 = 'USER_LOGOUT';            3003 = 'USER_CREATE'
    3004 = 'USER_MODIFY';             3005 = 'USER_SETTINGS';          3006 = 'USER_RENAME'
    3007 = 'USER_DELETE';             3008 = 'USER_DISABLE';           3009 = 'USER_ENABLE'
    # Groups
    4001 = 'GROUP_CREATE';            4002 = 'GROUP_MODIFY';           4003 = 'GROUP_ADD'
    4004 = 'GROUP_REMOVE';            4005 = 'GROUP_RENAME';           4006 = 'GROUP_DELETE'
    4007 = 'GROUP_ACL_ASSIGN';        4008 = 'GROUP_ACL_MODIFY';       4009 = 'GROUP_ACL_REMOVE'
    # User lists
    5001 = 'USER_LIST_CREATE';        5002 = 'USER_LIST_MODIFY';       5003 = 'USER_LIST_ADD'
    5004 = 'USER_LIST_REMOVE';        5005 = 'USER_LIST_RENAME';       5006 = 'USER_LIST_DELETE'
    5007 = 'USER_LIST_ACL_ASSIGN';    5008 = 'USER_LIST_ACL_MODIFY';   5009 = 'USER_LIST_ACL_REMOVE'
    # States
    6001 = 'STATE_CREATE';            6002 = 'STATE_RENAME';           6003 = 'STATE_DELETE'
    # Workflows
    7001 = 'WORKFLOW_CREATE';         7002 = 'WORKFLOW_ADD';           7003 = 'WORKFLOW_REMOVE'
    7004 = 'WORKFLOW_RENAME';         7005 = 'WORKFLOW_DELETE';        7006 = 'WORKFLOW_MOVE'
    # Interfaces / environments / views
    8001 = 'INTERFACE_CREATE';        8002 = 'INTERFACE_RENAME';       8003 = 'INTERFACE_DELETE'
    8101 = 'ENVIRONMENT_CREATE';      8102 = 'ENVIRONMENT_MODIFY';     8103 = 'ENVIRONMENT_RENAME'
    8104 = 'ENVIRONMENT_DELETE';      8105 = 'ENVIRONMENT_VIEW_REMOVE'; 8106 = 'ENVIRONMENT_VIEW_CHANGE'
    9001 = 'VIEW_CREATE';             9002 = 'VIEW_RENAME';            9003 = 'VIEW_DELETE'
    9004 = 'VIEW_CHANGE'
    # Applications / departments / environment attributes
    9101 = 'APPLICATION_CREATE';      9102 = 'APPLICATION_RENAME';     9103 = 'APPLICATION_DELETE'
    9104 = 'APPLICATION_VIEWER'
    9201 = 'DEPARTMENT_CREATE';       9202 = 'DEPARTMENT_RENAME';      9203 = 'DEPARTMENT_DELETE'
    9301 = 'ENVIRONMENT_ATTR_ADD';    9302 = 'ENVIRONMENT_ATTR_REMOVE'; 9303 = 'ENVIRONMENT_ATTR_MODIFY'
    # Work areas
    20001 = 'WORKAREA_CREATE';        20002 = 'WORKAREA_PROP_ADD';     20003 = 'WORKAREA_PROP_REMOVE'
    20004 = 'WORKAREA_RENAME';        20005 = 'WORKAREA_DELETE'
}

# Session caches: avoid repeated Get-PWDocumentsByGUIDs / Get-PWFoldersByGUIDs for missing or known objects.
$script:AuditPoller_DocFolderCache = @{}
$script:AuditPoller_UnresolvedGuids = @{}
$script:AuditPoller_FolderGuidCache = @{}
$script:AuditPoller_UnresolvedFolderGuids = @{}
$script:AuditPoller_WatchFolderCacheWarmed = $false

function _AuditPoller-LoadDocFolderCache {
    param([hashtable]$Config)
    if ($script:AuditPoller_DocFolderCache.Count -eq 0 -and (Get-Command -Name 'Get-QCDocumentFolderCache' -ErrorAction SilentlyContinue)) {
        try {
            $res = Get-QCDocumentFolderCache -Config $Config
            if ($res.IsSuccess -and $res.Data -and $res.Data.cache) {
                foreach ($k in $res.Data.cache.Keys) { $script:AuditPoller_DocFolderCache[$k] = $res.Data.cache[$k] }
            }
        } catch { }
    }
    return $script:AuditPoller_DocFolderCache
}

function _AuditPoller-GetAuditPollerBool {
    param([hashtable]$Config, [string]$Key, [bool]$Default = $true)
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable] -and $ap.ContainsKey($Key) -and $null -ne $ap[$Key]) { return [bool]$ap[$Key] }
            if ($ap.PSObject -and $null -ne $ap.$Key) { return [bool]$ap.$Key }
            if ($ap -is [hashtable] -and $ap.ContainsKey('folderGuidCache') -and $ap.folderGuidCache) {
                $fg = $ap.folderGuidCache
                if ($fg -is [hashtable] -and $fg.ContainsKey($Key) -and $null -ne $fg[$Key]) { return [bool]$fg[$Key] }
                if ($fg.PSObject -and $null -ne $fg.$Key) { return [bool]$fg.$Key }
            }
            elseif ($ap.PSObject -and $ap.folderGuidCache) {
                $fg = $ap.folderGuidCache
                if ($fg -is [hashtable] -and $fg.ContainsKey($Key) -and $null -ne $fg[$Key]) { return [bool]$fg[$Key] }
                if ($fg.PSObject -and $null -ne $fg.$Key) { return [bool]$fg.$Key }
            }
        }
    } catch { }
    return $Default
}

function _AuditPoller-GetPwObjectProperty {
    param([object]$Object, [string]$Name)
    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    try {
        if ($Object -is [hashtable] -and $Object.ContainsKey($Name)) { return $Object[$Name] }
        if ($Object.PSObject -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    } catch { }
    return $null
}

function _AuditPoller-NormalizeFolderGuidKey {
    param([AllowNull()][string]$Guid)
    $g = ([string]$Guid).Trim().Trim('{}').Trim()
    if ([string]::IsNullOrWhiteSpace($g)) { return $null }
    return $g.ToLowerInvariant()
}

function _AuditPoller-GetFolderGuidLookupVariants {
    param([AllowNull()][string]$Guid)
    $key = _AuditPoller-NormalizeFolderGuidKey -Guid $Guid
    if (-not $key) { return @() }
    $variants = [System.Collections.Generic.List[string]]::new()
    $raw = ([string]$Guid).Trim()
    [void]$variants.Add($raw)
    [void]$variants.Add($key)
    [void]$variants.Add('{' + $key + '}')
    [void]$variants.Add($key.ToUpperInvariant())
    [void]$variants.Add('{' + $key.ToUpperInvariant() + '}')
    return @($variants | Select-Object -Unique)
}

function _AuditPoller-SqlCastGuidList {
    param([string[]]$Guids)
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($g in @($Guids)) {
        $k = _AuditPoller-NormalizeFolderGuidKey -Guid $g
        if (-not $k) { continue }
        $escaped = $k -replace '''', ''''''
        [void]$parts.Add("CAST('$escaped' AS UNIQUEIDENTIFIER)")
    }
    return ($parts -join ',')
}

function _AuditPoller-PreparePwFolderObject {
    param([object]$Folder)
    if ($null -eq $Folder) { return $null }
    try {
        $m = $Folder.GetType().GetMethod('GetFullPath')
        if ($m) { $null = $m.Invoke($Folder, @()) }
    } catch { }
    return $Folder
}

function _AuditPoller-ExtractFolderPathFromPwFolder {
    param([object]$Folder)
    if ($null -eq $Folder) { return $null }
    $Folder = _AuditPoller-PreparePwFolderObject -Folder $Folder
    if ($Folder -is [string]) {
        $s = ([string]$Folder).Trim()
        if ($s -match '\\') { return _AuditPoller-NormalizeFolderPath -FolderPath $s }
        return $null
    }
    foreach ($prop in @(
        'FolderPath', 'Path', 'FullPath', 'folderPath', 'CanonicalPath', 'PWPath',
        'FullName', 'Name', 'FolderName', 'ParentPath', 'DocumentPath'
    )) {
        $v = _AuditPoller-GetPwObjectProperty -Object $Folder -Name $prop
        if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) {
            $s = [string]$v
            if ($prop -in @('Name', 'FolderName') -and $s -notmatch '\\') { continue }
            $norm = _AuditPoller-NormalizeFolderPath -FolderPath $s
            if ($norm) { return $norm }
        }
    }
    try {
        if ($Folder.PSObject) {
            foreach ($p in $Folder.PSObject.Properties) {
                $n = [string]$p.Name
                if ($n -match '(?i)path|folder') {
                    $v = $p.Value
                    if ($null -ne $v -and "$v" -match '\\') {
                        $norm = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$v)
                        if ($norm) { return $norm }
                    }
                }
            }
        }
    } catch { }
    return $null
}

function _AuditPoller-ExtractFolderGuidFromPwFolder {
    param([object]$Folder)
    if ($null -eq $Folder) { return $null }
    if ($Folder -is [string]) { return $null }
    foreach ($prop in @('FolderGUID', 'GUID', 'ObjectGUID', 'o_objguid', 'folderGUID', 'FolderGuid')) {
        $v = _AuditPoller-GetPwObjectProperty -Object $Folder -Name $prop
        if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) {
            return ([string]$v).Trim().Trim('{}')
        }
    }
    return $null
}

function _AuditPoller-TryApplyFolderLookup {
    param(
        [hashtable]$Config,
        [string]$RequestedGuid,
        [object]$FolderObject,
        [hashtable]$ParentToFolder,
        [hashtable]$FoundByGuid,
        [ref]$StatsRef
    )
    $fp = _AuditPoller-ExtractFolderPathFromPwFolder -Folder $FolderObject
    if (-not $fp -and $FolderObject -is [string]) {
        $fp = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$FolderObject)
    }
    if (-not $fp) { return $false }
    $fg = _AuditPoller-ExtractFolderGuidFromPwFolder -Folder $FolderObject
    $reqKey = _AuditPoller-NormalizeFolderGuidKey -Guid $RequestedGuid
    $keys = [System.Collections.Generic.List[string]]::new()
    if ($reqKey) { [void]$keys.Add($reqKey) }
    if ($fg) { [void]$keys.Add((_AuditPoller-NormalizeFolderGuidKey -Guid $fg)) }
    foreach ($k in @($keys | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $FoundByGuid[$k] = $fp
        _AuditPoller-RegisterFolderGuidPath -Config $Config -FolderGuid $k -FolderPath $fp -StatsRef $StatsRef
        if ($RequestedGuid) { $ParentToFolder[$RequestedGuid] = $fp }
        foreach ($v in @(_AuditPoller-GetFolderGuidLookupVariants -Guid $RequestedGuid)) {
            if (-not [string]::IsNullOrWhiteSpace($v)) { $ParentToFolder[$v] = $fp }
        }
    }
    return ($reqKey -and $FoundByGuid.ContainsKey($reqKey))
}

function _AuditPoller-TestPwDatasourceLoggedIn {
    if (-not (Get-Command -Name 'Select-PWSQL' -ErrorAction SilentlyContinue)) { return $false }
    try {
        $r = Select-PWSQL -SQLSelectStatement 'SELECT 1 AS ok' -ErrorAction Stop
        return ($null -ne $r)
    } catch {
        return $false
    }
}

function _AuditPoller-FetchPwFolderObjectsViaSqlProjectNo {
    param([string[]]$ChunkGuids)

    $byKey = @{}
    if (-not (Get-Command -Name 'Select-PWSQL' -ErrorAction SilentlyContinue)) { return $byKey }
    if (-not (Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue)) { return $byKey }

    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($g in @($ChunkGuids)) {
        $k = _AuditPoller-NormalizeFolderGuidKey -Guid $g
        if ($k) { [void]$keys.Add($k) }
    }
    $unique = @($keys | Sort-Object -Unique)
    if ($unique.Count -eq 0) { return $byKey }

    $inList = _AuditPoller-SqlCastGuidList -Guids $unique
    if (-not $inList) { return $byKey }
    $sql = "SELECT o_projguid, o_projectno FROM dms_proj WHERE o_projguid IN ($inList)"
    try {
        $dt = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
    } catch {
        return $byKey
    }
    if (-not $dt -or -not $dt.Rows) { return $byKey }

    foreach ($row in $dt.Rows) {
        $g = _AuditPoller-NormalizeFolderGuidKey -Guid ([string]$row.o_projguid)
        if (-not $g) { continue }
        try {
            $projNo = [int]$row.o_projectno
            $f = Get-PWFolders -FolderID $projNo -JustOne -ErrorAction Stop
            if (-not $f) { continue }
            $byKey[$g] = _AuditPoller-PreparePwFolderObject -Folder $f
        } catch { }
    }
    return $byKey
}

function _AuditPoller-RegisterFetchedFolderInByKey {
    param(
        [object]$FolderObject,
        [string]$RequestedGuid,
        [hashtable]$ByKey
    )
    if ($null -eq $FolderObject) { return }
    $fg = _AuditPoller-ExtractFolderGuidFromPwFolder -Folder $FolderObject
    $nk = _AuditPoller-NormalizeFolderGuidKey -Guid $fg
    $rk = _AuditPoller-NormalizeFolderGuidKey -Guid $RequestedGuid
    if ($nk) { $ByKey[$nk] = $FolderObject }
    if ($rk -and -not $ByKey.ContainsKey($rk)) { $ByKey[$rk] = $FolderObject }
}

function _AuditPoller-FetchPwFolderByGuidCmdlet {
    param([string]$Guid)
    if (-not (Get-Command -Name 'Get-PWFoldersByGUIDs' -ErrorAction SilentlyContinue)) { return $null }
    foreach ($v in @(_AuditPoller-GetFolderGuidLookupVariants -Guid $Guid)) {
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        try {
            $one = @(Get-PWFoldersByGUIDs -FolderGUIDs @($v) -ErrorAction Stop) | Where-Object { $null -ne $_ }
            if ($one.Count -gt 0) {
                return (_AuditPoller-PreparePwFolderObject -Folder $one[0])
            }
        } catch { }
    }
    return $null
}

function _AuditPoller-FetchPwFoldersForGuidChunk {
    param([string[]]$ChunkGuids)

    $byKey = @{}
    $batchGuids = @($ChunkGuids | ForEach-Object {
        $k = _AuditPoller-NormalizeFolderGuidKey -Guid $_
        if ($k) { $k } else { ([string]$_).Trim() }
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($batchGuids.Count -eq 0) { return $byKey }

    # pwps_dab: batch with many GUID variants often returns nothing; use one canonical guid per parent.
    if (Get-Command -Name 'Get-PWFoldersByGUIDs' -ErrorAction SilentlyContinue) {
        try {
            $folders = @(Get-PWFoldersByGUIDs -FolderGUIDs $batchGuids -ErrorAction Stop) | Where-Object { $null -ne $_ }
            foreach ($f in $folders) {
                $prepared = _AuditPoller-PreparePwFolderObject -Folder $f
                _AuditPoller-RegisterFetchedFolderInByKey -FolderObject $prepared -RequestedGuid '' -ByKey $byKey | Out-Null
            }
        } catch { }
    }

    foreach ($req in @($ChunkGuids)) {
        $rk = _AuditPoller-NormalizeFolderGuidKey -Guid $req
        if (-not $rk -or $byKey.ContainsKey($rk)) { continue }
        $f = _AuditPoller-FetchPwFolderByGuidCmdlet -Guid $req
        if ($f) {
            _AuditPoller-RegisterFetchedFolderInByKey -FolderObject $f -RequestedGuid $req -ByKey $byKey | Out-Null
        }
    }

    $stillNeed = @($ChunkGuids | Where-Object {
        $rk = _AuditPoller-NormalizeFolderGuidKey -Guid $_
        $rk -and -not $byKey.ContainsKey($rk)
    })
    if ($stillNeed.Count -gt 0) {
        $sqlByKey = _AuditPoller-FetchPwFolderObjectsViaSqlProjectNo -ChunkGuids $stillNeed
        foreach ($k in @($sqlByKey.Keys)) {
            if (-not $byKey.ContainsKey($k)) { $byKey[$k] = $sqlByKey[$k] }
        }
    }

    $stillNeed2 = @($ChunkGuids | Where-Object {
        $rk = _AuditPoller-NormalizeFolderGuidKey -Guid $_
        $rk -and -not $byKey.ContainsKey($rk)
    })
    if ($stillNeed2.Count -gt 0 -and (Get-Command -Name 'Get-PWDocumentsByGUIDs' -ErrorAction SilentlyContinue)) {
        try {
            $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs $stillNeed2 -ErrorAction Stop) | Where-Object { $null -ne $_ }
            foreach ($doc in $docs) {
                $dg = [string]$doc.DocumentGUID
                if (-not $dg) { continue }
                $fp = $null
                if ($doc.FolderPath) { $fp = [string]$doc.FolderPath }
                elseif ($doc.FullPath) {
                    $full = [string]$doc.FullPath
                    $fp = [System.IO.Path]::GetDirectoryName($full) -replace '/', '\'
                }
                if (-not $fp) { continue }
                $nk = _AuditPoller-NormalizeFolderGuidKey -Guid $dg
                if ($nk) { $byKey[$nk] = $fp }
            }
        } catch { }
    }

    return $byKey
}

function _AuditPoller-RegisterFolderGuidPath {
    param(
        [hashtable]$Config,
        [string]$FolderGuid,
        [string]$FolderPath,
        [string]$WatchRoot = '',
        [ref]$StatsRef = $null
    )
    if ([string]::IsNullOrWhiteSpace($FolderGuid) -or [string]::IsNullOrWhiteSpace($FolderPath)) { return }
    $key = _AuditPoller-NormalizeFolderGuidKey -Guid $FolderGuid
    if (-not $key) { return }
    $canonical = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $canonical) { return }
    $script:AuditPoller_FolderGuidCache[$key] = $canonical
    if ($script:AuditPoller_UnresolvedFolderGuids.ContainsKey($key)) {
        $script:AuditPoller_UnresolvedFolderGuids.Remove($key) | Out-Null
    }
    if ($StatsRef.Value) { $StatsRef.Value.parentFoldersResolved++ }
    $ttl = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'folderCacheTtlSeconds' -Default 86400
    if (Get-Command -Name 'Set-QCPwFolderCacheEntry' -ErrorAction SilentlyContinue) {
        Set-QCPwFolderCacheEntry -Config $Config -FolderGuid $FolderGuid -FolderPath $canonical -WatchRoot $WatchRoot -TtlSeconds $ttl | Out-Null
    }
}

function _AuditPoller-LoadFolderGuidCache {
    param([hashtable]$Config)
    if ($script:AuditPoller_FolderGuidCache.Count -eq 0 -and (Get-Command -Name 'Get-QCPwFolderGuidCache' -ErrorAction SilentlyContinue)) {
        try {
            $res = Get-QCPwFolderGuidCache -Config $Config
            if ($res.IsSuccess -and $res.Data -and $res.Data.cache) {
                foreach ($k in $res.Data.cache.Keys) { $script:AuditPoller_FolderGuidCache[$k] = $res.Data.cache[$k] }
            }
        } catch { }
    }
    return $script:AuditPoller_FolderGuidCache
}

function _AuditPoller-ResolveParentFoldersBatched {
    param(
        [hashtable]$Config,
        [string[]]$ParentGuids,
        [hashtable]$ParentToFolder,
        [ref]$StatsRef
    )

    $ttlSeconds = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'folderCacheTtlSeconds' -Default 86400
    $negTtlSeconds = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'negativeCacheTtlSeconds' -Default 1800
    [void](_AuditPoller-LoadFolderGuidCache -Config $Config)

    if (Get-Command -Name 'Get-QCPwFolderCacheBatch' -ErrorAction SilentlyContinue) {
        try {
            $cacheRes = Get-QCPwFolderCacheBatch -Config $Config -FolderGuids @($ParentGuids)
            if ($cacheRes.IsSuccess -and $cacheRes.Data.cache) {
                foreach ($k in $cacheRes.Data.cache.Keys) {
                    $entry = $cacheRes.Data.cache[$k]
                    if ($entry.resolveFailed) {
                        $script:AuditPoller_UnresolvedFolderGuids[$k] = $true
                        if ($StatsRef.Value) { $StatsRef.Value.failedFolderGuidCacheHits++ }
                        continue
                    }
                    if ($entry.folderPath) {
                        $canonical = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$entry.folderPath)
                        if ($canonical) {
                            $script:AuditPoller_FolderGuidCache[$k] = $canonical
                            foreach ($pg in @($ParentGuids)) {
                                if ((_AuditPoller-NormalizeFolderGuidKey -Guid $pg) -eq $k) { $ParentToFolder[$pg] = $canonical; break }
                            }
                            if ($StatsRef.Value) { $StatsRef.Value.parentFolderGuidCacheHits++ }
                        }
                    }
                }
            }
        } catch { }
    }

    $needPw = [System.Collections.Generic.List[string]]::new()
    foreach ($pg in @($ParentGuids)) {
        if ([string]::IsNullOrWhiteSpace($pg)) { continue }
        $key = $pg.Trim().ToLowerInvariant()
        if ($ParentToFolder.ContainsKey($pg)) { continue }
        if ($script:AuditPoller_FolderGuidCache.ContainsKey($key)) {
            $ParentToFolder[$pg] = $script:AuditPoller_FolderGuidCache[$key]
            if ($StatsRef.Value) { $StatsRef.Value.parentFolderGuidCacheHits++ }
            continue
        }
        if ($script:AuditPoller_UnresolvedFolderGuids.ContainsKey($key)) {
            if ($StatsRef.Value) { $StatsRef.Value.parentFolderGuidResolveSkipped++ }
            continue
        }
        [void]$needPw.Add($pg)
        if ($StatsRef.Value) { $StatsRef.Value.parentFolderGuidCacheMisses++ }
    }

    $canResolveFolders = (Get-Command -Name 'Get-PWFoldersByGUIDs' -ErrorAction SilentlyContinue) -or
        ((Get-Command -Name 'Select-PWSQL' -ErrorAction SilentlyContinue) -and (Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue))
    if (-not $canResolveFolders) { return }

    if (-not (_AuditPoller-TestPwDatasourceLoggedIn)) {
        if ($StatsRef.Value) { $StatsRef.Value.parentFolderResolveSkippedNotLoggedIn++ }
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_FOLDER_GUID_SKIP_NOT_LOGGED_IN' -Message 'Skipping parent folder GUID resolution: not logged into ProjectWise datasource.' -Data @{
                parentGuidCount = @($ParentGuids).Count
                logicVersion = $script:AuditPollerLogicVersion
            }
        }
        return
    }

    $batchSize = 200
    for ($i = 0; $i -lt $needPw.Count; $i += $batchSize) {
        $chunk = @($needPw[$i..[Math]::Min($i + $batchSize - 1, $needPw.Count - 1)])
        try {
            $folderByKey = _AuditPoller-FetchPwFoldersForGuidChunk -ChunkGuids $chunk
            $foundByGuid = @{}
            foreach ($pg in $chunk) {
                $rk = _AuditPoller-NormalizeFolderGuidKey -Guid $pg
                if (-not $rk) { continue }
                $folderObj = $null
                if ($folderByKey.ContainsKey($rk)) { $folderObj = $folderByKey[$rk] }
                if (-not $folderObj) { continue }
                [void](_AuditPoller-TryApplyFolderLookup -Config $Config -RequestedGuid $pg -FolderObject $folderObj `
                    -ParentToFolder $ParentToFolder -FoundByGuid $foundByGuid -StatsRef $StatsRef)
            }
            foreach ($pg in $chunk) {
                if ($ParentToFolder.ContainsKey($pg)) { continue }
                $key = _AuditPoller-NormalizeFolderGuidKey -Guid $pg
                if (-not $key) { continue }
                $hadPwObject = $folderByKey.ContainsKey($key)
                $script:AuditPoller_UnresolvedFolderGuids[$key] = $true
                if (Get-Command -Name 'Set-QCPwFolderCacheEntry' -ErrorAction SilentlyContinue) {
                    Set-QCPwFolderCacheEntry -Config $Config -FolderGuid $pg -ResolveFailed -TtlSeconds $negTtlSeconds | Out-Null
                }
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    $logCode = if ($hadPwObject) { 'AUDIT_FOLDER_GUID_NO_PATH' } else { 'AUDIT_FOLDER_GUID_NOT_FOUND' }
                    Write-QCJsonLog -Flush -Level 'Warning' -Code $logCode -Message "Could not resolve folder path for parent GUID $pg." -Data @{
                        folderGuid = $pg
                        hadPwObject = $hadPwObject
                        logicVersion = $script:AuditPollerLogicVersion
                    }
                }
            }
        } catch {
            foreach ($pg in $chunk) {
                $nk = _AuditPoller-NormalizeFolderGuidKey -Guid $pg
                if ($nk) { $script:AuditPoller_UnresolvedFolderGuids[$nk] = $true }
            }
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_FOLDER_GUID_RESOLVE_ERROR' -Message $_.Exception.Message -Data @{
                    folderGuids = @($chunk)
                    logicVersion = $script:AuditPollerLogicVersion
                }
            }
        }
    }
}

function Sync-AuditPollerWatchFolderGuidCache {
    <#
    .SYNOPSIS
    Pre-warms folder GUID -> path cache from projectWise.watchList.roots (and optional Sheets subfolders).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [array]$WatchRootConfigs = @()
    )

    if (-not (_AuditPoller-GetAuditPollerBool -Config $Config -Key 'warmWatchRootsOnStart' -Default $true)) {
        return New-QCSuccessResult -Code 'FOLDER_CACHE_WARM_SKIPPED' -Message 'warmWatchRootsOnStart is false.' -Data @{ warmed = 0 }
    }

    $warmed = 0
    $statsRef = [ref]@{ parentFoldersResolved = 0 }
    $normalized = @(_AuditPoller-NormalizeWatchRootConfigs -WatchRootConfigs $WatchRootConfigs -Config $Config)
    $warmSheets = _AuditPoller-GetAuditPollerBool -Config $Config -Key 'warmSheetsFoldersOnStart' -Default $true
    $maxSheets = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxSheetsFoldersToWarm' -Default 250

    foreach ($cfg in $normalized) {
        $rootPath = _AuditPoller-GetWatchRootPathFromConfig -Cfg $cfg
        if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }
        $apiRoot = $rootPath
        if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
            try {
                $converted = ConvertTo-PWCmdletFolderPath -InternalFolderPath $rootPath
                if (-not [string]::IsNullOrWhiteSpace($converted)) { $apiRoot = $converted }
            } catch { }
        }
        if (Get-Command -Name 'Get-PWFolders' -ErrorAction SilentlyContinue) {
            try {
                $f = Get-PWFolders -FolderPath $apiRoot -JustOne -ErrorAction SilentlyContinue
                $fg = _AuditPoller-ExtractFolderGuidFromPwFolder -Folder $f
                $fp = _AuditPoller-ExtractFolderPathFromPwFolder -Folder $f
                if (-not $fp) { $fp = _AuditPoller-NormalizeFolderPath -FolderPath $rootPath }
                if ($fg -and $fp) {
                    _AuditPoller-RegisterFolderGuidPath -Config $Config -FolderGuid $fg -FolderPath $fp -WatchRoot $rootPath -StatsRef $statsRef
                    $warmed++
                }
            } catch { }
        }

        if (-not $warmSheets) { continue }
        $pathsToWarm = [System.Collections.Generic.List[string]]::new()
        $suffix = 'CADD\Sheets'
        try {
            if ($cfg -is [hashtable] -and $cfg.ContainsKey('sheetsPathFromProject') -and $cfg.sheetsPathFromProject) {
                $suffix = [string]$cfg.sheetsPathFromProject
            } elseif ($cfg.PSObject -and $cfg.sheetsPathFromProject) { $suffix = [string]$cfg.sheetsPathFromProject }
        } catch { }
        $suffix = $suffix.Trim().TrimStart('\')
        $directSheets = _AuditPoller-NormalizeFolderPath -FolderPath (($rootPath.TrimEnd('\')) + '\' + $suffix)
        if ($directSheets) { [void]$pathsToWarm.Add($directSheets) }

        if ((Get-Command -Name 'Find-PWSheetsFoldersUnderRoot' -ErrorAction SilentlyContinue) -and $pathsToWarm.Count -lt $maxSheets) {
            try {
                $projectDepth = 1
                if ($cfg -is [hashtable] -and $cfg.ContainsKey('projectDepth')) { try { $projectDepth = [int]$cfg.projectDepth } catch { } }
                elseif ($cfg.PSObject -and $cfg.projectDepth) { try { $projectDepth = [int]$cfg.projectDepth } catch { } }
                $ds = ''
                if ($cfg -is [hashtable] -and $cfg.ContainsKey('datasourceName')) { $ds = [string]$cfg.datasourceName }
                elseif ($cfg.PSObject -and $cfg.datasourceName) { $ds = [string]$cfg.datasourceName }
                foreach ($sp in @(Find-PWSheetsFoldersUnderRoot -RootPath $rootPath -SheetsSuffix $suffix -DatasourceName $ds -ProjectDepth $projectDepth)) {
                    if ($pathsToWarm.Count -ge $maxSheets) { break }
                    $norm = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$sp)
                    if ($norm) { [void]$pathsToWarm.Add($norm) }
                }
            } catch { }
        }

        foreach ($sheetPath in $pathsToWarm) {
            if ($warmed -ge $maxSheets) { break }
            $apiPath = $sheetPath
            if (Get-Command -Name 'ConvertTo-PWCmdletFolderPath' -ErrorAction SilentlyContinue) {
                try {
                    $converted = ConvertTo-PWCmdletFolderPath -InternalFolderPath $sheetPath
                    if (-not [string]::IsNullOrWhiteSpace($converted)) { $apiPath = $converted }
                } catch { }
            }
            try {
                $f = Get-PWFolders -FolderPath $apiPath -JustOne -ErrorAction SilentlyContinue
                $fg = _AuditPoller-ExtractFolderGuidFromPwFolder -Folder $f
                $fp = _AuditPoller-ExtractFolderPathFromPwFolder -Folder $f
                if (-not $fp) { $fp = $sheetPath }
                if ($fg -and $fp) {
                    _AuditPoller-RegisterFolderGuidPath -Config $Config -FolderGuid $fg -FolderPath $fp -WatchRoot $rootPath -StatsRef $statsRef
                    $warmed++
                }
            } catch { }
        }
    }

    $script:AuditPoller_WatchFolderCacheWarmed = $true
    if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
        Write-QCJsonLog -Flush -Level 'Information' -Code 'AUDIT_FOLDER_CACHE_WARMED' -Message "Watch folder GUID cache warmed ($warmed folders)." -Data @{
            warmed = $warmed
            cacheSize = $script:AuditPoller_FolderGuidCache.Count
            warmSheets = $warmSheets
            maxSheetsFoldersToWarm = $maxSheets
        }
    }
    return New-QCSuccessResult -Code 'FOLDER_CACHE_WARM_OK' -Message "Warmed $warmed folder GUID mappings." -Data @{ warmed = $warmed; cacheSize = $script:AuditPoller_FolderGuidCache.Count }
}

function _AuditPoller-ResolveDocFoldersBatched {
    param(
        [hashtable]$Config,
        [string[]]$DocGuids,
        [hashtable]$DocToFolder,
        [ref]$StatsRef,
        [hashtable]$InvalidateGuids = @{}
    )

    $ttlSeconds = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'metadataCacheTtlSeconds' -Default 3600
    $negTtlSeconds = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'negativeCacheTtlSeconds' -Default 1800
    if (Get-Command -Name 'Get-QCPwDocumentCacheBatch' -ErrorAction SilentlyContinue) {
        try {
            $cacheRes = Get-QCPwDocumentCacheBatch -Config $Config -DocumentGuids @($DocGuids)
            if ($cacheRes.IsSuccess -and $cacheRes.Data.cache) {
                foreach ($k in $cacheRes.Data.cache.Keys) {
                    $entry = $cacheRes.Data.cache[$k]
                    if ($InvalidateGuids.ContainsKey($k)) { continue }
                    if ($entry.resolveFailed) {
                        $script:AuditPoller_UnresolvedGuids[$k] = $true
                        if ($StatsRef.Value) { $StatsRef.Value.failedGuidCacheHits++ }
                        continue
                    }
                    if ($entry.folderPath) {
                        $script:AuditPoller_DocFolderCache[$k] = $entry.folderPath
                        foreach ($dg in @($DocGuids)) {
                            if ($dg.Trim().ToLowerInvariant() -eq $k) { $DocToFolder[$dg] = $entry.folderPath; break }
                        }
                        if ($StatsRef.Value) { $StatsRef.Value.guidCacheHits++ }
                    }
                }
            }
        } catch { }
    }

    $needPw = [System.Collections.Generic.List[string]]::new()
    foreach ($dg in @($DocGuids)) {
        if ([string]::IsNullOrWhiteSpace($dg)) { continue }
        $key = $dg.Trim().ToLowerInvariant()
        if ($InvalidateGuids.ContainsKey($key)) {
            if ($script:AuditPoller_DocFolderCache.ContainsKey($key)) { $script:AuditPoller_DocFolderCache.Remove($key) | Out-Null }
            if ($script:AuditPoller_UnresolvedGuids.ContainsKey($key)) { $script:AuditPoller_UnresolvedGuids.Remove($key) | Out-Null }
        }
        if ($DocToFolder.ContainsKey($dg)) { continue }
        if ($script:AuditPoller_DocFolderCache.ContainsKey($key)) {
            $DocToFolder[$dg] = $script:AuditPoller_DocFolderCache[$key]
            if ($StatsRef.Value) { $StatsRef.Value.guidCacheHits++ }
            continue
        }
        if ($script:AuditPoller_UnresolvedGuids.ContainsKey($key)) {
            if ($StatsRef.Value) { $StatsRef.Value.guidResolveSkipped++ }
            continue
        }
        [void]$needPw.Add($dg)
        if ($StatsRef.Value) { $StatsRef.Value.guidCacheMisses++ }
    }

    $batchSize = 200
    for ($i = 0; $i -lt $needPw.Count; $i += $batchSize) {
        $chunk = @($needPw[$i..[Math]::Min($i + $batchSize - 1, $needPw.Count - 1)])
        try {
            $docs = @(Get-PWDocumentsByGUIDs -DocumentGUIDs $chunk -ErrorAction SilentlyContinue)
            $found = @{}
            foreach ($doc in $docs) {
                $dg = [string]$doc.DocumentGUID
                if (-not $dg) { continue }
                $fp = $null
                if ($doc.FolderPath) { $fp = [string]$doc.FolderPath }
                elseif ($doc.FullPath) {
                    $full = [string]$doc.FullPath
                    $fp = [System.IO.Path]::GetDirectoryName($full) -replace '/', '\'
                }
                if ($fp) {
                    $canonical = _AuditPoller-NormalizeFolderPath -FolderPath $fp
                    if ($canonical) {
                        $found[$dg] = $canonical
                        $DocToFolder[$dg] = $canonical
                        $script:AuditPoller_DocFolderCache[$dg.Trim().ToLowerInvariant()] = $canonical
                        if ($StatsRef.Value) { $StatsRef.Value.foldersResolved++ }
                        if (Get-Command -Name 'Upsert-QCDocumentActivityFolder' -ErrorAction SilentlyContinue) {
                            $dn = [string]$doc.Name
                            if (-not $dn) { $dn = [string]$doc.DocumentName }
                            Upsert-QCDocumentActivityFolder -Config $Config -DocumentGuid $dg -DocumentName $dn -FolderPath $canonical | Out-Null
                        }
                        if (Get-Command -Name 'Set-QCPwDocumentCacheEntry' -ErrorAction SilentlyContinue) {
                            $desc = ''
                            try { if ($doc.Description) { $desc = [string]$doc.Description } } catch { }
                            $wf = ''
                            try { if ($doc.WorkflowState) { $wf = [string]$doc.WorkflowState } } catch { }
                            Set-QCPwDocumentCacheEntry -Config $Config -DocumentGuid $dg -FolderPath $canonical -Description $desc -WorkflowState $wf -TtlSeconds $ttlSeconds | Out-Null
                        }
                    }
                }
            }
            foreach ($dg in $chunk) {
                if (-not $found.ContainsKey($dg)) {
                    $key = $dg.Trim().ToLowerInvariant()
                    $script:AuditPoller_UnresolvedGuids[$key] = $true
                    if (Get-Command -Name 'Set-QCPwDocumentCacheEntry' -ErrorAction SilentlyContinue) {
                        Set-QCPwDocumentCacheEntry -Config $Config -DocumentGuid $dg -ResolveFailed -TtlSeconds $negTtlSeconds | Out-Null
                    }
                }
            }
        } catch {
            foreach ($dg in $chunk) {
                $script:AuditPoller_UnresolvedGuids[$dg.Trim().ToLowerInvariant()] = $true
                if (Get-Command -Name 'Set-QCPwDocumentCacheEntry' -ErrorAction SilentlyContinue) {
                    Set-QCPwDocumentCacheEntry -Config $Config -DocumentGuid $dg -ResolveFailed -TtlSeconds $negTtlSeconds | Out-Null
                }
            }
        }
    }
}

function _AuditPoller-BuildCandidatesFromTriggerRows {
    param(
        [array]$Rows,
        [hashtable]$DocToFolder,
        [hashtable]$FolderMap,
        [array]$WatchRootConfigs,
        [hashtable]$Config
    )

    $WatchRootConfigs = @(_AuditPoller-NormalizeWatchRootConfigs -WatchRootConfigs $WatchRootConfigs -Config $Config)
    $watchRoots = @($WatchRootConfigs | ForEach-Object { _AuditPoller-GetWatchRootPathFromConfig -Cfg $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $matchRoots = _AuditPoller-BuildMatchRoots -WatchRoots $watchRoots
    $candidates = @()
    $folderUpdates = [System.Collections.Generic.List[object]]::new()
    $diagSkippedAction = 0
    $diagSkippedNoFolder = 0
    $diagSkippedNoWatch = 0

    foreach ($row in @($Rows)) {
        $auditId = $null
        try {
            $idVal = _AuditPoller-GetRowValue -Row $row -Name 'id'
            if ($null -ne $idVal) { $auditId = [long]$idVal }
        } catch { }
        $actionCode = _AuditPoller-GetTriggerActionCode -Row $row
        if (-not $script:QCRelevantActions.ContainsKey($actionCode)) {
            $diagSkippedAction++
            continue
        }
        $actionName = $script:QCRelevantActions[$actionCode]
        $objGuid = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_objguid')
        if ([string]::IsNullOrWhiteSpace($objGuid)) {
            $objGuid = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_objguid')
        }
        $parentGuid = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_parentguid')
        if ([string]::IsNullOrWhiteSpace($parentGuid)) {
            $parentGuid = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_parentguid')
        }

        $resolvedFolder = $null
        $resolvedFromDb = [string](_AuditPoller-GetRowValue -Row $row -Name 'resolved_folder')
        if (-not [string]::IsNullOrWhiteSpace($resolvedFromDb)) {
            $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath $resolvedFromDb
        }
        if (-not $resolvedFolder -and $objGuid -and $DocToFolder.ContainsKey($objGuid)) {
            $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$DocToFolder[$objGuid])
        }
        elseif (-not $resolvedFolder -and $parentGuid) {
            $pgKey = $parentGuid.Trim().ToLowerInvariant()
            if ($FolderMap.ContainsKey($parentGuid)) {
                $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$FolderMap[$parentGuid])
            }
            elseif ($FolderMap.ContainsKey($pgKey)) {
                $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$FolderMap[$pgKey])
            }
            elseif ($script:AuditPoller_FolderGuidCache.ContainsKey($pgKey)) {
                $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$script:AuditPoller_FolderGuidCache[$pgKey])
            }
        }

        $isWatchMatch = $false
        $isSheetsFolder = $false
        if ($resolvedFolder) {
            $isWatchMatch = _AuditPoller-MatchesWatchRoot -FolderPath $resolvedFolder -MatchRoots $matchRoots
            if ($isWatchMatch) {
                $isSheetsFolder = _AuditPoller-GetSheetsSubpath -FolderPath $resolvedFolder -WatchRootConfigs $WatchRootConfigs -MatchRoots $matchRoots
            }
        }

        $candidateType = if ($isWatchMatch) { 'WATCH_MATCH' } else { $null }
        if ($auditId -and $resolvedFolder) {
            [void]$folderUpdates.Add(@{ id = $auditId; resolvedFolder = $resolvedFolder; candidateType = $candidateType })
        }

        if (-not $resolvedFolder) {
            $diagSkippedNoFolder++
            continue
        }
        if (-not $isWatchMatch) {
            $diagSkippedNoWatch++
            continue
        }

        $enableQcPrepend = $false
        $enableQcCommentSync = $false
        $enableStatusSet = $false
        $watchRootPath = $null
        $rootCfg = _AuditPoller-GetWatchRootConfigForFolder -FolderPath $resolvedFolder -WatchRootConfigs $WatchRootConfigs
        if ($rootCfg) {
            if ($rootCfg -is [hashtable]) {
                try { if ($rootCfg.ContainsKey('enableQcPrepend')) { $enableQcPrepend = [bool]$rootCfg['enableQcPrepend'] } } catch { }
                try { if ($rootCfg.ContainsKey('enableQcCommentSync')) { $enableQcCommentSync = [bool]$rootCfg['enableQcCommentSync'] } } catch { }
                try { if ($rootCfg.ContainsKey('enableStatusSet')) { $enableStatusSet = [bool]$rootCfg['enableStatusSet'] } } catch { }
            } else {
                try { if ($rootCfg.enableQcPrepend) { $enableQcPrepend = [bool]$rootCfg.enableQcPrepend } } catch { }
                try { if ($rootCfg.enableQcCommentSync) { $enableQcCommentSync = [bool]$rootCfg.enableQcCommentSync } } catch { }
                try { if ($rootCfg.enableStatusSet) { $enableStatusSet = [bool]$rootCfg.enableStatusSet } } catch { }
            }
            $watchRootPath = _AuditPoller-GetWatchRootPathFromConfig -Cfg $rootCfg
        }

        $actTime = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_acttime')
        if ([string]::IsNullOrWhiteSpace($actTime)) { $actTime = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_acttime') }
        $userno = 0
        try {
            $u = _AuditPoller-GetRowValue -Row $row -Name 'pw_userno'
            if ($null -eq $u) { $u = _AuditPoller-GetRowValue -Row $row -Name 'o_userno' }
            if ($null -ne $u) { $userno = [int]$u }
        } catch { $userno = 0 }
        $username = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_username')
        $itemName = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_itemname')
        if ([string]::IsNullOrWhiteSpace($itemName)) { $itemName = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_itemname') }
        $itemDesc = _AuditPoller-GetRowValue -Row $row -Name 'pw_itemdesc'
        if ($null -eq $itemDesc) { $itemDesc = _AuditPoller-GetRowValue -Row $row -Name 'o_itemdesc' }
        $textParam = _AuditPoller-GetRowValue -Row $row -Name 'pw_textparam'
        if ($null -eq $textParam) { $textParam = _AuditPoller-GetRowValue -Row $row -Name 'o_textparam' }

        $candidates += @{
            auditEventId        = $auditId
            objGuid             = $objGuid
            parentGuid          = $parentGuid
            actionCode          = $actionCode
            actionName          = $actionName
            itemName            = $itemName
            actTime             = $actTime
            userno              = $userno
            username            = $username
            itemdesc            = if ($null -eq $itemDesc) { $null } else { [string]$itemDesc }
            textparam           = if ($null -eq $textParam) { $null } else { [string]$textParam }
            resolvedFolder      = $resolvedFolder
            isSheetsFolder      = $isSheetsFolder
            candidateType       = $candidateType
            enableQcPrepend     = $enableQcPrepend
            enableQcCommentSync = $enableQcCommentSync
            enableStatusSet     = $enableStatusSet
            watchRoot           = $watchRootPath
        }
    }

    return @{
        candidates          = $candidates
        folderUpdates       = @($folderUpdates)
        skippedActionCode   = $diagSkippedAction
        skippedNoFolder     = $diagSkippedNoFolder
        skippedNoWatchMatch = $diagSkippedNoWatch
    }
}

function _AuditPoller-NormalizeFolderPath {
    param([AllowNull()][string]$FolderPath)
    if (Get-Command -Name 'Normalize-QCDocumentsFolderPath' -ErrorAction SilentlyContinue) {
        $r = Normalize-QCDocumentsFolderPath -Path ([string]$FolderPath)
        if ($r.IsSuccess) { return [string]$r.Data.path }
    }
    $t = ($FolderPath -as [string]).Trim().TrimEnd('\').Replace('/', '\')
    $t = $t -replace '\\{2,}', '\'
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    if ($t -match '^(?i)pw:\\') {
        $idx = $t.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) { $t = $t.Substring($idx + 1) }
    }
    $t = $t.ToLowerInvariant()
    if ($t.StartsWith('documents\', [StringComparison]::Ordinal)) { return $t }
    return ('documents\' + $t.TrimStart('\'))
}

function _AuditPoller-GetWatchRootPathFromConfig {
    param([object]$Cfg)
    if ($null -eq $Cfg) { return '' }
    if ($Cfg -is [hashtable]) {
        if ($Cfg.ContainsKey('path') -and $Cfg['path']) { return ([string]$Cfg['path']).Trim() }
        return ''
    }
    try {
        if ($Cfg.PSObject.Properties['path'] -and $Cfg.path) { return ([string]$Cfg.path).Trim() }
    } catch { }
    return ''
}

function _AuditPoller-NormalizeWatchRootConfigs {
    param(
        [array]$WatchRootConfigs,
        [hashtable]$Config
    )
    $out = [System.Collections.Generic.List[object]]::new()
    if ($null -ne $WatchRootConfigs -and $WatchRootConfigs.Count -gt 0) {
        if ($WatchRootConfigs -is [hashtable] -and $WatchRootConfigs.ContainsKey('path')) {
            [void]$out.Add($WatchRootConfigs)
            return @($out)
        }
        foreach ($item in @($WatchRootConfigs)) {
            if ($null -eq $item) { continue }
            if (-not [string]::IsNullOrWhiteSpace((_AuditPoller-GetWatchRootPathFromConfig -Cfg $item))) {
                [void]$out.Add($item)
            }
        }
        if ($out.Count -gt 0) { return @($out) }
    }
    if ($Config.projectWise -and $Config.projectWise.watchList -and $Config.projectWise.watchList.roots) {
        foreach ($item in @($Config.projectWise.watchList.roots)) {
            if ($null -eq $item) { continue }
            if (-not [string]::IsNullOrWhiteSpace((_AuditPoller-GetWatchRootPathFromConfig -Cfg $item))) {
                [void]$out.Add($item)
            }
        }
    }
    return @($out)
}

function _AuditPoller-BuildMatchRoots {
    param([string[]]$WatchRoots)
    $matchRoots = [System.Collections.Generic.List[string]]::new()
    foreach ($root in @($WatchRoots)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $matchRoots.Add($root)
        if ($root.StartsWith('Documents\', [StringComparison]::OrdinalIgnoreCase)) {
            $matchRoots.Add($root.Substring('Documents\'.Length))
        } else {
            $matchRoots.Add("Documents\$root")
        }
    }
    return $matchRoots
}

function _AuditPoller-MatchesWatchRoot {
    param([string]$FolderPath, [System.Collections.Generic.List[string]]$MatchRoots)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $false }
    foreach ($root in $MatchRoots) {
        $r = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$root)
        if ($r -and $fp.StartsWith($r, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function _AuditPoller-GetWatchRootConfigForFolder {
    param([string]$FolderPath, [array]$WatchRootConfigs)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $null }
    foreach ($cfg in @($WatchRootConfigs)) {
        if (-not $cfg) { continue }
        $rootPath = _AuditPoller-GetWatchRootPathFromConfig -Cfg $cfg
        if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }
        $testRoots = @($rootPath)
        if ($rootPath.StartsWith('Documents\', [StringComparison]::OrdinalIgnoreCase)) {
            $testRoots += $rootPath.Substring('Documents\'.Length)
        } else {
            $testRoots += "Documents\$rootPath"
        }
        foreach ($tr in $testRoots) {
            $nr = _AuditPoller-NormalizeFolderPath -FolderPath $tr
            if ($nr -and $fp.StartsWith($nr, [StringComparison]::OrdinalIgnoreCase)) { return $cfg }
        }
    }
    return $null
}

function _AuditPoller-GetSheetsSubpath {
    param([string]$FolderPath, [array]$WatchRootConfigs, [System.Collections.Generic.List[string]]$MatchRoots)
    $fp = _AuditPoller-NormalizeFolderPath -FolderPath $FolderPath
    if (-not $fp) { return $false }
    foreach ($cfg in @($WatchRootConfigs)) {
        if (-not $cfg) { continue }
        $rootPath = _AuditPoller-GetWatchRootPathFromConfig -Cfg $cfg
        if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }
        $suffix = 'CADD\Sheets'
        if ($cfg -is [hashtable] -and $cfg.ContainsKey('sheetsPathFromProject') -and $cfg['sheetsPathFromProject']) {
            $suffix = [string]$cfg['sheetsPathFromProject']
        } elseif ($cfg.PSObject -and $cfg.PSObject.Properties['sheetsPathFromProject'] -and $cfg.sheetsPathFromProject) {
            $suffix = [string]$cfg.sheetsPathFromProject
        }
        $testRoots = @($rootPath)
        if ($rootPath.StartsWith('Documents\', [StringComparison]::OrdinalIgnoreCase)) {
            $testRoots += $rootPath.Substring('Documents\'.Length)
        } else {
            $testRoots += "Documents\$rootPath"
        }
        foreach ($tr in $testRoots) {
            $nr = _AuditPoller-NormalizeFolderPath -FolderPath $tr
            if ($nr -and $fp.StartsWith($nr, [StringComparison]::OrdinalIgnoreCase) -and
                $fp.IndexOf($suffix, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $true
            }
        }
    }
    return $false
}

function _AuditPoller-GetDisplayTimeZone {
    param([hashtable]$Config = @{})
    if ($Config -and $Config.Count -gt 0) {
        $id = Get-QCDisplayTimeZoneIdFromConfig -Config $Config
        return [TimeZoneInfo]::FindSystemTimeZoneById($id)
    }
    return Get-QCDisplayTimeZone
}

function _AuditPoller-AssumeUtcFromPw {
    param([DateTime]$DateTime)
    if ($DateTime.Kind -eq [DateTimeKind]::Utc) { return $DateTime }
    if ($DateTime.Kind -eq [DateTimeKind]::Local) { return $DateTime.ToUniversalTime() }
    # PW Select-PWSQL returns o_acttime as Unspecified with UTC wall-clock components.
    return [DateTime]::SpecifyKind($DateTime, [DateTimeKind]::Utc)
}

function _AuditPoller-ParseWatermarkToUtc {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    try {
        $s = $Raw.Trim().TrimEnd('Z')
        return [DateTime]::ParseExact(
            $s,
            'yyyy-MM-dd HH:mm:ss',
            $null,
            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        )
    } catch {
        try {
            $s = $Raw.Trim()
            if ($s.EndsWith('Z', [StringComparison]::OrdinalIgnoreCase)) {
                return [DateTime]::Parse(
                    $s,
                    $null,
                    [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
                )
            }
            return [DateTime]::Parse(
                $s,
                $null,
                [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
            )
        } catch { return $null }
    }
}

function _AuditPoller-FormatSqlUtc {
    param([DateTime]$Utc)
    $u = $Utc.ToUniversalTime()
    return $u.ToString('yyyy-MM-dd HH:mm:ss')
}

function _AuditPoller-FormatDisplayTime {
    param(
        [DateTime]$Utc,
        [hashtable]$Config = @{}
    )
    $tz = _AuditPoller-GetDisplayTimeZone -Config $Config
    $mt = [TimeZoneInfo]::ConvertTimeFromUtc($Utc.ToUniversalTime(), $tz)
    return $mt.ToString('yyyy-MM-dd HH:mm:ss')
}

function _AuditPoller-ParseActTime {
    param([string]$ActTime)
    if ([string]::IsNullOrWhiteSpace($ActTime)) { return $null }
    try { return [DateTime]::Parse($ActTime) } catch { return $null }
}

function _AuditPoller-TryAdvanceWatermarkAfter {
    param(
        [AllowNull()][string]$Current,
        [AllowNull()][string]$Candidate
    )
    if ([string]::IsNullOrWhiteSpace($Candidate)) {
        if ([string]::IsNullOrWhiteSpace($Current)) { return $null }
        return $Current
    }
    if ([string]::IsNullOrWhiteSpace($Current)) { return $Candidate }
    $curDt = _AuditPoller-ParseActTime -ActTime $Current
    $candDt = _AuditPoller-ParseActTime -ActTime $Candidate
    if ($candDt -and (-not $curDt -or $candDt -gt $curDt)) { return $Candidate }
    return $Current
}

function _AuditPoller-GetRowValue {
    param($Row, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Row) { return $null }
    try {
        if ($Row -is [System.Data.DataRow]) {
            foreach ($col in $Row.Table.Columns) {
                if ([string]::Equals([string]$col.ColumnName, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                    $v = $Row[$col]
                    if ($v -is [DBNull]) { return $null }
                    return $v
                }
            }
        }
        if ($Row -is [hashtable] -or $Row -is [System.Collections.IDictionary]) {
            foreach ($key in @($Row.Keys)) {
                if ([string]::Equals([string]$key, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                    $v = $Row[$key]
                    if ($v -is [DBNull]) { return $null }
                    return $v
                }
            }
        }
        foreach ($prop in $Row.PSObject.Properties) {
            if ([string]::Equals([string]$prop.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                $v = $prop.Value
                if ($v -is [DBNull]) { return $null }
                return $v
            }
        }
    } catch { }
    return $null
}

function _AuditPoller-FormatActTime {
    param(
        [AllowNull()][object]$Value,
        [hashtable]$Config = @{}
    )
    if ($null -eq $Value) { return $null }
    if ($Value -is [DBNull]) { return $null }
    if ($Value -is [DateTime]) {
        $utc = _AuditPoller-AssumeUtcFromPw -DateTime $Value
        return _AuditPoller-FormatDisplayTime -Utc $utc -Config $Config
    }
    $s = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try {
        $parsed = [DateTime]::Parse($s)
        $utc = _AuditPoller-AssumeUtcFromPw -DateTime $parsed
        return _AuditPoller-FormatDisplayTime -Utc $utc -Config $Config
    } catch { return $s }
}

function _AuditPoller-GetSqlResultRows {
    param($Result)
    if ($null -eq $Result) { return @() }
    try {
        if ($Result.PSObject.Properties.Name -contains 'Rows' -and $null -ne $Result.Rows) {
            return @($Result.Rows)
        }
    } catch { }
    if ($Result -is [System.Data.DataTable]) { return @($Result.Rows) }
    return @($Result)
}

function _AuditPoller-GetActionCode {
    param($Row)
    $v = _AuditPoller-GetRowValue -Row $Row -Name 'o_action'
    if ($null -eq $v) { return 0 }
    try { return [int]$v } catch { return 0 }
}

function _AuditPoller-TestAuditIngestAllowedActionCode {
    param([int]$ActionCode)
    return $script:QCRelevantActions.ContainsKey($ActionCode)
}

function _AuditPoller-TryAdvanceWatermarkFromAuditRow {
    param(
        $Row,
        [hashtable]$Config,
        [ref]$MaxPwActTime,
        [ref]$MaxPwActTimeUtc
    )
    $rawAct = _AuditPoller-GetRowValue -Row $Row -Name 'o_acttime'
    $actTime = _AuditPoller-FormatActTime -Value $rawAct -Config $Config
    if ($actTime) {
        $MaxPwActTime.Value = _AuditPoller-TryAdvanceWatermarkAfter -Current $MaxPwActTime.Value -Candidate $actTime
    }
    if ($null -ne $rawAct -and -not ($rawAct -is [DBNull])) {
        $utcStr = if ($rawAct -is [DateTime]) {
            _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $rawAct)
        } else {
            $parsed = _AuditPoller-ParseActTime -ActTime ([string]$rawAct)
            if ($parsed) { _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $parsed) } else { $null }
        }
        if ($utcStr) {
            $MaxPwActTimeUtc.Value = _AuditPoller-TryAdvanceWatermarkAfter -Current $MaxPwActTimeUtc.Value -Candidate $utcStr
        }
    }
}

function _AuditPoller-GetTriggerActionCode {
    param($Row)
    $code = 0
    $pwAct = _AuditPoller-GetRowValue -Row $Row -Name 'pw_action'
    if ($null -ne $pwAct) {
        try { $code = [int]$pwAct } catch { $code = 0 }
    }
    if ($code -eq 0) { $code = _AuditPoller-GetActionCode -Row $Row }
    return $code
}

function _AuditPoller-GetActionName {
    param([int]$ActionCode)
    if ($script:AuditActionNames.ContainsKey($ActionCode)) { return $script:AuditActionNames[$ActionCode] }
    return "UNKNOWN_$ActionCode"
}

function Get-PWAuditTrailActionName {
    <#
    .SYNOPSIS
    Resolves a ProjectWise dms_audt o_action / pw_action integer to its audit trail constant name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$ActionCode
    )
    return _AuditPoller-GetActionName -ActionCode $ActionCode
}

function _AuditPoller-NormalizeGuid {
    param([AllowNull()]$Value)
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.Trim()
}

function _AuditPoller-NewAuditEventDbRow {
    param(
        $Evt,
        [hashtable]$Config = @{},
        [AllowNull()][string]$ResolvedFolder = $null,
        [AllowNull()][string]$CandidateType = $null
    )
    $actionCode = _AuditPoller-GetActionCode -Row $Evt
    $objno = 0;  try { $objno = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_objno') } catch { $objno = 0 }
    $userno = 0; try { $userno = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_userno') } catch { $userno = 0 }
    $itemdesc = _AuditPoller-GetRowValue -Row $Evt -Name 'o_itemdesc'
    $textparam = _AuditPoller-GetRowValue -Row $Evt -Name 'o_textparam'
    $objtype = 0; try { $objtype = [int](_AuditPoller-GetRowValue -Row $Evt -Name 'o_objtype') } catch { $objtype = 0 }
    return @{
        acttime       = (_AuditPoller-FormatActTime -Value (_AuditPoller-GetRowValue -Row $Evt -Name 'o_acttime') -Config $Config)
        action        = $actionCode
        actionName    = (_AuditPoller-GetActionName -ActionCode $actionCode)
        objtype       = $objtype
        objno         = $objno
        objguid       = (_AuditPoller-NormalizeGuid (_AuditPoller-GetRowValue -Row $Evt -Name 'o_objguid'))
        parentguid    = (_AuditPoller-NormalizeGuid (_AuditPoller-GetRowValue -Row $Evt -Name 'o_parentguid'))
        userno        = $userno
        itemname      = [string](_AuditPoller-GetRowValue -Row $Evt -Name 'o_itemname')
        itemdesc      = if ($null -eq $itemdesc) { $null } else { [string]$itemdesc }
        textparam     = if ($null -eq $textparam) { $null } else { [string]$textparam }
        folder        = $ResolvedFolder
        candidateType = $CandidateType
    }
}

function Get-AuditTrailHighWaterMarkFromDatabase {
    <#
    .SYNOPSIS
    Reads the most recent successful poll_runs watermark_after from the database.
    Returns a DateTime or $null if no prior run exists.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)

    if (-not (Test-QCDatabaseEnabled -Config $Config)) { return $null }
    try {
        $res = Invoke-QCDatabaseScalar -Config $Config -Sql "SELECT TOP 1 watermark_after FROM poll_runs WHERE error_message IS NULL AND watermark_after IS NOT NULL ORDER BY started_at DESC"
        if ($res.IsSuccess -and $res.Data.value) {
            return _AuditPoller-ParseWatermarkToUtc -Raw ([string]$res.Data.value)
        }
    } catch { }
    return $null
}

function Get-AuditTrailCaptureWatermark {
    <#
    .SYNOPSIS
    Returns the latest successful audit watermark (DB watcher_state is primary when enabled).

    .DESCRIPTION
    Order: watcher_state.audit_watermark_utc, then local watermark file (both UTC with Z).
    poll_runs.watermark_after is telemetry only and is excluded (often lacks Z and skews MAX).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = ''
    )

    $found = @()
    if (Get-Command -Name 'Get-QCAuditWatermarkUtc' -ErrorAction SilentlyContinue) {
        try {
            $dbWm = Get-QCAuditWatermarkUtc -Config $Config
            if ($dbWm) { $found += $dbWm }
        } catch { }
    }
    if ($WatermarkPath -and (Test-Path -LiteralPath $WatermarkPath)) {
        try {
            $raw = (Get-Content -LiteralPath $WatermarkPath -Raw -ErrorAction Stop).Trim()
            $parsed = _AuditPoller-ParseWatermarkToUtc -Raw $raw
            if ($parsed) { $found += $parsed }
        } catch { }
    }
    if ($found.Count -eq 0) { return $null }
    return ($found | Sort-Object -Descending | Select-Object -First 1)
}

function _AuditPoller-EscapeSqlLiteral {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace("'", "''")
}

function _AuditPoller-GetAuditPollerInt {
    param([hashtable]$Config, [string]$Key, [int]$Default)
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable] -and $ap.ContainsKey($Key) -and $null -ne $ap[$Key]) { return [int]$ap[$Key] }
            if ($ap.PSObject -and $null -ne $ap.$Key) { return [int]$ap.$Key }
        }
    } catch { }
    return $Default
}

function Get-AuditTrailHighWaterMark {
    <#
    .SYNOPSIS
    Back-compat alias for Get-AuditTrailCaptureWatermark (database only when no path given).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = ''
    )
    return Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $WatermarkPath
}

function Set-AuditTrailCaptureWatermark {
    <#
    .SYNOPSIS
    Persists the audit capture high-water mark to watcher_state (when DB enabled) and the local file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WatermarkPath,
        [Parameter(Mandatory)][DateTime]$CapturedThrough,
        [hashtable]$Config = @{}
    )

    $utc = $CapturedThrough.ToUniversalTime()
    $ok = $true
    if ($Config.Count -gt 0 -and (Get-Command -Name 'Set-QCAuditWatermarkUtc' -ErrorAction SilentlyContinue)) {
        try { $ok = (Set-QCAuditWatermarkUtc -Config $Config -WatermarkUtc $utc) } catch { $ok = $false }
    }
    try {
        if ($WatermarkPath) {
            $dir = Split-Path -Parent $WatermarkPath
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $value = (_AuditPoller-FormatSqlUtc -Utc $utc) + 'Z'
            Set-Content -LiteralPath $WatermarkPath -Value $value -Encoding UTF8 -NoNewline
        }
        return $ok
    } catch {
        return $false
    }
}

function Get-AuditTrailPollWindow {
    <#
    .SYNOPSIS
    Computes the audit poll interval: (last successful watermark, now], or initialLookback on first run.

    .DESCRIPTION
    Steady-state queries use the last capture watermark only (optional small overlapSeconds).
    lookbackSeconds applies only when no watermark file exists yet (bootstrap).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$WatermarkPath = '',
        [int]$LookbackSeconds = 120,
        [switch]$UseRestartOverlap
    )

    # dms_audt o_acttime compares as UTC wall clock; do not use machine LocalDateTime for SQL bounds.
    $until = [DateTime]::UtcNow
    $lastCapture = Get-AuditTrailCaptureWatermark -Config $Config -WatermarkPath $WatermarkPath
    $initialLookbackSeconds = $LookbackSeconds
    $overlapSeconds = 0
    $restartOverlapSeconds = 300
    try {
        if ($Config.ContainsKey('auditPoller') -and $Config.auditPoller) {
            $ap = $Config.auditPoller
            if ($ap -is [hashtable]) {
                if ($ap.ContainsKey('initialLookbackSeconds') -and $null -ne $ap.initialLookbackSeconds) {
                    $initialLookbackSeconds = [int]$ap.initialLookbackSeconds
                }
                if ($ap.ContainsKey('overlapSeconds') -and $null -ne $ap.overlapSeconds) {
                    $overlapSeconds = [int]$ap.overlapSeconds
                }
                if ($ap.ContainsKey('restartOverlapSeconds') -and $null -ne $ap.restartOverlapSeconds) {
                    $restartOverlapSeconds = [int]$ap.restartOverlapSeconds
                }
            } elseif ($ap.PSObject) {
                if ($null -ne $ap.initialLookbackSeconds) { $initialLookbackSeconds = [int]$ap.initialLookbackSeconds }
                if ($null -ne $ap.overlapSeconds) { $overlapSeconds = [int]$ap.overlapSeconds }
                if ($null -ne $ap.restartOverlapSeconds) { $restartOverlapSeconds = [int]$ap.restartOverlapSeconds }
            }
        }
    } catch { }
    if ($initialLookbackSeconds -lt 1) { $initialLookbackSeconds = $LookbackSeconds }
    if ($overlapSeconds -lt 0) { $overlapSeconds = 0 }
    if ($restartOverlapSeconds -lt 0) { $restartOverlapSeconds = 0 }

    $effectiveOverlapSeconds = $overlapSeconds
    if ($lastCapture -and -not $UseRestartOverlap.IsPresent -and $effectiveOverlapSeconds -lt 1) {
        # dms_audt timestamps are second-granular in practice. Always re-read the
        # watermark second so rows sharing the high-water timestamp cannot be skipped;
        # audit_events natural-key dedupe absorbs the overlap.
        $effectiveOverlapSeconds = 1
    }
    $since = if ($lastCapture) {
        if ($UseRestartOverlap.IsPresent -and $restartOverlapSeconds -gt 0) {
            $lastCapture.AddSeconds(-$restartOverlapSeconds)
        } elseif ($effectiveOverlapSeconds -gt 0) {
            $lastCapture.AddSeconds(-$effectiveOverlapSeconds)
        } else {
            $lastCapture
        }
    } else {
        $until.AddSeconds(-$initialLookbackSeconds)
    }
    if ($since -ge $until) {
        $since = if ($overlapSeconds -gt 0) { $until.AddSeconds(-$overlapSeconds) } else { $until.AddSeconds(-1) }
    }
    $tz = _AuditPoller-GetDisplayTimeZone -Config $Config
    $watermarkBefore = if ($lastCapture) { _AuditPoller-FormatSqlUtc -Utc $lastCapture } else { $null }

    return @{
        since           = $since
        until           = $until
        sinceUtc        = _AuditPoller-FormatSqlUtc -Utc $since
        untilUtc        = _AuditPoller-FormatSqlUtc -Utc $until
        sinceDisplay    = _AuditPoller-FormatDisplayTime -Utc $since -Config $Config
        untilDisplay    = _AuditPoller-FormatDisplayTime -Utc $until -Config $Config
        watermarkBefore = $watermarkBefore
        isFirstCapture  = (-not $lastCapture)
        lookbackSecondsUsed = if ($lastCapture) {
            if ($UseRestartOverlap.IsPresent) { $restartOverlapSeconds } else { $effectiveOverlapSeconds }
        } else { $initialLookbackSeconds }
        overlapSecondsUsed  = if ($UseRestartOverlap.IsPresent) { $restartOverlapSeconds } else { $effectiveOverlapSeconds }
        restartOverlapUsed  = [bool]$UseRestartOverlap.IsPresent
        displayTimeZoneId = $tz.Id
    }
}

function Invoke-AuditTrailScan {
    <#
    .SYNOPSIS
    Queries the PW audit trail for recent QC-relevant events, resolves folders,
    matches against watch roots, and returns structured candidate data.

    .DESCRIPTION
    Production-ready extraction of the POC audit scan logic.
    Returns a QCResult with:
      - Data.events: all QC-relevant audit events
      - Data.candidates: events in watched Sheets folders (ready for job creation)
      - Data.docToFolder: document GUID -> folder path map
      - Data.stats: counts for telemetry
      - Data.watermarkAfter: latest event timestamp (for high-water mark)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][DateTime]$Since,
        [Parameter()][DateTime]$Until = [DateTime]::Now,
        [array]$WatchRootConfigs = @()
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stats = @{
        totalEvents        = 0
        relevantEvents     = 0
        foldersResolved    = 0
        watchMatches       = 0
        sheetsMatches      = 0
        dbWrites           = 0
        dbSkipped          = 0
        dbRowsPrepared     = 0
        dbRowsNullGuid     = 0
        pagesFetched       = 0
        eventsTruncated    = $false
        dbLastError        = $null
        dbWriteFailed      = 0
        dbUnprocessedLoaded = 0
        guidCacheHits      = 0
        guidCacheMisses    = 0
        guidResolveSkipped = 0
        failedGuidCacheHits = 0
        parentFoldersResolved = 0
        parentFolderGuidCacheHits = 0
        parentFolderGuidCacheMisses = 0
        parentFolderGuidResolveSkipped = 0
        failedFolderGuidCacheHits = 0
        triggerSource      = 'pw_batch'
        auditLogicVersion  = $script:AuditPollerLogicVersion
        ingestSkipped      = 0
        totalFetchedRaw    = 0
    }

    $maxPwActTime = $null
    $maxPwActTimeUtc = $null

    # 1. Query dms_audt — paginated ASC so busy servers are not stuck on the oldest TOP 500 only.
    # SQL bounds in UTC to match PW o_acttime (Unspecified DateTime with UTC wall-clock components).
    $sinceUtc = if ($Since.Kind -eq [DateTimeKind]::Utc) { $Since.ToUniversalTime() } else { _AuditPoller-AssumeUtcFromPw -DateTime $Since }
    $untilUtc = if ($Until.Kind -eq [DateTimeKind]::Utc) { $Until.ToUniversalTime() } else { _AuditPoller-AssumeUtcFromPw -DateTime $Until }
    $sinceStr = _AuditPoller-FormatSqlUtc -Utc $sinceUtc
    $queryUntilStr = _AuditPoller-FormatSqlUtc -Utc $untilUtc
    $pageSize = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'pageSize' -Default 500
    $maxPages = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxPagesPerPoll' -Default 100
    if ($pageSize -lt 1) { $pageSize = 500 }
    if ($maxPages -lt 1) { $maxPages = 100 }

    $allEvents = [System.Collections.Generic.List[object]]::new()
    $cursorSince = $sinceStr
    $cursorGuid = ''
    $cursorAction = $null
    $cursorObjNo = $null
    $pageNum = 0
    try {
        while ($pageNum -lt $maxPages) {
            $pageNum++
            $untilStr = $queryUntilStr
            $lowerBoundSql = if ($null -ne $cursorAction -and $null -ne $cursorObjNo) {
                $t = _AuditPoller-EscapeSqlLiteral -Value $cursorSince
                $g = _AuditPoller-EscapeSqlLiteral -Value $cursorGuid
                $a = [int]$cursorAction
                $n = [int]$cursorObjNo
                "(o_acttime > '$t' OR (o_acttime = '$t' AND (ISNULL(o_objguid,'') > '$g' OR (ISNULL(o_objguid,'') = '$g' AND (o_action > $a OR (o_action = $a AND o_objno > $n))))))"
            } else {
                "o_acttime > '$(_AuditPoller-EscapeSqlLiteral -Value $cursorSince)'"
            }
            $sql = "SELECT TOP $pageSize o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE $lowerBoundSql AND o_acttime <= '$(_AuditPoller-EscapeSqlLiteral -Value $untilStr)' ORDER BY o_acttime ASC, ISNULL(o_objguid,'') ASC, o_action ASC, o_objno ASC"
            $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
            $batch = @(_AuditPoller-GetSqlResultRows -Result $result)
            if ($batch.Count -eq 0) { break }
            foreach ($row in $batch) {
                $stats.totalFetchedRaw++
                _AuditPoller-TryAdvanceWatermarkFromAuditRow -Row $row -Config $Config -MaxPwActTime ([ref]$maxPwActTime) -MaxPwActTimeUtc ([ref]$maxPwActTimeUtc)
                $actionCode = _AuditPoller-GetActionCode -Row $row
                if (-not (_AuditPoller-TestAuditIngestAllowedActionCode -ActionCode $actionCode)) {
                    $stats.ingestSkipped++
                    continue
                }
                [void]$allEvents.Add($row)
            }
            $last = $batch[$batch.Count - 1]
            $lastAct = _AuditPoller-GetRowValue -Row $last -Name 'o_acttime'
            $cursorSince = if ($lastAct -is [DateTime]) {
                _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $lastAct)
            } else {
                _AuditPoller-FormatActTime -Value $lastAct -Config $Config
            }
            if ([string]::IsNullOrWhiteSpace($cursorSince)) { break }
            $cursorGuid = [string](_AuditPoller-GetRowValue -Row $last -Name 'o_objguid')
            if ($null -eq $cursorGuid) { $cursorGuid = '' }
            try { $cursorAction = [int](_AuditPoller-GetRowValue -Row $last -Name 'o_action') } catch { $cursorAction = 0 }
            try { $cursorObjNo = [int](_AuditPoller-GetRowValue -Row $last -Name 'o_objno') } catch { $cursorObjNo = 0 }
            if ($batch.Count -lt $pageSize) { break }
        }
        if ($pageNum -ge $maxPages -and [int]$stats.totalFetchedRaw -gt 0) {
            $stats.eventsTruncated = $true
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_QUERY_TRUNCATED' -Message "dms_audt page cap reached ($maxPages x $pageSize); watermark will not advance so the capped window is retried." -Data @{
                    since = $sinceStr; until = $untilStr; pagesFetched = $pageNum; eventsFetched = $allEvents.Count
                }
            }
        }
    } catch {
        $sw.Stop()
        return New-QCFailureResult -Code 'AUDIT_QUERY_FAILED' -Message "dms_audt query failed: $($_.Exception.Message)" -Data @{ durationMs = [int]$sw.ElapsedMilliseconds }
    }
    $stats.pagesFetched = $pageNum
    $stats.totalEvents = $allEvents.Count

    if ($allEvents.Count -eq 0) {
        $sw.Stop()
        $watermarkAfterEmpty = $null
        if ($maxPwActTimeUtc -and -not [bool]$stats.eventsTruncated) {
            $watermarkAfterEmpty = $maxPwActTimeUtc
            $stats.maxPwActTime = $maxPwActTime
            $stats.maxPwActTimeUtc = $maxPwActTimeUtc
        }
        if ($stats.ingestSkipped -gt 0 -and $watermarkAfterEmpty) {
            return New-QCSuccessResult -Code 'AUDIT_SCAN_OK' -Message "No ingestable audit events ($($stats.ingestSkipped) non-QC rows skipped)." -Data @{
                events = @(); candidates = @(); docToFolder = @{}; stats = $stats
                watermarkAfter = $watermarkAfterEmpty; durationMs = [int]$sw.ElapsedMilliseconds
                pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
            }
        }
        return New-QCSuccessResult -Code 'AUDIT_NO_EVENTS' -Message 'No audit events in window.' -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $watermarkAfterEmpty; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    # 2. Ingest QC-relevant rows into audit_events (QCRelevantActions allowlist only).
    $watermarkAfter = $null
    $dbRows = @()
    if (Test-QCDatabaseEnabled -Config $Config) {
        foreach ($evt in $allEvents) {
            $dbRows += (_AuditPoller-NewAuditEventDbRow -Evt $evt -Config $Config)
        }
    } else {
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_INGEST_DISABLED' -Message 'database.enabled is false; audit_events not written.' -Data @{ eventsFetched = $allEvents.Count }
        }
    }

    $stats.dbRowsPrepared = $dbRows.Count
    $stats.dbRowsNullGuid = @($dbRows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.objguid) }).Count

    $userNumbersToSync = [System.Collections.Generic.HashSet[int]]::new()
    if ($dbRows.Count -gt 0) {
        try {
            $dbRes = Write-QCAuditEventRows -Config $Config -Rows $dbRows
            if ($dbRes.IsSuccess -and $dbRes.Data) {
                $stats.dbWrites += [int]$dbRes.Data.written
                $stats.dbSkipped += [int]$dbRes.Data.skipped
                if ($null -ne $dbRes.Data.failed) { $stats.dbWriteFailed += [int]$dbRes.Data.failed }
                if ($dbRes.Data.lastError) { $stats.dbLastError = [string]$dbRes.Data.lastError }
            } else {
                $stats.dbSkipped += $dbRows.Count
                $stats.dbWriteFailed += $dbRows.Count
                $stats.dbLastError = [string]$dbRes.Message
                if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                    Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_FAILED' -Message "Write-QCAuditEventRows failed: $($dbRes.Message)" -Data @{
                        code = [string]$dbRes.Code
                        rowCount = $dbRows.Count
                        nullGuid = $stats.dbRowsNullGuid
                    }
                }
            }
        } catch {
            $stats.dbSkipped += $dbRows.Count
            $stats.dbWriteFailed += $dbRows.Count
            $stats.dbLastError = [string]$_.Exception.Message
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_EVENTS_WRITE_EXCEPTION' -Message $_.Exception.Message -Data @{ rowCount = $dbRows.Count; nullGuid = $stats.dbRowsNullGuid }
            }
        }
        if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
            $ingestSkippedNoGuid = 0
            try {
                if ($dbRes -and $dbRes.Data -and $null -ne $dbRes.Data.skippedNoGuid) { $ingestSkippedNoGuid = [int]$dbRes.Data.skippedNoGuid }
            } catch { }
            $ingestLevel = if ($stats.dbLastError) { 'Warning' } else { 'Information' }
            Write-QCJsonLog -Flush -Level $ingestLevel -Code 'AUDIT_EVENTS_INGEST' -Message "audit_events ingest: $($stats.dbWrites) written, $($stats.dbSkipped) skipped/duplicate." -Data @{
                eventsFetched = $stats.totalEvents
                totalFetchedRaw = [int]$stats.totalFetchedRaw
                ingestSkipped = [int]$stats.ingestSkipped
                rowsPrepared  = $stats.dbRowsPrepared
                rowsNullGuid  = $stats.dbRowsNullGuid
                written       = $stats.dbWrites
                skipped       = $stats.dbSkipped
                failed        = [int]$stats.dbWriteFailed
                skippedNoGuid = $ingestSkippedNoGuid
                lastError     = $stats.dbLastError
            }
        }
        foreach ($row in $dbRows) {
            $u = 0
            try { $u = [int]$row.userno } catch { $u = 0 }
            if ($u -gt 0) { [void]$userNumbersToSync.Add($u) }
        }
    }

    if ([int]$stats.dbWriteFailed -gt 0) {
        $sw.Stop()
        return New-QCFailureResult -Code 'AUDIT_EVENTS_WRITE_INCOMPLETE' -Message "audit_events ingest failed for $($stats.dbWriteFailed) rows; watermark not advanced." -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $null; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    if ($maxPwActTime) { $stats.maxPwActTime = $maxPwActTime }
    if ($maxPwActTimeUtc -and -not [bool]$stats.eventsTruncated) {
        $stats.maxPwActTimeUtc = $maxPwActTimeUtc
        $watermarkAfter = $maxPwActTimeUtc
    }

    # 3. QC trigger pipeline — database is source of truth for unprocessed rows.
    $triggerRows = @()
    $useDbTriggers = $false
    if (Test-QCDatabaseEnabled -Config $Config) {
        $maxUnprocessed = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxUnprocessedPerPoll' -Default 500
        if ($maxUnprocessed -lt 1) { $maxUnprocessed = 500 }
        $maxUnprocessedBatches = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxUnprocessedBatchesPerPoll' -Default 1
        if ($maxUnprocessedBatches -lt 1) { $maxUnprocessedBatches = 1 }
        if (Get-Command -Name 'Get-QCUnprocessedAuditEvents' -ErrorAction SilentlyContinue) {
            $dbRows = [System.Collections.Generic.List[object]]::new()
            $loadedIds = [System.Collections.Generic.List[long]]::new()
            for ($batchNo = 1; $batchNo -le $maxUnprocessedBatches; $batchNo++) {
                $unprocRes = Get-QCUnprocessedAuditEvents -Config $Config -MaxRows $maxUnprocessed -ExcludeEventIds @($loadedIds)
                if (-not ($unprocRes.IsSuccess -and $null -ne $unprocRes.Data)) { break }
                $batchRows = @()
                if ($unprocRes.Data.rows) { $batchRows = @($unprocRes.Data.rows) }
                if ($batchRows.Count -eq 0) { break }
                foreach ($r in $batchRows) {
                    [void]$dbRows.Add($r)
                    try {
                        if ($null -ne $r.id) { [void]$loadedIds.Add([long]$r.id) }
                    } catch { }
                }
                if ($batchRows.Count -lt $maxUnprocessed) { break }
            }
            if ($dbRows.Count -gt 0) {
                $triggerRows = @($dbRows)
                $stats.dbUnprocessedLoaded = $triggerRows.Count
                $stats.dbUnprocessedBatchesLoaded = [int][Math]::Ceiling($triggerRows.Count / [double]$maxUnprocessed)
                $stats.triggerSource = 'audit_events_db'
                $useDbTriggers = $true
            }
        }
    }
    if (-not $useDbTriggers) {
        $triggerRows = @($allEvents | Where-Object { $script:QCRelevantActions.ContainsKey((_AuditPoller-GetActionCode -Row $_)) })
        $stats.triggerSource = 'pw_batch'
    }
    $stats.relevantEvents = $triggerRows.Count

    if ($triggerRows.Count -eq 0) {
        if ($userNumbersToSync.Count -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
            try { Sync-PWUserDirectory -Config $Config -UserNumbers @($userNumbersToSync) -MaxUsers 25 | Out-Null } catch { }
        }
        $sw.Stop()
        return New-QCSuccessResult -Code 'AUDIT_SCAN_OK' -Message "Audit ingest complete: $($stats.totalEvents) fetched, 0 unprocessed QC-relevant events." -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $watermarkAfter; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    if (-not $script:AuditPoller_WatchFolderCacheWarmed) {
        try { Sync-AuditPollerWatchFolderGuidCache -Config $Config -WatchRootConfigs $WatchRootConfigs | Out-Null } catch { }
    }

    [void](_AuditPoller-LoadDocFolderCache -Config $Config)
    [void](_AuditPoller-LoadFolderGuidCache -Config $Config)
    $docToFolder = @{}
    foreach ($row in $triggerRows) {
        $og = [string]$row.pw_objguid
        if ([string]::IsNullOrWhiteSpace($og)) { $og = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_objguid') }
        if (-not $og) { continue }
        $key = $og.Trim().ToLowerInvariant()
        if ($script:AuditPoller_DocFolderCache.ContainsKey($key)) {
            $docToFolder[$og] = $script:AuditPoller_DocFolderCache[$key]
        }
    }

    $docGuids = @($triggerRows | ForEach-Object {
        $g = [string]$_.pw_objguid
        if ([string]::IsNullOrWhiteSpace($g)) { $g = [string](_AuditPoller-GetRowValue -Row $_ -Name 'o_objguid') }
        $g
    } | Where-Object { $_ -and $_ -ne '' } | Select-Object -Unique)

    $invalidate = @{}
    foreach ($row in $triggerRows) {
        $ac = _AuditPoller-GetTriggerActionCode -Row $row
        if ($ac -eq 1003) {
            $og = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_objguid')
            if ([string]::IsNullOrWhiteSpace($og)) { $og = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_objguid') }
            if (-not [string]::IsNullOrWhiteSpace($og)) { $invalidate[$og.Trim().ToLowerInvariant()] = $true }
        }
    }
    $statsRef = [ref]$stats
    _AuditPoller-ResolveDocFoldersBatched -Config $Config -DocGuids $docGuids -DocToFolder $docToFolder -StatsRef $statsRef -InvalidateGuids $invalidate | Out-Null

    $folderMap = @{}
    foreach ($row in $triggerRows) {
        $og = [string]$row.pw_objguid
        if ([string]::IsNullOrWhiteSpace($og)) { $og = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_objguid') }
        $pg = [string]$row.pw_parentguid
        if ([string]::IsNullOrWhiteSpace($pg)) { $pg = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_parentguid') }
        if ($og -and $docToFolder.ContainsKey($og) -and $pg -and -not $folderMap.ContainsKey($pg)) {
            $folderMap[$pg] = $docToFolder[$og]
        }
    }

    $parentGuids = @($triggerRows | ForEach-Object {
        $pg = [string]$_.pw_parentguid
        if ([string]::IsNullOrWhiteSpace($pg)) { $pg = [string](_AuditPoller-GetRowValue -Row $_ -Name 'o_parentguid') }
        $pg
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    $parentToFolder = @{}
    _AuditPoller-ResolveParentFoldersBatched -Config $Config -ParentGuids $parentGuids -ParentToFolder $parentToFolder -StatsRef $statsRef | Out-Null
    foreach ($pg in @($parentToFolder.Keys)) {
        if (-not $folderMap.ContainsKey($pg)) { $folderMap[$pg] = $parentToFolder[$pg] }
        $pgKey = $pg.Trim().ToLowerInvariant()
        if (-not $folderMap.ContainsKey($pgKey)) { $folderMap[$pgKey] = $parentToFolder[$pg] }
    }

    $normalizedWatchRoots = @(_AuditPoller-NormalizeWatchRootConfigs -WatchRootConfigs $WatchRootConfigs -Config $Config)
    $built = _AuditPoller-BuildCandidatesFromTriggerRows -Rows $triggerRows -DocToFolder $docToFolder -FolderMap $folderMap -WatchRootConfigs $normalizedWatchRoots -Config $Config
    $candidates = @($built.candidates)
    $stats.watchMatches = @($candidates).Count
    $stats.sheetsMatches = @($candidates | Where-Object { [bool]$_.isSheetsFolder }).Count
    try {
        $stats.candidateSkippedActionCode = [int]$built.skippedActionCode
        $stats.candidateSkippedNoFolder = [int]$built.skippedNoFolder
        $stats.candidateSkippedNoWatchMatch = [int]$built.skippedNoWatchMatch
    } catch { }

    if ($stats.relevantEvents -gt 0 -and $stats.watchMatches -eq 0 -and (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue)) {
        $watchRootPaths = @($normalizedWatchRoots | ForEach-Object { _AuditPoller-GetWatchRootPathFromConfig -Cfg $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $diagMatchRoots = _AuditPoller-BuildMatchRoots -WatchRoots $watchRootPaths
        $resolvedSamples = [System.Collections.Generic.List[string]]::new()
        $sampleMatchChecks = [System.Collections.Generic.List[string]]::new()
        foreach ($row in $triggerRows) {
            if ($resolvedSamples.Count -ge 5) { break }
            $og = [string]$row.pw_objguid
            if ([string]::IsNullOrWhiteSpace($og)) { $og = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_objguid') }
            $fp = $null
            if ($og -and $docToFolder.ContainsKey($og)) { $fp = [string]$docToFolder[$og] }
            else {
                $rf = [string](_AuditPoller-GetRowValue -Row $row -Name 'resolved_folder')
                if (-not [string]::IsNullOrWhiteSpace($rf)) { $fp = $rf }
            }
            if ($fp) {
                [void]$resolvedSamples.Add($fp)
                if ($sampleMatchChecks.Count -lt 3) {
                    $normFp = _AuditPoller-NormalizeFolderPath -FolderPath $fp
                    $matched = _AuditPoller-MatchesWatchRoot -FolderPath $fp -MatchRoots $diagMatchRoots
                    $acDiag = _AuditPoller-GetTriggerActionCode -Row $row
                    [void]$sampleMatchChecks.Add("action=$acDiag $normFp => $matched")
                }
            }
        }
        Write-QCJsonLog -Flush -Level 'Information' -Code 'WATCH_AUDIT_NO_WATCH_MATCH' -Message 'QC-relevant audit events did not match any projectWise.watchList.roots path.' -Data @{
            relevantEvents = $stats.relevantEvents
            triggerSource = [string]$stats.triggerSource
            auditLogicVersion = [string]$stats.auditLogicVersion
            watchRootCount = $watchRootPaths.Count
            matchRootsCount = $diagMatchRoots.Count
            watchRoots = @($watchRootPaths | Select-Object -First 5)
            resolvedFolderSamples = @($resolvedSamples)
            sampleMatchChecks = @($sampleMatchChecks)
            candidateSkippedActionCode = [int]$stats.candidateSkippedActionCode
            candidateSkippedNoFolder = [int]$stats.candidateSkippedNoFolder
            candidateSkippedNoWatchMatch = [int]$stats.candidateSkippedNoWatchMatch
            foldersResolved = [int]$stats.foldersResolved
            guidCacheMisses = [int]$stats.guidCacheMisses
            guidResolveSkipped = [int]$stats.guidResolveSkipped
        }
    }

    if ($built.folderUpdates.Count -gt 0 -and (Get-Command -Name 'Update-QCAuditEventsResolvedFolders' -ErrorAction SilentlyContinue)) {
        try { Update-QCAuditEventsResolvedFolders -Config $Config -Updates $built.folderUpdates | Out-Null } catch { }
    }

    if ($userNumbersToSync.Count -gt 0 -and (Get-Command -Name 'Sync-PWUserDirectory' -ErrorAction SilentlyContinue)) {
        try {
            Sync-PWUserDirectory -Config $Config -UserNumbers @($userNumbersToSync) -MaxUsers 25 | Out-Null
        } catch { }
    }

    $sw.Stop()
    return New-QCSuccessResult -Code 'AUDIT_SCAN_OK' -Message "Audit scan complete: $($stats.relevantEvents) unprocessed, $($stats.watchMatches) watch candidates." -Data @{
        events         = $triggerRows
        candidates     = $candidates
        docToFolder    = $docToFolder
        stats          = $stats
        watermarkAfter = $watermarkAfter
        durationMs     = [int]$sw.ElapsedMilliseconds
        pollWindow     = @{ since = $sinceStr; until = $queryUntilStr }
    }
}

function Get-AuditPollCycleCounter {
    <#
    .SYNOPSIS
    Reads/increments a cycle counter from a file to determine when to run reconciliation.
    Returns the current cycle number (1-based).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CounterPath)

    # 1-based: first tick returns 1 so cycle 0 does not force an immediate full-folder reconciliation.
    $counter = 1
    if (Test-Path -LiteralPath $CounterPath) {
        try { $counter = [int](Get-Content -LiteralPath $CounterPath -Raw -ErrorAction Stop).Trim() + 1 }
        catch { $counter = 1 }
    }
    try {
        $dir = Split-Path -Parent $CounterPath
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-Content -LiteralPath $CounterPath -Value ([string]$counter) -Encoding UTF8
    } catch { }
    return $counter
}

function Reset-AuditPollCycleCounter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CounterPath)
    try { Set-Content -LiteralPath $CounterPath -Value '0' -Encoding UTF8 } catch { }
}

function Get-AuditPollerLogicVersion {
    return $script:AuditPollerLogicVersion
}

function Test-QCAuditIngestAllowedActionCode {
    param([Parameter(Mandatory)][int]$ActionCode)
    return _AuditPoller-TestAuditIngestAllowedActionCode -ActionCode $ActionCode
}

Export-ModuleMember -Function Invoke-AuditTrailScan, Sync-AuditPollerWatchFolderGuidCache, Get-AuditTrailHighWaterMark, Get-AuditTrailHighWaterMarkFromDatabase, Get-AuditTrailCaptureWatermark, Set-AuditTrailCaptureWatermark, Get-AuditTrailPollWindow, Get-AuditPollCycleCounter, Reset-AuditPollCycleCounter, Get-AuditPollerLogicVersion, Get-PWAuditTrailActionName, Test-QCAuditIngestAllowedActionCode
