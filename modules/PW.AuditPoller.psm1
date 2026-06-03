# PW.AuditPoller.psm1
# Responsibility: Query ProjectWise audit trail (dms_audt), resolve folders,
# match against watch roots, and return candidate events for job creation.
# Extracted from the POC Test-AuditEventIngestion.ps1 for production use.

# Dependencies (Core.Results, Core.Runtime, Core.Database) must be imported by the
# caller before this module. Re-importing with -Force here would clobber their
# global-scope exports.

# Bump when trigger/candidate logic changes; appears in WATCH_AUDIT_SCAN_DONE logs.
$script:AuditPollerLogicVersion = '2026-06-03-hashtable-row-v3'

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

# Full PW action map for audit_events ingestion (telemetry is unfiltered; triggers use QCRelevantActions only).
$script:AuditActionNames = @{
    1    = 'FOLDER_CREATE';     2    = 'FOLDER_MODIFY';     3    = 'FOLDER_WFLOW'
    4    = 'FOLDER_DELETE';     5    = 'FOLDER_STATE'
    1001 = 'DOCUMENT_CREATE';   1002 = 'DOCUMENT_MODIFY';   1003 = 'DOCUMENT_ATTR'
    1004 = 'DOCUMENT_FILE_ADD'; 1005 = 'DOCUMENT_FILE_REM'; 1006 = 'DOCUMENT_FILE_REP'
    1007 = 'DOCUMENT_CIN';     1008 = 'DOCUMENT_VIEW';     1009 = 'DOCUMENT_CHOUT'
    1010 = 'DOCUMENT_CPOUT';   1011 = 'DOCUMENT_GOUT';     1012 = 'DOCUMENT_STATE'
    1013 = 'DOCUMENT_FINAL_S'; 1014 = 'DOCUMENT_FINAL_R'; 1015 = 'DOCUMENT_VERSION'
    1016 = 'DOCUMENT_MOVE';    1020 = 'DOCUMENT_DELETE';    1022 = 'DOCUMENT_FREE'
    1027 = 'DOCUMENT_IMPORT'; 3001 = 'USER_LOGIN';        3002 = 'USER_LOGOUT'
}

# Session caches: avoid repeated Get-PWDocumentsByGUIDs for missing or known documents.
$script:AuditPoller_DocFolderCache = @{}
$script:AuditPoller_UnresolvedGuids = @{}

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
        elseif (-not $resolvedFolder -and $parentGuid -and $FolderMap.ContainsKey($parentGuid)) {
            $resolvedFolder = _AuditPoller-NormalizeFolderPath -FolderPath ([string]$FolderMap[$parentGuid])
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
        $itemName = [string](_AuditPoller-GetRowValue -Row $row -Name 'pw_itemname')
        if ([string]::IsNullOrWhiteSpace($itemName)) { $itemName = [string](_AuditPoller-GetRowValue -Row $row -Name 'o_itemname') }

        $candidates += @{
            auditEventId        = $auditId
            objGuid             = $objGuid
            parentGuid          = $parentGuid
            actionCode          = $actionCode
            actionName          = $actionName
            itemName            = $itemName
            actTime             = $actTime
            userno              = $userno
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
    $t = ($FolderPath -as [string]).Trim().TrimEnd('\').Replace('/', '\')
    $t = $t -replace '\\{2,}', '\'
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    # Align with Core.Paths: strip pw:\datasource\ prefix so watch roots can match.
    if ($t -match '^(?i)pw:\\') {
        $idx = $t.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) { $t = $t.Substring($idx + 1) }
    }
    if ($t -match '^(?i)documents\\') { return $t }
    return ('Documents\' + $t)
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
    if ($script:QCRelevantActions.ContainsKey($ActionCode)) { return $script:QCRelevantActions[$ActionCode] }
    if ($script:AuditActionNames.ContainsKey($ActionCode)) { return $script:AuditActionNames[$ActionCode] }
    return "UNKNOWN_$ActionCode"
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

    $since = if ($lastCapture) {
        if ($UseRestartOverlap.IsPresent -and $restartOverlapSeconds -gt 0) {
            $lastCapture.AddSeconds(-$restartOverlapSeconds)
        } elseif ($overlapSeconds -gt 0) {
            $lastCapture.AddSeconds(-$overlapSeconds)
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
            if ($UseRestartOverlap.IsPresent) { $restartOverlapSeconds } else { $overlapSeconds }
        } else { $initialLookbackSeconds }
        overlapSecondsUsed  = if ($UseRestartOverlap.IsPresent) { $restartOverlapSeconds } else { $overlapSeconds }
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
        dbUnprocessedLoaded = 0
        guidCacheHits      = 0
        guidCacheMisses    = 0
        guidResolveSkipped = 0
        failedGuidCacheHits = 0
        triggerSource      = 'pw_batch'
        auditLogicVersion  = $script:AuditPollerLogicVersion
    }

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
    $pageNum = 0
    try {
        while ($pageNum -lt $maxPages) {
            $pageNum++
            $untilStr = $queryUntilStr
            $lowerBoundSql = if ($cursorGuid) {
                $t = _AuditPoller-EscapeSqlLiteral -Value $cursorSince
                $g = _AuditPoller-EscapeSqlLiteral -Value $cursorGuid
                "(o_acttime > '$t' OR (o_acttime = '$t' AND o_objguid > '$g'))"
            } else {
                "o_acttime > '$(_AuditPoller-EscapeSqlLiteral -Value $cursorSince)'"
            }
            $sql = "SELECT TOP $pageSize o_acttime, o_action, o_objtype, o_objno, o_objguid, o_parentguid, o_userno, o_itemname, o_itemdesc, o_textparam FROM dms_audt WHERE $lowerBoundSql AND o_acttime <= '$(_AuditPoller-EscapeSqlLiteral -Value $untilStr)' ORDER BY o_acttime ASC, o_objguid ASC"
            $result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
            $batch = @(_AuditPoller-GetSqlResultRows -Result $result)
            if ($batch.Count -eq 0) { break }
            foreach ($row in $batch) { [void]$allEvents.Add($row) }
            $last = $batch[$batch.Count - 1]
            $lastAct = _AuditPoller-GetRowValue -Row $last -Name 'o_acttime'
            $cursorSince = if ($lastAct -is [DateTime]) {
                _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $lastAct)
            } else {
                _AuditPoller-FormatActTime -Value $lastAct -Config $Config
            }
            if ([string]::IsNullOrWhiteSpace($cursorSince)) { break }
            $cursorGuid = [string](_AuditPoller-GetRowValue -Row $last -Name 'o_objguid')
            if ($batch.Count -lt $pageSize) { break }
        }
        if ($pageNum -ge $maxPages -and $allEvents.Count -gt 0) {
            $stats.eventsTruncated = $true
            if (Get-Command -Name 'Write-QCJsonLog' -ErrorAction SilentlyContinue) {
                Write-QCJsonLog -Flush -Level 'Warning' -Code 'AUDIT_QUERY_TRUNCATED' -Message "dms_audt page cap reached ($maxPages x $pageSize); re-run will continue from watermark cursor." -Data @{
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
        return New-QCSuccessResult -Code 'AUDIT_NO_EVENTS' -Message 'No audit events in window.' -Data @{
            events = @(); candidates = @(); docToFolder = @{}; stats = $stats
            watermarkAfter = $null; durationMs = [int]$sw.ElapsedMilliseconds
            pollWindow = @{ since = $sinceStr; until = $queryUntilStr }
        }
    }

    # 2. Ingest every fetched row into audit_events (no QC/watch/action filtering).
    $maxPwActTime = $null
    $maxPwActTimeUtc = $null
    $watermarkAfter = $null
    $dbRows = @()
    if (Test-QCDatabaseEnabled -Config $Config) {
        foreach ($evt in $allEvents) {
            $rawAct = _AuditPoller-GetRowValue -Row $evt -Name 'o_acttime'
            $actTime = _AuditPoller-FormatActTime -Value $rawAct -Config $Config
            if ($actTime) { $maxPwActTime = _AuditPoller-TryAdvanceWatermarkAfter -Current $maxPwActTime -Candidate $actTime }
            if ($null -ne $rawAct -and -not ($rawAct -is [DBNull])) {
                $utcStr = if ($rawAct -is [DateTime]) {
                    _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $rawAct)
                } else {
                    $parsed = _AuditPoller-ParseActTime -ActTime ([string]$rawAct)
                    if ($parsed) { _AuditPoller-FormatSqlUtc -Utc (_AuditPoller-AssumeUtcFromPw -DateTime $parsed) } else { $null }
                }
                if ($utcStr) { $maxPwActTimeUtc = _AuditPoller-TryAdvanceWatermarkAfter -Current $maxPwActTimeUtc -Candidate $utcStr }
            }
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
                if ($dbRes.Data.lastError) { $stats.dbLastError = [string]$dbRes.Data.lastError }
            } else {
                $stats.dbSkipped += $dbRows.Count
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
                rowsPrepared  = $stats.dbRowsPrepared
                rowsNullGuid  = $stats.dbRowsNullGuid
                written       = $stats.dbWrites
                skipped       = $stats.dbSkipped
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

    if ($maxPwActTime) { $stats.maxPwActTime = $maxPwActTime }
    if ($maxPwActTimeUtc) {
        $stats.maxPwActTimeUtc = $maxPwActTimeUtc
        $watermarkAfter = $maxPwActTimeUtc
    }

    # 3. QC trigger pipeline — database is source of truth for unprocessed rows.
    $triggerRows = @()
    $useDbTriggers = $false
    if (Test-QCDatabaseEnabled -Config $Config) {
        $maxUnprocessed = _AuditPoller-GetAuditPollerInt -Config $Config -Key 'maxUnprocessedPerPoll' -Default 500
        if (Get-Command -Name 'Get-QCUnprocessedAuditEvents' -ErrorAction SilentlyContinue) {
            $unprocRes = Get-QCUnprocessedAuditEvents -Config $Config -MaxRows $maxUnprocessed
            if ($unprocRes.IsSuccess -and $null -ne $unprocRes.Data) {
                $dbRows = @()
                if ($unprocRes.Data.rows) { $dbRows = @($unprocRes.Data.rows) }
                if ($dbRows.Count -gt 0) {
                    $triggerRows = $dbRows
                    $stats.dbUnprocessedLoaded = $triggerRows.Count
                    $stats.triggerSource = 'audit_events_db'
                    $useDbTriggers = $true
                }
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

    [void](_AuditPoller-LoadDocFolderCache -Config $Config)
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

Export-ModuleMember -Function Invoke-AuditTrailScan, Get-AuditTrailHighWaterMark, Get-AuditTrailHighWaterMarkFromDatabase, Get-AuditTrailCaptureWatermark, Set-AuditTrailCaptureWatermark, Get-AuditTrailPollWindow, Get-AuditPollCycleCounter, Reset-AuditPollCycleCounter, Get-AuditPollerLogicVersion
