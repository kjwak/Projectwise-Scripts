# Offline process, catalog and query contract checks for the WinForms operator console.
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$ops = Join-Path $root 'modules\Ops'
foreach ($module in @('QC.OpsConsole','QC.OpsDiagnostics','QC.OpsProcess','QC.OpsSqlExplorer')) {
    Import-Module (Join-Path $ops "$module.psm1") -DisableNameChecking
}
function Assert-True($Value, [string]$Message) { if (-not $Value) { throw $Message } }
function Assert-Rejected([scriptblock]$Test, [string]$Message) {
    $rejected = $false
    try { & $Test | Out-Null } catch { $rejected = $true }
    Assert-True $rejected $Message
}

$tempRoot = Join-Path $env:TEMP ('QC Ops Fixture ' + [guid]::NewGuid().ToString('N'))
New-Item -Path $tempRoot -ItemType Directory | Out-Null
try {
    $child = Join-Path $tempRoot 'Child Test.ps1'
    @'
param([string]$FolderPath, [switch]$Preview, [int]$Delay, [int]$ExitCode)
[Console]::Out.WriteLine("out:$FolderPath|$($Preview.IsPresent)")
[Console]::Out.WriteLine("pid:$PID")
[Console]::Error.WriteLine('test stderr')
Write-Host 'test host output'
if ($Delay) { Start-Sleep -Milliseconds $Delay }
exit $ExitCode
'@ | Set-Content -LiteralPath $child -Encoding utf8
    $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $folder = 'Documents\A "quoted" folder\'
    $result = Invoke-QCOpsChildProcess -ScriptPath $child -Parameters @{ FolderPath = $folder; Preview = $true } -OutputQueue $queue
    Assert-True ($result.IsSuccess -and $result.Data.exitCode -eq 0 -and $result.Data.elapsedMs -gt 0) "Typed process failed: $($result.Message) $($result.Data.output)"
    Assert-True ($queue.ToArray() -contains "out:$folder|True") 'Named arguments lost spaces, quotes, switch or trailing slash'
    Assert-True ($queue.ToArray() -contains '[stderr] test stderr') 'stderr was not captured independently'
    Assert-True ($result.Data.output -match 'test host output') 'Write-Host output was not readable'
    $result = Invoke-QCOpsChildProcess -ScriptPath $child -ArgumentList @('-FolderPath','Documents\Folder with spaces\','-ExitCode','7')
    Assert-True ($result.Code -eq 'OPS_CHILD_FAILED' -and $result.Data.exitCode -eq 7) 'Nonzero exit was not surfaced'
    Assert-True ($result.Data.output.Contains('Documents\Folder with spaces\')) 'Legacy argument quoting lost the path'
    $result = Invoke-QCOpsChildProcess -ScriptPath $child -Parameters @{ Delay = 9000 } -TimeoutMs 2200
    Assert-True ($result.Code -eq 'OPS_CHILD_TIMEOUT' -and $result.Data.elapsedMs -lt 7500) 'Child timeout did not stop the wait'
    if ($result.Data.output -match '(?m)^pid:(\d+)') {
        Assert-True (-not (Get-Process -Id ([int]$Matches[1]) -ErrorAction SilentlyContinue)) 'The timed-out child is still running'
    }
    $result = Invoke-QCOpsChildProcess -ScriptPath (Join-Path $tempRoot 'Missing.ps1')
    Assert-True ($result.Code -eq 'OPS_CHILD_MISSING') 'Missing script should produce a typed error'

    $catalog = @(Get-QCOpsMaintenanceCatalog)
    foreach ($entry in $catalog) {
        $metadata = Get-Command $entry.script
        foreach ($field in $entry.fields) { Assert-True ($metadata.Parameters.ContainsKey($field.name)) "Invalid field $($field.name) on $($entry.id)" }
        foreach ($mode in @('apply','preview')) {
            if (-not $entry.ContainsKey($mode)) { continue }
            foreach ($arg in $entry[$mode].Keys) { Assert-True ($metadata.Parameters.ContainsKey($arg)) "Invalid mode parameter $arg on $($entry.id)" }
        }
    }
    $action = $catalog | Where-Object id -eq 'sync-sheet-index'
    $preview = Get-QCOpsMaintenanceParameters -Action $action -Values @{ FolderPath = $folder; OneLevelDeep = $true } -AppSettingsPath 'C:\config with spaces\local.json'
    $apply = Get-QCOpsMaintenanceParameters -Action $action -Values @{ FolderPath = $folder } -Apply
    Assert-True ($preview.DryRun -and $preview.OneLevelDeep -and -not $preview.ConfirmWrites) 'Preview lost its write guard'
    Assert-True ($apply.ConfirmWrites -and -not $apply.DryRun) 'Apply did not use the script confirmation contract'
    Assert-Rejected { Get-QCOpsMaintenanceParameters -Action $action -Values @{} } 'Missing required folder was accepted'
    $presets = @(Get-QCOpsFolderPresets -Config @{
        projectWise = @{ watchList = @{ roots = @(@{ path = 'Documents\Workspace'; projectDepth = 1; sheetsPathFromProject = 'CADD\Sheets' }) } }
        watchFolders = @('Documents\Workspace\Project\CADD\Sheets')
    })
    Assert-True ($presets.Count -eq 2) 'Configured watch root and concrete folder were not represented'

    $logs = @(
        @{ ts = '2026-11-02T09:00:01-07:00'; code = 'TEST_OLD'; level = 'Warning'; logFile = 'watch.jsonl'; message = "a`r`nb" }
        @{ ts = '2026-11-02T16:01:00Z'; code = 'TEST_NEW'; level = 'Error'; logFile = 'worker.jsonl'; data = @{ jobId = 'job-1' } }
    )
    $rows = @(ConvertTo-QCOpsLogRows -Events $logs)
    Assert-True ($rows[0].code -eq 'TEST_NEW' -and $rows[0].job -eq 'job-1') 'Logs should sort by instant and expose IDs'
    Assert-True ($rows[1].detail.message -eq "a`r`nb") 'Multiline log order was changed'
    Assert-True ((@(ConvertTo-QCOpsLogRows -Events $logs -Ascending))[0].code -eq 'TEST_OLD') 'Oldest-first log view is broken'
    $fixtureFolders = @('pending','_logs','_prepend_holds')
    foreach ($folderName in $fixtureFolders) { [void](New-Item -Path (Join-Path $tempRoot $folderName) -ItemType Directory) }
    @{ id = 'pending-example'; type = 'STATUS_SET_GEN' } | ConvertTo-Json -Compress |
        Set-Content -LiteralPath (Join-Path $tempRoot 'pending\pending-example.json')
    @{ jobId = 'held-example'; lane = @{ folder = 'Documents\Example'; document = 'test-rev.pdf' } } | ConvertTo-Json -Compress |
        Set-Content -LiteralPath (Join-Path $tempRoot '_prepend_holds\sample.json')
    $logs[0] | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $tempRoot '_logs\Watch-QCTrigger_fixture.jsonl')
    $snapshotArgs = @{
        Config = @{ queue = @{ root = $tempRoot }; database = @{ enabled = $false } }
        TaskName = 'QC-OpsConsole-Fixture-DoesNotExist'; IncludeQueue = $true; IncludeLogs = $true
    }
    $snapshot = Get-QCOpsSnapshot @snapshotArgs
    Assert-True ($snapshot.queueAvailable -and $snapshot.oldest.BaseName -eq 'pending-example') 'Snapshot lost its configured queue or oldest job'
    Assert-True ($snapshot.holds -eq 1 -and @($snapshot.queue | Where-Object state -eq 'held').Count -eq 1) 'Durable lane hold is missing'
    Assert-True ($snapshot.logs.Count -eq 1 -and $snapshot.logs[0].source -eq 'watcher') 'Snapshot did not read its watcher log'

    $dataset = [pscustomobject]@{
        schema = 'dbo'; name = 'unsafe]name'; time = 'created_at'; folder = 'source_folder'
        columns = @(
            @{ name = 'id'; type = 'int'; sortable = $true; key = 1 }
            @{ name = 'created_at'; type = 'datetimeoffset'; sortable = $true; key = 0 }
            @{ name = 'source_folder'; type = 'nvarchar'; sortable = $true; key = 0 }
            @{ name = 'status'; type = 'nvarchar'; sortable = $true; key = 0 }
        ) | ForEach-Object { [pscustomobject]$_ }
    }
    $query = New-QCOpsExplorerQuery -Dataset $dataset -Filter @{ Folder = 'A_100%'; Column = 'status'; Text = "value'; text" } -Page 2 -PageSize 50
    Assert-True ($query.sql.Contains('[unsafe]]name]') -and $query.sql.Contains('[created_at] DESC, [id] DESC')) 'Identifiers or newest-first ordering are wrong'
    Assert-True ($query.parameters.skip -eq 100 -and $query.parameters.take -eq 51) 'Server-side paging offset or extra-row detection is wrong'
    Assert-True ($query.parameters.underFolder -eq 'A[_]100[%]\%') 'Folder selector did not escape LIKE wildcards'
    Assert-True (-not $query.sql.Contains("value'; text")) 'Filter values must only appear in parameters'
    Assert-Rejected { New-QCOpsExplorerQuery -Dataset $dataset -Sort 'bad name' } 'Invalid sort identifier was accepted'
    $exportQuery = New-QCOpsExplorerQuery -Dataset $dataset -Projection @('status') -Export
    Assert-True ($exportQuery.sql.StartsWith('SELECT [status] ') -and $exportQuery.parameters.take -eq 5001) 'Export reads more than its selected columns or limit'
    $sqlModule = Get-Module QC.OpsSqlExplorer
    $metadata = & $sqlModule {
        function Invoke-QCDatabaseQuery {
            $table = New-Object Data.DataTable
            foreach ($name in @('schema_name','object_name','object_type','column_name','data_type')) { [void]$table.Columns.Add($name) }
            [void]$table.Columns.Add('max_length',[int])
            [void]$table.Columns.Add('key_ordinal',[int])
            [void]$table.Rows.Add('dbo','processing_jobs','U','id','int',4,1)
            [void]$table.Rows.Add('dbo','processing_jobs','U','created_at','datetimeoffset',10,0)
            [void]$table.Rows.Add('dbo','processing_jobs','U','source_folder','nvarchar',2000,0)
            [void]$table.Rows.Add('dbo','outside_view','V','event_time','datetime2',8,0)
            New-QCSuccessResult -Data @{ table = $table }
        }
        Get-QCOpsExplorerCatalog -Config @{ database = @{ enabled = $true } }
    }
    Assert-True ($metadata.IsSuccess -and $metadata.Data.items.Count -eq 2) 'Metadata did not create distinct table and view options'
    Assert-True ($metadata.Data.items[0].curated -and $metadata.Data.items[0].time -eq 'created_at') 'Curated timestamp dataset did not sort first'
    Assert-True ($metadata.Data.items[1].time -eq 'event_time') 'Generic dataset did not detect its available timestamp'
    Assert-True ((New-QCOpsCustomSelect "SELECT 'INSERT; DROP' AS [sample]") -match 'TOP \(251\)') 'A literal-only SELECT should be accepted'
    foreach ($invalid in @('SELECT a INTO dbo.b FROM dbo.c','SELECT 1; SELECT 2','SELECT * FROM OPENROWSET(...)','UPDATE dbo.foo SET x=1','SELECT NEXT VALUE FOR dbo.foo')) {
        Assert-Rejected { New-QCOpsCustomSelect $invalid } "Rejected custom SELECT check failed: $invalid"
    }
    $table = New-Object Data.DataTable
    [void]$table.Columns.Add('visible')
    [void]$table.Columns.Add('hidden')
    [void]$table.Rows.Add('=1+1','not selected')
    $csv = Join-Path $tempRoot 'sample.csv'
    Export-QCOpsExplorerCsv -Table $table -Columns @('visible') -Path $csv
    $data = Import-Csv -LiteralPath $csv
    Assert-True ($data.visible -eq "'=1+1" -and -not $data.PSObject.Properties['hidden']) 'CSV omitted its selected column or formula protection'
    Write-Output 'test_ops_console_overhaul: passed'
} finally {
    # This test only removes its own GUID-named temporary fixture.
    foreach ($folderName in @('pending','_logs','_prepend_holds')) {
        $folderPath = Join-Path $tempRoot $folderName
        if (Test-Path -LiteralPath $folderPath) {
            Get-ChildItem -LiteralPath $folderPath -File | Remove-Item -Force
            Remove-Item -LiteralPath $folderPath -Force
        }
    }
    Get-ChildItem -LiteralPath $tempRoot -File | Remove-Item -Force
    Remove-Item -LiteralPath $tempRoot -Force
}
