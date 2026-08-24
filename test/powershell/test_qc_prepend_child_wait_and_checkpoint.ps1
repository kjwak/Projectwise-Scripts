# Child-wait (PID only) and prepend/writeback checkpoint resume tests.
$ErrorActionPreference = 'Stop'

function Assert-True($Condition, $Message) {
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}
function Assert-Eq($Actual, $Expected, $Message) {
    if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot '_Resolve-ModuleImplPath.ps1')

Import-Module "$repoRoot/modules/Core/Core.Results.psm1" -Force
Import-Module "$repoRoot/modules/Processing/QC.Processors.psm1" -Force

$procText = Get-Content -LiteralPath (Resolve-ModuleImplPath -ModuleName 'QC.Processors.psm1') -Raw -Encoding UTF8
Assert-True ($procText -match '_QCP-WaitForLaunchedProcess') 'Processors should wait on the launched PowerShell PID'
Assert-True ($procText -match '_QCP-StartAndWaitLaunchedProcess') 'Processors should start prepend child without Start-Process -Wait'
Assert-True ($procText -match '_QCP-TestPrependWritebackOnlyResume') 'Processors should resume writeback from durable checkpoint'
Assert-True ($procText -match "checkpoint = 'prepend_complete'" -or $procText -match 'prepend_complete') 'Processors persist prepend_complete checkpoint'

Assert-True (-not (_QCP-TestPrependWritebackOnlyResume -Job @{})) 'Missing checkpoint is not writeback-only resume'
Assert-True (-not (_QCP-TestPrependWritebackOnlyResume -Job @{ checkpoint = 'notification_send' })) 'Unrelated checkpoint is not writeback-only resume'
Assert-True (_QCP-TestPrependWritebackOnlyResume -Job @{ checkpoint = 'prepend_complete' }) 'prepend_complete resumes writeback only'
Assert-True (_QCP-TestPrependWritebackOnlyResume -Job @{ checkpoint = 'writeback_running' }) 'writeback_running resumes writeback only'
Assert-True (_QCP-TestPrependAlreadyComplete -Job @{ checkpoint = 'writeback_complete' }) 'writeback_complete is already complete'
Assert-True (-not (_QCP-TestPrependAlreadyComplete -Job @{ checkpoint = 'prepend_complete' })) 'prepend_complete is not already complete'

# Wait returns when the launched PowerShell process exits even if a descendant keeps running.
$tmpWait = Join-Path $env:TEMP ('qc-prepend-wait-' + ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $tmpWait -Force | Out-Null
$grandPid = 0
try {
    $childScript = Join-Path $tmpWait 'child.ps1'
    $pidFile = Join-Path $tmpWait 'grand.pid'
    $stdoutPath = Join-Path $tmpWait 'out.txt'
    $stderrPath = Join-Path $tmpWait 'err.txt'
    @(
        '$ErrorActionPreference = "Stop"'
        '$g = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -Command Start-Sleep -Seconds 25" -PassThru -WindowStyle Hidden'
        ('Set-Content -LiteralPath "' + $pidFile.Replace('\', '\\') + '" -Value $g.Id -Encoding ASCII')
        '[Environment]::Exit(11)'
    ) | Set-Content -LiteralPath $childScript -Encoding ASCII

    $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-File', $childScript) `
        -NoNewWindow -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    Assert-True ($null -ne $p) 'Child process should start'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $wait = _QCP-WaitForLaunchedProcess -Process $p -PollMilliseconds 100 -TimeoutSeconds 15 -HeartbeatSeconds 0
    $sw.Stop()
    Assert-True $wait.exited 'Wait should return after the launched PowerShell process exits'
    Assert-True (-not $wait.timedOut) 'Wait should not time out on a short child'
    Assert-Eq ([int]$wait.exitCode) 11 'Wait should surface the launched process exit code'
    Assert-True ($sw.Elapsed.TotalSeconds -lt 12) 'Wait must not block on a long-lived descendant'
    Assert-True (Test-Path -LiteralPath $pidFile) 'Child should have recorded descendant PID'
    $grandPid = [int]((Get-Content -LiteralPath $pidFile -ErrorAction Stop).Trim())
    $grand = Get-Process -Id $grandPid -ErrorAction SilentlyContinue
    Assert-True ($null -ne $grand) 'Descendant should still be running after parent wait returned'
} finally {
    if ($grandPid -gt 0) {
        Stop-Process -Id $grandPid -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $tmpWait -Recurse -Force -ErrorAction SilentlyContinue
}

# Simulated prepend success + writeback failure must resume writeback without spawning prepend again.
$script:spawnCalls = 0
$script:writebackCalls = 0
InModuleScope -ModuleName QC.Processors {
    $script:spawnCalls = 0
    $script:writebackCalls = 0
    function _QCP-StartAndWaitLaunchedProcess {
        param(
            [string]$FilePath,
            [string]$ArgumentList,
            [string]$StdoutPath,
            [string]$StderrPath,
            [hashtable]$Job = $null,
            [hashtable]$Config = $null,
            [int]$TimeoutSeconds = 0,
            [int]$PollMilliseconds = 500,
            [int]$HeartbeatSeconds = 15
        )
        $script:spawnCalls++
        return @{
            process = $null
            processId = 4242
            exited = $true
            timedOut = $false
            exitCode = 0
            elapsedMs = 10
            stdout = 'ok'
            stderr = ''
            startFailed = $false
        }
    }
    function _QCP-InvokePrependPostSuccessWriteback {
        param(
            [hashtable]$Job,
            [hashtable]$Config,
            [object]$SuccessResult,
            [string]$IncomingFolder,
            [string]$IncomingDocName,
            [string]$DatasourceName,
            [hashtable]$ProjectWiseCfg,
            [switch]$ClearTriggerTag
        )
        $script:writebackCalls++
        if ($script:writebackCalls -eq 1) {
            return New-QCFailureResult -Code 'QC_PREPEND_POST_WRITEBACK_FAILED' -Message 'simulated writeback failure' -Data @{
                resumedFromCheckpoint = [bool]$SuccessResult.Data.resumedFromCheckpoint
            }
        }
        return New-QCSuccessResult -Code 'QC_PREPEND_OK' -Message 'writeback ok' -Data @{
            resumedFromCheckpoint = [bool]$SuccessResult.Data.resumedFromCheckpoint
            writebackCalls = $script:writebackCalls
        }
    }

    $job = @{
        id = 'qc_qcprepend_checkpoint_resume'
        type = 'QC_PREPEND'
        sourceFolder = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\n_seg'
        sourceName = '080J082001ab001.pdf'
        sourcePath = 'documents\caltrans\cafwy2200-i-15_elpse\cadd\sheets\n_seg\080J082001ab001.pdf'
        metadata = @{
            qcProcessType = 'review'
            prependTrigger = 'initialQcPdf'
            expectedLanePdfName = '080J082001ab001-rev.pdf'
            sourceDocumentGuid = 'ba1a9c32-2adf-4f39-9d0b-9dac20d0a286'
        }
    }
    $cfg = @{
        dryRun = $false
        qcPrepend = @{ mode = 'projectWise' }
        projectWise = @{ datasourceName = 'typsa-us-pw.bentley.com:typsa-us-pw-03' }
    }

    $first = Invoke-QCPrependProcessor -Job $job -Config $cfg
    if ($first.IsSuccess) { throw 'ASSERT FAILED: first pass should fail at simulated writeback' }
    if ($script:spawnCalls -ne 1) { throw "ASSERT FAILED: prepend child should run once on first pass (got $($script:spawnCalls))" }
    if ($script:writebackCalls -ne 1) { throw "ASSERT FAILED: writeback should run once on first pass (got $($script:writebackCalls))" }
    $cp = [string]$job.checkpoint
    if ($cp -ne 'prepend_complete' -and $cp -ne 'writeback_running') {
        throw "ASSERT FAILED: first pass must persist prepend_complete or writeback_running (got '$cp')"
    }

    $second = Invoke-QCPrependProcessor -Job $job -Config $cfg
    if (-not $second.IsSuccess) { throw "ASSERT FAILED: second pass should succeed writeback-only. code=$($second.Code) msg=$($second.Message)" }
    if ($script:spawnCalls -ne 1) { throw "ASSERT FAILED: prepend child must not run again after prepend_complete (got $($script:spawnCalls))" }
    if ($script:writebackCalls -ne 2) { throw "ASSERT FAILED: writeback should run again on resume (got $($script:writebackCalls))" }
    if (-not [bool]$second.Data.resumedFromCheckpoint) { throw 'ASSERT FAILED: resumed writeback should set resumedFromCheckpoint' }
}

Write-Host 'All QC_PREPEND child-wait and checkpoint tests passed.' -ForegroundColor Green
