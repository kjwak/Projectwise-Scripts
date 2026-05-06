# QC.Worker.psm1
# Responsibility: Shared worker retry and state-transition policy.

function Move-QCJobWithLockRetries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$FromState,
        [Parameter(Mandatory)][string]$ToState,
        [Parameter(Mandatory)][hashtable]$Config,
        [hashtable]$Job,
        [int]$MaxTries = 8,
        [int]$SleepMs = 3000
    )

    $last = $null
    for ($t = 1; $t -le $MaxTries; $t++) {
        $last = Move-QCJob -JobId $JobId -FromState $FromState -ToState $ToState -Config $Config -Job $Job
        if ($last.IsSuccess) { return $last }
        if ([string]$last.Code -ne 'QUEUE_LOCK_TIMEOUT') { return $last }
        if ($t -ge $MaxTries) { return $last }
        Start-Sleep -Milliseconds $SleepMs
    }
    return $last
}

Export-ModuleMember -Function *
