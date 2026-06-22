<#
.SYNOPSIS
Promote commits from dev to main (fast-forward when possible) and push main.

.DESCRIPTION
Use after committing and pushing work to dev. This script fetches origin,
shows commits on dev that are not yet on main, merges dev into main, pushes
main, and returns you to dev.

.PARAMETER DevBranch
Source branch. Default 'dev'.

.PARAMETER MainBranch
Target branch. Default 'main'.

.PARAMETER Remote
Remote name. Default 'origin'.

.PARAMETER SkipFetch
Skip 'git fetch' before comparing branches.

.PARAMETER WhatIf
Show what would happen without checking out, merging, or pushing.

.EXAMPLE
.\scripts\Promote-DevToMain.ps1

.EXAMPLE
.\scripts\Promote-DevToMain.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string]$DevBranch = 'dev',

    [Parameter(Mandatory = $false)]
    [string]$MainBranch = 'main',

    [Parameter(Mandatory = $false)]
    [string]$Remote = 'origin',

    [Parameter(Mandatory = $false)]
    [switch]$SkipFetch
)

$ErrorActionPreference = 'Stop'

function Invoke-Git {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    # Git writes routine status text to stderr; with $ErrorActionPreference = 'Stop'
    # PowerShell would treat those lines as terminating errors.
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & git @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousEap
    }

    $lines = @(
        $output | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.ToString() } else { [string]$_ }
        }
    )

    if ($exitCode -ne 0) {
        $msg = if ($lines.Count -gt 0) { ($lines -join [Environment]::NewLine).Trim() } else { "git $($Arguments -join ' ') failed with exit code $exitCode." }
        throw $msg
    }

    return $lines
}

function Write-Step {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path (Split-Path $scriptDir -Parent) -Parent

Push-Location $repoRoot
$originalBranch = $null
try {
    $null = Invoke-Git @('rev-parse', '--is-inside-work-tree')

    $status = Invoke-Git @('status', '--porcelain')
    if ($status) {
        throw "Working tree is not clean. Commit or stash changes before promoting dev to main."
    }

    $originalBranch = [string](Invoke-Git @('rev-parse', '--abbrev-ref', 'HEAD'))

    if (-not $SkipFetch) {
        Write-Step "Fetching $Remote..."
        Invoke-Git @('fetch', $Remote) | Out-Null
    }

    $devRef = "$Remote/$DevBranch"
    $mainRef = "$Remote/$MainBranch"

    foreach ($ref in @($devRef, $mainRef)) {
        $exists = Invoke-Git @('rev-parse', '--verify', $ref)
        if (-not $exists) { throw "Missing remote ref: $ref" }
    }

    $counts = [string](Invoke-Git @('rev-list', '--left-right', '--count', "$mainRef...$devRef"))
    $parts = $counts.Trim() -split '\s+'
    if ($parts.Count -ne 2) { throw "Unexpected rev-list output: $counts" }

    $mainAhead = [int]$parts[0]
    $devAhead = [int]$parts[1]

    if ($devAhead -eq 0 -and $mainAhead -eq 0) {
        Write-Host "$DevBranch and $MainBranch are already in sync on $Remote." -ForegroundColor Green
        return
    }

    if ($mainAhead -gt 0) {
        throw "$MainBranch is $mainAhead commit(s) ahead of $DevBranch on $Remote. Merge or rebase $DevBranch onto $MainBranch before promoting."
    }

    Write-Step "Commits on $DevBranch not yet on ${MainBranch}:"
    Invoke-Git @('log', '--oneline', "$mainRef..$devRef") | ForEach-Object { Write-Host "  $_" }

    if (-not $PSCmdlet.ShouldProcess("$Remote/$MainBranch", "Merge $DevBranch and push")) {
        Write-Host "(WhatIf - no branches changed, nothing pushed)" -ForegroundColor DarkGray
        return
    }

    Write-Step "Checking out $MainBranch..."
    Invoke-Git @('checkout', $MainBranch) | ForEach-Object { Write-Host $_ }

    Write-Step "Pulling latest $MainBranch from $Remote..."
    Invoke-Git @('pull', $Remote, $MainBranch) | ForEach-Object { Write-Host $_ }

    Write-Step "Merging $DevBranch into $MainBranch..."
    Invoke-Git @('merge', $DevBranch, '-m', "Merge branch '$DevBranch' into $MainBranch") | ForEach-Object { Write-Host $_ }

    Write-Step "Pushing $MainBranch to $Remote..."
    Invoke-Git @('push', $Remote, $MainBranch) | ForEach-Object { Write-Host $_ }

    Write-Step "Returning to $DevBranch..."
    Invoke-Git @('checkout', $DevBranch) | ForEach-Object { Write-Host $_ }

    $newTip = [string](Invoke-Git @('rev-parse', '--short', 'HEAD'))
    Write-Host ""
    Write-Host "Done. $MainBranch now matches $DevBranch at $newTip." -ForegroundColor Green
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    try {
        $current = [string](Invoke-Git @('rev-parse', '--abbrev-ref', 'HEAD'))
        if ($current -ne $originalBranch -and $originalBranch) {
            Write-Step "Restoring original branch '$originalBranch'..."
            Invoke-Git @('checkout', $originalBranch) | Out-Null
        }
    } catch {
        Write-Host "Could not restore original branch: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    exit 1
}
finally {
    Pop-Location
}
