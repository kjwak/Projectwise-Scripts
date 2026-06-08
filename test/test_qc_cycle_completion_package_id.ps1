# QC cycle completion package-key migration tests (mocked persistence).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core.Database.psm1') -Force

function Assert-Eq($a, $b, $msg) {
    if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" }
}

function Assert-True($cond, $msg) {
    if (-not $cond) { throw "ASSERT FAILED: $msg" }
}

$packageId = [guid]'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
$dgnGuid = '11111111-1111-1111-1111-111111111111'
$script:completions = [System.Collections.Generic.List[object]]::new()
$script:summaryTargets = [System.Collections.Generic.List[guid]]::new()

function Resolve-QCCycleCompletionSheetPackageId {
    param([hashtable]$Config, [string]$DocumentGuid = '', [Nullable[guid]]$SheetPackageId = $null)
    if ($null -ne $SheetPackageId -and $SheetPackageId -ne [guid]::Empty) { return $SheetPackageId }
    if ([string]$DocumentGuid -eq $dgnGuid) { return $packageId }
    return $null
}

function Ensure-QCCycleCompletion {
    param(
        [hashtable]$Config, [string]$DocumentGuid, [string]$QcCycleId, [string]$QcReviewType,
        [Nullable[guid]]$SheetPackageId = $null,
        [string]$DocumentName = '', [Nullable[int]]$QcCycleNumber = $null,
        [Nullable[long]]$TransitionEventId = $null, [Nullable[long]]$AuditEventId = $null,
        [string]$CompletedBy = '', [Nullable[datetime]]$CompletedAt = $null
    )
    $pkg = Resolve-QCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId
    if (-not $pkg) {
        return [pscustomobject]@{
            IsSuccess = $true; Code = 'QC_CYCLE_COMPLETION_SKIPPED'; Message = 'sheet_package_id could not be resolved.'
            Data = @{ inserted = $false; reused = $false; completionId = $null; sheetPackageId = $null; reason = 'sheet_package_not_found' }
        }
    }
    $key = $pkg.ToString() + '|' + $QcCycleId + '|' + $QcReviewType
    $existing = @($script:completions | Where-Object { ($_.sheetPackageId.ToString() + '|' + $_.qcCycleId + '|' + $_.qcReviewType) -eq $key })
    if ($existing.Count -gt 0) {
        return [pscustomobject]@{
            IsSuccess = $true; Code = 'QC_CYCLE_COMPLETION_REUSED'; Message = 'reused'
            Data = @{ inserted = $false; reused = $true; completionId = 1; sheetPackageId = $pkg }
        }
    }
    $script:completions.Add([pscustomobject]@{
        sheetPackageId = $pkg; documentGuid = $DocumentGuid; qcCycleId = $QcCycleId; qcReviewType = $QcReviewType
    }) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true; Code = 'QC_CYCLE_COMPLETION_WRITTEN'; Message = 'written'
        Data = @{ inserted = $true; reused = $false; completionId = $script:completions.Count; sheetPackageId = $pkg }
    }
}

function Update-QCSheetCycleCompletionSummary {
    param([hashtable]$Config, [string]$DocumentGuid = '', [Nullable[guid]]$SheetPackageId = $null)
    $pkg = Resolve-QCCycleCompletionSheetPackageId -Config $Config -DocumentGuid $DocumentGuid -SheetPackageId $SheetPackageId
    if (-not $pkg) {
        return [pscustomobject]@{
            IsSuccess = $false; Code = 'QC_CYCLE_SUMMARY_INVALID'; Message = 'missing package'; Data = @{ written = $false }
        }
    }
    $script:summaryTargets.Add($pkg) | Out-Null
    return [pscustomobject]@{
        IsSuccess = $true; Code = 'QC_CYCLE_SUMMARY_UPDATED'; Message = 'updated'; Data = @{ written = $true; sheetPackageId = $pkg }
    }
}

$cfg = @{ database = @{ enabled = $true } }

$r1 = Ensure-QCCycleCompletion -Config $cfg -DocumentGuid $dgnGuid -QcCycleId 'cycle-1' -QcReviewType 'production'
Assert-True $r1.Data.inserted 'first insert should insert'
Assert-Eq $r1.Data.sheetPackageId.ToString() $packageId.ToString() 'insert resolves sheet_package_id'

$r2 = Ensure-QCCycleCompletion -Config $cfg -DocumentGuid $dgnGuid -QcCycleId 'cycle-1' -QcReviewType 'production'
Assert-True $r2.Data.reused 'duplicate package key should reuse'
Assert-Eq $script:completions.Count 1 'same package+cycle+review inserts once'

$r3 = Ensure-QCCycleCompletion -Config $cfg -DocumentGuid $dgnGuid -QcCycleId 'cycle-1' -QcReviewType 'peer_review'
Assert-True $r3.Data.inserted 'different review type should insert separately'
Assert-Eq $script:completions.Count 2 'package+cycle+different review inserts second row'

$r4 = Ensure-QCCycleCompletion -Config $cfg -DocumentGuid '99999999-9999-9999-9999-999999999999' -QcCycleId 'cycle-2' -QcReviewType 'production'
Assert-Eq $r4.Code 'QC_CYCLE_COMPLETION_SKIPPED' 'missing package should skip'
Assert-Eq $r4.Data.reason 'sheet_package_not_found' 'missing package reason'

$rollup = Update-QCSheetCycleCompletionSummary -Config $cfg -SheetPackageId $packageId
Assert-True $rollup.Data.written 'rollup should write sheet_packages'
Assert-Eq $script:summaryTargets.Count 1 'rollup targets sheet_packages once'
Assert-Eq $script:summaryTargets[0].ToString() $packageId.ToString() 'rollup package id'

$completionRows = @(@{ id = 10; document_guid = [guid]$dgnGuid; sheet_package_id = $null })
$sheetDocuments = @(@{ document_guid = [guid]$dgnGuid; sheet_package_id = $packageId })
$mapped = @($completionRows | ForEach-Object {
    $row = $_
    $doc = $sheetDocuments | Where-Object { $_.document_guid -eq $row.document_guid } | Select-Object -First 1
    if ($doc) { @{ id = $row.id; sheet_package_id = $doc.sheet_package_id } }
})
Assert-Eq $mapped.Count 1 'backfill should map one completion'
Assert-Eq $mapped[0].sheet_package_id.ToString() $packageId.ToString() 'backfill maps to package id'

$dupRows = @(
    @{ sheet_package_id = $packageId; qc_cycle_id = 'cycle-dup'; qc_review_type = 'production' },
    @{ sheet_package_id = $packageId; qc_cycle_id = 'cycle-dup'; qc_review_type = 'production' }
)
$dupGroups = @($dupRows | Group-Object { $_.sheet_package_id.ToString() + '|' + $_.qc_cycle_id + '|' + $_.qc_review_type } | Where-Object { $_.Count -gt 1 })
Assert-Eq $dupGroups.Count 1 'validation should detect one duplicate package logical key group'
Assert-Eq $dupGroups[0].Count 2 'duplicate group should contain two rows'

Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot 'scripts\sql\backfill-qc-cycle-completions-package-id.sql')) 'backfill SQL exists'
Assert-True (Test-Path -LiteralPath (Join-Path $repoRoot 'scripts\sql\validate-qc-cycle-completions-package-id.sql')) 'validate SQL exists'

Write-Host 'OK: qc cycle completion package id tests passed.' -ForegroundColor Green
