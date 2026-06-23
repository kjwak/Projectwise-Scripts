# User-initiated workflow notifications must target the lane QC PDF with qcProcessType set.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force

function Assert-True($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Assert-Eq($a, $b, $msg) { if ($a -ne $b) { throw "ASSERT FAILED: $msg (got '$a', expected '$b')" } }

$mod = Import-Module (Join-Path $repoRoot 'modules\Workflow\QC.AuditTriggers.psm1') -Force -PassThru
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Discovery.psm1') -Force -WarningAction SilentlyContinue
Assert-True ($null -ne $mod) 'QC.AuditTriggers module loaded'

$cfg = @{
    auditPoller = @{
        workflowTriggers = @{
            enabled = $true
            notifyOnStateChange = $true
            qcPdfNotificationsOnly = $true
        }
    }
}

$result = & $mod {
    param($Config)
    $revMember = @{
        documentGuid = '7f484f1b-54d2-42cb-8068-ee1a9ca9fa09'
        documentName = '0818000063ea501-rev.pdf'
        document = $null
    }
    $resolved = _QCAT-ResolveSheetPackageNotificationMember -Members @($revMember) `
        -TriggerDocumentGuid $revMember.documentGuid -TriggerDocumentName $revMember.documentName `
        -Config $Config -FolderPath 'documents\test\seg_1' -SheetStem '0818000063ea501'

    $ctx = _QCAT-BuildWorkflowStateNotifyContext -Config $Config -FolderPath 'documents\test\seg_1' `
        -NotifyDocumentName '0818000063ea501-rev.pdf' -NotifyDocumentGuid $revMember.documentGuid `
        -TransitionSource 'user_audit' -StateTransitionKey 'audit:1' -ChangedByUser 187 `
        -ChangedByUsername 'user@example.com' -AuditEventId 1

    $stemMembers = @(@{
        documentGuid = '05550892-10dd-4e6c-b057-45b200b8a9a8'
        documentName = '0818000063ea501.pdf'
        document = $null
    })
    $allMembers = @(
        $stemMembers[0],
        @{ documentGuid = 'a051c3d3-caf0-4fc1-bc29-da40fd945dc9'; documentName = '0818000063ea501-prod.pdf'; document = $null },
        @{ documentGuid = 'fcbe7d2c-6fa0-47c2-8926-80b4c0908178'; documentName = '0818000063ea501-chk.pdf'; document = $null },
        $revMember
    )
    $laneFromStem = _QCAT-ResolveSheetPackageNotificationMember -Members $stemMembers `
        -TriggerDocumentGuid $stemMembers[0].documentGuid -TriggerDocumentName $stemMembers[0].documentName `
        -Config $Config -FolderPath 'documents\test\seg_1' -SheetStem '0818000063ea501' -AllMembers $allMembers

    @{
        resolvedName = [string]$resolved.documentName
        ctxLaneName = [string]$ctx.notificationLaneDocumentName
        ctxRoleSource = [string]$ctx.roleSourceDocumentName
        ctxProcessType = [string]$ctx.qcProcessType
        laneFromStemName = [string]$laneFromStem.documentName
    }
} -Config $cfg

Assert-Eq $result.resolvedName '0818000063ea501-rev.pdf' 'lane trigger resolves to lane PDF notify member'
Assert-Eq $result.ctxLaneName '0818000063ea501-rev.pdf' 'notify context keeps lane document name'
Assert-Eq $result.ctxRoleSource '0818000063ea501.pdf' 'notify context uses stem PDF for role lookup'
Assert-Eq $result.ctxProcessType 'review' 'notify context includes lane qcProcessType'
Assert-True (Test-QCIsQcPdfDocumentName -DocumentName $result.laneFromStemName) `
    'stem trigger with qcPdfNotificationsOnly resolves to a lane QC PDF'

Write-Host 'test_qc_user_audit_notification_context: OK' -ForegroundColor Green
