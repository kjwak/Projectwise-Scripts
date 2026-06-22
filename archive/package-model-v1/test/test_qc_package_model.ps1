$ErrorActionPreference = 'Stop'
$archiveRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $archiveRoot)
Import-Module "$archiveRoot/modules/QC.PackageResolver.psm1" -Force -Global
Import-Module "$archiveRoot/modules/QC.AttributePolicy.psm1" -Force -Global
Import-Module "$archiveRoot/modules/QC.StatePolicy.psm1" -Force -Global
Import-Module "$archiveRoot/modules/QC.PackageSync.psm1" -Force -Global
Import-Module "$archiveRoot/modules/QC.Package.Database.psm1" -Force -Global

function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw "ASSERT FAILED: $Message" } }
function Assert-Eq($Actual, $Expected, [string]$Message) { if ($Actual -ne $Expected) { throw "ASSERT FAILED: $Message`nExpected: $Expected`nActual:   $Actual" } }
function New-Doc([string]$Name,[string]$Guid,[string]$State='In Production',[hashtable]$Attrs=@{}) {
    [pscustomobject]@{ Name=$Name; Guid=$Guid; FolderPath='Documents\Proj\CADD\Sheets'; WorkflowState=$State; Environment='Caltrans'; Attributes=$Attrs }
}
$config = @{ qcPackage = @{ enabled=$true; canonicalDocumentRole='ProductionPdf'; fallbackOrder=@('ProductionPdf','QcPdf','Dgn'); namingRules=@{ dgnExtensions=@('.dgn'); productionPdfExtensions=@('.pdf'); qcPdfSuffixes=@('_QC','-QC','_QC_HISTORY') }; statePrecedence=@('Error Needs Attention','QC Initiated','QC Finalizing','Ready for QC','Corrections Received','Redlines Received','In Production'); userOwnedAttributes=@('QC_Review_Type','QC_Designer_Email','QC_Reviewer_Email','QC_Checker_Email','QC_Originator','QC_Due_Date'); automationOwnedAttributes=@('QC_Prepend_Status','QC_Last_Automation_Run','QC_Last_Email_Sent','QC_Automation_JobId','QC_Error_Message','QC_PackageId','QC_Source_Dgn_Guid','QC_Production_Pdf_Guid','QC_QcPdf_Guid'); syncUserAttributesFromNonCanonical=$true; writeAutomationStatusTo=@('ProductionPdf','QcPdf') } }
$dgn = New-Doc 'A101.dgn' 'dgn-1' 'In Production' @{ QC_Review_Type='Peer Review'; QC_Reviewer_Email='reviewer@example.com' }
$pdf = New-Doc 'A101.pdf' 'pdf-1' 'QC Received' @{ QC_Review_Type='Production QC'; QC_Reviewer_Email='lead@example.com' }
$qc  = New-Doc 'A101_QC.pdf' 'qc-1' 'QC Initiated' @{}
$docs = @($dgn,$pdf,$qc)

foreach($doc in $docs){
    $r = Resolve-QCPackage -Document $doc -CandidateDocuments $docs -Config $config -DryRun
    Assert-True $r.IsSuccess "$($doc.Name) audit event resolves package"
    Assert-Eq $r.Data.package.PackageId (Resolve-QCPackage -Document $pdf -CandidateDocuments $docs -Config $config).Data.package.PackageId 'All related audit events resolve same package id'
}
$pkg = (Resolve-QCPackage -Document $dgn -CandidateDocuments $docs -Config $config).Data.package
$canon = Get-QCPackageCanonicalDocument -Package $pkg -Config $config
Assert-True $canon.IsSuccess 'Canonical selection should succeed'
Assert-Eq $canon.Data.role 'ProductionPdf' 'Production PDF selected as canonical'

$missingPdfPkg = (Resolve-QCPackage -Document $qc -CandidateDocuments @($dgn,$qc) -Config $config).Data.package
$fallback = Get-QCPackageCanonicalDocument -Package $missingPdfPkg -Config $config
Assert-Eq $fallback.Data.role 'QcPdf' 'Missing production PDF falls back predictably to QC PDF'
Assert-True ($fallback.Data.warnings.Count -gt 0) 'Fallback should warn when production PDF is missing'

$state = Resolve-QCPackageState -Package $pkg -Config $config
Assert-True $state.Data.conflict 'Conflicting states are detected'
Assert-Eq $state.Data.state 'QC Initiated' 'Configured state precedence selects highest-priority present state'
$conf = Test-QCPackageStateConflict -Package $pkg -Config $config
Assert-True $conf.Data.hasConflict 'Conflict test reports conflict'

$script:writeCalls = 0
function Update-PWDocumentAttributes { param($InputDocument, [hashtable]$Attributes) $script:writeCalls++ }
$sync = Sync-QCPackageAttributes -Package $pkg -Config $config -AutomationValues @{ QC_Prepend_Status='Done'; QC_Error_Message='' } -DryRun
Assert-True $sync.IsSuccess 'Attribute sync dry-run should succeed'
Assert-Eq $script:writeCalls 0 'Dry-run performs no attribute mutation'
$userWrites = @($sync.Data.writes | Where-Object { $_.Ownership -eq 'UserOwned' })
Assert-Eq $userWrites.Count 0 'Canonical user-owned attributes are not overwritten blindly when already populated'
$autoWrites = @($sync.Data.writes | Where-Object { $_.Ownership -eq 'AutomationOwned' })
Assert-Eq $autoWrites.Count 2 'Automation-owned attributes written only to configured documents'
Assert-True (($autoWrites.Role -contains 'ProductionPdf') -and ($autoWrites.Role -contains 'QcPdf')) 'Automation writes target PDF and QC PDF'
Remove-Item function:\Update-PWDocumentAttributes -ErrorAction SilentlyContinue

$cache = Write-QCPackageCache -Package $pkg -Config $config -DryRun
Assert-Eq $cache.Code 'PACKAGE_CACHE_DRY_RUN' 'Dry-run does not mutate SQL package cache'

function Get-PWEnvironmentColumns { param([string]$EnvironmentName) @([pscustomobject]@{Name='QC_Review_Type'},[pscustomobject]@{Name='QC_Reviewer_Email'},[pscustomobject]@{Name='QC_Prepend_Status'}) }
$invalid = Test-QCPackageAttributeValidity -Document (New-Doc 'B101.pdf' 'bad-1' 'In Production' @{ QC_Reviewer_Email='not-an-email' }) -Config $config -AttributeNames @('QC_Review_Type','QC_Reviewer_Email','QC_Missing')
Assert-True (-not $invalid.IsSuccess) 'Invalid/missing environment attributes are reported clearly'
Assert-True (($invalid.Data.errors -join '|') -match 'QC_Missing') 'Missing attribute is named in validation errors'
Assert-True (($invalid.Data.errors -join '|') -match 'Invalid email') 'Invalid email is named in validation errors'
Remove-Item function:\Get-PWEnvironmentColumns -ErrorAction SilentlyContinue

Write-Host 'All QC package model tests passed.' -ForegroundColor Green
