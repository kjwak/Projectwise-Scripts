<#
.SYNOPSIS
Renders the QC HTML email template and optionally sends a test message via Microsoft Graph.

.DESCRIPTION
Loads email/sample_notification_data.json, renders HTML with ConvertTo-QCEmailHtml,
builds a Graph sendMail payload with New-QCGraphEmailMessage, and writes preview files
under output/. Use -Send -Live to deliver a test email (default To: jflint@aztec.us; requires appsettings.secrets.json).

.EXAMPLE
.\scripts\Test-QCEmailTemplate.ps1

.EXAMPLE
.\scripts\Test-QCEmailTemplate.ps1 -Send -Live
#>
[CmdletBinding()]
param(
    [string]$To = 'jflint@aztec.us',
    [string[]]$Cc = @(),
    [string]$AppSettingsPath = '',
    [string]$SampleDataPath = '',
    [string]$TemplatePath = 'email/templates/qc_notification.html',
    [string]$LogoPath = 'email/typsalogo.png.webp',
    [string]$OutputDir = 'output',
    [switch]$Send,
    [switch]$Live,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    throw 'PSScriptRoot is empty. Run from the repo root.'
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
if ([string]::IsNullOrWhiteSpace($AppSettingsPath)) {
    $AppSettingsPath = Join-Path $repoRoot 'appsettings.json'
}
if ([string]::IsNullOrWhiteSpace($SampleDataPath)) {
    $SampleDataPath = Join-Path $repoRoot 'email\sample_notification_data.json'
}

function _TestEmail-GetGraphSettings([hashtable]$Config) {
    if (-not $Config -or -not $Config.ContainsKey('notifications')) { return $null }
    $notifications = $Config['notifications']
    if (-not $notifications) { return $null }
    if ($notifications -isnot [hashtable] -and (Get-Command -Name 'ConvertTo-HashtableDeep' -ErrorAction SilentlyContinue)) {
        $notifications = ConvertTo-HashtableDeep -Value $notifications
    }
    if (-not $notifications -or -not $notifications.ContainsKey('graph')) { return $null }
    $graph = $notifications['graph']
    if (-not $graph) { return $null }
    if ($graph -isnot [hashtable] -and (Get-Command -Name 'ConvertTo-HashtableDeep' -ErrorAction SilentlyContinue)) {
        $graph = ConvertTo-HashtableDeep -Value $graph
    }
    return $graph
}

foreach ($mod in @('Core.Results.psm1', 'Core.Runtime.psm1', 'QC.NotificationGraph.psm1', 'QC.Notifications.psm1')) {
    $modPath = Join-Path $repoRoot "modules\$mod"
    if (-not (Test-Path -LiteralPath $modPath)) { throw "Module not found: $modPath" }
    Import-Module $modPath -Force
}
Import-Module (Join-Path $repoRoot 'modules\QC.NotificationGraph.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\QC.NotificationTemplates.psm1') -Force

if (-not (Test-Path -LiteralPath $SampleDataPath)) {
    throw "Sample data not found: $SampleDataPath"
}

$rawSample = Get-Content -LiteralPath $SampleDataPath -Raw -Encoding UTF8
$sampleObj = $rawSample | ConvertFrom-Json -ErrorAction Stop
$sampleData = @{}
foreach ($p in $sampleObj.PSObject.Properties) {
    $sampleData[$p.Name] = if ($null -eq $p.Value) { '' } else { [string]$p.Value }
}
if ($sampleData.Count -eq 0) { throw 'Sample data is empty.' }

$html = ConvertTo-QCEmailHtml -TemplatePath $TemplatePath -Data $sampleData
$subject = if ($sampleData.NotificationTitle) { [string]$sampleData.NotificationTitle } else { 'QC Email Template Test' }

$outDir = Join-Path $repoRoot $OutputDir
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$htmlPath = Join-Path $outDir 'test_qc_email.html'
$html | Set-Content -LiteralPath $htmlPath -Encoding UTF8 -ErrorAction Stop

$graphBody = New-QCGraphEmailMessage -ToRecipients @('preview@example.com') -Subject $subject `
    -HtmlBody $html -LogoPath $LogoPath -CcRecipients @()
$jsonPath = Join-Path $outDir 'test_qc_graph_payload.json'
($graphBody | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $jsonPath -Encoding UTF8 -ErrorAction Stop

Write-Host "Repo root:     $repoRoot" -ForegroundColor Cyan
Write-Host "HTML preview:  $htmlPath" -ForegroundColor Green
Write-Host "Graph payload: $jsonPath" -ForegroundColor Green

if (-not $NoOpen) {
    try {
        Start-Process -FilePath $htmlPath -ErrorAction Stop
        Write-Host 'Opened HTML preview in default browser.' -ForegroundColor Gray
    } catch {
        Write-Host "Could not open HTML automatically: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($Send) {
    if ([string]::IsNullOrWhiteSpace($To)) {
        throw '-To is required when using -Send (default is jflint@aztec.us).'
    }
    Import-Module (Join-Path $repoRoot 'modules\Core.Results.psm1') -Force
    Import-Module (Join-Path $repoRoot 'modules\Core.Runtime.psm1') -Force
    $configResult = Read-QCAppSettings -Path $AppSettingsPath
    if (-not $configResult.IsSuccess) {
        throw ('Failed to load appsettings: ' + $configResult.Message)
    }
    $config = $configResult.Data.config
    if ($configResult.Data.mergeChain) {
        Write-Host ('Config layers: ' + ($configResult.Data.mergeChain -join '; ')) -ForegroundColor Cyan
    }

    $graph = _TestEmail-GetGraphSettings -Config $config
    if (-not $graph) {
        throw 'notifications.graph missing after merge; ensure appsettings.secrets.json exists beside appsettings.json.'
    }

    $validation = Test-QCNotificationGraphConfigured -GraphSettings $graph
    if (-not $validation.configured) {
        throw ('Graph not configured. Missing: ' + ($validation.missing -join ', '))
    }

    $notifications = $config['notifications']
    if ($notifications -isnot [hashtable] -and (Get-Command -Name 'ConvertTo-HashtableDeep' -ErrorAction SilentlyContinue)) {
        $notifications = ConvertTo-HashtableDeep -Value $notifications
    }
    $dryRun = $true
    if ($notifications -and $notifications.ContainsKey('dryRun')) { $dryRun = [bool]$notifications['dryRun'] }
    if ($Live) { $dryRun = $false }

    $payload = @{
        eventType = 'EMAIL_TEMPLATE_TEST'
        documentName = if ($sampleData.DocumentName) { [string]$sampleData.DocumentName } else { 'test.pdf' }
        subject = $subject
        body = 'HTML template test (see htmlBody).'
        htmlBody = $html
        logoPath = $LogoPath
        to = @($To.Trim())
        cc = @($Cc | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    Write-Host "Sending to: $($payload.to -join ', ')  DryRun: $dryRun" -ForegroundColor Cyan
    $result = Send-QCNotificationGraph -GraphSettings $graph -Payload $payload -DryRun:$dryRun
    Write-Host ($result | ConvertTo-Json -Depth 8)
    if (-not $result.IsSuccess) { throw $result.Message }
    if ($dryRun) {
        Write-Host 'Dry run OK. Re-run with -Send -Live to send a real email.' -ForegroundColor Green
    } else {
        Write-Host 'Graph sendMail completed.' -ForegroundColor Green
    }
}
