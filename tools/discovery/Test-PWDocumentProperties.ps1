<#
.SYNOPSIS
Discovers available PW document properties, environment attributes, and cmdlets.
Run this to determine correct property names for StateName and email attributes.
#>
[CmdletBinding()]
param([int]$Hours = 1)

$ErrorActionPreference = 'Continue'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $repoRoot 'modules\Core\Core.Results.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\Core\Core.Runtime.psm1') -Force
Import-Module (Join-Path $repoRoot 'modules\ProjectWise\PW.Connection.psm1') -Force

$config = Get-Content -LiteralPath (Join-Path $repoRoot 'appsettings.json') -Raw | ConvertFrom-Json
$pw = $config.projectWise
$ds = if ($pw.datasourceName) { [string]$pw.datasourceName } elseif ($pw.datasource) { [string]$pw.datasource } else { $null }
$credPath = if ($pw.credentialPath) { [string]$pw.credentialPath } else { 'C:\PW_QC_LOCAL\pw_cred.txt' }

$credRes = Get-PWCredentialFromFile -CredentialPath $credPath
if (-not $credRes.IsSuccess) { Write-Host "Credential failed: $($credRes.Message)" -ForegroundColor Red; return }
$connRes = Connect-PW -DatasourceName $ds -Credential ([pscredential]$credRes.Data.credential)
Write-Host "Connected to PW: $ds" -ForegroundColor Green

# Find a recent document GUID from audit trail
$since = (Get-Date).AddHours(-$Hours).ToString('yyyy-MM-dd HH:mm:ss')
$sql = "SELECT TOP 5 o_objguid, o_itemname, o_action FROM dms_audt WHERE o_acttime >= '$since' AND o_objtype = 2 ORDER BY o_acttime DESC"
$result = Select-PWSQL -SQLSelectStatement $sql -ErrorAction Stop
$rows = @($result.Rows)
if ($rows.Count -eq 0) { Write-Host "No recent audit events found." -ForegroundColor Yellow; return }

$testGuid = [string]$rows[0].o_objguid
$testName = [string]$rows[0].o_itemname
Write-Host "`n=== Testing with document: $testName (GUID: $testGuid) ===" -ForegroundColor Cyan

# 1. Get document via Get-PWDocumentsByGUIDs
Write-Host "`n[1] Get-PWDocumentsByGUIDs properties:" -ForegroundColor Yellow
$doc = Get-PWDocumentsByGUIDs -DocumentGUIDs @($testGuid) -ErrorAction SilentlyContinue
if (-not $doc) { Write-Host "  Document not found by GUID (may be deleted)" -ForegroundColor Red; return }

$props = @($doc.PSObject.Properties | Sort-Object Name)
Write-Host "  Total properties: $($props.Count)" -ForegroundColor Gray
foreach ($p in $props) {
    $val = try { [string]$p.Value } catch { '(error)' }
    if ($val -and $val -ne '' -and $val -ne '0' -and $val -ne '0001-01-01') {
        Write-Host "    $($p.Name) = $val" -ForegroundColor Gray
    }
}

# 2. Link-related properties (for QC email QCPdfUrl resolution)
Write-Host "`n[2] Link-related properties:" -ForegroundColor Yellow
foreach ($name in @('DocumentGUID', 'DocumentGuid', 'DocumentURN', 'ProjectURN', 'ProjectWiseWebLink', 'projectWiseWebLink')) {
    $v = $null
    try { if ($doc.PSObject.Properties[$name]) { $v = $doc.$name } } catch { }
    $display = if ($v) { [string]$v } else { '(null)' }
    $color = if ($v) { 'Green' } else { 'DarkGray' }
    Write-Host "    $name = $display" -ForegroundColor $color
}

# 3. Check specific state-related properties
Write-Host "`n[3] State-related properties:" -ForegroundColor Yellow
foreach ($name in @('StateName','State','StateId','WorkflowName','Workflow','o_stateno','StatusName','DocumentState')) {
    $v = $null
    try { if ($doc.PSObject.Properties[$name]) { $v = $doc.$name } } catch { }
    $display = if ($v) { [string]$v } else { '(null)' }
    $color = if ($v) { 'Green' } else { 'DarkGray' }
    Write-Host "    $name = $display" -ForegroundColor $color
}

# 4. Check if Get-PWRichProperties exists
Write-Host "`n[4] Get-PWRichProperties:" -ForegroundColor Yellow
$richCmd = Get-Command -Name 'Get-PWRichProperties' -ErrorAction SilentlyContinue
if ($richCmd) {
    Write-Host "  Cmdlet EXISTS" -ForegroundColor Green
    try {
        $rich = Get-PWRichProperties -InputDocuments @($doc) -ErrorAction Stop
        $richProps = @($rich.PSObject.Properties | Sort-Object Name)
        Write-Host "  Rich properties count: $($richProps.Count)" -ForegroundColor Gray
        foreach ($rp in $richProps) {
            $val = try { [string]$rp.Value } catch { '(error)' }
            if ($val -and $val -ne '' -and $val -ne '0') {
                Write-Host "    $($rp.Name) = $val" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  Get-PWRichProperties FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  Cmdlet NOT FOUND" -ForegroundColor Red
}

# 5. Check other attribute-related cmdlets
Write-Host "`n[5] Available PW attribute cmdlets:" -ForegroundColor Yellow
$attrCmds = @(Get-Command -Name '*PW*Attribute*','*PW*Rich*','*PW*Env*','*PW*Property*' -ErrorAction SilentlyContinue | Sort-Object Name)
foreach ($c in $attrCmds) {
    Write-Host "    $($c.Name) ($($c.CommandType))" -ForegroundColor Cyan
}

# 6. Try to find email attributes via SQL (environment tables)
Write-Host "`n[6] Searching for EM_Designer_Email via PW SQL:" -ForegroundColor Yellow
$docNo = $null
try { $docNo = $doc.DocumentID } catch { }
if (-not $docNo) { try { $docNo = $doc.o_docno } catch { } }
if ($docNo) {
    Write-Host "  DocumentID (o_docno): $docNo" -ForegroundColor Gray
    # Find which environment table this doc uses
    try {
        $envSql = "SELECT e.o_envname, e.o_tablename FROM dms_env e INNER JOIN dms_doc d ON d.o_envid = e.o_envid WHERE d.o_docno = $docNo"
        $envRes = Select-PWSQL -SQLSelectStatement $envSql -ErrorAction Stop
        $envRows = @($envRes.Rows)
        if ($envRows.Count -gt 0) {
            $envTable = [string]$envRows[0].o_tablename
            $envName = [string]$envRows[0].o_envname
            Write-Host "  Environment: $envName (table: $envTable)" -ForegroundColor Green

            # List columns in the env table
            $colSql = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '$envTable' ORDER BY ORDINAL_POSITION"
            $colRes = Select-PWSQL -SQLSelectStatement $colSql -ErrorAction Stop
            $cols = @($colRes.Rows | ForEach-Object { [string]$_.COLUMN_NAME })
            Write-Host "  Columns ($($cols.Count)):" -ForegroundColor Gray
            foreach ($col in $cols) {
                $isEmail = ($col -match '(?i)email|designer|reviewer')
                $color = if ($isEmail) { 'Green' } else { 'Gray' }
                Write-Host "    $col" -ForegroundColor $color
            }

            # Query actual values for this doc
            $emailCols = @($cols | Where-Object { $_ -match '(?i)email|designer|reviewer|state' })
            if ($emailCols.Count -gt 0) {
                $selectCols = ($emailCols -join ', ')
                $valSql = "SELECT $selectCols FROM $envTable WHERE o_docno = $docNo"
                $valRes = Select-PWSQL -SQLSelectStatement $valSql -ErrorAction Stop
                $valRows = @($valRes.Rows)
                if ($valRows.Count -gt 0) {
                    Write-Host "  Attribute values for doc $docNo`:" -ForegroundColor Yellow
                    foreach ($ec in $emailCols) {
                        $ev = $valRows[0].$ec
                        $evStr = if ($ev -is [DBNull] -or $null -eq $ev) { '(null)' } else { [string]$ev }
                        Write-Host "    $ec = $evStr" -ForegroundColor $(if ($evStr -ne '(null)') { 'Green' } else { 'DarkGray' })
                    }
                } else {
                    Write-Host "  No rows in $envTable for docno $docNo" -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "  No environment assigned to this document." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  SQL query failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  Could not determine DocumentID" -ForegroundColor Red
}

# 7. Check document state via SQL
Write-Host "`n[7] Document state via SQL:" -ForegroundColor Yellow
if ($docNo) {
    try {
        $stateSql = "SELECT d.o_stateno, s.o_statename FROM dms_doc d LEFT JOIN dms_stat s ON d.o_stateno = s.o_stateno WHERE d.o_docno = $docNo"
        $stateRes = Select-PWSQL -SQLSelectStatement $stateSql -ErrorAction Stop
        $stateRows = @($stateRes.Rows)
        if ($stateRows.Count -gt 0) {
            $sno = $stateRows[0].o_stateno
            $sname = $stateRows[0].o_statename
            $snameStr = if ($sname -is [DBNull]) { '(no state)' } else { [string]$sname }
            Write-Host "  o_stateno = $sno, o_statename = $snameStr" -ForegroundColor Green
        }
    } catch {
        Write-Host "  State query failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Disconnect-PW | Out-Null
Write-Host "`n=== Done ===" -ForegroundColor Green
