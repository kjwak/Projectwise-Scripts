$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot 'Workflow\QC.Workflow.psm1'
if (-not (Test-Path -LiteralPath $target)) {
    throw "Module implementation not found: $target"
}
Import-Module $target -Force -Global
