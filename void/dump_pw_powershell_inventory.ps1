#requires -Version 5.1
<#
Dumps ProjectWise/Bentley PowerShell modules, commands, and parameters.

Outputs:
  - modules.txt
  - commands.txt
  - commands.csv
  - commands.json
  - per-command syntax in commands_syntax.txt
  - per-command parameter details in commands_parameters.txt

Run from 32-bit PowerShell if your PW module requires x86.
#>

$ErrorActionPreference = "Stop"

# ---------- CONFIG ----------
$OutRoot = Join-Path $PWD "pwps_dump_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$IncludeNameRegex = '(?i)pw|projectwise|bentley'
# Optional: hard import of a known pwps module path (uncomment if you want)
# $ForceImportPwpsPath = "C:\Program Files (x86)\Bentley\ProjectWise\bin\PowerShell\pwps\pwps.psd1"
# ---------------------------

New-Item -ItemType Directory -Force -Path $OutRoot | Out-Null

function Write-TextFile([string]$name, [string[]]$lines) {
  $path = Join-Path $OutRoot $name
  $lines | Out-File -FilePath $path -Encoding UTF8
  return $path
}

# ---------- 1) Module discovery ----------
$allModules = Get-Module -ListAvailable |
  Where-Object {
    $_.Name -match $IncludeNameRegex -or
    $_.Path -match $IncludeNameRegex
  } |
  Sort-Object Name, Version -Descending

# Try to import pwps if present (best effort)
try {
  $pwps = $allModules | Where-Object Name -eq 'pwps' | Select-Object -First 1
  if ($pwps) {
    Import-Module $pwps.Path -Force -ErrorAction Stop
  }
} catch {
  # keep going
}

# Optional forced import by explicit path
# if ($ForceImportPwpsPath -and (Test-Path $ForceImportPwpsPath)) {
#   Import-Module $ForceImportPwpsPath -Force
# }

$loadedModules = Get-Module |
  Where-Object { $_.Name -match $IncludeNameRegex -or $_.Path -match $IncludeNameRegex } |
  Sort-Object Name

$moduleLines = @()
$moduleLines += "=== Modules (ListAvailable filter: $IncludeNameRegex) ==="
$moduleLines += ""
$moduleLines += ($allModules | Select-Object Name, Version, Path | Format-Table -AutoSize | Out-String).TrimEnd()
$moduleLines += ""
$moduleLines += "=== Loaded Modules (after imports) ==="
$moduleLines += ""
$moduleLines += ($loadedModules | Select-Object Name, Version, Path | Format-Table -AutoSize | Out-String).TrimEnd()

Write-TextFile "modules.txt" $moduleLines | Out-Null

# ---------- 2) Command discovery ----------
$targetModuleNames = $loadedModules.Name | Select-Object -Unique
if (-not $targetModuleNames -or $targetModuleNames.Count -eq 0) {
  # fallback: include pwps even if it didn't load
  $targetModuleNames = @('pwps')
}

$cmds = foreach ($mn in $targetModuleNames) {
  try { Get-Command -Module $mn -ErrorAction Stop } catch { @() }
}

# Also include commands in the current session that match regex even if module metadata is weird
$cmds2 = Get-Command | Where-Object {
  ($_.Name -match $IncludeNameRegex) -or
  ($_.ModuleName -match $IncludeNameRegex)
}

$allCmds = @($cmds + $cmds2) |
  Where-Object { $_ } |
  Sort-Object ModuleName, Name -Unique

# ---------- 3) Build structured data ----------
$rows = foreach ($c in $allCmds) {
  $meta = [ordered]@{
    Name            = $c.Name
    CommandType     = $c.CommandType.ToString()
    ModuleName      = $c.ModuleName
    Source          = $c.Source
    Version         = $null
    DllOrPath       = $null
    ParameterSets   = @()
  }

  # Module details if available
  $m = $loadedModules | Where-Object Name -eq $c.ModuleName | Select-Object -First 1
  if ($m) {
    $meta.Version = $m.Version.ToString()
    $meta.DllOrPath = $m.Path
  }

  # Parameter set + parameter details
  $psets = @()
  foreach ($ps in $c.ParameterSets) {
    $pinfo = @()
    foreach ($p in $ps.Parameters) {
      $pinfo += [ordered]@{
        Name                = $p.Name
        ParameterType       = $p.ParameterType.FullName
        IsMandatory         = $p.IsMandatory
        Position            = $p.Position
        ValueFromPipeline   = $p.ValueFromPipeline
        ValueFromPipelineByPropertyName = $p.ValueFromPipelineByPropertyName
        ValueFromRemainingArguments     = $p.ValueFromRemainingArguments
        Aliases             = @($p.Aliases)
      }
    }

    $psets += [ordered]@{
      ParameterSetName = $ps.Name
      IsDefault        = $ps.IsDefault
      Parameters       = $pinfo
    }
  }
  $meta.ParameterSets = $psets

  [pscustomobject]$meta
}

# ---------- 4) Human-readable dumps ----------
$cmdLines = @()
$cmdLines += "=== Commands (filtered) ==="
$cmdLines += ""
$cmdLines += ($allCmds | Select-Object Name, CommandType, ModuleName, Source | Format-Table -AutoSize | Out-String).TrimEnd()
Write-TextFile "commands.txt" $cmdLines | Out-Null

# Per-command syntax
$syntaxLines = New-Object System.Collections.Generic.List[string]
$syntaxLines.Add("=== Command Syntax ===")
$syntaxLines.Add("")
foreach ($c in $allCmds) {
  $syntaxLines.Add("----- $($c.ModuleName)\$($c.Name)  ($($c.CommandType)) -----")
  try {
    $syn = (Get-Command $c.Name -ErrorAction Stop).Syntax
    if ($syn) {
      foreach ($s in $syn) { $syntaxLines.Add($s) }
    } else {
      $syntaxLines.Add("(No syntax available)")
    }
  } catch {
    $syntaxLines.Add("ERROR retrieving syntax: $($_.Exception.Message)")
  }
  $syntaxLines.Add("")
}
Write-TextFile "commands_syntax.txt" $syntaxLines | Out-Null

# Per-command parameter details (expanded)
$paramLines = New-Object System.Collections.Generic.List[string]
$paramLines.Add("=== Command Parameters (detailed) ===")
$paramLines.Add("")
foreach ($r in $rows) {
  $paramLines.Add("===== $($r.ModuleName)\$($r.Name)  [$($r.CommandType)] =====")
  if (-not $r.ParameterSets -or $r.ParameterSets.Count -eq 0) {
    $paramLines.Add("(No parameter set info)")
    $paramLines.Add("")
    continue
  }

  foreach ($ps in $r.ParameterSets) {
    $paramLines.Add("  ParameterSet: $($ps.ParameterSetName)  Default=$($ps.IsDefault)")
    foreach ($p in $ps.Parameters) {
      $paramLines.Add(("    -{0}  Type={1}  Mandatory={2}  Pos={3}  Pipe={4}  PipeByProp={5}  RemainingArgs={6}  Aliases={7}" -f `
        $p.Name, $p.ParameterType, $p.IsMandatory, $p.Position, $p.ValueFromPipeline, $p.ValueFromPipelineByPropertyName, `
        $p.ValueFromRemainingArguments, ($p.Aliases -join ',')))
    }
  }
  $paramLines.Add("")
}
Write-TextFile "commands_parameters.txt" $paramLines | Out-Null

# ---------- 5) Machine-readable exports ----------
$csvPath  = Join-Path $OutRoot "commands.csv"
$jsonPath = Join-Path $OutRoot "commands.json"

# Flatten a CSV-friendly view (one row per command; parameter sets omitted)
$rows |
  Select-Object Name, CommandType, ModuleName, Source, Version, DllOrPath |
  Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

$rows | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8

"`nDONE. Output folder: $OutRoot"
"modules.txt, commands.txt, commands_syntax.txt, commands_parameters.txt, commands.csv, commands.json"
