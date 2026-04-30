# Shared helpers for STA -> MTA relaunch via powershell.exe -File (native exe: no hashtable splat of PSBoundParameters).
# Dot-source from scripts that need pwps_dab MTA: . "$PSScriptRoot\StaMtaRelaunch.ps1"

# Child -File processes receive bools as "-Name:1" / "-Name:0" (strings); [bool] params then fail to bind.
function ConvertTo-BoolLoose {
  param([Parameter(Mandatory = $true)][object] $Value)
  if ($null -eq $Value) { return $false }
  if ($Value -is [bool]) { return $Value }
  $s = [string]$Value
  if ($s -match '^(?i)(true|1|yes)$') { return $true }
  if ($s -match '^(?i)(false|0|no)$') { return $false }
  return $false
}

function Build-PowerShellExeFileArgs {
  param(
    [Parameter(Mandatory = $true)]
    [string] $ScriptPath,
    [Parameter(Mandatory = $true)]
    [hashtable] $BoundParameters
  )
  $list = New-Object System.Collections.ArrayList
  [void]$list.AddRange([string[]]@('-MTA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath))
  foreach ($name in @($BoundParameters.Keys)) {
    $val = $BoundParameters[$name]
    if ($null -eq $val) { continue }
    if ($val -is [System.Management.Automation.SwitchParameter]) {
      if ($val.IsPresent) { [void]$list.Add("-$name") }
      continue
    }
    if ($val -is [bool]) {
      # Single argv token so the child parser binds [bool] reliably (-Name 1 as two tokens can still coerce wrong).
      [void]$list.Add("-${name}:$(if ($val) { '1' } else { '0' })")
      continue
    }
    if ($val -is [System.Array]) {
      [void]$list.Add("-$name")
      foreach ($item in $val) { [void]$list.Add([string]$item) }
      continue
    }
    [void]$list.Add("-$name")
    [void]$list.Add([string]$val)
  }
  return [string[]]$list.ToArray()
}
