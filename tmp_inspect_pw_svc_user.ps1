$ErrorActionPreference = 'Stop'
Import-Module pwps -Force
Import-Module pwps_dab -Force

Write-Output '=== Open-PWConnection parameter sets ==='
$cmd = Get-Command Open-PWConnection
foreach ($ps in $cmd.ParameterSets) {
  $params = ($ps.Parameters | ForEach-Object {
    $n = $_.Name
    if ($_.IsMandatory) { "$n*" } else { $n }
  }) -join ', '
  Write-Output ("  [{0}] {1}" -f $ps.Name, $params)
}

Write-Output ''
Write-Output '=== Related user/identity cmdlets ==='
Get-Command -Module pwps,pwps_dab | Where-Object {
  $_.Name -match 'PWUser|Identity|Login|Session|Credential'
} | Select-Object Name, ModuleName | Sort-Object Name | Format-Table -AutoSize

$credPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
if (-not (Test-Path -LiteralPath $credPath)) {
  Write-Output "Credential file missing: $credPath"
  exit 1
}

# Parse key=value cred file without echoing secrets
$map = @{}
Get-Content -LiteralPath $credPath | ForEach-Object {
  $line = $_.Trim()
  if (-not $line -or $line.StartsWith('#')) { return }
  $i = $line.IndexOf('=')
  if ($i -lt 1) { return }
  $k = $line.Substring(0, $i).Trim().ToLowerInvariant()
  $v = $line.Substring($i + 1).Trim()
  $map[$k] = $v
}
$userName = if ($map['username']) { $map['username'] } elseif ($map['user']) { $map['user'] } else { $null }
$pwdPlain = if ($map['password']) { $map['password'] } else { $null }
if (-not $userName -or -not $pwdPlain) { throw 'credential file missing username/password keys' }

Write-Output ("Credential file userName: {0}" -f $userName)
Write-Output ("Password present: {0} (len={1})" -f (-not [string]::IsNullOrWhiteSpace($pwdPlain)), $pwdPlain.Length)

$ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
$sec = ConvertTo-SecureString $pwdPlain -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($userName, $sec)

Write-Output ''
Write-Output '=== Connecting with UserName/Password (same as Connect-PW) ==='
try {
  Open-PWConnection -DatasourceName $ds -UserName $cred.UserName -Password $cred.Password -WarningAction SilentlyContinue | Out-Null
  Write-Output 'Connect OK'
} catch {
  Write-Output ("Connect FAILED: {0}" -f $_.Exception.Message)
  exit 2
}

try {
  Write-Output ''
  Write-Output '=== Current connection / login info (best-effort) ==='
  foreach ($name in @('Get-PWLogin','Get-PWConnection','Get-PWCurrentUser','Get-PWLoggedInUser','Get-PWDatasource')) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if (-not $c) { continue }
    Write-Output ("--- {0} ---" -f $name)
    try {
      & $name | Format-List * | Out-String | Write-Output
    } catch {
      Write-Output ("  error: {0}" -f $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Get-PWUser for service account ==='
  $candidates = @($userName, 'SVC_TYPSA_Archivist', 'SVC_TYPSA_ARCHIVIST', 'srv_typsa_archivist') | Select-Object -Unique
  foreach ($u in $candidates) {
    Write-Output ("-- lookup: {0} --" -f $u)
    try {
      $users = @(Get-PWUser -UserName $u -ErrorAction Stop)
      if ($users.Count -eq 0) {
        Write-Output '  (no rows)'
        continue
      }
      foreach ($usr in $users) {
        $props = $usr.PSObject.Properties.Name
        Write-Output ("  properties: {0}" -f ($props -join ', '))
        # Dump non-secret-ish properties
        $interesting = @(
          'Name','UserName','Description','Email','Type','UserType','Disabled','IsDisabled',
          'Identity','Identities','WindowsUser','Domain','CreateDate','LastLogin','PasswordExpires',
          'PasswordNeverExpires','PasswordExpiration','ExpirePassword','DaysToExpire','MaxFailures',
          'UserID','ID','Guid','GUID','Federated','Provider','AuthType','AuthenticationType',
          'CanChangePassword','MustChangePassword','AccountExpires','Flags','UserFlags'
        )
        foreach ($p in $interesting) {
          if ($props -contains $p) {
            $val = $usr.$p
            if ($null -eq $val) { continue }
            if ($val -is [securestring]) { continue }
            Write-Output ("  {0} = {1}" -f $p, ($val | Out-String).Trim())
          }
        }
        # Also print any property matching expire/identity/auth/disable/pass
        foreach ($p in $props) {
          if ($p -match '(?i)expir|ident|auth|disabl|pass|federat|provider|window|sso|ims|flag|type|login|lock') {
            if ($interesting -contains $p) { continue }
            try {
              $val = $usr.$p
              if ($null -eq $val) { continue }
              if ($val -is [securestring]) { continue }
              $s = ($val | Out-String).Trim()
              if ($s.Length -gt 300) { $s = $s.Substring(0,300) + '...' }
              Write-Output ("  {0} = {1}" -f $p, $s)
            } catch {}
          }
        }
      }
    } catch {
      Write-Output ("  error: {0}" -f $_.Exception.Message)
    }
  }

  Write-Output ''
  Write-Output '=== Identity provider / Windows identity lookups ==='
  foreach ($cmdName in @('Get-PWUserIdentityProvider','Get-PWUserIdentityProviders','Get-PWUserWindowsIdentity','Get-PWUserIdentities','Get-PWIdentityProvider')) {
    $c = Get-Command $cmdName -ErrorAction SilentlyContinue
    if (-not $c) { Write-Output ("(missing) {0}" -f $cmdName); continue }
    Write-Output ("--- {0} ---" -f $cmdName)
    try {
      $all = @(& $cmdName -ErrorAction Stop)
      Write-Output ("  count={0}" -f $all.Count)
      $all | Select-Object -First 30 | ForEach-Object {
        ($_ | Format-List * | Out-String).Trim()
      } | Write-Output
    } catch {
      # try filtered by username
      try {
        $filtered = @(& $cmdName -UserName $userName -ErrorAction Stop)
        Write-Output ("  filtered count={0}" -f $filtered.Count)
        $filtered | Select-Object -First 20 | ForEach-Object {
          ($_ | Format-List * | Out-String).Trim()
        } | Write-Output
      } catch {
        Write-Output ("  error: {0}" -f $_.Exception.Message)
      }
    }
  }

  # Try identity on specific user object methods
  Write-Output ''
  Write-Output '=== Get-PWUserIdentityProvider parameter sets (if present) ==='
  $idCmd = Get-Command Get-PWUserIdentityProvider -ErrorAction SilentlyContinue
  if ($idCmd) {
    foreach ($ps in $idCmd.ParameterSets) {
      $params = ($ps.Parameters | ForEach-Object { $_.Name }) -join ', '
      Write-Output ("  [{0}] {1}" -f $ps.Name, $params)
    }
  }
}
finally {
  try { Close-PWConnection -ErrorAction SilentlyContinue | Out-Null } catch {}
  Write-Output ''
  Write-Output 'Disconnected.'
}
