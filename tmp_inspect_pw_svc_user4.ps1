$ErrorActionPreference = 'Stop'
Import-Module pwps -Force
Import-Module pwps_dab -Force

$credPath = 'C:\PW_QC_LOCAL\pw_cred.txt'
$map = @{}
Get-Content -LiteralPath $credPath | ForEach-Object {
  $line = $_.Trim()
  if (-not $line -or $line.StartsWith('#')) { return }
  $i = $line.IndexOf('=')
  if ($i -lt 1) { return }
  $map[$line.Substring(0, $i).Trim().ToLowerInvariant()] = $line.Substring($i + 1).Trim()
}
$sec = ConvertTo-SecureString $map['password'] -AsPlainText -Force
$ds = 'typsa-us-pw.bentley.com:typsa-us-pw-03'
Open-PWConnection -DatasourceName $ds -UserName $map['username'] -Password $sec -Admin -WarningAction SilentlyContinue | Out-Null

try {
  Write-Output '=== Full General user settings for SVC_TYPSA_Archivist ==='
  Get-PWUserSetting -UserID 790 -UserSettingCategory General | Format-Table UserSettingName, UserSettingValue -AutoSize | Out-String | Write-Output

  Write-Output '=== UI settings (dialogs) ==='
  Get-PWUserSetting -UserID 790 -UserSettingCategory UI |
    Where-Object { $_.UserSettingName -match 'Dialog|Progress|LocalDoc|Error' } |
    Format-Table UserSettingName, UserSettingValue -AutoSize | Out-String | Write-Output

  Write-Output '=== User.Settings.All credential-related ==='
  $usr = Get-PWUser -UserName 'SVC_TYPSA_Archivist'
  $usr.Settings.All.GetEnumerator() | Where-Object {
    $_.Key -match '(?i)cred|expir|pass|login|dialog|auth|session' -or $_.Value -match '(?i)cred|expir|999'
  } | ForEach-Object { Write-Output ("{0} = {1}" -f $_.Key, $_.Value) }

  Write-Output ''
  Write-Output '=== Set-PWUserSetting / Update-PWUserSetting parameter info ==='
  foreach ($name in @('Set-PWUserSetting','Update-PWUserSetting')) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if (-not $c) { Write-Output ("(missing) {0}" -f $name); continue }
    Write-Output ("--- {0} ---" -f $name)
    foreach ($ps in $c.ParameterSets) {
      Write-Output ("  [{0}] {1}" -f $ps.Name, (($ps.Parameters | Where-Object { -not $_.Name.StartsWith('Verbose') -and $_.Name -notmatch 'Action|Variable|Buffer|Debug|Warning|Information|Pipeline|OutVariable' }) | ForEach-Object Name) -join ', ')
    }
  }

  # Help text if any
  Write-Output ''
  Write-Output '=== Help snippets for CredentialExpirationPolicy ==='
  try {
    $h = Get-Help Set-PWUserSetting -ErrorAction SilentlyContinue
    if ($h) { ($h | Out-String).Substring(0, [Math]::Min(2000, ($h | Out-String).Length)) | Write-Output }
  } catch {}

  Write-Output ''
  Write-Output '=== Compare a typical federated user vs service account (sample) ==='
  # Find any user with IdentityProvider set
  try {
    $sample = Get-PWUser | Where-Object { $_.IdentityProvider -or $_.Identity -or $_.Type -ne 'Logical' } | Select-Object -First 5
    Write-Output ("non-logical/federated sample count returned: {0}" -f @($sample).Count)
    $sample | Select-Object UserName, Type, IdentityProvider, Identity, IsDisabled | Format-Table -AutoSize | Out-String | Write-Output
  } catch {
    Write-Output ("enumerate users error: {0}" -f $_.Exception.Message)
  }
}
finally {
  Close-PWConnection -ErrorAction SilentlyContinue | Out-Null
}
