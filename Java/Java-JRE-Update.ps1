#Requires -Version 5.1
<#
.SYNOPSIS
  Checks Java JRE (for a chosen major family) and updates via winget if out-of-date.
  Optional removal of older JREs and cleanup of env vars / PATH.

.PARAMETER TargetFamily
  Java major family to manage: 8, 11, 17, 21

.PARAMETER Vendor
  'Temurin' (default) or 'Oracle'
  Note: Oracle JRE is generally only available for Java 8 via winget.

.PARAMETER ReportOnly
  Dry mode: report status only, no installs/uninstalls/cleanup.

.PARAMETER RemoveOlder
  Attempts to uninstall older JRE installs of the SAME family.
  Works best for MSI-based installs (QuietUninstallString or msiexec).

.PARAMETER Cleanup
  Removes stale JAVA_HOME/JRE_HOME and PATH entries that point to missing folders.

.PARAMETER Force
  If set, proceeds even if java/javaw processes are running (not recommended).

.EXAMPLE
  .\Update-JavaJRE.ps1 -TargetFamily 17 -ReportOnly

.EXAMPLE
  .\Update-JavaJRE.ps1 -TargetFamily 17 -RemoveOlder -Cleanup

.EXAMPLE
  .\Update-JavaJRE.ps1 -TargetFamily 8 -Vendor Oracle -RemoveOlder -Cleanup
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
  [ValidateSet('Temurin','Oracle')]
  [string]$Vendor = 'Temurin',

  [ValidateSet(8,11,17,21)]
  [int]$TargetFamily = 17,

  [switch]$ReportOnly,
  [switch]$RemoveOlder,
  [switch]$Cleanup,
  [switch]$Force,

  [string]$LogPath = "$env:ProgramData\JavaUpdate\JavaJRE-Update.log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------------------- helpers --------------------
function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-LogFolder {
  $dir = Split-Path -Parent $LogPath
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

function Normalize-JavaVersion {
  param([Parameter(Mandatory)][string]$VersionString)

  $v = $VersionString.Trim()

  # java.runtime.version: 1.8.0_481-b10
  if ($v -match '^1\.8\.0_(\d+)(?:-b(\d+))?$') {
    $upd = [int]$Matches[1]
    $bld = if ($Matches[2]) { [int]$Matches[2] } else { 0 }
    return [pscustomobject]@{
      Raw=$v; Major=8; Minor=0; SecOrUpd=$upd; Build=$bld
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f 8,0,$upd,$bld
    }
  }

  # 17.0.10+7
  if ($v -match '^(\d+)\.(\d+)\.(\d+)\+(\d+)$') {
    return [pscustomobject]@{
      Raw=$v; Major=[int]$Matches[1]; Minor=[int]$Matches[2]; SecOrUpd=[int]$Matches[3]; Build=[int]$Matches[4]
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3],[int]$Matches[4]
    }
  }

  # 17.0.10.7
  if ($v -match '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
    return [pscustomobject]@{
      Raw=$v; Major=[int]$Matches[1]; Minor=[int]$Matches[2]; SecOrUpd=[int]$Matches[3]; Build=[int]$Matches[4]
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3],[int]$Matches[4]
    }
  }

  # "8 Update 451"
  if ($v -match '^\s*8\s+Update\s+(\d+)\s*$') {
    $upd = [int]$Matches[1]
    return [pscustomobject]@{
      Raw=$v; Major=8; Minor=0; SecOrUpd=$upd; Build=0
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f 8,0,$upd,0
    }
  }

  throw "Cannot normalize Java version string: '$VersionString'"
}

function Try-Normalize {
  param([string]$VersionString)
  if (-not $VersionString) { return $null }
  try { return Normalize-JavaVersion -VersionString $VersionString } catch { return $null }
}

function Get-JavaSettings {
  param([Parameter(Mandatory)][string]$JavaExePath)

  $out = & $JavaExePath -XshowSettings:properties -version 2>&1
  $runtime = ($out | Where-Object { $_ -match '^\s*java\.runtime\.version\s*=' } | Select-Object -First 1)
  $home    = ($out | Where-Object { $_ -match '^\s*java\.home\s*=' } | Select-Object -First 1)

  $runtimeVer = if ($runtime) { (($runtime -split '=',2)[1]).Trim() } else { $null }
  $javaHome   = if ($home)    { (($home    -split '=',2)[1]).Trim() } else { $null }

  [pscustomobject]@{ RuntimeVersion = $runtimeVer; JavaHome = $javaHome }
}

function Get-ImageTypeFromHome {
  param([string]$JavaHome)
  if (-not $JavaHome) { return 'unknown' }
  $h = $JavaHome.ToLowerInvariant()
  if ($h -match '\\jdk-|\\jdk\\' -or $h -match '\\java\\jdk') { return 'jdk' }
  if ($h -match '\\jre-|\\jre\\' -or $h -match '\\java\\jre') { return 'jre' }
  # Temurin commonly uses jre- / jdk- folders; Oracle 8 uses jre1.8.0_XXX
  if ($h -match '\\jre1\.8\.0_') { return 'jre' }
  if ($h -match '\\jdk1\.8\.0_') { return 'jdk' }
  return 'unknown'
}

function Find-JavaCandidates {
  $candidates = New-Object System.Collections.Generic.HashSet[string]

  if ($env:JAVA_HOME) {
    $p = Join-Path $env:JAVA_HOME 'bin\java.exe'
    if (Test-Path $p) { [void]$candidates.Add($p) }
  }

  $cmd = Get-Command java.exe -ErrorAction SilentlyContinue
  if ($cmd -and (Test-Path $cmd.Source)) { [void]$candidates.Add($cmd.Source) }

  $roots = @(
    "$env:ProgramFiles\Java",
    "$env:ProgramFiles\Eclipse Adoptium",
    "$env:ProgramFiles (x86)\Java",
    "$env:ProgramFiles (x86)\Eclipse Adoptium"
  )

  foreach ($r in $roots) {
    if (-not (Test-Path $r)) { continue }
    Get-ChildItem -Path $r -Directory -ErrorAction SilentlyContinue | ForEach-Object {
      $p = Join-Path $_.FullName 'bin\java.exe'
      if (Test-Path $p) { [void]$candidates.Add($p) }
    }
  }

  return $candidates.ToArray()
}

function Get-InstalledJavaForFamilyAndType {
  param(
    [Parameter(Mandatory)][int]$Family,
    [Parameter(Mandatory)][ValidateSet('jre','jdk')] [string]$ImageType
  )

  $best = $null
  $all  = @()

  foreach ($javaExe in (Find-JavaCandidates)) {
    $settings = $null
    try { $settings = Get-JavaSettings -JavaExePath $javaExe } catch { continue }
    if (-not $settings.RuntimeVersion) { continue }

    $norm = Try-Normalize -VersionString $settings.RuntimeVersion
    if (-not $norm) { continue }

    $type = Get-ImageTypeFromHome -JavaHome $settings.JavaHome
    $entry = [pscustomobject]@{
      JavaExe  = $javaExe
      JavaHome = $settings.JavaHome
      Type     = $type
      Version  = $settings.RuntimeVersion
      Norm     = $norm
    }
    $all += $entry

    if ($norm.Major -eq $Family -and $type -eq $ImageType) {
      if (-not $best -or $norm.Key -gt $best.Norm.Key) { $best = $entry }
    }
  }

  [pscustomobject]@{ Best = $best; All = $all }
}

function Get-WingetId {
  param(
    [Parameter(Mandatory)][string]$Vendor,
    [Parameter(Mandatory)][int]$Family
  )

  if ($Vendor -eq 'Temurin') {
    return "EclipseAdoptium.Temurin.$Family.JRE"
  }

  # Oracle: typically only JRE 8
  if ($Vendor -eq 'Oracle') {
    if ($Family -ne 8) { throw "Oracle JRE via winget is typically only available for Java 8. Use -TargetFamily 8 or switch Vendor to Temurin." }
    return "Oracle.JavaRuntimeEnvironment"
  }

  throw "Unknown vendor: $Vendor"
}

function Get-WingetAvailableVersion {
  param([Parameter(Mandatory)][string]$WingetId)

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { throw "winget.exe not found. Install 'App Installer' or use a different deployment method." }

  # Try JSON output first (newer winget)
  try {
    $json = & $winget.Source show --exact --id $WingetId --output json 2>$null
    if ($LASTEXITCODE -eq 0 -and $json) {
      $obj = $json | ConvertFrom-Json
      if ($obj.Version) { return [string]$obj.Version }
      if ($obj.Versions -and $obj.Versions[0]) { return [string]$obj.Versions[0] }
    }
  } catch { }

  # Fallback: parse text
  $txt = & $winget.Source show --exact --id $WingetId 2>&1
  $line = $txt | Where-Object { $_ -match '^\s*Version:\s*' } | Select-Object -First 1
  if ($line) { return ($line -replace '^\s*Version:\s*','').Trim() }

  throw "Unable to determine available version for winget id '$WingetId'."
}

function InstallOrUpgrade-WithWinget {
  param([Parameter(Mandatory)][string]$WingetId)

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { throw "winget.exe not found." }

  # If present in winget list, use upgrade; otherwise install
  $isInstalled = $false
  try {
    $listTxt = & $winget.Source list --exact --id $WingetId 2>$null
    if ($LASTEXITCODE -eq 0 -and ($listTxt -join "`n") -match [regex]::Escape($WingetId)) { $isInstalled = $true }
  } catch { }

  $common = @(
    '--exact','--id', $WingetId,
    '--accept-package-agreements','--accept-source-agreements',
    '--silent','--disable-interactivity','--scope','machine'
  )

  if ($isInstalled) {
    if ($PSCmdlet.ShouldProcess("winget upgrade $WingetId","Upgrade")) {
      $p = Start-Process -FilePath $winget.Source -ArgumentList @('upgrade') + $common -Wait -PassThru -NoNewWindow
      if ($p.ExitCode -ne 0) { throw "winget upgrade failed (exit $($p.ExitCode)) for $WingetId" }
    }
  } else {
    if ($PSCmdlet.ShouldProcess("winget install $WingetId","Install")) {
      $p = Start-Process -FilePath $winget.Source -ArgumentList @('install') + $common -Wait -PassThru -NoNewWindow
      if ($p.ExitCode -ne 0) { throw "winget install failed (exit $($p.ExitCode)) for $WingetId" }
    }
  }
}

function Get-UninstallEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  $items = foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
      Where-Object { $_.DisplayName -and -not $_.SystemComponent }
  }

  $items | ForEach-Object {
    [pscustomobject]@{
      DisplayName          = $_.DisplayName
      DisplayVersion       = $_.DisplayVersion
      Publisher            = $_.Publisher
      UninstallString      = $_.UninstallString
      QuietUninstallString = $_.QuietUninstallString
      InstallLocation      = $_.InstallLocation
      PSPath               = $_.PSPath
    }
  }
}

function Get-MsiProductCodeFromUninstallString {
  param([string]$UninstallString)
  if (-not $UninstallString) { return $null }

  # Common patterns:
  # MsiExec.exe /I{GUID}
  # MsiExec.exe /X{GUID}
  $m = [regex]::Match($UninstallString, '\{[0-9A-Fa-f\-]{36}\}')
  if ($m.Success) { return $m.Value }
  return $null
}

function Uninstall-EntrySilently {
  param([Parameter(Mandatory)]$Entry)

  # Prefer QuietUninstallString
  if ($Entry.QuietUninstallString) {
    if ($PSCmdlet.ShouldProcess($Entry.DisplayName, "Uninstall (QuietUninstallString)")) {
      $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $Entry.QuietUninstallString -Wait -PassThru -NoNewWindow
      return $p.ExitCode
    }
    return 0
  }

  # MSI product code
  $code = Get-MsiProductCodeFromUninstallString -UninstallString $Entry.UninstallString
  if ($code) {
    $args = "/x $code /qn /norestart"
    if ($PSCmdlet.ShouldProcess($Entry.DisplayName, "msiexec $args")) {
      $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
      return $p.ExitCode
    }
    return 0
  }

  Write-Warning "Cannot silently uninstall (non-MSI / no quiet string): $($Entry.DisplayName)"
  return 0
}

function Cleanup-EnvVarsAndPath {
  param([string]$KeepJavaHome)

  $envKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'

  $javaHome = (Get-ItemProperty -Path $envKey -Name JAVA_HOME -ErrorAction SilentlyContinue).JAVA_HOME
  $jreHome  = (Get-ItemProperty -Path $envKey -Name JRE_HOME  -ErrorAction SilentlyContinue).JRE_HOME

  if ($javaHome -and -not (Test-Path $javaHome)) {
    if ($PSCmdlet.ShouldProcess("JAVA_HOME", "Remove stale value '$javaHome'")) {
      Remove-ItemProperty -Path $envKey -Name JAVA_HOME -ErrorAction SilentlyContinue
    }
  }
  if ($jreHome -and -not (Test-Path $jreHome)) {
    if ($PSCmdlet.ShouldProcess("JRE_HOME", "Remove stale value '$jreHome'")) {
      Remove-ItemProperty -Path $envKey -Name JRE_HOME -ErrorAction SilentlyContinue
    }
  }

  if ($KeepJavaHome -and (Test-Path $KeepJavaHome)) {
    if ($PSCmdlet.ShouldProcess("JAVA_HOME", "Set to '$KeepJavaHome'")) {
      Set-ItemProperty -Path $envKey -Name JAVA_HOME -Value $KeepJavaHome
    }
  }

  $pathVal = (Get-ItemProperty -Path $envKey -Name Path -ErrorAction SilentlyContinue).Path
  if ($pathVal) {
    $parts = $pathVal -split ';' | Where-Object { $_ -and $_.Trim() -ne '' }
    $newParts = @()

    foreach ($p in $parts) {
      $pp = $p.Trim()
      # Drop java-ish entries that no longer exist
      if ($pp -match '\\java\\' -or $pp -match '\\eclipse adoptium\\' -or $pp -match '\\temurin\\') {
        if (-not (Test-Path $pp)) { continue }
      }
      $newParts += $pp
    }

    $newPath = ($newParts | Select-Object -Unique) -join ';'
    if ($newPath -ne $pathVal) {
      if ($PSCmdlet.ShouldProcess("PATH", "Remove stale Java entries")) {
        Set-ItemProperty -Path $envKey -Name Path -Value $newPath
      }
    }
  }
}

# -------------------- main --------------------
if (-not (Test-IsAdmin)) { throw "Run this script as Administrator." }

Ensure-LogFolder
Start-Transcript -Path $LogPath -Append | Out-Null

try {
  if (-not $Force) {
    $javaProcs = Get-Process -Name java,javaw,javaws -ErrorAction SilentlyContinue
    if ($javaProcs) { throw "Java processes are running. Close apps using Java or re-run with -Force." }
  }

  $wingetId = Get-WingetId -Vendor $Vendor -Family $TargetFamily
  Write-Host "Managing JRE: Vendor=$Vendor  Family=$TargetFamily  wingetId=$wingetId"

  $scan = Get-InstalledJavaForFamilyAndType -Family $TargetFamily -ImageType 'jre'
  if ($scan.All.Count -gt 0) {
    Write-Host "`nDetected Java runtimes on this machine:"
    $scan.All | Sort-Object { $_.Norm.Key } -Descending | ForEach-Object {
      Write-Host ("- {0}  {1}  ({2})" -f $_.Version, $_.JavaHome, $_.Type)
    }
  }

  $installed = $scan.Best
  if ($installed) {
    Write-Host "`nBest matching installed JRE $TargetFamily: $($installed.Version) @ $($installed.JavaHome)"
  } else {
    Write-Host "`nNo matching JRE $TargetFamily found."
  }

  $availableVer = $null
  try { $availableVer = Get-WingetAvailableVersion -WingetId $wingetId } catch { Write-Warning $_.Exception.Message }
  if ($availableVer) { Write-Host "Latest available via winget: $availableVer" }

  $needsUpdate = $true
  if ($installed -and $availableVer) {
    $i = $installed.Norm
    $a = Try-Normalize -VersionString $availableVer
    if (-not $a) { Write-Warning "Could not normalize available version '$availableVer' — will attempt upgrade."; $needsUpdate = $true }
    else { $needsUpdate = ($i.Key -lt $a.Key) }
  } elseif ($installed -and -not $availableVer) {
    Write-Warning "Cannot determine available version. Will not auto-update without winget metadata."
    $needsUpdate = $false
  }

  if ($ReportOnly) {
    if ($availableVer) {
      Write-Host "`nReportOnly: Update required? $needsUpdate"
    } else {
      Write-Host "`nReportOnly: Available version unknown; cannot decide update reliably."
    }

    if ($RemoveOlder -or $Cleanup) {
      Write-Host "ReportOnly: Would also perform RemoveOlder=$RemoveOlder Cleanup=$Cleanup"
    }
    exit ($(if ($needsUpdate) { 2 } else { 0 }))
  }

  if ($needsUpdate) {
    Write-Host "`nUpdating/Installing via winget..."
    InstallOrUpgrade-WithWinget -WingetId $wingetId
  } else {
    Write-Host "`nNo update required."
  }

  # Re-scan after install/upgrade
  $scanAfter = Get-InstalledJavaForFamilyAndType -Family $TargetFamily -ImageType 'jre'
  $keep = $scanAfter.Best
  $keepNorm = if ($keep) { $keep.Norm } else { $null }

  if ($keep) {
    Write-Host "Post-action best JRE $TargetFamily: $($keep.Version) @ $($keep.JavaHome)"
  } else {
    Write-Warning "After update, still no matching JRE $TargetFamily detected."
  }

  if ($RemoveOlder -and $keepNorm) {
    Write-Host "`nRemoving older JRE installs for family $TargetFamily (best-effort)..."
    $entries = Get-UninstallEntries

    # Heuristics: filter likely JRE entries for this vendor/family
    $candidates = $entries | Where-Object {
      ($_.DisplayName -match 'JRE|Runtime Environment' -or $_.DisplayName -match 'Temurin.*JRE' -or $_.DisplayName -match 'Java\(TM\) SE Runtime') -and
      ( ($Vendor -eq 'Temurin' -and $_.DisplayName -match 'Temurin|Eclipse Adoptium|Adoptium') -or
        ($Vendor -eq 'Oracle'  -and $_.DisplayName -match 'Oracle|Java\(TM\)') )
    }

    foreach ($e in $candidates) {
      $n = Try-Normalize -VersionString ($e.DisplayVersion)
      if (-not $n) { continue }
      if ($n.Major -ne $TargetFamily) { continue }
      if ($n.Key -lt $keepNorm.Key) {
        Write-Host ("- Uninstalling older: {0} ({1})" -f $e.DisplayName, $e.DisplayVersion)
        $code = Uninstall-EntrySilently -Entry $e
        if ($code -ne 0) { Write-Warning "Uninstall exit code $code for $($e.DisplayName)" }
      }
    }
  }

  if ($Cleanup -or $RemoveOlder) {
    Write-Host "`nCleanup: env vars + PATH..."
    Cleanup-EnvVarsAndPath -KeepJavaHome ($keep.JavaHome)
  }

  exit ($(if ($needsUpdate) { 1 } else { 0 }))
}
finally {
  Stop-Transcript | Out-Null
}