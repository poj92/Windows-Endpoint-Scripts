#Requires -Version 5.1
<#
VCRedist-v14-UrlInstall-MinKeepCleanup.ps1

- Uses Datto RMM env vars (component input variables) to:
  1) Download target VC++ v14 redistributable installers (x86/x64) from provided URLs
  2) Determine target version from installer FileVersion
  3) Uninstall any installed VC++ v14 redists below MinKeepVersion
  4) Install target if it's higher than installed (or missing)
  5) After install, re-run cleanup below MinKeepVersion
  6) Cleans up downloaded installers (does NOT delete Windows Installer cache / Package Cache)

Scope: VC++ "v14" family, typically displayed as:
  "Microsoft Visual C++ 2015-2019/2022/2026 Redistributable (x86/x64)"

Datto env vars:
  VCRedist_TargetUrl_X64
  VCRedist_TargetUrl_X86
  VCRedist_MinKeepVersion
  VCRedist_ReportOnly
  VCRedist_IncludeX64
  VCRedist_IncludeX86
  VCRedist_ForceMSI
  VCRedist_LogPath

Exit codes:
  0 = no changes needed
  1 = changes made
  2 = report only; changes would be made
  3 = error
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$TargetUrlX64,
  [string]$TargetUrlX86,
  [string]$MinKeepVersion,
  [switch]$ReportOnly,
  [switch]$IncludeX64,
  [switch]$IncludeX86,
  [switch]$ForceMSI,
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\VCRedist\VCRedistUpdate.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# ---------------- Datto env var helpers ----------------
function Get-Env([string]$Name) {
  try { return (Get-Item "Env:$Name" -ErrorAction SilentlyContinue).Value } catch { return $null }
}
function Get-EnvBool([string]$Name, [bool]$Default=$false) {
  $v = Get-Env $Name
  if ($null -eq $v -or $v -eq '') { return $Default }
  switch (($v.ToString()).Trim().ToLowerInvariant()) {
    '1' { $true } 'true' { $true } 'yes' { $true } 'y' { $true } 'on' { $true }
    '0' { $false } 'false' { $false } 'no' { $false } 'n' { $false } 'off' { $false }
    default { $Default }
  }
}

# Datto overrides only if not explicitly provided
if (-not $PSBoundParameters.ContainsKey('TargetUrlX64')) { $TargetUrlX64 = Get-Env 'VCRedist_TargetUrl_X64' }
if (-not $PSBoundParameters.ContainsKey('TargetUrlX86')) { $TargetUrlX86 = Get-Env 'VCRedist_TargetUrl_X86' }
if (-not $PSBoundParameters.ContainsKey('MinKeepVersion')) { $MinKeepVersion = Get-Env 'VCRedist_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly')) { $ReportOnly = Get-EnvBool 'VCRedist_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX64')) { $IncludeX64 = Get-EnvBool 'VCRedist_IncludeX64' $true }
if (-not $PSBoundParameters.ContainsKey('IncludeX86')) { $IncludeX86 = Get-EnvBool 'VCRedist_IncludeX86' $true }
if (-not $PSBoundParameters.ContainsKey('ForceMSI')) { $ForceMSI = Get-EnvBool 'VCRedist_ForceMSI' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'VCRedist_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- logging / utils ----------------
function Write-Log([string]$Message) {
  $dir = Split-Path -Parent $LogPath
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] {1}" -f $ts, $Message
  Write-Host $line
  try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
}

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Parse-Version4([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)\.(\d+)\.(\d+)')
  if (-not $m.Success) { return $null }
  return [Version]("{0}.{1}.{2}.{3}" -f $m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value, $m.Groups[4].Value)
}

function Is-V14DisplayName([string]$displayName) {
  if (-not $displayName) { return $false }
  # Covers "2015-2019", "2015-2022", "2015-2026" etc.
  return ($displayName -match 'Microsoft Visual C\+\+\s+2015-\d{4}\s+Redistributable') -and ($displayName -match '\(x86\)|\(x64\)')
}

function Get-ArpEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )
  foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | ForEach-Object {
      if (-not $_.DisplayName) { return }
      [pscustomobject]@{
        DisplayName = $_.DisplayName
        DisplayVersion = $_.DisplayVersion
        QuietUninstallString = $_.QuietUninstallString
        UninstallString = $_.UninstallString
        PSPath = $_.PSPath
      }
    }
  }
}

function Get-V14Installed([ValidateSet('x86','x64')]$Arch) {
  $all = @()
  foreach ($e in (Get-ArpEntries)) {
    if (-not (Is-V14DisplayName $e.DisplayName)) { continue }
    if ($Arch -eq 'x64' -and $e.DisplayName -notmatch '\(x64\)') { continue }
    if ($Arch -eq 'x86' -and $e.DisplayName -notmatch '\(x86\)') { continue }

    $vv = Parse-Version4 $e.DisplayVersion
    if (-not $vv) { continue }

    $all += [pscustomobject]@{
      Arch = $Arch
      Version = $vv
      Entry = $e
    }
  }
  $all | Sort-Object Version
}

function Get-MaxVersion($list) {
  if (-not $list -or $list.Count -eq 0) { return $null }
  ($list | Sort-Object Version -Descending | Select-Object -First 1).Version
}

function Normalize-MsiUninstall([string]$cmd) {
  if (-not $cmd) { return $null }

  if ($cmd -match '(?i)msiexec(\.exe)?\s') {
    $args = $cmd -replace '(?i)^.*?msiexec(\.exe)?\s*', ''
    $args = $args -replace '(?i)/I', '/X'
    if ($ForceMSI) {
      if ($args -notmatch '(?i)/qn') { $args += ' /qn' }
    } else {
      if ($args -notmatch '(?i)/quiet') { $args += ' /quiet' }
    }
    if ($args -notmatch '(?i)/norestart') { $args += ' /norestart' }
    return @{ Exe='msiexec.exe'; Args=$args }
  }

  return @{ Exe='cmd.exe'; Args="/c `"$cmd`"" }
}

function Uninstall-Entry($obj) {
  $e = $obj.Entry
  $cmd = $e.QuietUninstallString
  if (-not $cmd) { $cmd = $e.UninstallString }
  if (-not $cmd) {
    Write-Log "WARNING: No uninstall string for '$($e.DisplayName)'"
    return
  }

  $norm = Normalize-MsiUninstall $cmd
  if (-not $norm) {
    Write-Log "WARNING: Could not normalize uninstall for '$($e.DisplayName)'"
    return
  }

  Write-Log ("Uninstalling: {0} ({1})" -f $e.DisplayName, $e.DisplayVersion)
  $p = Start-Process -FilePath $norm.Exe -ArgumentList $norm.Args -Wait -PassThru -NoNewWindow
  Write-Log ("Uninstall exit code: {0}" -f $p.ExitCode)
}

function Download-File([string]$Url, [string]$OutFile) {
  Ensure-Tls12
  Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
}

function Get-InstallerFileVersion([string]$FilePath) {
  $fv = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($FilePath).FileVersion
  $v = Parse-Version4 $fv
  if (-not $v) { throw "Could not parse installer FileVersion '$fv' from $FilePath" }
  return $v
}

function Install-Redist([string]$InstallerPath, [string]$Label) {
  Write-Log "Installing $Label (silent)..."
  $args = "/install /quiet /norestart"
  $p = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru -NoNewWindow
  Write-Log ("Install exit code: {0}" -f $p.ExitCode)
  if ($p.ExitCode -ne 0) { throw "$Label installer failed (exit $($p.ExitCode))." }
}

function Cleanup-BelowMin([Version]$minKeep) {
  $removedAny = $false
  if ($IncludeX64) {
    $inst = @(Get-V14Installed -Arch x64)
    foreach ($x in ($inst | Where-Object { $_.Version -lt $minKeep })) {
      Uninstall-Entry $x
      $removedAny = $true
    }
  }
  if ($IncludeX86) {
    $inst = @(Get-V14Installed -Arch x86)
    foreach ($x in ($inst | Where-Object { $_.Version -lt $minKeep })) {
      Uninstall-Entry $x
      $removedAny = $true
    }
  }
  return $removedAny
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator / SYSTEM)." }

  Write-Log "Starting VC++ v14 management (URL-driven)..."
  Write-Log ("Options: ReportOnly={0} IncludeX64={1} IncludeX86={2} ForceMSI={3}" -f $ReportOnly, $IncludeX64, $IncludeX86, $ForceMSI)

  if (-not $IncludeX64 -and -not $IncludeX86) { throw "Both IncludeX64 and IncludeX86 are false; nothing to do." }
  if ($IncludeX64 -and -not $TargetUrlX64) { throw "Missing VCRedist_TargetUrl_X64." }
  if ($IncludeX86 -and -not $TargetUrlX86) { throw "Missing VCRedist_TargetUrl_X86." }

  $tmp = Join-Path $env:TEMP ("VCRedist_Target_" + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null

  $fileX64 = Join-Path $tmp "vc_redist.x64.exe"
  $fileX86 = Join-Path $tmp "vc_redist.x86.exe"

  # Download target installers and determine their versions
  $targetVerX64 = $null
  $targetVerX86 = $null

  if ($IncludeX64) {
    Write-Log "Downloading target x64 installer..."
    Download-File -Url $TargetUrlX64 -OutFile $fileX64
    $targetVerX64 = Get-InstallerFileVersion $fileX64
    Write-Log ("Target x64 installer version: {0}" -f $targetVerX64)
  }
  if ($IncludeX86) {
    Write-Log "Downloading target x86 installer..."
    Download-File -Url $TargetUrlX86 -OutFile $fileX86
    $targetVerX86 = Get-InstallerFileVersion $fileX86
    Write-Log ("Target x86 installer version: {0}" -f $targetVerX86)
  }

  # If MinKeepVersion not set, default to the target version (highest of provided)
  if (-not $MinKeepVersion) {
    $defaultMin = $null
    foreach ($v in @($targetVerX64, $targetVerX86)) {
      if ($v -and (-not $defaultMin -or $v -gt $defaultMin)) { $defaultMin = $v }
    }
    if (-not $defaultMin) { throw "Could not determine a default MinKeepVersion." }
    $MinKeepVersion = $defaultMin.ToString()
    Write-Log ("MinKeepVersion not provided; defaulting to target version: {0}" -f $MinKeepVersion)
  }

  $minKeep = Parse-Version4 $MinKeepVersion
  if (-not $minKeep) { throw "Invalid VCRedist_MinKeepVersion '$MinKeepVersion' (expected 4-part version e.g. 14.38.33135.0)." }

  # Current installed state
  $instX64 = if ($IncludeX64) { @(Get-V14Installed -Arch x64) } else { @() }
  $instX86 = if ($IncludeX86) { @(Get-V14Installed -Arch x86) } else { @() }
  $maxX64  = Get-MaxVersion $instX64
  $maxX86  = Get-MaxVersion $instX86

  Write-Log ("Installed max versions: x64={0} x86={1}" -f ($maxX64 ?? '<none>'), ($maxX86 ?? '<none>'))
  Write-Log ("MinKeepVersion (remove below): {0}" -f $minKeep)

  # Determine below-min entries
  $belowMinCount = 0
  if ($IncludeX64) { $belowMinCount += @($instX64 | Where-Object { $_.Version -lt $minKeep }).Count }
  if ($IncludeX86) { $belowMinCount += @($instX86 | Where-Object { $_.Version -lt $minKeep }).Count }
  Write-Log ("Installed v14 entries below MinKeepVersion: {0}" -f $belowMinCount)

  # Determine if install needed (target > installedMax OR missing)
  $needInstall = $false
  if ($IncludeX64 -and $targetVerX64) {
    if (-not $maxX64 -or $targetVerX64 -gt $maxX64) { $needInstall = $true }
    elseif ($targetVerX64 -lt $maxX64) { Write-Log ("WARNING: Target x64 ({0}) is older than installed ({1}); will NOT downgrade." -f $targetVerX64, $maxX64) }
  }
  if ($IncludeX86 -and $targetVerX86) {
    if (-not $maxX86 -or $targetVerX86 -gt $maxX86) { $needInstall = $true }
    elseif ($targetVerX86 -lt $maxX86) { Write-Log ("WARNING: Target x86 ({0}) is older than installed ({1}); will NOT downgrade." -f $targetVerX86, $maxX86) }
  }

  Write-Log ("Install needed? {0}" -f $needInstall)

  $wouldChange = ($belowMinCount -gt 0) -or $needInstall
  if ($ReportOnly) {
    Write-Log "ReportOnly: no changes will be made."
    if ($belowMinCount -gt 0) { Write-Log "ReportOnly: would uninstall versions below MinKeepVersion." }
    if ($needInstall) { Write-Log "ReportOnly: would install target VC++ redistributable(s)." }
    try { Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    exit ($(if ($wouldChange) { 2 } else { 0 }))
  }

  $changed = $false

  # 1) Remove anything below min first
  if ($belowMinCount -gt 0) {
    Write-Log "Removing installed versions below MinKeepVersion..."
    if (Cleanup-BelowMin -minKeep $minKeep) { $changed = $true }
  }

  # Refresh after cleanup
  $maxX64 = if ($IncludeX64) { Get-MaxVersion @(Get-V14Installed -Arch x64) } else { $null }
  $maxX86 = if ($IncludeX86) { Get-MaxVersion @(Get-V14Installed -Arch x86) } else { $null }

  # 2) Install target if higher than installed (or missing)
  if ($needInstall) {
    if ($IncludeX64 -and $targetVerX64 -and (-not $maxX64 -or $targetVerX64 -gt $maxX64)) {
      Install-Redist -InstallerPath $fileX64 -Label ("VC++ v14 x64 " + $targetVerX64)
      $changed = $true
    }
    if ($IncludeX86 -and $targetVerX86 -and (-not $maxX86 -or $targetVerX86 -gt $maxX86)) {
      Install-Redist -InstallerPath $fileX86 -Label ("VC++ v14 x86 " + $targetVerX86)
      $changed = $true
    }
  }

  # 3) After install, remove anything below MinKeepVersion again (in case multiple entries exist)
  Write-Log "Post-install cleanup: removing any remaining versions below MinKeepVersion..."
  if (Cleanup-BelowMin -minKeep $minKeep) { $changed = $true }

  # NOTE on "remove files":
  # MSI uninstall removes installed binaries. We do NOT delete Windows Installer cache/Package Cache because it can break repair/uninstall.

  # Cleanup downloaded installer files
  try { Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }

  # Final state logging
  $finalX64 = if ($IncludeX64) { Get-MaxVersion @(Get-V14Installed -Arch x64) } else { $null }
  $finalX86 = if ($IncludeX86) { Get-MaxVersion @(Get-V14Installed -Arch x86) } else { $null }
  Write-Log ("Final installed max versions: x64={0} x86={1}" -f ($finalX64 ?? '<n/a>'), ($finalX86 ?? '<n/a>'))

  Write-Log "Done."
  exit ($(if ($changed) { 1 } else { 0 }))
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}