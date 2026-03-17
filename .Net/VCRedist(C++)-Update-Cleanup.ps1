#Requires -Version 5.1
<#
VCRedist-v14-Update-Cleanup.ps1

Targets: Microsoft Visual C++ v14 Redistributable line (2015-2022 / 2015-2026 naming), x86 and x64.

Rules:
1) Remove any installed v14 redists below MinKeepVersion.
2) Determine "latest" from official installers and compare to installed.
3) Install latest if higher than installed (or missing).
4) After successful install, uninstall any older v14 versions (< latest) and delete downloaded installer files.
5) Handle 32-bit and 64-bit.

Datto env vars:
  VCRedist_MinKeepVersion   (e.g. 14.38.33135.0)
  VCRedist_ReportOnly
  VCRedist_IncludeX86
  VCRedist_IncludeX64
  VCRedist_LogPath
  VCRedist_ForceMSI

Exit codes:
  0 = no changes needed
  1 = changes made
  2 = report only; changes would be made
  3 = error
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$MinKeepVersion,
  [switch]$ReportOnly,
  [switch]$IncludeX86,
  [switch]$IncludeX64,
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

# Datto overrides only if param wasn't explicitly provided
if (-not $PSBoundParameters.ContainsKey('MinKeepVersion')) { $MinKeepVersion = Get-Env 'VCRedist_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))     { $ReportOnly     = Get-EnvBool 'VCRedist_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))     { $IncludeX86     = Get-EnvBool 'VCRedist_IncludeX86' $true }
if (-not $PSBoundParameters.ContainsKey('IncludeX64'))     { $IncludeX64     = Get-EnvBool 'VCRedist_IncludeX64' $true }
if (-not $PSBoundParameters.ContainsKey('ForceMSI'))       { $ForceMSI       = Get-EnvBool 'VCRedist_ForceMSI' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath'))        {
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

function Is-V14Name([string]$displayName) {
  if (-not $displayName) { return $false }
  # Match common modern naming (2015-2019, 2015-2022, 2015-2026) and include Redistributable.
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
    if (-not (Is-V14Name $e.DisplayName)) { continue }
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
  # Returns @{ Exe="msiexec.exe"; Args="..." } or @{ Exe="cmd.exe"; Args="/c ..." }
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

  # Non-MSI: run as-is (best effort)
  return @{ Exe='cmd.exe'; Args="/c `"$cmd`"" }
}

function Uninstall-Entry($entryObj) {
  $e = $entryObj.Entry
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

  Write-Log ("Uninstalling: {0} {1} [{2}]" -f $e.DisplayName, $e.DisplayVersion, $entryObj.Arch)
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

# Official permalinks for latest supported v14 redist (x86/x64)
$UrlX64 = "https://aka.ms/vc14/vc_redist.x64.exe"
$UrlX86 = "https://aka.ms/vc14/vc_redist.x86.exe"

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator)." }

  Write-Log "Starting VC++ v14 Redistributable check..."
  Write-Log ("Options: ReportOnly={0} IncludeX64={1} IncludeX86={2} ForceMSI={3}" -f $ReportOnly, $IncludeX64, $IncludeX86, $ForceMSI)

  # Minimum keep version (required for your policy)
  if (-not $MinKeepVersion) { throw "VCRedist_MinKeepVersion is required (e.g. 14.38.33135.0)." }
  $minKeep = Parse-Version4 $MinKeepVersion
  if (-not $minKeep) { throw "Invalid MinKeepVersion '$MinKeepVersion' (expected 4-part version like 14.38.33135.0)." }

  # Detect installed
  $instX64 = if ($IncludeX64) { @(Get-V14Installed -Arch x64) } else { @() }
  $instX86 = if ($IncludeX86) { @(Get-V14Installed -Arch x86) } else { @() }
  $maxX64  = Get-MaxVersion $instX64
  $maxX86  = Get-MaxVersion $instX86

  Write-Log ("Installed max: x64={0} x86={1}" -f ($maxX64 ?? '<none>'), ($maxX86 ?? '<none>'))

  # Determine which installed entries are below minimum keep
  $belowMin = @()
  $belowMin += @($instX64 | Where-Object { $_.Version -lt $minKeep })
  $belowMin += @($instX86 | Where-Object { $_.Version -lt $minKeep })
  Write-Log ("Entries below MinKeepVersion ({0}): {1}" -f $minKeep, $belowMin.Count)

  # Download installers to read latest version
  $tmp = Join-Path $env:TEMP "VCRedistLatest"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $fileX64 = Join-Path $tmp "vc_redist.x64.exe"
  $fileX86 = Join-Path $tmp "vc_redist.x86.exe"

  Write-Log "Downloading installers to determine latest version..."
  if ($IncludeX64) { Download-File -Url $UrlX64 -OutFile $fileX64 }
  if ($IncludeX86) { Download-File -Url $UrlX86 -OutFile $fileX86 }

  $latest = $null
  if ($IncludeX64) { $latest = Get-InstallerFileVersion $fileX64 }
  elseif ($IncludeX86) { $latest = Get-InstallerFileVersion $fileX86 }

  Write-Log ("Latest installer version detected: {0}" -f $latest)

  # Decide if install needed (latest > installed max OR missing)
  $needInstall = $false
  if ($IncludeX64 -and ((-not $maxX64) -or ($latest -gt $maxX64))) { $needInstall = $true }
  if ($IncludeX86 -and ((-not $maxX86) -or ($latest -gt $maxX86))) { $needInstall = $true }
  Write-Log ("Install latest needed? {0}" -f $needInstall)

  # Plan actions
  $wouldChange = ($belowMin.Count -gt 0) -or $needInstall
  if ($ReportOnly) {
    Write-Log "ReportOnly enabled."
    if ($belowMin.Count -gt 0) { Write-Log "Would uninstall below-min entries." }
    if ($needInstall) { Write-Log "Would install latest x64/x86." }
    if ($wouldChange) { exit 2 } else { exit 0 }
  }

  $changed = $false

  # 2) Remove below minimum first (as requested)
  foreach ($e in $belowMin) {
    if ($PSCmdlet.ShouldProcess($e.Entry.DisplayName, "Uninstall below MinKeepVersion")) {
      Uninstall-Entry $e
      $changed = $true
    }
  }

  # Refresh installed after min cleanup
  $instX64 = if ($IncludeX64) { @(Get-V14Installed -Arch x64) } else { @() }
  $instX86 = if ($IncludeX86) { @(Get-V14Installed -Arch x86) } else { @() }
  $maxX64  = Get-MaxVersion $instX64
  $maxX86  = Get-MaxVersion $instX86

  # 3) Install latest if higher than installed
  if ($needInstall) {
    Write-Log "Installing latest VC++ Redistributable(s)..."
    $args = "/install /quiet /norestart"

    if ($IncludeX64) {
      $p64 = Start-Process -FilePath $fileX64 -ArgumentList $args -Wait -PassThru -NoNewWindow
      Write-Log ("x64 install exit code: {0}" -f $p64.ExitCode)
      if ($p64.ExitCode -ne 0) { throw "x64 installer failed (exit $($p64.ExitCode))." }
      $changed = $true
    }
    if ($IncludeX86) {
      $p86 = Start-Process -FilePath $fileX86 -ArgumentList $args -Wait -PassThru -NoNewWindow
      Write-Log ("x86 install exit code: {0}" -f $p86.ExitCode)
      if ($p86.ExitCode -ne 0) { throw "x86 installer failed (exit $($p86.ExitCode))." }
      $changed = $true
    }
  }

  # 4) After successful install, remove any old v14 versions (< latest)
  $instX64 = if ($IncludeX64) { @(Get-V14Installed -Arch x64) } else { @() }
  $instX86 = if ($IncludeX86) { @(Get-V14Installed -Arch x86) } else { @() }

  $olderThanLatest = @()
  $olderThanLatest += @($instX64 | Where-Object { $_.Version -lt $latest })
  $olderThanLatest += @($instX86 | Where-Object { $_.Version -lt $latest })

  if ($olderThanLatest.Count -gt 0) {
    Write-Log ("Removing older v14 entries below latest ({0}): {1}" -f $latest, $olderThanLatest.Count)
    foreach ($e in $olderThanLatest) {
      Uninstall-Entry $e
      $changed = $true
    }
  } else {
    Write-Log "No older v14 entries below latest were found."
  }

  # 4/5) Cleanup downloaded installer files (safe)
  try {
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Cleaned up downloaded installer files."
  } catch { }

  Write-Log "Done."
  exit ($(if ($changed) { 1 } else { 0 }))
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}