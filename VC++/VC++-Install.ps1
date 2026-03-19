#Requires -Version 5.1
<#
VC++ Any-Version Conditional Install + MinKeep Cleanup (Improved detection/logging)

Rules:
1) Detect installed "Microsoft Visual C++" entries (x86/x64) from ARP.
2) If NO installed VC++ entry is below MinKeepVersion => DO NOTHING (no install).
3) If any installed VC++ entry is below MinKeepVersion:
   - Install provided URL installer(s) ONLY for architectures that have below-min items
   - Uninstall all VC++ entries below MinKeepVersion
   - Rescan & remove any remaining below-min entries
4) Clean up downloaded installers.

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

# ---------------- Datto env helpers ----------------
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

if (-not $PSBoundParameters.ContainsKey('TargetUrlX64'))   { $TargetUrlX64   = Get-Env 'VCRedist_TargetUrl_X64' }
if (-not $PSBoundParameters.ContainsKey('TargetUrlX86'))   { $TargetUrlX86   = Get-Env 'VCRedist_TargetUrl_X86' }
if (-not $PSBoundParameters.ContainsKey('MinKeepVersion')) { $MinKeepVersion = Get-Env 'VCRedist_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))     { $ReportOnly     = Get-EnvBool 'VCRedist_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX64'))     { $IncludeX64     = Get-EnvBool 'VCRedist_IncludeX64' $true }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))     { $IncludeX86     = Get-EnvBool 'VCRedist_IncludeX86' $true }
if (-not $PSBoundParameters.ContainsKey('ForceMSI'))       { $ForceMSI       = Get-EnvBool 'VCRedist_ForceMSI' $false }
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
function Ensure-Tls12 { try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { } }

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Turn any result into a REAL array with no null elements
function Normalize-List($obj) {
  if ($null -eq $obj) { return @() }
  $arr = @($obj) | Where-Object { $_ -ne $null }
  return $arr
}

function Parse-VersionFlexible([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?')
  if (-not $m.Success) { return $null }
  $a = $m.Groups[1].Value
  $b = $m.Groups[2].Value
  $c = if ($m.Groups[3].Success) { $m.Groups[3].Value } else { '0' }
  $d = if ($m.Groups[4].Success) { $m.Groups[4].Value } else { '0' }
  return [Version]("$a.$b.$c.$d")
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
      }
    }
  }
}

function Get-VcEntries([ValidateSet('x86','x64')]$Arch) {
  $all = @()
  foreach ($e in (Get-ArpEntries)) {
    if ($e.DisplayName -notmatch '^Microsoft Visual C\+\+') { continue }
    if ($e.DisplayName -notmatch '\(x86\)|\(x64\)') { continue }
    if ($Arch -eq 'x64' -and $e.DisplayName -notmatch '\(x64\)') { continue }
    if ($Arch -eq 'x86' -and $e.DisplayName -notmatch '\(x86\)') { continue }

    $vv = Parse-VersionFlexible $e.DisplayVersion
    if (-not $vv) {
      $m = [regex]::Match($e.DisplayName, '(\d+\.\d+(?:\.\d+){0,2})')
      if ($m.Success) { $vv = Parse-VersionFlexible $m.Groups[1].Value }
    }
    if (-not $vv) { continue }

    $all += [pscustomobject]@{ Arch=$Arch; Version=$vv; Entry=$e }
  }
  return Normalize-List ($all | Where-Object { $_.Entry -ne $null })
}

function Get-MaxVersion($list) {
  $list = Normalize-List $list
  if ($list.Count -eq 0) { return $null }
  return ($list | Sort-Object Version -Descending | Select-Object -First 1).Version
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
  if (-not $obj -or -not $obj.Entry) { return }

  $e = $obj.Entry
  $cmd = $e.QuietUninstallString
  if (-not $cmd) { $cmd = $e.UninstallString }
  if (-not $cmd) { Write-Log "WARNING: No uninstall string for '$($e.DisplayName)'"; return }

  $norm = Normalize-MsiUninstall $cmd
  if (-not $norm) { Write-Log "WARNING: Could not normalize uninstall for '$($e.DisplayName)'"; return }

  Write-Log ("Uninstalling: {0} ({1})" -f $e.DisplayName, $e.DisplayVersion)
  $p = Start-Process -FilePath $norm.Exe -ArgumentList $norm.Args -Wait -PassThru -NoNewWindow
  Write-Log ("Uninstall exit code: {0}" -f $p.ExitCode)
}

function Download-File([string]$Url, [string]$OutFile) {
  Ensure-Tls12
  Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
}

function Install-FromExe([string]$InstallerPath, [string]$Label) {
  Write-Log "Installing $Label (silent)..."
  $p = Start-Process -FilePath $InstallerPath -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  Write-Log ("Install exit code: {0}" -f $p.ExitCode)

  if (@(0,3010,1638,1641) -notcontains $p.ExitCode) {
    throw "$Label installer failed (exit $($p.ExitCode))."
  }
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator / SYSTEM)." }

  Write-Log "Starting VC++ (any-version) conditional install + MinKeep cleanup..."
  Write-Log ("Options: ReportOnly={0} IncludeX64={1} IncludeX86={2} ForceMSI={3}" -f $ReportOnly, $IncludeX64, $IncludeX86, $ForceMSI)

  if (-not $IncludeX64 -and -not $IncludeX86) { throw "Both IncludeX64 and IncludeX86 are false; nothing to do." }
  if (-not $MinKeepVersion) { throw "VCRedist_MinKeepVersion is required." }

  $minKeep = Parse-VersionFlexible $MinKeepVersion
  if (-not $minKeep) { throw "Invalid MinKeepVersion '$MinKeepVersion'." }

  # Inventory (always arrays)
  $instX64 = if ($IncludeX64) { Normalize-List (Get-VcEntries -Arch x64) } else { @() }
  $instX86 = if ($IncludeX86) { Normalize-List (Get-VcEntries -Arch x86) } else { @() }

  $maxX64 = Get-MaxVersion $instX64
  $maxX86 = Get-MaxVersion $instX86

  Write-Log ("VC++ entries detected: x64={0} x86={1}" -f $instX64.Count, $instX86.Count)
  Write-Log ("Highest detected versions: x64={0} x86={1}" -f ($(if ($maxX64) { $maxX64 } else { "<none>" })), ($(if ($maxX86) { $maxX86 } else { "<none>" })))
  Write-Log ("MinKeepVersion (remove below): {0}" -f $minKeep)

  $belowMinX64 = if ($IncludeX64) { Normalize-List ($instX64 | Where-Object { $_.Version -lt $minKeep }) } else { @() }
  $belowMinX86 = if ($IncludeX86) { Normalize-List ($instX86 | Where-Object { $_.Version -lt $minKeep }) } else { @() }

  $belowMinAll = Normalize-List (@($belowMinX64 + $belowMinX86) | Where-Object { $_ -ne $null -and $_.Entry -ne $null })

  Write-Log ("Entries below MinKeepVersion: x64={0} x86={1} total={2}" -f $belowMinX64.Count, $belowMinX86.Count, $belowMinAll.Count)

  if ($belowMinAll.Count -eq 0) {
    Write-Log "No installed VC++ entries below MinKeepVersion. No install/uninstall will be performed."
    exit 0
  }

  if ($IncludeX64 -and $belowMinX64.Count -gt 0 -and -not $TargetUrlX64) { throw "Missing VCRedist_TargetUrl_X64 (needed because x64 has below-min entries)." }
  if ($IncludeX86 -and $belowMinX86.Count -gt 0 -and -not $TargetUrlX86) { throw "Missing VCRedist_TargetUrl_X86 (needed because x86 has below-min entries)." }

  if ($ReportOnly) {
    Write-Log "ReportOnly: would install from provided URL(s) for affected arch(es) and remove below-min entries."
    exit 2
  }

  $tmp = Join-Path $env:TEMP ("VCRedist_Target_" + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null

  try {
    if ($IncludeX64 -and $belowMinX64.Count -gt 0) {
      $fileX64 = Join-Path $tmp "vc_redist.x64.exe"
      Write-Log "Downloading target x64 installer..."
      Download-File -Url $TargetUrlX64 -OutFile $fileX64
      Install-FromExe -InstallerPath $fileX64 -Label "VC++ target x64 (from URL)"
    } else {
      Write-Log "x64: no below-min entries; skipping install."
    }

    if ($IncludeX86 -and $belowMinX86.Count -gt 0) {
      $fileX86 = Join-Path $tmp "vc_redist.x86.exe"
      Write-Log "Downloading target x86 installer..."
      Download-File -Url $TargetUrlX86 -OutFile $fileX86
      Install-FromExe -InstallerPath $fileX86 -Label "VC++ target x86 (from URL)"
    } else {
      Write-Log "x86: no below-min entries; skipping install."
    }

    Write-Log "Removing installed VC++ entries below MinKeepVersion..."
    foreach ($x in ($belowMinAll | Sort-Object Version)) { Uninstall-Entry $x }

    # Rescan & remove any remaining below-min entries
    $instX64b = if ($IncludeX64) { Normalize-List (Get-VcEntries -Arch x64) } else { @() }
    $instX86b = if ($IncludeX86) { Normalize-List (Get-VcEntries -Arch x86) } else { @() }

    $below2 = @()
    if ($IncludeX64) { $below2 += Normalize-List ($instX64b | Where-Object { $_.Version -lt $minKeep }) }
    if ($IncludeX86) { $below2 += Normalize-List ($instX86b | Where-Object { $_.Version -lt $minKeep }) }
    $below2 = Normalize-List ($below2 | Where-Object { $_ -ne $null -and $_.Entry -ne $null })

    if ($below2.Count -gt 0) {
      Write-Log ("Removing remaining below-min VC++ entries: {0}" -f $below2.Count)
      foreach ($x in ($below2 | Sort-Object Version)) { Uninstall-Entry $x }
    }

    Write-Log "Done."
    exit 1
  }
  finally {
    try { Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }
  }
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}