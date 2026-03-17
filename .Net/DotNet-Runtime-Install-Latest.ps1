#Requires -Version 5.1
<#
DotNetRuntime-MinKeep-Cleanup.ps1

Key rules:
1) Remove anything below MinKeepVersion.
2) Install the latest .NET Runtime only if there is something below MinKeepVersion to remove.
3) Also handles (remove below min):
   - .NET Runtime (Microsoft.NETCore.App) via dotnet-core-uninstall --runtime
   - ASP.NET Core Shared Framework via dotnet-core-uninstall --aspnet-runtime
   - ASP.NET Core Hosting Bundle via dotnet-core-uninstall --hosting-bundle
   - .NET Host and HostFX Resolver via Apps & Features uninstall entries (optional)

Datto env vars (optional):
  DotNet_MinKeepVersion            (e.g. 8.0.10)
  DotNet_ReportOnly
  DotNet_IncludeX86
  DotNet_LatestLTSOnly
  DotNet_TargetChannel             (e.g. "8.0")
  DotNet_RemoveHostComponents      (true/false)
  DotNet_RemoveHostingBundle       (true/false)
  DotNet_ForceUninstallTool        (true/false)
  DotNet_LogPath

Exit codes:
  0 = no changes needed
  1 = changed (installed and/or removed)
  2 = report only and changes would be made
  3 = error
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$MinKeepVersion,                 # e.g. 8.0.10
  [switch]$ReportOnly,
  [switch]$IncludeX86,
  [switch]$LatestLTSOnly,
  [string]$TargetChannel,                  # e.g. 8.0
  [switch]$RemoveHostComponents,
  [switch]$RemoveHostingBundle,
  [switch]$ForceUninstallTool,
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\DotNet\DotNetRuntimeUpdate.log"
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

# Datto overrides only if parameter wasn't explicitly provided
if (-not $PSBoundParameters.ContainsKey('MinKeepVersion'))        { $MinKeepVersion        = Get-Env 'DotNet_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))           { $ReportOnly           = Get-EnvBool 'DotNet_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))           { $IncludeX86           = Get-EnvBool 'DotNet_IncludeX86' $false }
if (-not $PSBoundParameters.ContainsKey('LatestLTSOnly'))        { $LatestLTSOnly        = Get-EnvBool 'DotNet_LatestLTSOnly' $false }
if (-not $PSBoundParameters.ContainsKey('TargetChannel'))        { $TargetChannel        = Get-Env 'DotNet_TargetChannel' }
if (-not $PSBoundParameters.ContainsKey('RemoveHostComponents')) { $RemoveHostComponents = Get-EnvBool 'DotNet_RemoveHostComponents' $true }
if (-not $PSBoundParameters.ContainsKey('RemoveHostingBundle'))  { $RemoveHostingBundle  = Get-EnvBool 'DotNet_RemoveHostingBundle' $true }
if (-not $PSBoundParameters.ContainsKey('ForceUninstallTool'))   { $ForceUninstallTool   = Get-EnvBool 'DotNet_ForceUninstallTool' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath'))              {
  $lp = Get-Env 'DotNet_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- Logging / utils ----------------
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

function Parse-Version3([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)(?:\.(\d+))?')
  if (-not $m.Success) { return $null }
  $maj = $m.Groups[1].Value
  $min = $m.Groups[2].Value
  $pat = if ($m.Groups[3].Success) { $m.Groups[3].Value } else { '0' }
  return [Version]("$maj.$min.$pat")
}

# ---------------- Microsoft release metadata ----------------
function Get-LatestStableDotNetRelease {
  Ensure-Tls12
  $indexUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  $channels = @($idx.'releases-index') | Where-Object { $_.product -eq '.NET' -and $_.'support-phase' -eq 'active' }

  if ($LatestLTSOnly) {
    $channels = @($channels | Where-Object { $_.'release-type' -eq 'lts' })
  }
  if ($TargetChannel) {
    $channels = @($channels | Where-Object { $_.'channel-version' -eq $TargetChannel })
  }
  if (-not $channels -or $channels.Count -eq 0) {
    throw "No matching active .NET channel found. LatestLTSOnly=$LatestLTSOnly TargetChannel='$TargetChannel'."
  }

  $best = $channels | Sort-Object @{Expression={ [decimal]$_.('channel-version') }} -Descending | Select-Object -First 1

  [pscustomobject]@{
    ChannelVersion  = [string]$best.'channel-version'
    LatestRelease   = [string]$best.'latest-release'
    ReleasesJsonUrl = [string]$best.'releases.json'
  }
}

function Get-LatestInstallersForRelease {
  param(
    [Parameter(Mandatory)][string]$ReleasesJsonUrl,
    [Parameter(Mandatory)][string]$LatestRelease
  )

  Ensure-Tls12
  $rj = Invoke-RestMethod -Uri $ReleasesJsonUrl -UseBasicParsing
  $release = @($rj.releases) | Where-Object { $_.'release-version' -eq $LatestRelease } | Select-Object -First 1
  if (-not $release) { throw "Could not find release-version $LatestRelease in releases.json." }

  $runtimeObj = $release.runtime

  function Pick-ExeUrl($obj, [string]$rid, [string]$nameLike) {
    if (-not $obj) { return $null }
    $files = @($obj.files)
    $f = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like $nameLike -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f) { return [string]$f.url }
    $f2 = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f2) { return [string]$f2.url }
    return $null
  }

  [pscustomobject]@{
    RuntimeVersion        = [string]$runtimeObj.version
    RuntimeInstallerUrlX64 = Pick-ExeUrl $runtimeObj "win-x64" "dotnet-runtime*"
    RuntimeInstallerUrlX86 = Pick-ExeUrl $runtimeObj "win-x86" "dotnet-runtime*"
  }
}

# ---------------- Installed versions (sharedfx) ----------------
function Get-InstalledSharedFxVersions {
  param(
    [ValidateSet('x64','x86')][string]$Arch,
    [Parameter(Mandatory)][string]$FxName
  )

  $base = "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\$Arch\sharedfx\$FxName"
  if (-not (Test-Path $base)) { return @() }

  $props = (Get-ItemProperty -Path $base -ErrorAction SilentlyContinue).PSObject.Properties.Name
  $versions = @()
  foreach ($p in $props) { if ($p -match '^\d+\.\d+(\.\d+)?') { $versions += $p } }
  $versions | Sort-Object -Unique
}

function Get-VersionsToRemoveSharedFx {
  param(
    [Parameter(Mandatory)][string]$FxName,
    [Parameter(Mandatory)][Version]$MinKeep
  )

  $all = @()
  foreach ($v in (Get-InstalledSharedFxVersions -Arch x64 -FxName $FxName)) {
    $vv = Parse-Version3 $v
    if ($vv -and $vv -lt $MinKeep) { $all += $vv }
  }
  if ($IncludeX86) {
    foreach ($v in (Get-InstalledSharedFxVersions -Arch x86 -FxName $FxName)) {
      $vv = Parse-Version3 $v
      if ($vv -and $vv -lt $MinKeep) { $all += $vv }
    }
  }
  ($all | Sort-Object -Unique)
}

# ---------------- ARP (Apps & Features) helper for Host/HostFX/Hosting bundle ----------------
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

function Get-ArpToRemove {
  param(
    [Parameter(Mandatory)][Version]$MinKeep
  )

  $patterns = @(
    'Microsoft\.NET Host FX Resolver',
    'Microsoft\.NET Host(?!ing)',              # Host (not Hosting Bundle)
    'Microsoft ASP\.NET Core Hosting Bundle',
    'Microsoft ASP\.NET Core Shared Framework',
    'Microsoft \.NET Runtime'
  )

  $all = @()
  foreach ($e in (Get-ArpEntries)) {
    foreach ($pat in $patterns) {
      if ($e.DisplayName -match $pat) {
        $vv = Parse-Version3 $e.DisplayVersion
        if ($vv -and $vv -lt $MinKeep) {
          $all += [pscustomobject]@{ Entry=$e; Version=$vv; Pattern=$pat }
        }
        break
      }
    }
  }

  $all | Sort-Object Version
}

function Invoke-UninstallEntry {
  param(
    [Parameter(Mandatory)]$Entry
  )

  $cmd = $Entry.QuietUninstallString
  if (-not $cmd) { $cmd = $Entry.UninstallString }

  if (-not $cmd) {
    Write-Log "WARNING: No uninstall string for '$($Entry.DisplayName)'"
    return
  }

  # Common MSI forms: "MsiExec.exe /I{GUID}" -> convert to /X and add quiet
  $exe = $null; $args = $null

  if ($cmd -match 'msiexec(\.exe)?\s') {
    $exe = "msiexec.exe"
    $args = $cmd
    $args = $args -replace '(?i)msiexec(\.exe)?\s*', ''
    $args = $args -replace '(?i)/I', '/X'
    if ($args -notmatch '(?i)/quiet') { $args += ' /quiet' }
    if ($args -notmatch '(?i)/norestart') { $args += ' /norestart' }
  } else {
    # Non-MSI: best effort add /quiet /norestart if not present
    $exe = "cmd.exe"
    $args = "/c `"$cmd`""
  }

  Write-Log "Uninstalling: $($Entry.DisplayName) ($($Entry.DisplayVersion))"
  Start-Process -FilePath $exe -ArgumentList $args -Wait -NoNewWindow | Out-Null
}

# ---------------- dotnet-core-uninstall tool ----------------
function Ensure-DotNetUninstallTool {
  $cmd = Get-Command dotnet-core-uninstall.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { return $null }

  Write-Log "dotnet-core-uninstall not found. Installing via winget (Microsoft.DotNet.UninstallTool)..."
  $p = Start-Process -FilePath $winget.Source -ArgumentList @(
    "install","-e","--id","Microsoft.DotNet.UninstallTool",
    "--accept-package-agreements","--accept-source-agreements",
    "--silent","--disable-interactivity"
  ) -Wait -PassThru -NoNewWindow

  if ($p.ExitCode -ne 0) {
    Write-Log "WARNING: winget install failed (exit $($p.ExitCode))."
    return $null
  }

  $cmd2 = Get-Command dotnet-core-uninstall.exe -ErrorAction SilentlyContinue
  if ($cmd2) { return $cmd2.Source }
  return $null
}

function Uninstall-Tool-RemoveBelow {
  param(
    [Parameter(Mandatory)][string]$TargetSwitch,   # --runtime | --aspnet-runtime | --hosting-bundle
    [Parameter(Mandatory)][string]$MinKeepVerText  # e.g. 8.0.10
  )

  $tool = Ensure-DotNetUninstallTool
  if (-not $tool) {
    Write-Log "WARNING: dotnet-core-uninstall unavailable; skipping uninstall-tool removal for $TargetSwitch."
    return
  }

  $args = @("remove", $TargetSwitch, "--all-below", $MinKeepVerText, "-y")
  if (-not $IncludeX86) { $args = @("remove", $TargetSwitch, "--x64", "--all-below", $MinKeepVerText, "-y") }
  if ($ForceUninstallTool) { $args += "--force" }

  Write-Log "dotnet-core-uninstall $($args -join ' ')"
  Start-Process -FilePath $tool -ArgumentList $args -Wait -NoNewWindow | Out-Null
}

function Remove-OlderSharedFoldersBelowMin {
  param(
    [Parameter(Mandatory)][Version]$MinKeep
  )

  $roots = @("C:\Program Files\dotnet")
  if ($IncludeX86) { $roots += "C:\Program Files (x86)\dotnet" }

  $fxFolders = @(
    "shared\Microsoft.NETCore.App",
    "shared\Microsoft.AspNetCore.App",
    "shared\Microsoft.WindowsDesktop.App"
  )

  foreach ($root in $roots) {
    foreach ($fx in $fxFolders) {
      $path = Join-Path $root $fx
      if (-not (Test-Path $path)) { continue }

      Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-Version3 $_.Name
        if ($v -and $v -lt $MinKeep) {
          Write-Log "Deleting folder below min: $($_.FullName)"
          Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
      }
    }
  }
}

function Download-And-InstallExe {
  param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][string]$Label)

  Ensure-Tls12
  $tmp = Join-Path $env:TEMP "DotNetRuntimeUpdate"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $file = Join-Path $tmp ([IO.Path]::GetFileName($Url))

  Write-Log "Downloading $Label..."
  Invoke-WebRequest -Uri $Url -OutFile $file -UseBasicParsing

  Write-Log "Installing $Label (silent)..."
  $p = Start-Process -FilePath $file -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  if ($p.ExitCode -ne 0) { throw "$Label install failed (exit $($p.ExitCode))." }
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator)." }

  Write-Log ("Starting. ReportOnly={0} IncludeX86={1} LatestLTSOnly={2} TargetChannel='{3}' RemoveHostComponents={4} RemoveHostingBundle={5} ForceUninstallTool={6}" -f `
    $ReportOnly, $IncludeX86, $LatestLTSOnly, $TargetChannel, $RemoveHostComponents, $RemoveHostingBundle, $ForceUninstallTool)

  $latestInfo = Get-LatestStableDotNetRelease
  Write-Log ("Selected channel {0} latest release {1}" -f $latestInfo.ChannelVersion, $latestInfo.LatestRelease)

  $latest = Get-LatestInstallersForRelease -ReleasesJsonUrl $latestInfo.ReleasesJsonUrl -LatestRelease $latestInfo.LatestRelease
  $latestRuntimeVer = Parse-Version3 $latest.RuntimeVersion
  if (-not $latestRuntimeVer) { throw "Could not parse latest runtime version: $($latest.RuntimeVersion)" }

  # If MinKeepVersion not set, default it to "latest runtime" (previous behaviour)
  if (-not $MinKeepVersion) {
    $MinKeepVersion = $latest.RuntimeVersion
    Write-Log "MinKeepVersion not set; defaulting MinKeepVersion=$MinKeepVersion"
  }

  $minKeepObj = Parse-Version3 $MinKeepVersion
  if (-not $minKeepObj) { throw "MinKeepVersion '$MinKeepVersion' is invalid. Use e.g. 8.0.10" }

  # Determine if anything will be removed (< min keep)
  $toRemoveRuntime = Get-VersionsToRemoveSharedFx -FxName "Microsoft.NETCore.App" -MinKeep $minKeepObj
  $toRemoveAspNet  = Get-VersionsToRemoveSharedFx -FxName "Microsoft.AspNetCore.App" -MinKeep $minKeepObj
  $toRemoveDesk    = Get-VersionsToRemoveSharedFx -FxName "Microsoft.WindowsDesktop.App" -MinKeep $minKeepObj

  $arpToRemove = @()
  if ($RemoveHostComponents -or $RemoveHostingBundle) {
    $arpToRemove = Get-ArpToRemove -MinKeep $minKeepObj
    if (-not $RemoveHostComponents) {
      $arpToRemove = @($arpToRemove | Where-Object { $_.Entry.DisplayName -notmatch 'Microsoft\.NET Host' -and $_.Entry.DisplayName -notmatch 'Host FX Resolver' })
    }
    if (-not $RemoveHostingBundle) {
      $arpToRemove = @($arpToRemove | Where-Object { $_.Entry.DisplayName -notmatch 'Hosting Bundle' })
    }
  }

  $cleanupNeeded =
    ($toRemoveRuntime.Count -gt 0) -or
    ($toRemoveAspNet.Count -gt 0) -or
    ($toRemoveDesk.Count -gt 0) -or
    ($arpToRemove.Count -gt 0)

  Write-Log ("MinKeepVersion={0} ; CleanupNeeded={1}" -f $minKeepObj, $cleanupNeeded)
  Write-Log ("Below-min sharedfx to remove: Runtime={0} AspNet={1} Desktop={2} ; ARP packages={3}" -f `
    $toRemoveRuntime.Count, $toRemoveAspNet.Count, $toRemoveDesk.Count, $arpToRemove.Count)

  # Rule #2: only install latest if cleanup is happening
  if (-not $cleanupNeeded) {
    Write-Log "No versions below MinKeepVersion were found. Skipping install of latest by design."
    exit 0
  }

  if ($ReportOnly) {
    Write-Log "ReportOnly: Would install latest runtime and remove everything below MinKeepVersion."
    exit 2
  }

  $changed = $false

  # Install latest runtime (x64 and optionally x86) BEFORE removals
  if (-not $latest.RuntimeInstallerUrlX64) { throw "Could not locate latest x64 runtime installer URL." }
  Download-And-InstallExe -Url $latest.RuntimeInstallerUrlX64 -Label ".NET Runtime x64 $($latest.RuntimeVersion)"
  $changed = $true

  if ($IncludeX86 -and $latest.RuntimeInstallerUrlX86) {
    Download-And-InstallExe -Url $latest.RuntimeInstallerUrlX86 -Label ".NET Runtime x86 $($latest.RuntimeVersion)"
    $changed = $true
  }

  # Remove below MinKeep using dotnet-core-uninstall where possible
  Uninstall-Tool-RemoveBelow -TargetSwitch "--runtime"        -MinKeepVerText $MinKeepVersion
  Uninstall-Tool-RemoveBelow -TargetSwitch "--aspnet-runtime" -MinKeepVerText $MinKeepVersion
  if ($RemoveHostingBundle) {
    Uninstall-Tool-RemoveBelow -TargetSwitch "--hosting-bundle" -MinKeepVerText $MinKeepVersion
  }

  # Optional: remove Host / HostFX / Runtime / Shared Framework entries via ARP (for those below min)
  if ($arpToRemove.Count -gt 0) {
    foreach ($x in $arpToRemove) {
      Invoke-UninstallEntry -Entry $x.Entry
      $changed = $true
    }
  }

  # Remove leftover files below MinKeep
  Remove-OlderSharedFoldersBelowMin -MinKeep $minKeepObj
  $changed = $true

  Write-Log "Done."
  exit ($(if ($changed) { 1 } else { 0 }))
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}