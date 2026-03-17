#Requires -Version 5.1
<#
DotNetRuntime-AutoUpdate-Cleanup.ps1

Datto RMM Input Variables supported (env vars):
  DotNet_ReportOnly
  DotNet_IncludeX86
  DotNet_LatestLTSOnly
  DotNet_ForceUninstallTool
  DotNet_LogPath

Behavior:
- Finds latest stable (or latest active LTS) .NET release from official Microsoft release metadata.
- Compares installed runtimes.
- Installs latest runtime (and updates ASP.NET Core + Desktop runtime only if they already exist).
- Removes older versions (< latest) using:
   1) dotnet-core-uninstall tool (preferred) if available/installed
   2) deletes older shared runtime folders (best-effort file cleanup)

Exit codes:
  0 = already at latest (or no action needed)
  1 = installed latest and/or removed older
  2 = report only and changes would be made
  3 = error
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [switch]$ReportOnly,
  [switch]$IncludeX86,
  [switch]$LatestLTSOnly,
  [switch]$ForceUninstallTool,
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\DotNet\DotNetRuntimeUpdate.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# ----------------- Datto env var helpers -----------------
function Get-Env([string]$Name) {
  try { return (Get-Item "Env:$Name" -ErrorAction SilentlyContinue).Value } catch { return $null }
}

function Get-EnvBool([string]$Name, [bool]$Default=$false) {
  $v = Get-Env $Name
  if ($null -eq $v -or $v -eq '') { return $Default }
  switch (($v.ToString()).Trim().ToLowerInvariant()) {
    '1' { return $true }
    'true' { return $true }
    'yes' { return $true }
    'y' { return $true }
    'on' { return $true }
    '0' { return $false }
    'false' { return $false }
    'no' { return $false }
    'n' { return $false }
    'off' { return $false }
    default { return $Default }
  }
}

# Datto variables override only if the parameter was NOT explicitly provided.
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))        { $ReportOnly        = Get-EnvBool 'DotNet_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))        { $IncludeX86        = Get-EnvBool 'DotNet_IncludeX86' $false }
if (-not $PSBoundParameters.ContainsKey('LatestLTSOnly'))     { $LatestLTSOnly     = Get-EnvBool 'DotNet_LatestLTSOnly' $false }
if (-not $PSBoundParameters.ContainsKey('ForceUninstallTool')){ $ForceUninstallTool= Get-EnvBool 'DotNet_ForceUninstallTool' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'DotNet_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ----------------- logging & utilities -----------------
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

function Parse-SemVer([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)\.(\d+)')
  if (-not $m.Success) { return $null }
  return [Version]("{0}.{1}.{2}" -f $m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value)
}

# ----------------- Microsoft release metadata -----------------
function Get-LatestStableDotNetRelease {
  Ensure-Tls12
  $indexUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  $channels = @($idx.'releases-index') | Where-Object {
    $_.product -eq '.NET' -and $_.'support-phase' -eq 'active'
  }

  if ($LatestLTSOnly) {
    $channels = @($channels | Where-Object { $_.'release-type' -eq 'lts' })
  }

  if (-not $channels -or $channels.Count -eq 0) {
    throw "Could not find an active .NET channel in releases-index.json."
  }

  $best = $channels |
    Sort-Object @{ Expression = { [decimal]$_.('channel-version') } } -Descending |
    Select-Object -First 1

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
  if (-not $release) { throw "Could not find release-version $LatestRelease in $ReleasesJsonUrl" }

  $runtimeObj = $release.runtime
  $aspnetObj  = $release.'aspnetcore-runtime'
  $desktopObj = $release.windowsdesktop

  function Pick-ExeUrl($obj, [string]$rid, [string]$nameLike) {
    if (-not $obj) { return $null }
    $files = @($obj.files)
    $f = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like $nameLike -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f) { return [string]$f.url }
    $f2 = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f2) { return [string]$f2.url }
    return $null
  }

  $rid = "win-x64"
  [pscustomobject]@{
    RuntimeVersion      = [string]$runtimeObj.version
    AspNetCoreVersion   = if ($aspnetObj)  { [string]$aspnetObj.version } else { $null }
    WindowsDesktopVer   = if ($desktopObj) { [string]$desktopObj.version } else { $null }
    RuntimeInstallerUrl = Pick-ExeUrl $runtimeObj $rid "dotnet-runtime*"
    AspNetInstallerUrl  = Pick-ExeUrl $aspnetObj  $rid "aspnetcore-runtime*"
    DesktopInstallerUrl = Pick-ExeUrl $desktopObj $rid "windowsdesktop-runtime*"
  }
}

# ----------------- installed versions -----------------
function Get-InstalledSharedFxVersions {
  param(
    [ValidateSet('x64','x86')][string]$Arch,
    [Parameter(Mandatory)][string]$FxName
  )

  $base = "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\$Arch\sharedfx\$FxName"
  if (-not (Test-Path $base)) { return @() }

  $props = (Get-ItemProperty -Path $base -ErrorAction SilentlyContinue).PSObject.Properties.Name
  $versions = @()
  foreach ($p in $props) {
    if ($p -match '^\d+\.\d+\.\d+') { $versions += $p }
  }
  return ($versions | Sort-Object -Unique)
}

function Test-HasAnyRuntimeFamily([string]$FxName) {
  $x64 = Get-InstalledSharedFxVersions -Arch x64 -FxName $FxName
  if ($x64.Count -gt 0) { return $true }
  if ($IncludeX86) {
    $x86 = Get-InstalledSharedFxVersions -Arch x86 -FxName $FxName
    if ($x86.Count -gt 0) { return $true }
  }
  return $false
}

# ----------------- install / uninstall -----------------
function Download-And-InstallExe {
  param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][string]$Label)

  Ensure-Tls12
  $tmp = Join-Path $env:TEMP "DotNetRuntimeUpdate"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $file = Join-Path $tmp ([IO.Path]::GetFileName($Url))

  Write-Log "Downloading $Label installer..."
  Invoke-WebRequest -Uri $Url -OutFile $file -UseBasicParsing

  Write-Log "Installing $Label (silent)..."
  $p = Start-Process -FilePath $file -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  if ($p.ExitCode -ne 0) { throw "$Label installer failed (exit code $($p.ExitCode))." }
}

function Ensure-DotNetUninstallTool {
  $cmd = Get-Command dotnet-core-uninstall.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { return $null }

  Write-Log "dotnet-core-uninstall not found. Installing via winget: Microsoft.DotNet.UninstallTool"
  $p = Start-Process -FilePath $winget.Source -ArgumentList @(
    "install","-e","--id","Microsoft.DotNet.UninstallTool",
    "--accept-package-agreements","--accept-source-agreements",
    "--silent","--disable-interactivity"
  ) -Wait -PassThru -NoNewWindow

  if ($p.ExitCode -ne 0) {
    Write-Log "WARNING: winget install Microsoft.DotNet.UninstallTool failed (exit $($p.ExitCode))."
    return $null
  }

  $cmd2 = Get-Command dotnet-core-uninstall.exe -ErrorAction SilentlyContinue
  if ($cmd2) { return $cmd2.Source }
  return $null
}

function Remove-OlderWithUninstallTool([string]$LatestVersion) {
  $tool = Ensure-DotNetUninstallTool
  if (-not $tool) {
    Write-Log "WARNING: dotnet-core-uninstall unavailable. Will only delete older shared runtime folders (best-effort)."
    return
  }

  $force = if ($ForceUninstallTool) { "--force" } else { $null }

  Write-Log "Removing older .NET runtimes below $LatestVersion using dotnet-core-uninstall..."
  $argsRuntime = @("remove","--runtime","--all-below",$LatestVersion,"-y")
  if (-not $IncludeX86) { $argsRuntime = @("remove","--runtime","--x64","--all-below",$LatestVersion,"-y") }
  if ($force) { $argsRuntime += $force }
  Start-Process -FilePath $tool -ArgumentList $argsRuntime -Wait -PassThru -NoNewWindow | Out-Null

  Write-Log "Removing older ASP.NET Core runtimes below $LatestVersion using dotnet-core-uninstall..."
  $argsAsp = @("remove","--aspnet-runtime","--all-below",$LatestVersion,"-y")
  if (-not $IncludeX86) { $argsAsp = @("remove","--aspnet-runtime","--x64","--all-below",$LatestVersion,"-y") }
  if ($force) { $argsAsp += $force }
  Start-Process -FilePath $tool -ArgumentList $argsAsp -Wait -PassThru -NoNewWindow | Out-Null
}

function Remove-OlderSharedFolders([Version]$LatestVerObj) {
  $dotnetRoots = @("C:\Program Files\dotnet")
  if ($IncludeX86) { $dotnetRoots += "C:\Program Files (x86)\dotnet" }

  $fxFolders = @(
    "shared\Microsoft.NETCore.App",
    "shared\Microsoft.AspNetCore.App",
    "shared\Microsoft.WindowsDesktop.App"
  )

  foreach ($root in $dotnetRoots) {
    foreach ($fx in $fxFolders) {
      $path = Join-Path $root $fx
      if (-not (Test-Path $path)) { continue }

      Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-SemVer $_.Name
        if ($v -and $v -lt $LatestVerObj) {
          Write-Log "Deleting older runtime folder: $($_.FullName)"
          if ($PSCmdlet.ShouldProcess($_.FullName, "Remove-Item (old runtime folder)")) {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
          }
        }
      }
    }
  }
}

# ----------------- main -----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator)." }

  Write-Log "Starting .NET runtime check/update..."
  Write-Log ("Options: ReportOnly={0} IncludeX86={1} LatestLTSOnly={2} ForceUninstallTool={3}" -f $ReportOnly, $IncludeX86, $LatestLTSOnly, $ForceUninstallTool)

  $latestInfo = Get-LatestStableDotNetRelease
  Write-Log ("Selected channel {0} latest release {1}" -f $latestInfo.ChannelVersion, $latestInfo.LatestRelease)

  $installers = Get-LatestInstallersForRelease -ReleasesJsonUrl $latestInfo.ReleasesJsonUrl -LatestRelease $latestInfo.LatestRelease
  $latestVerObj = Parse-SemVer $installers.RuntimeVersion
  if (-not $latestVerObj) { throw "Could not parse latest runtime version: $($installers.RuntimeVersion)" }

  $hasAspNet  = Test-HasAnyRuntimeFamily -FxName "Microsoft.AspNetCore.App"
  $hasDesktop = Test-HasAnyRuntimeFamily -FxName "Microsoft.WindowsDesktop.App"

  $installedMax = $null
  foreach ($v in (Get-InstalledSharedFxVersions -Arch x64 -FxName "Microsoft.NETCore.App")) {
    $vv = Parse-SemVer $v
    if ($vv -and (-not $installedMax -or $vv -gt $installedMax)) { $installedMax = $vv }
  }
  if ($IncludeX86) {
    foreach ($v in (Get-InstalledSharedFxVersions -Arch x86 -FxName "Microsoft.NETCore.App")) {
      $vv = Parse-SemVer $v
      if ($vv -and (-not $installedMax -or $vv -gt $installedMax)) { $installedMax = $vv }
    }
  }

  $needsUpdate = $true
  if ($installedMax) {
    $needsUpdate = ($installedMax -lt $latestVerObj)
    Write-Log ("Highest installed runtime: {0} ; Latest: {1} ; NeedsUpdate={2}" -f $installedMax, $latestVerObj, $needsUpdate)
  } else {
    Write-Log "No Microsoft.NETCore.App runtime detected. Will install latest."
  }

  if ($ReportOnly) {
    if ($needsUpdate) {
      Write-Log "ReportOnly: Would install latest runtime and remove older versions."
      exit 2
    } else {
      Write-Log "ReportOnly: Already at latest runtime; would remove older (<latest) if any."
      exit 0
    }
  }

  $changed = $false

  if ($needsUpdate) {
    if (-not $installers.RuntimeInstallerUrl) { throw "Could not find runtime installer URL in release metadata." }
    Download-And-InstallExe -Url $installers.RuntimeInstallerUrl -Label ".NET Runtime $($installers.RuntimeVersion)"
    $changed = $true
  }

  if ($hasAspNet -and $installers.AspNetInstallerUrl) {
    Write-Log "ASP.NET Core runtime detected; ensuring latest is installed."
    Download-And-InstallExe -Url $installers.AspNetInstallerUrl -Label "ASP.NET Core Runtime $($installers.AspNetCoreVersion)"
    $changed = $true
  }

  if ($hasDesktop -and $installers.DesktopInstallerUrl) {
    Write-Log ".NET Desktop runtime detected; ensuring latest is installed."
    Download-And-InstallExe -Url $installers.DesktopInstallerUrl -Label ".NET Desktop Runtime $($installers.WindowsDesktopVer)"
    $changed = $true
  }

  Write-Log "Removing older versions below latest..."
  Remove-OlderWithUninstallTool -LatestVersion $installers.RuntimeVersion
  Remove-OlderSharedFolders -LatestVerObj $latestVerObj
  $changed = $true

  Write-Log "Done."
  exit ($(if ($changed) { 1 } else { 0 }))
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}