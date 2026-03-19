#Requires -Version 5.1
<#
VC++ Any-Version Conditional Install + MinKeep Cleanup (fix: op_Addition + correct counts)

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
  [string]$MinKeepVersion,
  [string]$TargetChannel,
  [switch]$ReportOnly,
  [switch]$IncludeX86,
  [switch]$LatestLTSOnly,
  [switch]$ForceUninstallTool,
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\DotNet\DotNet-MinKeep-Update.log"
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

if (-not $PSBoundParameters.ContainsKey('MinKeepVersion'))     { $MinKeepVersion = Get-Env 'DotNet_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('TargetChannel'))      { $TargetChannel = Get-Env 'DotNet_TargetChannel' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))         { $ReportOnly = Get-EnvBool 'DotNet_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))         { $IncludeX86 = Get-EnvBool 'DotNet_IncludeX86' $false }
# IMPORTANT CHANGE: Default LTS-only to TRUE
if (-not $PSBoundParameters.ContainsKey('LatestLTSOnly'))      { $LatestLTSOnly = Get-EnvBool 'DotNet_LatestLTSOnly' $true }
if (-not $PSBoundParameters.ContainsKey('ForceUninstallTool')) { $ForceUninstallTool = Get-EnvBool 'DotNet_ForceUninstallTool' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'DotNet_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- logging / utilities ----------------
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

function Normalize-List($obj) { @($obj) | Where-Object { $_ -ne $null } }

function Parse-Version3Or4([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?')
  if (-not $m.Success) { return $null }
  $a = $m.Groups[1].Value
  $b = $m.Groups[2].Value
  $c = if ($m.Groups[3].Success) { $m.Groups[3].Value } else { '0' }
  $d = if ($m.Groups[4].Success) { $m.Groups[4].Value } else { '0' }
  return [Version]("$a.$b.$c.$d")
}

function Derive-ChannelFromMinKeep([Version]$minKeepObj) {
  return ("{0}.{1}" -f $minKeepObj.Major, $minKeepObj.Minor)
}

# ---------------- release metadata ----------------
function Get-LatestChannelInfo {
  param([string]$ChannelVersion, [switch]$LatestLTSOnly)

  Ensure-Tls12
  $indexUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  $channels = @($idx.'releases-index') | Where-Object {
    $_.product -eq '.NET' -and $_.'support-phase' -eq 'active'
  }

  if ($LatestLTSOnly) {
    $channels = @($channels | Where-Object { $_.'release-type' -eq 'lts' })
  }

  if ($ChannelVersion) {
    $channels = @($channels | Where-Object { $_.'channel-version' -eq $ChannelVersion })
  }

  if (-not $channels -or $channels.Count -eq 0) {
    return $null
  }

  $best = $channels | Sort-Object @{Expression={ [decimal]$_.('channel-version') }} -Descending | Select-Object -First 1

  [pscustomobject]@{
    ChannelVersion  = [string]$best.'channel-version'
    LatestRelease   = [string]$best.'latest-release'
    ReleasesJsonUrl = [string]$best.'releases.json'
    ReleaseType     = [string]$best.'release-type'
    SupportPhase    = [string]$best.'support-phase'
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

  function Pick-ExeUrl($obj, [string]$rid, [string]$nameLike) {
    if (-not $obj) { return $null }
    $files = @($obj.files)
    $f = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like $nameLike -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f) { return [string]$f.url }
    $f2 = $files | Where-Object { $_.rid -eq $rid -and $_.url -and $_.name -like "*.exe" } | Select-Object -First 1
    if ($f2) { return [string]$f2.url }
    return $null
  }

  $runtimeObj = $release.runtime
  $aspnetObj  = $release.'aspnetcore-runtime'
  $desktopObj = $release.windowsdesktop
  $sdkObj     = $release.sdk

  [pscustomobject]@{
    ReleaseVersion = [string]$release.'release-version'

    RuntimeVersion  = if ($runtimeObj) { [string]$runtimeObj.version } else { $null }
    AspNetVersion   = if ($aspnetObj)  { [string]$aspnetObj.version } else { $null }
    DesktopVersion  = if ($desktopObj) { [string]$desktopObj.version } else { $null }
    SdkVersion      = if ($sdkObj)     { [string]$sdkObj.version } else { $null }

    RuntimeUrlX64   = Pick-ExeUrl $runtimeObj "win-x64" "dotnet-runtime*"
    RuntimeUrlX86   = Pick-ExeUrl $runtimeObj "win-x86" "dotnet-runtime*"

    AspNetUrlX64    = Pick-ExeUrl $aspnetObj  "win-x64" "aspnetcore-runtime*"
    AspNetUrlX86    = Pick-ExeUrl $aspnetObj  "win-x86" "aspnetcore-runtime*"

    DesktopUrlX64   = Pick-ExeUrl $desktopObj "win-x64" "windowsdesktop-runtime*"
    DesktopUrlX86   = Pick-ExeUrl $desktopObj "win-x86" "windowsdesktop-runtime*"

    SdkUrlX64       = Pick-ExeUrl $sdkObj     "win-x64" "dotnet-sdk*"
    SdkUrlX86       = Pick-ExeUrl $sdkObj     "win-x86" "dotnet-sdk*"
  }
}

# ---------------- inventory from filesystem ----------------
function Get-DotNetRoot([ValidateSet('x64','x86')]$Arch) {
  if ($Arch -eq 'x86') { return "C:\Program Files (x86)\dotnet" }
  return "C:\Program Files\dotnet"
}

function Get-VersionsFromDirs([string]$Path) {
  if (-not (Test-Path $Path)) { return @() }
  $dirs = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
  $vers = @()
  foreach ($d in $dirs) {
    $v = Parse-Version3Or4 $d.Name
    if ($v) { $vers += $v }
  }
  return ($vers | Sort-Object -Unique)
}

function Get-DotNetInventoryForArch([ValidateSet('x64','x86')]$Arch) {
  $root = Get-DotNetRoot $Arch
  [pscustomobject]@{
    Arch    = $Arch
    Root    = $root
    Sdk     = Get-VersionsFromDirs (Join-Path $root "sdk")
    Runtime = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.NETCore.App")
    AspNet  = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.AspNetCore.App")
    Desktop = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.WindowsDesktop.App")
  }
}

function Format-VersionList($arr) {
  $arr = Normalize-List $arr
  if ($arr.Count -eq 0) { return "<none>" }
  return (($arr | Sort-Object | ForEach-Object { $_.ToString() }) -join ", ")
}

function Get-BelowMin($versions, [Version]$minKeepObj) {
  $versions = Normalize-List $versions
  @($versions | Where-Object { $_ -lt $minKeepObj } | Sort-Object -Unique)
}

# ---------------- uninstall tool + installer helpers (same as previous script) ----------------
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
    Write-Log "WARNING: winget install Microsoft.DotNet.UninstallTool failed."
    return $null
  }

  $cmd2 = Get-Command dotnet-core-uninstall.exe -ErrorAction SilentlyContinue
  if ($cmd2) { return $cmd2.Source }
  return $null
}

function Try-UninstallToolRemoveBelow {
  param(
    [string]$ToolPath,
    [string]$SwitchName,
    [string]$MinKeepText,
    [switch]$IncludeX86,
    [switch]$Force
  )

  if (-not $ToolPath) { return $false }

  $args = @("remove", $SwitchName, "--all-below", $MinKeepText, "-y")
  if (-not $IncludeX86) { $args = @("remove", $SwitchName, "--x64", "--all-below", $MinKeepText, "-y") }
  if ($Force) { $args += "--force" }

  Write-Log ("dotnet-core-uninstall {0}" -f ($args -join " "))
  try {
    $p = Start-Process -FilePath $ToolPath -ArgumentList $args -Wait -PassThru -NoNewWindow
    Write-Log ("dotnet-core-uninstall exit={0}" -f $p.ExitCode)
    return $true
  } catch {
    Write-Log ("WARNING: dotnet-core-uninstall failed for {0}: {1}" -f $SwitchName, $_.Exception.Message)
    return $false
  }
}

function Download-And-InstallExe {
  param([Parameter(Mandatory)][string]$Url,[Parameter(Mandatory)][string]$Label)

  Ensure-Tls12
  $tmp = Join-Path $env:TEMP "DotNetMinKeepUpdate"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $file = Join-Path $tmp ([IO.Path]::GetFileName($Url))

  Write-Log "Downloading: $Label"
  Invoke-WebRequest -Uri $Url -OutFile $file -UseBasicParsing

  Write-Log "Installing: $Label (silent)"
  $p = Start-Process -FilePath $file -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  if (@(0,3010) -notcontains $p.ExitCode) { throw "$Label installer failed (exit $($p.ExitCode))." }
}

function Remove-BelowMinFolders([Version]$minKeepObj, [switch]$IncludeX86, [hashtable]$FamiliesToClean) {
  $arches = @('x64'); if ($IncludeX86) { $arches += 'x86' }
  foreach ($arch in $arches) {
    $root = Get-DotNetRoot $arch
    if (-not (Test-Path $root)) { continue }

    if ($FamiliesToClean.SDK) {
      $p = Join-Path $root "sdk"
      if (Test-Path $p) { Get-ChildItem $p -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-Version3Or4 $_.Name
        if ($v -and $v -lt $minKeepObj) { Write-Log "Deleting SDK folder below min: $($_.FullName)"; Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
      }}
    }
    if ($FamiliesToClean.Runtime) {
      $p = Join-Path $root "shared\Microsoft.NETCore.App"
      if (Test-Path $p) { Get-ChildItem $p -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-Version3Or4 $_.Name
        if ($v -and $v -lt $minKeepObj) { Write-Log "Deleting Runtime folder below min: $($_.FullName)"; Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
      }}
    }
    if ($FamiliesToClean.AspNet) {
      $p = Join-Path $root "shared\Microsoft.AspNetCore.App"
      if (Test-Path $p) { Get-ChildItem $p -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-Version3Or4 $_.Name
        if ($v -and $v -lt $minKeepObj) { Write-Log "Deleting ASP.NET Core folder below min: $($_.FullName)"; Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
      }}
    }
    if ($FamiliesToClean.Desktop) {
      $p = Join-Path $root "shared\Microsoft.WindowsDesktop.App"
      if (Test-Path $p) { Get-ChildItem $p -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $v = Parse-Version3Or4 $_.Name
        if ($v -and $v -lt $minKeepObj) { Write-Log "Deleting Desktop folder below min: $($_.FullName)"; Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
      }}
    }
  }
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run elevated (Administrator / SYSTEM)." }
  if (-not $MinKeepVersion) { throw "DotNet_MinKeepVersion is required (e.g. 8.0.0)." }

  $minKeepObj = Parse-Version3Or4 $MinKeepVersion
  if (-not $minKeepObj) { throw "MinKeepVersion '$MinKeepVersion' is invalid." }

  # IMPORTANT CHANGE: if TargetChannel not supplied, default to LTS
  $derivedChannel = Derive-ChannelFromMinKeep $minKeepObj
  if (-not $TargetChannel) {
    $TargetChannel = $derivedChannel
    $LatestLTSOnly = $true
    Write-Log ("TargetChannel not set. Defaulting to LTS channel selection. First choice='{0}'" -f $TargetChannel)
  }

  Write-Log ("Starting. ReportOnly={0} IncludeX86={1} LatestLTSOnly={2} TargetChannel='{3}' ForceUninstallTool={4}" -f `
    $ReportOnly, $IncludeX86, $LatestLTSOnly, $TargetChannel, $ForceUninstallTool)

  # Inventory report
  $invX64 = Get-DotNetInventoryForArch -Arch x64
  $invX86 = if ($IncludeX86) { Get-DotNetInventoryForArch -Arch x86 } else { $null }

  Write-Log "Installed .NET inventory (x64):"
  Write-Log ("  SDK     : {0}" -f (Format-VersionList $invX64.Sdk))
  Write-Log ("  Runtime : {0}" -f (Format-VersionList $invX64.Runtime))
  Write-Log ("  AspNet  : {0}" -f (Format-VersionList $invX64.AspNet))
  Write-Log ("  Desktop : {0}" -f (Format-VersionList $invX64.Desktop))

  if ($IncludeX86) {
    Write-Log "Installed .NET inventory (x86):"
    Write-Log ("  SDK     : {0}" -f (Format-VersionList $invX86.Sdk))
    Write-Log ("  Runtime : {0}" -f (Format-VersionList $invX86.Runtime))
    Write-Log ("  AspNet  : {0}" -f (Format-VersionList $invX86.AspNet))
    Write-Log ("  Desktop : {0}" -f (Format-VersionList $invX86.Desktop))
  }

  # Below-min per family
  $below = @{
    SDK     = @()
    Runtime = @()
    AspNet  = @()
    Desktop = @()
  }

  $below.SDK     += Get-BelowMin $invX64.Sdk     $minKeepObj
  $below.Runtime += Get-BelowMin $invX64.Runtime $minKeepObj
  $below.AspNet  += Get-BelowMin $invX64.AspNet  $minKeepObj
  $below.Desktop += Get-BelowMin $invX64.Desktop $minKeepObj

  if ($IncludeX86) {
    $below.SDK     += Get-BelowMin $invX86.Sdk     $minKeepObj
    $below.Runtime += Get-BelowMin $invX86.Runtime $minKeepObj
    $below.AspNet  += Get-BelowMin $invX86.AspNet  $minKeepObj
    $below.Desktop += Get-BelowMin $invX86.Desktop $minKeepObj
  }

  $below.SDK     = @($below.SDK     | Sort-Object -Unique)
  $below.Runtime = @($below.Runtime | Sort-Object -Unique)
  $below.AspNet  = @($below.AspNet  | Sort-Object -Unique)
  $below.Desktop = @($below.Desktop | Sort-Object -Unique)

  Write-Log ("MinKeepVersion={0}" -f $minKeepObj)
  Write-Log ("Below-min detected: SDK={0} Runtime={1} AspNet={2} Desktop={3}" -f `
    $below.SDK.Count, $below.Runtime.Count, $below.AspNet.Count, $below.Desktop.Count)

  $need = @{
    SDK     = ($below.SDK.Count     -gt 0)
    Runtime = ($below.Runtime.Count -gt 0)
    AspNet  = ($below.AspNet.Count  -gt 0)
    Desktop = ($below.Desktop.Count -gt 0)
  }

  $cleanupNeeded = ($need.SDK -or $need.Runtime -or $need.AspNet -or $need.Desktop)
  if (-not $cleanupNeeded) {
    Write-Log "No versions below MinKeepVersion were found. Skipping install by design."
    exit 0
  }

  # Channel selection with fallback:
  # 1) Try requested TargetChannel with LTS-only if enabled
  # 2) If not found, fallback to highest active LTS
  $ch = Get-LatestChannelInfo -ChannelVersion $TargetChannel -LatestLTSOnly:$LatestLTSOnly
  if (-not $ch -and $LatestLTSOnly) {
    Write-Log ("WARNING: No active LTS channel found for '{0}'. Falling back to highest active LTS channel." -f $TargetChannel)
    $ch = Get-LatestChannelInfo -ChannelVersion $null -LatestLTSOnly:$true
  }
  if (-not $ch) {
    throw "Could not resolve a .NET channel (TargetChannel='$TargetChannel' LatestLTSOnly=$LatestLTSOnly)."
  }

  Write-Log ("Selected channel {0} ({1}) latest release {2}" -f $ch.ChannelVersion, $ch.ReleaseType, $ch.LatestRelease)
  $latest = Get-LatestInstallersForRelease -ReleasesJsonUrl $ch.ReleasesJsonUrl -LatestRelease $ch.LatestRelease

  Write-Log ("Latest versions for channel {0}: SDK={1} Runtime={2} AspNet={3} Desktop={4}" -f `
    $ch.ChannelVersion, $latest.SdkVersion, $latest.RuntimeVersion, $latest.AspNetVersion, $latest.DesktopVersion)

  if ($ReportOnly) {
    if ($need.SDK)     { Write-Log ("ReportOnly: would install latest SDK {0} and remove SDK below {1}: {2}" -f $latest.SdkVersion, $minKeepObj, (Format-VersionList $below.SDK)) }
    if ($need.Runtime) { Write-Log ("ReportOnly: would install latest Runtime {0} and remove Runtime below {1}: {2}" -f $latest.RuntimeVersion, $minKeepObj, (Format-VersionList $below.Runtime)) }
    if ($need.AspNet)  { Write-Log ("ReportOnly: would install latest AspNet {0} and remove AspNet below {1}: {2}" -f $latest.AspNetVersion, $minKeepObj, (Format-VersionList $below.AspNet)) }
    if ($need.Desktop) { Write-Log ("ReportOnly: would install latest Desktop {0} and remove Desktop below {1}: {2}" -f $latest.DesktopVersion, $minKeepObj, (Format-VersionList $below.Desktop)) }
    exit 2
  }

  # Install latest only for families that need cleanup
  if ($need.SDK) {
    if (-not $latest.SdkUrlX64) { throw "Missing SDK installer URL (x64) for channel $($ch.ChannelVersion)." }
    Download-And-InstallExe -Url $latest.SdkUrlX64 -Label ".NET SDK x64 $($latest.SdkVersion)"
    if ($IncludeX86 -and $latest.SdkUrlX86) { Download-And-InstallExe -Url $latest.SdkUrlX86 -Label ".NET SDK x86 $($latest.SdkVersion)" }
  }
  if ($need.Runtime) {
    if (-not $latest.RuntimeUrlX64) { throw "Missing Runtime installer URL (x64) for channel $($ch.ChannelVersion)." }
    Download-And-InstallExe -Url $latest.RuntimeUrlX64 -Label ".NET Runtime x64 $($latest.RuntimeVersion)"
    if ($IncludeX86 -and $latest.RuntimeUrlX86) { Download-And-InstallExe -Url $latest.RuntimeUrlX86 -Label ".NET Runtime x86 $($latest.RuntimeVersion)" }
  }
  if ($need.AspNet) {
    if (-not $latest.AspNetUrlX64) { throw "Missing AspNet installer URL (x64) for channel $($ch.ChannelVersion)." }
    Download-And-InstallExe -Url $latest.AspNetUrlX64 -Label "ASP.NET Core Runtime x64 $($latest.AspNetVersion)"
    if ($IncludeX86 -and $latest.AspNetUrlX86) { Download-And-InstallExe -Url $latest.AspNetUrlX86 -Label "ASP.NET Core Runtime x86 $($latest.AspNetVersion)" }
  }
  if ($need.Desktop) {
    if (-not $latest.DesktopUrlX64) { throw "Missing Desktop installer URL (x64) for channel $($ch.ChannelVersion)." }
    Download-And-InstallExe -Url $latest.DesktopUrlX64 -Label ".NET Desktop Runtime x64 $($latest.DesktopVersion)"
    if ($IncludeX86 -and $latest.DesktopUrlX86) { Download-And-InstallExe -Url $latest.DesktopUrlX86 -Label ".NET Desktop Runtime x86 $($latest.DesktopVersion)" }
  }

  # Remove below-min (best effort)
  $tool = Ensure-DotNetUninstallTool
  if ($tool) {
    if ($need.SDK)     { [void](Try-UninstallToolRemoveBelow -ToolPath $tool -SwitchName "--sdk"                    -MinKeepText $MinKeepVersion -IncludeX86:$IncludeX86 -Force:$ForceUninstallTool) }
    if ($need.Runtime) { [void](Try-UninstallToolRemoveBelow -ToolPath $tool -SwitchName "--runtime"                -MinKeepText $MinKeepVersion -IncludeX86:$IncludeX86 -Force:$ForceUninstallTool) }
    if ($need.AspNet)  { [void](Try-UninstallToolRemoveBelow -ToolPath $tool -SwitchName "--aspnet-runtime"         -MinKeepText $MinKeepVersion -IncludeX86:$IncludeX86 -Force:$ForceUninstallTool) }
    if ($need.Desktop) { [void](Try-UninstallToolRemoveBelow -ToolPath $tool -SwitchName "--windowsdesktop-runtime" -MinKeepText $MinKeepVersion -IncludeX86:$IncludeX86 -Force:$ForceUninstallTool) }
  } else {
    Write-Log "WARNING: dotnet-core-uninstall unavailable; relying on folder cleanup only."
  }

  Remove-BelowMinFolders -minKeepObj $minKeepObj -IncludeX86:$IncludeX86 -FamiliesToClean $need

  # Final inventory
  $finalX64 = Get-DotNetInventoryForArch -Arch x64
  Write-Log "Final .NET inventory (x64):"
  Write-Log ("  SDK     : {0}" -f (Format-VersionList $finalX64.Sdk))
  Write-Log ("  Runtime : {0}" -f (Format-VersionList $finalX64.Runtime))
  Write-Log ("  AspNet  : {0}" -f (Format-VersionList $finalX64.AspNet))
  Write-Log ("  Desktop : {0}" -f (Format-VersionList $finalX64.Desktop))

  if ($IncludeX86) {
    $finalX86 = Get-DotNetInventoryForArch -Arch x86
    Write-Log "Final .NET inventory (x86):"
    Write-Log ("  SDK     : {0}" -f (Format-VersionList $finalX86.Sdk))
    Write-Log ("  Runtime : {0}" -f (Format-VersionList $finalX86.Runtime))
    Write-Log ("  AspNet  : {0}" -f (Format-VersionList $finalX86.AspNet))
    Write-Log ("  Desktop : {0}" -f (Format-VersionList $finalX86.Desktop))
  }

  Write-Log "Done."
  exit 1
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}