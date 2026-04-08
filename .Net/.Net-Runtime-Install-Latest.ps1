#Requires -Version 5.1
<#
DotNet-MinKeep-ConditionalInstall-Cleanup.ps1 (CORRECTED)

Behavior:
- MinKeepVersion applies to SDK + Runtime + ASP.NET Core + Windows Desktop
  using both folder and ARP views.
- TargetChannel defaults to the LTS channel derived from MinKeepVersion (major.minor)
  when not explicitly set.
- Installs latest for a family/arch ONLY if:
    (a) there are below-min items for that family/arch (ARP or folder), AND
    (b) there is NOT already a compliant (>= MinKeepVersion) folder version
        for that same family/arch.
- Removes below-min ARP packages for managed families only:
    SDK / Runtime / ASP.NET Core / Windows Desktop.
- Dedupes ARP uninstalls to avoid duplicate 1605 spam.
- Treats MSI exit codes 0 / 1605 / 1614 / 3010 as OK.
- Performs folder cleanup (best-effort) for x64 and optionally x86.
- Post-check logs any remaining below-min ARP items.

Datto RMM env vars:
  DotNet_MinKeepVersion           (required)
  DotNet_TargetChannel            (optional)
  DotNet_ReportOnly               (optional, default false)
  DotNet_IncludeX86               (optional, default false)
  DotNet_LatestLTSOnly            (optional, default true)
  DotNet_ForceUninstallTool       (optional, default false)
  DotNet_LogPath                  (optional)

Exit codes:
  0 = success / no action / report only
  3 = error
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
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
  try { (Get-Item "Env:$Name" -ErrorAction SilentlyContinue).Value } catch { $null }
}
function Get-EnvBool([string]$Name, [bool]$Default = $false) {
  $v = Get-Env $Name
  if ($null -eq $v -or $v -eq '') { return $Default }
  switch (($v.ToString()).Trim().ToLowerInvariant()) {
    '1' { $true }
    'true' { $true }
    'yes' { $true }
    'y' { $true }
    'on' { $true }
    '0' { $false }
    'false' { $false }
    'no' { $false }
    'n' { $false }
    'off' { $false }
    default { $Default }
  }
}

if (-not $PSBoundParameters.ContainsKey('MinKeepVersion'))     { $MinKeepVersion      = Get-Env 'DotNet_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('TargetChannel'))      { $TargetChannel       = Get-Env 'DotNet_TargetChannel' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))         { $ReportOnly          = Get-EnvBool 'DotNet_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))         { $IncludeX86          = Get-EnvBool 'DotNet_IncludeX86' $false }
if (-not $PSBoundParameters.ContainsKey('LatestLTSOnly'))      { $LatestLTSOnly       = Get-EnvBool 'DotNet_LatestLTSOnly' $true }
if (-not $PSBoundParameters.ContainsKey('ForceUninstallTool')) { $ForceUninstallTool  = Get-EnvBool 'DotNet_ForceUninstallTool' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'DotNet_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- logging / utils ----------------
function Write-Log([string]$Message) {
  $dir = Split-Path -Parent $LogPath
  if ($dir -and -not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
  }
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] {1}" -f $ts, $Message
  Write-Host $line
  try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
}

function Normalize-List($obj) {
  @($obj) | Where-Object { $_ -ne $null }
}

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Parse-Version3Or4([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?')
  if (-not $m.Success) { return $null }
  $a = $m.Groups[1].Value
  $b = $m.Groups[2].Value
  $c = if ($m.Groups[3].Success) { $m.Groups[3].Value } else { '0' }
  $d = if ($m.Groups[4].Success) { $m.Groups[4].Value } else { '0' }
  [Version]("$a.$b.$c.$d")
}

function Derive-Channel([Version]$MinKeepObj) {
  "{0}.{1}" -f $MinKeepObj.Major, $MinKeepObj.Minor
}

function Format-VersionList($arr) {
  $arr = Normalize-List $arr
  if ($arr.Count -eq 0) { "<none>" }
  else { (($arr | Sort-Object | ForEach-Object { $_.ToString() }) -join ", ") }
}

function BelowMin($versions, [Version]$Min) {
  @((Normalize-List $versions) | Where-Object { $_ -lt $Min } | Sort-Object -Unique)
}

function Get-MaxVersion($versions) {
  $v = @($versions) | Where-Object { $_ -ne $null } | Sort-Object -Descending
  if ($v.Count -eq 0) { return $null }
  $v[0]
}

# ---------------- Release metadata ----------------
function Get-LatestChannelInfo {
  param(
    [string]$ChannelVersion,
    [switch]$LtsOnly
  )

  Ensure-Tls12
  $indexUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  $channels = @($idx.'releases-index') | Where-Object {
    $_.product -eq '.NET' -and $_.'support-phase' -eq 'active'
  }

  if ($LtsOnly) {
    $channels = @($channels | Where-Object { $_.'release-type' -eq 'lts' })
  }

  if ($ChannelVersion) {
    $channels = @($channels | Where-Object { $_.'channel-version' -eq $ChannelVersion })
  }

  if (-not $channels -or $channels.Count -eq 0) { return $null }

  $best = $channels |
    Sort-Object @{ Expression = { [decimal]$_.('channel-version') } } -Descending |
    Select-Object -First 1

  [pscustomobject]@{
    ChannelVersion  = [string]$best.'channel-version'
    LatestRelease   = [string]$best.'latest-release'
    ReleasesJsonUrl = [string]$best.'releases.json'
    ReleaseType     = [string]$best.'release-type'
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
  if (-not $release) {
    throw "Could not find release-version $LatestRelease in releases.json."
  }

  function Pick($obj, [string]$rid, [string]$nameLike) {
    if (-not $obj) { return $null }
    $files = @($obj.files)
    $f = $files | Where-Object {
      $_.rid -eq $rid -and $_.url -and $_.name -like $nameLike -and $_.name -like "*.exe"
    } | Select-Object -First 1
    if ($f) { return [string]$f.url }

    $f2 = $files | Where-Object {
      $_.rid -eq $rid -and $_.url -and $_.name -like "*.exe"
    } | Select-Object -First 1
    if ($f2) { return [string]$f2.url }

    $null
  }

  $runtimeObj = $release.runtime
  $aspnetObj  = $release.'aspnetcore-runtime'
  $desktopObj = $release.windowsdesktop
  $sdkObj     = $release.sdk

  [pscustomobject]@{
    RuntimeVersion = if ($runtimeObj) { [string]$runtimeObj.version } else { $null }
    AspNetVersion  = if ($aspnetObj)  { [string]$aspnetObj.version } else { $null }
    DesktopVersion = if ($desktopObj) { [string]$desktopObj.version } else { $null }
    SdkVersion     = if ($sdkObj)     { [string]$sdkObj.version } else { $null }

    RuntimeUrlX64  = Pick $runtimeObj "win-x64" "dotnet-runtime*"
    RuntimeUrlX86  = Pick $runtimeObj "win-x86" "dotnet-runtime*"
    AspNetUrlX64   = Pick $aspnetObj  "win-x64" "aspnetcore-runtime*"
    AspNetUrlX86   = Pick $aspnetObj  "win-x86" "aspnetcore-runtime*"
    DesktopUrlX64  = Pick $desktopObj "win-x64" "windowsdesktop-runtime*"
    DesktopUrlX86  = Pick $desktopObj "win-x86" "windowsdesktop-runtime*"
    SdkUrlX64      = Pick $sdkObj     "win-x64" "dotnet-sdk*"
    SdkUrlX86      = Pick $sdkObj     "win-x86" "dotnet-sdk*"
  }
}

# ---------------- Folder inventory ----------------
function Get-DotNetRoot([ValidateSet('x64', 'x86')]$Arch) {
  if ($Arch -eq 'x86') { "C:\Program Files (x86)\dotnet" } else { "C:\Program Files\dotnet" }
}

function Get-VersionsFromDirs([string]$Path) {
  if (-not (Test-Path $Path)) { return @() }
  $vers = @()
  Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $v = Parse-Version3Or4 $_.Name
    if ($v) { $vers += $v }
  }
  ($vers | Sort-Object -Unique)
}

function Get-DotNetInventory([ValidateSet('x64', 'x86')]$Arch) {
  $root = Get-DotNetRoot $Arch
  [pscustomobject]@{
    Arch    = $Arch
    Root    = $root
    SDK     = Get-VersionsFromDirs (Join-Path $root "sdk")
    Runtime = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.NETCore.App")
    AspNet  = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.AspNetCore.App")
    Desktop = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.WindowsDesktop.App")
  }
}

function Has-CompliantFolderVersion {
  param(
    [ValidateSet('SDK', 'Runtime', 'AspNet', 'Desktop')]$Family,
    [ValidateSet('x64', 'x86')]$Arch,
    [Version]$MinKeepObj
  )

  $inv = Get-DotNetInventory $Arch
  $list = switch ($Family) {
    'SDK'     { $inv.SDK }
    'Runtime' { $inv.Runtime }
    'AspNet'  { $inv.AspNet }
    'Desktop' { $inv.Desktop }
  }

  $max = Get-MaxVersion $list
  return ($max -and $max -ge $MinKeepObj)
}

# ---------------- ARP inventory ----------------
function Get-ArpEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | ForEach-Object {
      if (-not $_.DisplayName) { return }
      [pscustomobject]@{
        DisplayName           = $_.DisplayName
        DisplayVersion        = $_.DisplayVersion
        QuietUninstallString  = $_.QuietUninstallString
        UninstallString       = $_.UninstallString
      }
    }
  }
}

function Get-ArchFromName([string]$Name) {
  if (-not $Name) { return $null }
  if ($Name -match '(?i)\barm64\b') { return 'arm64' }
  if ($Name -match '(?i)\(x64\)|\bx64\b') { return 'x64' }
  if ($Name -match '(?i)\(x86\)|\bx86\b') { return 'x86' }
  $null
}

function Classify-DotNetArpFamily([string]$DisplayName) {
  if (-not $DisplayName) { return $null }

  if ($DisplayName -match '(?i)Windows Desktop Runtime') { return 'Desktop' }
  if ($DisplayName -match '(?i)ASP\.NET Core')           { return 'AspNet' }
  if ($DisplayName -match '(?i)\.NET Runtime|Microsoft \.NET Runtime') { return 'Runtime' }
  if ($DisplayName -match '(?i)\.NET SDK|\bSDK\b')       { return 'SDK' }

  # Intentionally exclude Host / HostFxr / Hosting Bundle from action logic.
  $null
}

# IMPORTANT: return $out (NOT ,$out). Call sites wrap with @(...).
function Get-DotNetArpBelowMin {
  param(
    [Version]$MinKeep,
    [switch]$IncludeX86
  )

  $want = @('x64')
  if ($IncludeX86) { $want += 'x86' }

  $out = @()

  foreach ($e in Get-ArpEntries) {
    $dn = [string]$e.DisplayName

    if ($dn -notmatch '(?i)\bMicrosoft\b') { continue }
    if ($dn -notmatch '(?i)\.NET|ASP\.NET|Windows Desktop Runtime') { continue }

    $fam = Classify-DotNetArpFamily $dn
    if (-not $fam) { continue }

    $arch = Get-ArchFromName $dn
    if ($arch -eq 'arm64') { continue }
    if ($arch -and ($want -notcontains $arch)) { continue }

    $ver = Parse-Version3Or4 $e.DisplayVersion
    if (-not $ver) {
      $m = [regex]::Match($dn, '(\d+\.\d+\.\d+(?:\.\d+)?)')
      if ($m.Success) { $ver = Parse-Version3Or4 $m.Value }
    }
    if (-not $ver) { continue }

    if ($ver -lt $MinKeep) {
      $archOut = if ($arch) { $arch } else { '' }
      $out += [pscustomobject]@{
        Family  = $fam
        Arch    = $archOut
        Version = $ver
        Entry   = $e
      }
    }
  }

  return $out
}

function Normalize-MsiUninstall([string]$Cmd, [switch]$Force) {
  if (-not $Cmd) { return $null }

  if ($Cmd -match '(?i)msiexec(\.exe)?\s') {
    $args = $Cmd -replace '(?i)^.*?msiexec(\.exe)?\s*', ''

    # Replace only a leading /I or -I switch, not arbitrary text
    $args = $args -replace '^(?i)\s*([/-])i\b', '$1X'

    if ($args -notmatch '(?i)(^|\s)(/quiet|/qn)\b') {
      $args += $(if ($Force) { ' /qn' } else { ' /quiet' })
    }
    if ($args -notmatch '(?i)(^|\s)/norestart\b') {
      $args += ' /norestart'
    }

    return @{
      Exe  = 'msiexec.exe'
      Args = $args
    }
  }

  @{
    Exe  = 'cmd.exe'
    Args = "/c `"$Cmd`""
  }
}

function Uninstall-ArpEntry($Entry, [switch]$Force) {
  $cmd = $Entry.QuietUninstallString
  if (-not $cmd) { $cmd = $Entry.UninstallString }
  if (-not $cmd) {
    Write-Log "WARNING: No uninstall string for '$($Entry.DisplayName)'"
    return
  }

  if (-not $PSCmdlet.ShouldProcess($Entry.DisplayName, "Uninstall ARP package")) {
    Write-Log "WhatIf/Confirm prevented uninstall of '$($Entry.DisplayName)'"
    return
  }

  $n = Normalize-MsiUninstall $cmd -Force:$Force
  Write-Log ("ARP uninstall: {0} ({1})" -f $Entry.DisplayName, $Entry.DisplayVersion)
  $p = Start-Process -FilePath $n.Exe -ArgumentList $n.Args -Wait -PassThru -NoNewWindow

  if (@(0,1605,1614,3010) -contains $p.ExitCode) {
    Write-Log ("ARP uninstall exit={0} (ok)" -f $p.ExitCode)
  }
  else {
    Write-Log ("WARNING: ARP uninstall exit={0}" -f $p.ExitCode)
  }
}

function Download-And-InstallExe([string]$Url, [string]$Label) {
  Ensure-Tls12

  if ([string]::IsNullOrWhiteSpace($Url)) {
    throw "Installer URL missing for $Label."
  }

  $tmp = Join-Path $env:TEMP "DotNetMinKeepUpdate"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $file = Join-Path $tmp ([IO.Path]::GetFileName($Url))

  if (-not $PSCmdlet.ShouldProcess($Label, "Download and install")) {
    Write-Log "WhatIf/Confirm prevented install of '$Label'"
    return
  }

  Write-Log "Downloading: $Label"
  Invoke-WebRequest -Uri $Url -OutFile $file -UseBasicParsing

  Write-Log "Installing: $Label (silent)"
  $p = Start-Process -FilePath $file -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  if (@(0,3010) -notcontains $p.ExitCode) {
    throw "$Label installer failed (exit $($p.ExitCode))."
  }
}

function Remove-FolderIfBelowMin([string]$Path, [Version]$Min) {
  if (-not (Test-Path $Path)) { return }

  Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $v = Parse-Version3Or4 $_.Name
    if ($v -and $v -lt $Min) {
      if ($PSCmdlet.ShouldProcess($_.FullName, "Delete below-min folder")) {
        try {
          Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop
          Write-Log "Deleted folder below min: $($_.FullName)"
        }
        catch {
          Write-Log "WARNING: Failed to delete folder $($_.FullName) : $($_.Exception.Message)"
        }
      }
      else {
        Write-Log "WhatIf/Confirm prevented deletion of folder '$($_.FullName)'"
      }
    }
  }
}

function Get-FamilyVersionsFromInventory($Inventory, [string]$Family) {
  switch ($Family) {
    'SDK'     { $Inventory.SDK }
    'Runtime' { $Inventory.Runtime }
    'AspNet'  { $Inventory.AspNet }
    'Desktop' { $Inventory.Desktop }
    default   { @() }
  }
}

function Test-NeedsActionForFamilyArch {
  param(
    [Parameter(Mandatory)]$BelowArp,
    [Parameter(Mandatory)]$Inventory,
    [Parameter(Mandatory)][ValidateSet('SDK','Runtime','AspNet','Desktop')]$Family,
    [Parameter(Mandatory)][ValidateSet('x64','x86')]$Arch,
    [Parameter(Mandatory)][Version]$MinKeepObj
  )

  $arpHit = @(
    $BelowArp | Where-Object {
      $_.Family -eq $Family -and (
        $_.Arch -eq $Arch -or
        [string]::IsNullOrWhiteSpace($_.Arch)
      )
    }
  ).Count -gt 0

  $folderHit = (BelowMin (Get-FamilyVersionsFromInventory -Inventory $Inventory -Family $Family) $MinKeepObj).Count -gt 0

  return ($arpHit -or $folderHit)
}

function Install-IfNeeded {
  param(
    [Parameter(Mandatory)][ValidateSet('SDK','Runtime','AspNet','Desktop')]$Family,
    [Parameter(Mandatory)][ValidateSet('x64','x86')]$Arch,
    [Parameter(Mandatory)][bool]$Needed,
    [Parameter(Mandatory)][Version]$MinKeepObj,
    [Parameter(Mandatory)]$Latest
  )

  if (-not $Needed) { return }

  if (Has-CompliantFolderVersion -Family $Family -Arch $Arch -MinKeepObj $MinKeepObj) {
    Write-Log ("Skipping {0} install ({1}): compliant version already present." -f $Family, $Arch)
    return
  }

  switch ("$Family|$Arch") {
    'Runtime|x64' { Download-And-InstallExe $Latest.RuntimeUrlX64 ".NET Runtime x64 $($Latest.RuntimeVersion)" }
    'Runtime|x86' { Download-And-InstallExe $Latest.RuntimeUrlX86 ".NET Runtime x86 $($Latest.RuntimeVersion)" }
    'AspNet|x64'  { Download-And-InstallExe $Latest.AspNetUrlX64  "ASP.NET Core Runtime x64 $($Latest.AspNetVersion)" }
    'AspNet|x86'  { Download-And-InstallExe $Latest.AspNetUrlX86  "ASP.NET Core Runtime x86 $($Latest.AspNetVersion)" }
    'Desktop|x64' { Download-And-InstallExe $Latest.DesktopUrlX64 ".NET Desktop Runtime x64 $($Latest.DesktopVersion)" }
    'Desktop|x86' { Download-And-InstallExe $Latest.DesktopUrlX86 ".NET Desktop Runtime x86 $($Latest.DesktopVersion)" }
    'SDK|x64'     { Download-And-InstallExe $Latest.SdkUrlX64     ".NET SDK x64 $($Latest.SdkVersion)" }
    'SDK|x86'     { Download-And-InstallExe $Latest.SdkUrlX86     ".NET SDK x86 $($Latest.SdkVersion)" }
    default       { throw "Unhandled install target: $Family $Arch" }
  }
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) {
    throw "Run elevated (Administrator / SYSTEM)."
  }
  if (-not $MinKeepVersion) {
    throw "DotNet_MinKeepVersion is required."
  }

  $minKeepObj = Parse-Version3Or4 $MinKeepVersion
  if (-not $minKeepObj) {
    throw "MinKeepVersion '$MinKeepVersion' is invalid."
  }

  if (-not $TargetChannel) {
    $TargetChannel = Derive-Channel $minKeepObj
    $LatestLTSOnly = $true
    Write-Log ("TargetChannel not set. Defaulting to LTS channel derived from MinKeepVersion: '{0}'" -f $TargetChannel)
  }

  Write-Log ("Starting. ReportOnly={0} IncludeX86={1} LatestLTSOnly={2} TargetChannel='{3}'" -f `
    $ReportOnly, $IncludeX86, $LatestLTSOnly, $TargetChannel)

  $invX64 = Get-DotNetInventory x64
  $invX86 = if ($IncludeX86) { Get-DotNetInventory x86 } else { $null }

  Write-Log "Installed .NET inventory (folders) (x64):"
  Write-Log ("  SDK     : {0}" -f (Format-VersionList $invX64.SDK))
  Write-Log ("  Runtime : {0}" -f (Format-VersionList $invX64.Runtime))
  Write-Log ("  AspNet  : {0}" -f (Format-VersionList $invX64.AspNet))
  Write-Log ("  Desktop : {0}" -f (Format-VersionList $invX64.Desktop))

  if ($IncludeX86) {
    Write-Log "Installed .NET inventory (folders) (x86):"
    Write-Log ("  SDK     : {0}" -f (Format-VersionList $invX86.SDK))
    Write-Log ("  Runtime : {0}" -f (Format-VersionList $invX86.Runtime))
    Write-Log ("  AspNet  : {0}" -f (Format-VersionList $invX86.AspNet))
    Write-Log ("  Desktop : {0}" -f (Format-VersionList $invX86.Desktop))
  }

  $belowArp = @(Get-DotNetArpBelowMin -MinKeep $minKeepObj -IncludeX86:$IncludeX86)
  Write-Log ("MinKeepVersion={0}" -f $minKeepObj)
  Write-Log ("Below-min (ARP, managed families only): {0} item(s)" -f $belowArp.Count)
  foreach ($x in ($belowArp | Sort-Object Family, Arch, Version)) {
    Write-Log ("  Below-min ARP: {0} {1} {2} :: {3}" -f $x.Family, $x.Arch, $x.Version, $x.Entry.DisplayName)
  }

  $need = [ordered]@{
    Runtime = [ordered]@{
      x64 = Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX64 -Family Runtime -Arch x64 -MinKeepObj $minKeepObj
      x86 = if ($IncludeX86) { Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX86 -Family Runtime -Arch x86 -MinKeepObj $minKeepObj } else { $false }
    }
    AspNet = [ordered]@{
      x64 = Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX64 -Family AspNet -Arch x64 -MinKeepObj $minKeepObj
      x86 = if ($IncludeX86) { Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX86 -Family AspNet -Arch x86 -MinKeepObj $minKeepObj } else { $false }
    }
    Desktop = [ordered]@{
      x64 = Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX64 -Family Desktop -Arch x64 -MinKeepObj $minKeepObj
      x86 = if ($IncludeX86) { Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX86 -Family Desktop -Arch x86 -MinKeepObj $minKeepObj } else { $false }
    }
    SDK = [ordered]@{
      x64 = Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX64 -Family SDK -Arch x64 -MinKeepObj $minKeepObj
      x86 = if ($IncludeX86) { Test-NeedsActionForFamilyArch -BelowArp $belowArp -Inventory $invX86 -Family SDK -Arch x86 -MinKeepObj $minKeepObj } else { $false }
    }
  }

  Write-Log ("Need matrix: Runtime[x64={0},x86={1}] AspNet[x64={2},x86={3}] Desktop[x64={4},x86={5}] SDK[x64={6},x86={7}]" -f `
    $need.Runtime.x64, $need.Runtime.x86, `
    $need.AspNet.x64,  $need.AspNet.x86, `
    $need.Desktop.x64, $need.Desktop.x86, `
    $need.SDK.x64,     $need.SDK.x86)

  $anyNeed = @(
    $need.Runtime.x64, $need.Runtime.x86,
    $need.AspNet.x64,  $need.AspNet.x86,
    $need.Desktop.x64, $need.Desktop.x86,
    $need.SDK.x64,     $need.SDK.x86
  ) -contains $true

  if (-not $anyNeed) {
    Write-Log "No versions below MinKeepVersion were found in managed families. Skipping install/uninstall by design."
    exit 0
  }

  $ch = Get-LatestChannelInfo -ChannelVersion $TargetChannel -LtsOnly:$LatestLTSOnly
  if (-not $ch -and $LatestLTSOnly) {
    Write-Log ("WARNING: No active LTS channel found for '{0}'. Falling back to highest active LTS." -f $TargetChannel)
    $ch = Get-LatestChannelInfo -ChannelVersion $null -LtsOnly:$true
  }
  if (-not $ch) {
    throw "Could not resolve .NET channel."
  }

  Write-Log ("Selected channel {0} ({1}) latest release {2}" -f $ch.ChannelVersion, $ch.ReleaseType, $ch.LatestRelease)
  $latest = Get-LatestInstallersForRelease -ReleasesJsonUrl $ch.ReleasesJsonUrl -LatestRelease $ch.LatestRelease

  if ($ReportOnly) {
    Write-Log "ReportOnly: would install latest for family/arch pairs with below-min items and no compliant folder version; then uninstall below-min ARP entries for managed families only."
    exit 0
  }

  # Install latest only if needed for that same arch and not already compliant on disk
  Install-IfNeeded -Family Runtime -Arch x64 -Needed $need.Runtime.x64 -MinKeepObj $minKeepObj -Latest $latest
  if ($IncludeX86) { Install-IfNeeded -Family Runtime -Arch x86 -Needed $need.Runtime.x86 -MinKeepObj $minKeepObj -Latest $latest }

  Install-IfNeeded -Family AspNet -Arch x64 -Needed $need.AspNet.x64 -MinKeepObj $minKeepObj -Latest $latest
  if ($IncludeX86) { Install-IfNeeded -Family AspNet -Arch x86 -Needed $need.AspNet.x86 -MinKeepObj $minKeepObj -Latest $latest }

  Install-IfNeeded -Family Desktop -Arch x64 -Needed $need.Desktop.x64 -MinKeepObj $minKeepObj -Latest $latest
  if ($IncludeX86) { Install-IfNeeded -Family Desktop -Arch x86 -Needed $need.Desktop.x86 -MinKeepObj $minKeepObj -Latest $latest }

  Install-IfNeeded -Family SDK -Arch x64 -Needed $need.SDK.x64 -MinKeepObj $minKeepObj -Latest $latest
  if ($IncludeX86) { Install-IfNeeded -Family SDK -Arch x86 -Needed $need.SDK.x86 -MinKeepObj $minKeepObj -Latest $latest }

  # Re-evaluate below-min ARP after installs, then dedupe uninstall by normalized uninstall command
  $belowArp2 = @(Get-DotNetArpBelowMin -MinKeep $minKeepObj -IncludeX86:$IncludeX86)

  if ($belowArp2.Count -gt 0) {
    $groups = $belowArp2 | Group-Object -Property @{
      Expression = {
        $cmd = $_.Entry.QuietUninstallString
        if ([string]::IsNullOrWhiteSpace($cmd)) { $cmd = $_.Entry.UninstallString }
        if ([string]::IsNullOrWhiteSpace($cmd)) { $cmd = $_.Entry.DisplayName }
        $cmd
      }
    }

    Write-Log ("ARP packages below min to remove: {0} (dedup groups={1})" -f $belowArp2.Count, $groups.Count)

    foreach ($g in $groups) {
      $item = $g.Group | Select-Object -First 1
      Uninstall-ArpEntry $item.Entry -Force:$ForceUninstallTool
    }
  }
  else {
    Write-Log "No below-min ARP packages found to uninstall."
  }

  # Folder cleanup (best effort)
  $root64 = Get-DotNetRoot x64
  if ($need.Runtime.x64) { Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.NETCore.App")       $minKeepObj }
  if ($need.AspNet.x64)  { Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.AspNetCore.App")     $minKeepObj }
  if ($need.Desktop.x64) { Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.WindowsDesktop.App") $minKeepObj }
  if ($need.SDK.x64)     { Remove-FolderIfBelowMin (Join-Path $root64 "sdk")                                  $minKeepObj }

  if ($IncludeX86) {
    $root86 = Get-DotNetRoot x86
    if ($need.Runtime.x86) { Remove-FolderIfBelowMin (Join-Path $root86 "shared\Microsoft.NETCore.App")       $minKeepObj }
    if ($need.AspNet.x86)  { Remove-FolderIfBelowMin (Join-Path $root86 "shared\Microsoft.AspNetCore.App")     $minKeepObj }
    if ($need.Desktop.x86) { Remove-FolderIfBelowMin (Join-Path $root86 "shared\Microsoft.WindowsDesktop.App") $minKeepObj }
    if ($need.SDK.x86)     { Remove-FolderIfBelowMin (Join-Path $root86 "sdk")                                  $minKeepObj }
  }

  Start-Sleep -Seconds 2

  # Post-check: remaining below-min ARP
  $after = @(Get-DotNetArpBelowMin -MinKeep $minKeepObj -IncludeX86:$IncludeX86)
  Write-Log ("Post-check: Below-min (ARP, managed families only) remaining: {0}" -f $after.Count)
  foreach ($x in ($after | Sort-Object Family, Arch, Version)) {
    Write-Log ("  Remaining below-min ARP: {0} {1} {2} :: {3}" -f $x.Family, $x.Arch, $x.Version, $x.Entry.DisplayName)
  }

  Write-Log "Done."
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}