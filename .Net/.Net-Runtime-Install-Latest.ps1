#Requires -Version 5.1
<#
DotNet-MinKeep-ConditionalInstall-Cleanup.ps1

What it does:
- Reports installed .NET inventory: SDK + Runtime + ASP.NET Core + Windows Desktop (x64 + optional x86).
- Removes versions below MinKeepVersion (best effort: uninstall tool + folder cleanup).
- Installs latest for a family ONLY if that family has below-min versions detected (i.e. would be removed).

Channel selection:
- TargetChannel defaults to major.minor of MinKeepVersion (prevents unintended upgrades to higher channels like 10.0).

Datto env vars:
  DotNet_MinKeepVersion (required)
  DotNet_TargetChannel
  DotNet_ReportOnly
  DotNet_IncludeX86
  DotNet_LatestLTSOnly
  DotNet_ForceUninstallTool
  DotNet_LogPath

Exit codes:
  0 = no action needed (no below-min versions found)
  1 = changes made (installed and/or removed)
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
function Get-Env([string]$Name) { try { (Get-Item "Env:$Name" -ErrorAction SilentlyContinue).Value } catch { $null } }
function Get-EnvBool([string]$Name, [bool]$Default=$false) {
  $v = Get-Env $Name
  if ($null -eq $v -or $v -eq '') { return $Default }
  switch (($v.ToString()).Trim().ToLowerInvariant()) {
    '1' { $true } 'true' { $true } 'yes' { $true } 'y' { $true } 'on' { $true }
    '0' { $false } 'false' { $false } 'no' { $false } 'n' { $false } 'off' { $false }
    default { $Default }
  }
}

if (-not $PSBoundParameters.ContainsKey('MinKeepVersion'))   { $MinKeepVersion = Get-Env 'DotNet_MinKeepVersion' }
if (-not $PSBoundParameters.ContainsKey('TargetChannel'))    { $TargetChannel  = Get-Env 'DotNet_TargetChannel' }
if (-not $PSBoundParameters.ContainsKey('ReportOnly'))       { $ReportOnly     = Get-EnvBool 'DotNet_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('IncludeX86'))       { $IncludeX86     = Get-EnvBool 'DotNet_IncludeX86' $false }
if (-not $PSBoundParameters.ContainsKey('LatestLTSOnly'))    { $LatestLTSOnly  = Get-EnvBool 'DotNet_LatestLTSOnly' $true }  # default LTS
if (-not $PSBoundParameters.ContainsKey('ForceUninstallTool')) { $ForceUninstallTool = Get-EnvBool 'DotNet_ForceUninstallTool' $false }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'DotNet_LogPath'
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
function Normalize-List($obj) { @($obj) | Where-Object { $_ -ne $null } }
function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Parse-Version3Or4([string]$v) {
  if (-not $v) { return $null }
  $m = [regex]::Match($v.Trim(), '^(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?')
  if (-not $m.Success) { return $null }
  $a=$m.Groups[1].Value; $b=$m.Groups[2].Value
  $c=if($m.Groups[3].Success){$m.Groups[3].Value}else{'0'}
  $d=if($m.Groups[4].Success){$m.Groups[4].Value}else{'0'}
  [Version]("$a.$b.$c.$d")
}
function Derive-Channel([Version]$minKeepObj) { "{0}.{1}" -f $minKeepObj.Major, $minKeepObj.Minor }
function Format-VersionList($arr) {
  $arr=Normalize-List $arr
  if($arr.Count -eq 0){ "<none>" } else { (($arr|Sort-Object|ForEach-Object{$_.ToString()}) -join ", ") }
}
function BelowMin($versions,[Version]$min){ @((Normalize-List $versions) | Where-Object { $_ -lt $min } | Sort-Object -Unique) }

# ---------------- Release metadata ----------------
function Get-LatestChannelInfo {
  param([string]$ChannelVersion, [switch]$LtsOnly)

  Ensure-Tls12
  $indexUrl = "https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json"
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  $channels = @($idx.'releases-index') | Where-Object { $_.product -eq '.NET' -and $_.'support-phase' -eq 'active' }
  if ($LtsOnly) { $channels = @($channels | Where-Object { $_.'release-type' -eq 'lts' }) }
  if ($ChannelVersion) { $channels = @($channels | Where-Object { $_.'channel-version' -eq $ChannelVersion }) }

  if (-not $channels -or $channels.Count -eq 0) { return $null }

  $best = $channels | Sort-Object @{Expression={ [decimal]$_.('channel-version') }} -Descending | Select-Object -First 1
  [pscustomobject]@{
    ChannelVersion  = [string]$best.'channel-version'
    LatestRelease   = [string]$best.'latest-release'
    ReleasesJsonUrl = [string]$best.'releases.json'
    ReleaseType     = [string]$best.'release-type'
  }
}

function Get-LatestInstallersForRelease {
  param([Parameter(Mandatory)][string]$ReleasesJsonUrl,[Parameter(Mandatory)][string]$LatestRelease)

  Ensure-Tls12
  $rj = Invoke-RestMethod -Uri $ReleasesJsonUrl -UseBasicParsing
  $release = @($rj.releases) | Where-Object { $_.'release-version' -eq $LatestRelease } | Select-Object -First 1
  if (-not $release) { throw "Could not find release-version $LatestRelease in releases.json." }

  function Pick($obj,[string]$rid,[string]$nameLike){
    if(-not $obj){return $null}
    $files=@($obj.files)
    $f=$files|Where-Object{ $_.rid -eq $rid -and $_.url -and $_.name -like $nameLike -and $_.name -like "*.exe"}|Select-Object -First 1
    if($f){ return [string]$f.url }
    $f2=$files|Where-Object{ $_.rid -eq $rid -and $_.url -and $_.name -like "*.exe"}|Select-Object -First 1
    if($f2){ return [string]$f2.url }
    $null
  }

  $runtimeObj=$release.runtime
  $aspnetObj=$release.'aspnetcore-runtime'
  $desktopObj=$release.windowsdesktop
  $sdkObj=$release.sdk

  [pscustomobject]@{
    RuntimeVersion  = if($runtimeObj){[string]$runtimeObj.version}else{$null}
    AspNetVersion   = if($aspnetObj){[string]$aspnetObj.version}else{$null}
    DesktopVersion  = if($desktopObj){[string]$desktopObj.version}else{$null}
    SdkVersion      = if($sdkObj){[string]$sdkObj.version}else{$null}

    RuntimeUrlX64   = Pick $runtimeObj "win-x64" "dotnet-runtime*"
    RuntimeUrlX86   = Pick $runtimeObj "win-x86" "dotnet-runtime*"
    AspNetUrlX64    = Pick $aspnetObj  "win-x64" "aspnetcore-runtime*"
    AspNetUrlX86    = Pick $aspnetObj  "win-x86" "aspnetcore-runtime*"
    DesktopUrlX64   = Pick $desktopObj "win-x64" "windowsdesktop-runtime*"
    DesktopUrlX86   = Pick $desktopObj "win-x86" "windowsdesktop-runtime*"
    SdkUrlX64       = Pick $sdkObj     "win-x64" "dotnet-sdk*"
    SdkUrlX86       = Pick $sdkObj     "win-x86" "dotnet-sdk*"
  }
}

# ---------------- Folder inventory ----------------
function Get-DotNetRoot([ValidateSet('x64','x86')]$Arch) { if($Arch -eq 'x86'){"C:\Program Files (x86)\dotnet"}else{"C:\Program Files\dotnet"} }
function Get-VersionsFromDirs([string]$Path) {
  if(-not (Test-Path $Path)){ return @() }
  $vers=@()
  Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $v=Parse-Version3Or4 $_.Name
    if($v){ $vers += $v }
  }
  ($vers | Sort-Object -Unique)
}
function Get-DotNetInventory([ValidateSet('x64','x86')]$Arch) {
  $root=Get-DotNetRoot $Arch
  [pscustomobject]@{
    Arch=$Arch; Root=$root
    SDK     = Get-VersionsFromDirs (Join-Path $root "sdk")
    Runtime = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.NETCore.App")
    AspNet  = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.AspNetCore.App")
    Desktop = Get-VersionsFromDirs (Join-Path $root "shared\Microsoft.WindowsDesktop.App")
  }
}

# ---------------- ARP below-min (includes Desktop Runtime) ----------------
function Get-ArpEntries {
  $paths=@(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )
  foreach($p in $paths){
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | ForEach-Object{
      if(-not $_.DisplayName){return}
      [pscustomobject]@{
        DisplayName=$_.DisplayName
        DisplayVersion=$_.DisplayVersion
        QuietUninstallString=$_.QuietUninstallString
        UninstallString=$_.UninstallString
      }
    }
  }
}

function Get-ArchFromName([string]$name){
  if(-not $name){return $null}
  if($name -match '(?i)\barm64\b'){ return 'arm64' }
  if($name -match '(?i)\(x64\)|\bx64\b'){ return 'x64' }
  if($name -match '(?i)\(x86\)|\bx86\b'){ return 'x86' }
  $null
}

function Classify-DotNetArpFamily([string]$dn){
  if(-not $dn){return $null}
  if($dn -match '(?i)Windows Desktop Runtime'){ return 'Desktop' }
  if($dn -match '(?i)ASP\.NET Core'){ return 'AspNet' }
  if($dn -match '(?i)\.NET Runtime|Microsoft \.NET Runtime'){ return 'Runtime' }
  if($dn -match '(?i)\.NET SDK|\bSDK\b'){ return 'SDK' }
  if($dn -match '(?i)\.NET Host FX Resolver'){ return 'HostFxr' }
  if($dn -match '(?i)\.NET Host\b'){ return 'Host' }
  if($dn -match '(?i)\bHosting Bundle\b'){ return 'HostingBundle' }
  $null
}

function Get-DotNetArpBelowMin([Version]$MinKeep,[switch]$IncludeX86){
  $want=@('x64'); if($IncludeX86){$want+= 'x86'}
  $out=@()
  foreach($e in Get-ArpEntries){
    $dn=[string]$e.DisplayName
    if($dn -notmatch '(?i)\bMicrosoft\b'){ continue }
    if($dn -notmatch '(?i)\.NET|ASP\.NET|Windows Desktop Runtime|Hosting Bundle'){ continue }

    $fam = Classify-DotNetArpFamily $dn
    if(-not $fam){ continue }

    $arch = Get-ArchFromName $dn
    if($arch -eq 'arm64'){ continue }
    if($arch -and ($want -notcontains $arch) -and $fam -ne 'HostingBundle'){ continue }

    $ver = Parse-Version3Or4 $e.DisplayVersion
    if(-not $ver){
      $m=[regex]::Match($dn,'(\d+\.\d+\.\d+(?:\.\d+)?)')
      if($m.Success){ $ver = Parse-Version3Or4 $m.Value }
    }
    if(-not $ver){ continue }

    if($ver -lt $MinKeep){
      $archOut = if($arch){ $arch } else { '' }
      $out += [pscustomobject]@{ Family=$fam; Arch=$archOut; Version=$ver; Entry=$e }
    }
  }
  $out
}

function Normalize-MsiUninstall([string]$cmd,[switch]$Force){
  if(-not $cmd){return $null}
  if($cmd -match '(?i)msiexec(\.exe)?\s'){
    $args=$cmd -replace '(?i)^.*?msiexec(\.exe)?\s*',''
    $args=$args -replace '(?i)/I','/X'
    if($args -notmatch '(?i)/quiet|/qn'){ $args += ($(if($Force){' /qn'}else{' /quiet'})) }
    if($args -notmatch '(?i)/norestart'){ $args += ' /norestart' }
    return @{Exe='msiexec.exe'; Args=$args}
  }
  @{Exe='cmd.exe'; Args="/c `"$cmd`""}
}

function Uninstall-ArpEntry($entry,[switch]$Force){
  $cmd=$entry.QuietUninstallString
  if(-not $cmd){ $cmd=$entry.UninstallString }
  if(-not $cmd){ Write-Log "WARNING: No uninstall string for '$($entry.DisplayName)'"; return }

  $n=Normalize-MsiUninstall $cmd -Force:$Force
  Write-Log ("ARP uninstall: {0} ({1})" -f $entry.DisplayName, $entry.DisplayVersion)
  $p=Start-Process -FilePath $n.Exe -ArgumentList $n.Args -Wait -PassThru -NoNewWindow

  if(@(0,1605,1614,3010) -contains $p.ExitCode){
    Write-Log ("ARP uninstall exit={0} (ok)" -f $p.ExitCode)
  } else {
    Write-Log ("WARNING: ARP uninstall exit={0}" -f $p.ExitCode)
  }
}

function Download-And-InstallExe([string]$Url,[string]$Label){
  Ensure-Tls12
  $tmp=Join-Path $env:TEMP "DotNetMinKeepUpdate"
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  $file=Join-Path $tmp ([IO.Path]::GetFileName($Url))
  Write-Log "Downloading: $Label"
  Invoke-WebRequest -Uri $Url -OutFile $file -UseBasicParsing
  Write-Log "Installing: $Label (silent)"
  $p=Start-Process -FilePath $file -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
  if(@(0,3010) -notcontains $p.ExitCode){ throw "$Label installer failed (exit $($p.ExitCode))." }
}

function Remove-FolderIfBelowMin([string]$path,[Version]$min){
  if(-not (Test-Path $path)){ return }
  Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object{
    $v=Parse-Version3Or4 $_.Name
    if($v -and $v -lt $min){
      try{
        Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop
        Write-Log "Deleted folder below min: $($_.FullName)"
      } catch {
        Write-Log "WARNING: Failed to delete folder $($_.FullName) : $($_.Exception.Message)"
      }
    }
  }
}

# ---------------- MAIN ----------------
try{
  if(-not (Test-IsAdmin)){ throw "Run elevated (Administrator / SYSTEM)." }
  if(-not $MinKeepVersion){ throw "DotNet_MinKeepVersion is required." }

  $minKeepObj = Parse-Version3Or4 $MinKeepVersion
  if(-not $minKeepObj){ throw "MinKeepVersion '$MinKeepVersion' is invalid." }

  if(-not $TargetChannel){
    $TargetChannel = Derive-Channel $minKeepObj
    $LatestLTSOnly = $true
    Write-Log ("TargetChannel not set. Defaulting to LTS channel derived from MinKeepVersion: '{0}'" -f $TargetChannel)
  }

  Write-Log ("Starting. ReportOnly={0} IncludeX86={1} LatestLTSOnly={2} TargetChannel='{3}'" -f `
    $ReportOnly, $IncludeX86, $LatestLTSOnly, $TargetChannel)

  $invX64=Get-DotNetInventory x64
  $invX86= if($IncludeX86){ Get-DotNetInventory x86 } else { $null }

  Write-Log "Installed .NET inventory (folders) (x64):"
  Write-Log ("  SDK     : {0}" -f (Format-VersionList $invX64.SDK))
  Write-Log ("  Runtime : {0}" -f (Format-VersionList $invX64.Runtime))
  Write-Log ("  AspNet  : {0}" -f (Format-VersionList $invX64.AspNet))
  Write-Log ("  Desktop : {0}" -f (Format-VersionList $invX64.Desktop))

  $belowArp = Get-DotNetArpBelowMin -MinKeep $minKeepObj -IncludeX86:$IncludeX86
  Write-Log ("MinKeepVersion={0}" -f $minKeepObj)
  Write-Log ("Below-min (ARP): {0} item(s)" -f $belowArp.Count)

  $need=@{
    SDK     = (@($belowArp|Where-Object{$_.Family -eq 'SDK'}).Count -gt 0)     -or ((BelowMin $invX64.SDK $minKeepObj).Count -gt 0)
    Runtime = (@($belowArp|Where-Object{$_.Family -eq 'Runtime'}).Count -gt 0) -or ((BelowMin $invX64.Runtime $minKeepObj).Count -gt 0)
    AspNet  = (@($belowArp|Where-Object{$_.Family -eq 'AspNet'}).Count -gt 0)  -or ((BelowMin $invX64.AspNet $minKeepObj).Count -gt 0)
    Desktop = (@($belowArp|Where-Object{$_.Family -eq 'Desktop'}).Count -gt 0) -or ((BelowMin $invX64.Desktop $minKeepObj).Count -gt 0)
  }

  if(-not ($need.SDK -or $need.Runtime -or $need.AspNet -or $need.Desktop)){
    Write-Log "No versions below MinKeepVersion were found. Skipping install by design."
    exit 0
  }

  $ch = Get-LatestChannelInfo -ChannelVersion $TargetChannel -LtsOnly:$LatestLTSOnly
  if(-not $ch -and $LatestLTSOnly){
    Write-Log ("WARNING: No active LTS channel found for '{0}'. Falling back to highest active LTS." -f $TargetChannel)
    $ch = Get-LatestChannelInfo -ChannelVersion $null -LtsOnly:$true
  }
  if(-not $ch){ throw "Could not resolve .NET channel." }

  Write-Log ("Selected channel {0} ({1}) latest release {2}" -f $ch.ChannelVersion, $ch.ReleaseType, $ch.LatestRelease)
  $latest = Get-LatestInstallersForRelease -ReleasesJsonUrl $ch.ReleasesJsonUrl -LatestRelease $ch.LatestRelease

  if($ReportOnly){
    Write-Log "ReportOnly: would install latest for families with below-min items, then uninstall below-min (including Desktop Runtime) via ARP."
    exit 2
  }

  if($need.Runtime){ Download-And-InstallExe $latest.RuntimeUrlX64 ".NET Runtime x64 $($latest.RuntimeVersion)" }
  if($need.AspNet){  Download-And-InstallExe $latest.AspNetUrlX64 "ASP.NET Core Runtime x64 $($latest.AspNetVersion)" }
  if($need.Desktop){ Download-And-InstallExe $latest.DesktopUrlX64 ".NET Desktop Runtime x64 $($latest.DesktopVersion)" }
  if($need.SDK){     Download-And-InstallExe $latest.SdkUrlX64 ".NET SDK x64 $($latest.SdkVersion)" }

  # Dedup ARP uninstalls by uninstall command to avoid duplicate 1605 spam
  $belowArp2 = Get-DotNetArpBelowMin -MinKeep $minKeepObj -IncludeX86:$IncludeX86
  if($belowArp2.Count -gt 0){
    $groups = $belowArp2 | Group-Object -Property @{Expression={
      $cmd = $_.Entry.QuietUninstallString
      if(-not $cmd){ $cmd = $_.Entry.UninstallString }
      if(-not $cmd){ $cmd = $_.Entry.DisplayName }
      $cmd
    }}
    Write-Log ("ARP packages below min to remove: {0} (dedup groups={1})" -f $belowArp2.Count, $groups.Count)
    foreach($g in $groups){
      $item = $g.Group | Select-Object -First 1
      Uninstall-ArpEntry $item.Entry -Force:$ForceUninstallTool
    }
  }

  # Folder cleanup (best effort)
  $root64 = Get-DotNetRoot x64
  if($need.Runtime){ Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.NETCore.App") $minKeepObj }
  if($need.AspNet){  Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.AspNetCore.App") $minKeepObj }
  if($need.Desktop){ Remove-FolderIfBelowMin (Join-Path $root64 "shared\Microsoft.WindowsDesktop.App") $minKeepObj }
  if($need.SDK){     Remove-FolderIfBelowMin (Join-Path $root64 "sdk") $minKeepObj }

  Write-Log "Done."
  exit 1
}
catch{
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}