#Requires -Version 5.1
<#
.Java-JRE-Update.ps1

Manages Java JRE (Temurin or Oracle*) for a given major family (8/11/17/21):
- Detect installed JRE(s) and pick the best (highest) for that family
- Compare against winget "latest available" version (when winget present)
- Install/upgrade if needed (unless -ReportOnly)
- Optional: remove older JREs of same family (-RemoveOlder)
- Optional: cleanup stale JAVA_HOME/JRE_HOME and stale Java PATH entries (-Cleanup)

Notes:
- Oracle JRE is generally only available for Java 8. For 11/17/21 use Temurin JRE.
- MSI fallback install is supported for Temurin only.

Exit codes (useful for RMM):
  0 = up-to-date or no action required
  1 = performed an update/install
  2 = ReportOnly and update would be required
  3 = update required but cannot proceed (e.g., winget missing and MSI fallback disabled)
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

  # If winget is missing, download and install Temurin MSI (Vendor must be Temurin)
  [switch]$UseMsiFallback = $true,

  [string]$LogPath = "$env:ProgramData\JavaUpdate\JavaJRE-Update.log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Tls12ForPs5 {
  if ($PSVersionTable.PSVersion.Major -lt 6) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  }
}

function Ensure-LogFolder {
  $dir = Split-Path -Parent $LogPath
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

function Normalize-JavaVersion {
  param([Parameter(Mandatory)][string]$VersionString)

  $v = $VersionString.Trim()

  # Java 8 runtime style: 1.8.0_381-b09
  if ($v -match '^1\.8\.0_(\d+)(?:-b(\d+))?$') {
    $upd = [int]$Matches[1]
    $bld = if ($Matches[2]) { [int]$Matches[2] } else { 0 }
    return [pscustomobject]@{
      Raw=$v; Major=8; Minor=0; SecOrUpd=$upd; Build=$bld
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f 8,0,$upd,$bld
    }
  }

  # Common modern style: 17.0.10+7
  if ($v -match '^(\d+)\.(\d+)\.(\d+)\+(\d+)$') {
    return [pscustomobject]@{
      Raw=$v; Major=[int]$Matches[1]; Minor=[int]$Matches[2]; SecOrUpd=[int]$Matches[3]; Build=[int]$Matches[4]
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3],[int]$Matches[4]
    }
  }

  # winget/registry often: 17.0.10.7 or 8.0.382.5
  if ($v -match '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
    return [pscustomobject]@{
      Raw=$v; Major=[int]$Matches[1]; Minor=[int]$Matches[2]; SecOrUpd=[int]$Matches[3]; Build=[int]$Matches[4]
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3],[int]$Matches[4]
    }
  }

  # sometimes: 17.0.10 (no build)
  if ($v -match '^(\d+)\.(\d+)\.(\d+)$') {
    return [pscustomobject]@{
      Raw=$v; Major=[int]$Matches[1]; Minor=[int]$Matches[2]; SecOrUpd=[int]$Matches[3]; Build=0
      Key = '{0:D3}.{1:D3}.{2:D5}.{3:D5}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3],0
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

  $runtimeLine = $out | Where-Object { $_ -match '^\s*java\.runtime\.version\s*=' } | Select-Object -First 1
  $homeLine    = $out | Where-Object { $_ -match '^\s*java\.home\s*=' } | Select-Object -First 1

  $runtimeVer = if ($runtimeLine) { (($runtimeLine -split '=',2)[1]).Trim() } else { $null }
  $javaHome   = if ($homeLine)    { (($homeLine    -split '=',2)[1]).Trim() } else { $null }

  [pscustomobject]@{ RuntimeVersion=$runtimeVer; JavaHome=$javaHome }
}

function Is-JreHome {
  param([Parameter(Mandatory)][string]$JavaHome)

  # If javac exists, it's a JDK (not a pure JRE)
  $javac = Join-Path $JavaHome 'bin\javac.exe'
  if (Test-Path $javac) { return $false }

  # Java 8 often: ...\jre1.8.0_xxx  (definitely JRE)
  if ($JavaHome -match '\\jre1\.8\.0_') { return $true }

  # Many JREs still have java.exe in bin
  $java = Join-Path $JavaHome 'bin\java.exe'
  if (Test-Path $java) { return $true }

  return $false
}

function Find-JavaCandidates {
  $candidates = @()

  if ($env:JAVA_HOME) {
    $p = Join-Path $env:JAVA_HOME 'bin\java.exe'
    if (Test-Path $p) { $candidates += $p }
  }
  if ($env:JRE_HOME) {
    $p = Join-Path $env:JRE_HOME 'bin\java.exe'
    if (Test-Path $p) { $candidates += $p }
  }

  $cmd = Get-Command java.exe -ErrorAction SilentlyContinue
  if ($cmd -and (Test-Path $cmd.Source)) { $candidates += $cmd.Source }

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
      if (Test-Path $p) { $candidates += $p }
    }
  }

  @($candidates | Sort-Object -Unique)
}

function Get-InstalledJreForFamily {
  param([Parameter(Mandatory)][int]$Family)

  $best = $null
  $all  = @()

  foreach ($javaExe in (Find-JavaCandidates)) {
    $settings = $null
    try { $settings = Get-JavaSettings -JavaExePath $javaExe } catch { continue }
    if (-not $settings.RuntimeVersion -or -not $settings.JavaHome) { continue }

    $norm = Try-Normalize -VersionString $settings.RuntimeVersion
    if (-not $norm) { continue }

    $isJre = Is-JreHome -JavaHome $settings.JavaHome

    $entry = [pscustomobject]@{
      JavaExe=$javaExe
      JavaHome=$settings.JavaHome
      IsJre=$isJre
      Version=$settings.RuntimeVersion
      Norm=$norm
    }
    $all += $entry

    if ($isJre -and $norm.Major -eq $Family) {
      if (-not $best -or $norm.Key -gt $best.Norm.Key) { $best = $entry }
    }
  }

  [pscustomobject]@{ Best=$best; All=$all }
}

function Get-WingetId {
  param([Parameter(Mandatory)][string]$Vendor, [Parameter(Mandatory)][int]$Family)

  if ($Vendor -eq 'Temurin') { return "EclipseAdoptium.Temurin.$Family.JRE" }

  # Oracle JRE is typically only for Java 8
  if ($Vendor -eq 'Oracle') {
    if ($Family -ne 8) { throw "Oracle JRE via winget is typically only available for Java 8. Use Vendor=Temurin for 11/17/21." }
    return "Oracle.JavaRuntimeEnvironment"
  }

  throw "Unknown vendor: $Vendor"
}

function Get-WingetAvailableVersion {
  param([Parameter(Mandatory)][string]$WingetId)
  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { throw "winget.exe not found." }

  try {
    $json = & $winget.Source show --exact --id $WingetId --output json 2>$null
    if ($LASTEXITCODE -eq 0 -and $json) {
      $obj = $json | ConvertFrom-Json
      if ($obj.Version) { return [string]$obj.Version }
      if ($obj.Versions -and $obj.Versions[0]) { return [string]$obj.Versions[0] }
    }
  } catch { }

  $txt = & $winget.Source show --exact --id $WingetId 2>&1
  $line = $txt | Where-Object { $_ -match '^\s*Version:\s*' } | Select-Object -First 1
  if ($line) { return ($line -replace '^\s*Version:\s*','').Trim() }

  throw "Unable to determine available version for winget id '$WingetId'."
}

function InstallOrUpgrade-WithWinget {
  param([Parameter(Mandatory)][string]$WingetId)

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) { throw "winget.exe not found." }

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

function Install-TemurinJreViaMsi {
  param([Parameter(Mandatory)][int]$Family)

  Ensure-Tls12ForPs5

  # Adoptium "latest installer" endpoint (MSI)
  $url = "https://api.adoptium.net/v3/installer/latest/$Family/ga/windows/x64/jre/hotspot/normal/eclipse?project=jdk"

  $dir = Join-Path $env:TEMP "JavaUpdate"
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  $msi = Join-Path $dir ("temurin-jre-{0}-latest.msi" -f $Family)

  if ($PSCmdlet.ShouldProcess($url, ("Download MSI to {0}" -f $msi))) {
    Invoke-WebRequest -Uri $url -OutFile $msi -UseBasicParsing
  }

  $args = "/i `"$msi`" /qn /norestart"
  if ($PSCmdlet.ShouldProcess(("msiexec.exe {0}" -f $args), "Install Temurin JRE MSI")) {
    $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) { throw "msiexec failed (exit $($p.ExitCode)) installing $msi" }
  }
}

function Get-UninstallEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
      Where-Object { $_.DisplayName -and -not $_.SystemComponent } |
      ForEach-Object {
        [pscustomobject]@{
          DisplayName=$_.DisplayName
          DisplayVersion=$_.DisplayVersion
          UninstallString=$_.UninstallString
          QuietUninstallString=$_.QuietUninstallString
        }
      }
  }
}

function Get-MsiProductCodeFromUninstallString {
  param([string]$UninstallString)
  if (-not $UninstallString) { return $null }
  $m = [regex]::Match($UninstallString, '\{[0-9A-Fa-f\-]{36}\}')
  if ($m.Success) { $m.Value } else { $null }
}

function Try-ExtractVersionFromUninstallEntry {
  param([Parameter(Mandatory)]$Entry)
  if ($Entry.DisplayVersion) { return [string]$Entry.DisplayVersion }

  $name = [string]$Entry.DisplayName
  $m8  = [regex]::Match($name, '(1\.8\.0_\d+(?:-b\d+)?)')
  if ($m8.Success) { return $m8.Groups[1].Value }

  $mV  = [regex]::Match($name, '(\d+\.\d+\.\d+(?:\.\d+)?(?:\+\d+)?)')
  if ($mV.Success) { return $mV.Groups[1].Value }

  $null
}

function Uninstall-EntrySilently {
  param([Parameter(Mandatory)]$Entry)

  if ($Entry.QuietUninstallString) {
    if ($PSCmdlet.ShouldProcess($Entry.DisplayName, "Uninstall (quiet)")) {
      $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $Entry.QuietUninstallString -Wait -PassThru -NoNewWindow
      return $p.ExitCode
    }
    return 0
  }

  $code = Get-MsiProductCodeFromUninstallString -UninstallString $Entry.UninstallString
  if ($code) {
    $args = "/x $code /qn /norestart"
    if ($PSCmdlet.ShouldProcess($Entry.DisplayName, ("msiexec {0}" -f $args))) {
      $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
      return $p.ExitCode
    }
    return 0
  }

  Write-Warning ("Cannot silently uninstall (no MSI code / no quiet string): {0}" -f $Entry.DisplayName)
  return 0
}

function Cleanup-EnvVarsAndPath {
  param([string]$KeepJavaHome)

  $envKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'

  $javaHome = (Get-ItemProperty -Path $envKey -Name JAVA_HOME -ErrorAction SilentlyContinue).JAVA_HOME
  $jreHome  = (Get-ItemProperty -Path $envKey -Name JRE_HOME  -ErrorAction SilentlyContinue).JRE_HOME

  if ($javaHome -and -not (Test-Path $javaHome)) {
    if ($PSCmdlet.ShouldProcess("JAVA_HOME", ("Remove stale value '{0}'" -f $javaHome))) {
      Remove-ItemProperty -Path $envKey -Name JAVA_HOME -ErrorAction SilentlyContinue
    }
  }
  if ($jreHome -and -not (Test-Path $jreHome)) {
    if ($PSCmdlet.ShouldProcess("JRE_HOME", ("Remove stale value '{0}'" -f $jreHome))) {
      Remove-ItemProperty -Path $envKey -Name JRE_HOME -ErrorAction SilentlyContinue
    }
  }

  if ($KeepJavaHome -and (Test-Path $KeepJavaHome)) {
    if ($PSCmdlet.ShouldProcess("JAVA_HOME", ("Set to '{0}'" -f $KeepJavaHome))) {
      Set-ItemProperty -Path $envKey -Name JAVA_HOME -Value $KeepJavaHome
    }
    if ($PSCmdlet.ShouldProcess("JRE_HOME", ("Set to '{0}'" -f $KeepJavaHome))) {
      Set-ItemProperty -Path $envKey -Name JRE_HOME -Value $KeepJavaHome
    }
  }

  $pathVal = (Get-ItemProperty -Path $envKey -Name Path -ErrorAction SilentlyContinue).Path
  if ($pathVal) {
    $parts = $pathVal -split ';' | Where-Object { $_ -and $_.Trim() -ne '' }
    $newParts = @()

    foreach ($p in $parts) {
      $pp = $p.Trim()
      $looksJava = ($pp -match '\\java\\' -or $pp -match '\\javapath' -or $pp -match '\\eclipse adoptium\\' -or $pp -match '\\temurin\\' -or $pp -match '\\adoptium\\')
      if ($looksJava -and -not (Test-Path $pp)) { continue }
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

# -------- main --------
if (-not (Test-IsAdmin)) { throw "Run this script as Administrator." }

Ensure-LogFolder
try { Start-Transcript -Path $LogPath -Append | Out-Null } catch { }

try {
  if (-not $Force) {
    $javaProcs = Get-Process -Name java,javaw,javaws -ErrorAction SilentlyContinue
    if ($javaProcs) { throw "Java processes are running. Close apps using Java or re-run with -Force." }
  }

  $wingetId = Get-WingetId -Vendor $Vendor -Family $TargetFamily
  Write-Host ("Managing JRE: Vendor={0} Family={1} wingetId={2}" -f $Vendor, $TargetFamily, $wingetId)
  Write-Host ""

  $scan = Get-InstalledJreForFamily -Family $TargetFamily
  $installed = $scan.Best

  if ($installed) {
    Write-Host ("Best matching installed JRE {0}: {1} @ {2}" -f $TargetFamily, $installed.Version, $installed.JavaHome)
  } else {
    Write-Host ("No matching JRE {0} found." -f $TargetFamily)
  }

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  $availableVer = $null
  if ($winget) {
    try {
      $availableVer = Get-WingetAvailableVersion -WingetId $wingetId
      Write-Host ("Latest available via winget: {0}" -f $availableVer)
    } catch {
      Write-Warning $_.Exception.Message
    }
  } else {
    Write-Warning "winget.exe not found."
  }

  $needsUpdate = $false
  if (-not $installed) {
    $needsUpdate = $true
  } elseif ($availableVer) {
    $i = $installed.Norm
    $a = Try-Normalize -VersionString $availableVer
    $needsUpdate = if ($a) { $i.Key -lt $a.Key } else { $true }
  } else {
    $needsUpdate = (-not $installed)
  }

  if ($ReportOnly) {
    Write-Host ""
    Write-Host ("ReportOnly: Update required? {0}" -f $needsUpdate)
    exit ($(if ($needsUpdate) { 2 } else { 0 }))
  }

  $didUpdate = $false

  if ($needsUpdate) {
    if ($winget) {
      Write-Host ""
      Write-Host "Updating/Installing via winget..."
      InstallOrUpgrade-WithWinget -WingetId $wingetId
      $didUpdate = $true
    } else {
      if ($Vendor -ne 'Temurin') { throw "winget missing and MSI fallback only supports Vendor=Temurin." }
      if (-not $UseMsiFallback) { throw "winget missing and -UseMsiFallback not set." }

      Write-Host ""
      Write-Host "winget missing - installing via Temurin MSI fallback..."
      Install-TemurinJreViaMsi -Family $TargetFamily
      $didUpdate = $true
    }
  } else {
    Write-Host ""
    Write-Host "No update required."
  }

  # Re-scan
  $scanAfter = Get-InstalledJreForFamily -Family $TargetFamily
  $keep = $scanAfter.Best
  $keepNorm = if ($keep) { $keep.Norm } else { $null }

  if ($keep) {
    Write-Host ("Post-action best JRE {0}: {1} @ {2}" -f $TargetFamily, $keep.Version, $keep.JavaHome)
  } else {
    Write-Warning ("After update, still no matching JRE {0} detected." -f $TargetFamily)
  }

  if ($RemoveOlder -and $keepNorm) {
    Write-Host ""
    Write-Host ("Removing older JRE installs for family {0} (best-effort)..." -f $TargetFamily)

    $entries = Get-UninstallEntries

    $candidates = $entries | Where-Object {
      ($_.DisplayName -match 'JRE|Runtime Environment|Temurin.*JRE|Adoptium.*JRE|Java\(TM\) SE Runtime') -and
      ( ($Vendor -eq 'Temurin' -and $_.DisplayName -match 'Temurin|Eclipse Adoptium|Adoptium') -or
        ($Vendor -eq 'Oracle'  -and $_.DisplayName -match 'Oracle|Java') )
    }

    foreach ($e in $candidates) {
      $verStr = Try-ExtractVersionFromUninstallEntry -Entry $e
      $n = Try-Normalize -VersionString $verStr
      if (-not $n) { continue }
      if ($n.Major -ne $TargetFamily) { continue }

      if ($n.Key -lt $keepNorm.Key) {
        Write-Host ("- Uninstalling older: {0} ({1})" -f $e.DisplayName, $verStr)
        $code = Uninstall-EntrySilently -Entry $e
        if ($code -ne 0) { Write-Warning ("Uninstall exit code {0} for {1}" -f $code, $e.DisplayName) }
      }
    }
  }

  if ($Cleanup -or $RemoveOlder) {
    Write-Host ""
    Write-Host "Cleanup: env vars + PATH..."
    $keepHome = if ($keep) { $keep.JavaHome } else { $null }
    Cleanup-EnvVarsAndPath -KeepJavaHome $keepHome
  }

  exit ($(if ($didUpdate) { 1 } else { 0 }))
}
finally {
  try { Stop-Transcript | Out-Null } catch { }
}