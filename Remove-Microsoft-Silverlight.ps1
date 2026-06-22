#Requires -Version 5.1

<#
Author: Peter Opeyemi James / Nexus Open Systems Ltd
Script: Remove-Microsoft-Silverlight-Datto.ps1
Date: 2026-05-07

Purpose:
- Datto RMM component script to detect and remove Microsoft Silverlight from Windows endpoints.
- Designed to reduce the attack surface from an end-of-support Microsoft component.
- Uses Add/Remove Programs registry inventory rather than Win32_Product.
- Handles MSI uninstall strings silently and does not force a reboot.
- Includes ReportOnly mode, optional process closing, optional residual cleanup, logging, and post-checks.

Datto RMM environment variables:
  Silverlight_ReportOnly           Optional. Default: false. Inventory/log only; no removal.
  Silverlight_RemoveSDK            Optional. Default: true. Remove Silverlight SDK/developer components as well.
  Silverlight_CloseProcesses       Optional. Default: false. Close browsers/Silverlight-related processes before uninstall.
  Silverlight_AggressiveCleanup    Optional. Default: true. Remove leftover Silverlight folders/registry keys after uninstall.
  Silverlight_MatchAnyPublisher    Optional. Default: false. If true, remove any ARP entry containing Silverlight, not only Microsoft-looking entries.
  Silverlight_FailIfRemaining      Optional. Default: true. Exit 3 if Silverlight ARP entries remain and no reboot-required code was returned.
  Silverlight_LogPath              Optional. Default: %ProgramData%\NexusOpenSystems\Silverlight\Remove-Silverlight.log

Exit codes:
  0 = success / no action / report only / reboot required but not forced
  3 = error or Silverlight still present after remediation when Silverlight_FailIfRemaining=true
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
  [switch]$ReportOnly,
  [switch]$RemoveSDK,
  [switch]$CloseProcesses,
  [switch]$AggressiveCleanup,
  [switch]$MatchAnyPublisher,
  [switch]$FailIfRemaining,
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\Silverlight\Remove-Silverlight.log"
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
    '1'     { return $true }
    'true'  { return $true }
    'yes'   { return $true }
    'y'     { return $true }
    'on'    { return $true }
    '0'     { return $false }
    'false' { return $false }
    'no'    { return $false }
    'n'     { return $false }
    'off'   { return $false }
    default { return $Default }
  }
}

if (-not $PSBoundParameters.ContainsKey('ReportOnly'))        { $ReportOnly        = Get-EnvBool 'Silverlight_ReportOnly' $false }
if (-not $PSBoundParameters.ContainsKey('RemoveSDK'))         { $RemoveSDK         = Get-EnvBool 'Silverlight_RemoveSDK' $true }
if (-not $PSBoundParameters.ContainsKey('CloseProcesses'))    { $CloseProcesses    = Get-EnvBool 'Silverlight_CloseProcesses' $false }
if (-not $PSBoundParameters.ContainsKey('AggressiveCleanup')) { $AggressiveCleanup = Get-EnvBool 'Silverlight_AggressiveCleanup' $true }
if (-not $PSBoundParameters.ContainsKey('MatchAnyPublisher')) { $MatchAnyPublisher = Get-EnvBool 'Silverlight_MatchAnyPublisher' $false }
if (-not $PSBoundParameters.ContainsKey('FailIfRemaining'))   { $FailIfRemaining   = Get-EnvBool 'Silverlight_FailIfRemaining' $true }
if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'Silverlight_LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- logging / utilities ----------------
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

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Normalize-List($Object) {
  @($Object) | Where-Object { $_ -ne $null }
}

function Get-FirstNonBlank {
  param([string[]]$Values)
  foreach ($v in $Values) {
    if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
  }
  return $null
}

function Get-ProductCodeFromString([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
  $m = [regex]::Match($Text, '\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}')
  if ($m.Success) { return $m.Value }
  return $null
}

function Format-ArpEntry($Entry) {
  $publisher = if ($Entry.Publisher) { $Entry.Publisher } else { '<publisher unknown>' }
  $version = if ($Entry.DisplayVersion) { $Entry.DisplayVersion } else { '<version unknown>' }
  "{0} | Version={1} | Publisher={2} | Registry={3}" -f $Entry.DisplayName, $version, $publisher, $Entry.RegistryPath
}

# ---------------- ARP inventory ----------------
function Get-ArpEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | ForEach-Object {
      if ($_.DisplayName) {
        [pscustomobject]@{
          DisplayName          = [string]$_.DisplayName
          DisplayVersion       = [string]$_.DisplayVersion
          Publisher            = [string]$_.Publisher
          QuietUninstallString = [string]$_.QuietUninstallString
          UninstallString      = [string]$_.UninstallString
          WindowsInstaller     = $_.WindowsInstaller
          RegistryKeyName      = [string]$_.PSChildName
          RegistryPath         = [string]$_.PSPath
        }
      }
    }
  }
}

function Test-IsMicrosoftSilverlightEntry {
  param(
    [Parameter(Mandatory)]$Entry,
    [switch]$RemoveSDK,
    [switch]$MatchAnyPublisher
  )

  $name = [string]$Entry.DisplayName
  if ([string]::IsNullOrWhiteSpace($name)) { return $false }
  if ($name -notmatch '(?i)\bsilverlight\b') { return $false }

  # Keep SDK/developer components configurable because some organisations still inventory them separately.
  if ($name -match '(?i)\b(SDK|Software Development Kit|Tools|Toolkit)\b' -and -not $RemoveSDK) {
    return $false
  }

  if ($MatchAnyPublisher) { return $true }

  $publisher = [string]$Entry.Publisher
  if ($publisher -match '(?i)\bMicrosoft\b') { return $true }
  if ($name -match '(?i)^\s*Microsoft\s+Silverlight\b') { return $true }

  return $false
}

function Get-SilverlightArpEntries {
  param(
    [switch]$RemoveSDK,
    [switch]$MatchAnyPublisher
  )

  $items = @()
  foreach ($e in Get-ArpEntries) {
    if (Test-IsMicrosoftSilverlightEntry -Entry $e -RemoveSDK:$RemoveSDK -MatchAnyPublisher:$MatchAnyPublisher) {
      $items += $e
    }
  }

  return ($items | Sort-Object DisplayName, DisplayVersion, RegistryPath -Unique)
}

# ---------------- uninstall handling ----------------
function Normalize-UninstallCommand {
  param([Parameter(Mandatory)]$Entry)

  $cmd = Get-FirstNonBlank @($Entry.QuietUninstallString, $Entry.UninstallString)
  $productCode = Get-ProductCodeFromString $cmd
  if (-not $productCode) { $productCode = Get-ProductCodeFromString $Entry.RegistryKeyName }

  $looksMsi = $false
  if ($cmd -match '(?i)\bmsiexec(\.exe)?\b') { $looksMsi = $true }
  if ($Entry.WindowsInstaller -eq 1) { $looksMsi = $true }
  if ($productCode -and $Entry.RegistryKeyName -eq $productCode) { $looksMsi = $true }

  if ($productCode -and $looksMsi) {
    return @{
      Exe  = 'msiexec.exe'
      Args = "/x $productCode /qn /norestart"
    }
  }

  if ($cmd -match '(?i)\bmsiexec(\.exe)?\b') {
    $args = $cmd -replace '(?i)^.*?msiexec(\.exe)?\s*', ''
    $args = $args -replace '^(?i)\s*([/-])i\b', '$1X'

    if ($args -notmatch '(?i)(^|\s)(/quiet|/qn)\b') { $args += ' /qn' }
    if ($args -notmatch '(?i)(^|\s)/norestart\b') { $args += ' /norestart' }

    return @{
      Exe  = 'msiexec.exe'
      Args = $args
    }
  }

  if ([string]::IsNullOrWhiteSpace($cmd)) {
    return $null
  }

  # Fallback for non-MSI uninstallers. Silverlight is normally MSI-based, but this avoids doing nothing
  # when a vendor-style uninstall string is present. Common silent flags are appended only if absent.
  $cmdToRun = $cmd.Trim()
  if ($cmdToRun -notmatch '(?i)(^|\s)(/quiet|/qn|/silent|/s)\b') { $cmdToRun += ' /quiet' }
  if ($cmdToRun -notmatch '(?i)(^|\s)/norestart\b') { $cmdToRun += ' /norestart' }

  return @{
    Exe  = $env:ComSpec
    Args = "/c `"$cmdToRun`""
  }
}

function Uninstall-ArpEntry {
  param([Parameter(Mandatory)]$Entry)

  $n = Normalize-UninstallCommand -Entry $Entry
  if (-not $n) {
    Write-Log "WARNING: No uninstall command found for: $(Format-ArpEntry $Entry)"
    return $null
  }

  if (-not $PSCmdlet.ShouldProcess($Entry.DisplayName, 'Uninstall Microsoft Silverlight component')) {
    Write-Log "WhatIf/Confirm prevented uninstall of '$($Entry.DisplayName)'"
    return $null
  }

  Write-Log "Uninstalling: $(Format-ArpEntry $Entry)"
  Write-Log "Command: $($n.Exe) $($n.Args)"

  $p = Start-Process -FilePath $n.Exe -ArgumentList $n.Args -Wait -PassThru -NoNewWindow

  if (@(0,1605,1614,3010) -contains $p.ExitCode) {
    Write-Log "Uninstall exit=$($p.ExitCode) (ok)"
  }
  else {
    Write-Log "WARNING: Uninstall exit=$($p.ExitCode) for '$($Entry.DisplayName)'"
  }

  return $p.ExitCode
}

# ---------------- optional process close ----------------
function Close-SilverlightRelatedProcesses {
  $names = @(
    'sllauncher',
    'iexplore',
    'msedge',
    'chrome',
    'firefox',
    'lync',
    'communicator'
  )

  $procs = @(Get-Process -Name $names -ErrorAction SilentlyContinue)
  if ($procs.Count -eq 0) {
    Write-Log 'No browser/Silverlight-related processes found to close.'
    return
  }

  foreach ($proc in $procs) {
    if ($PSCmdlet.ShouldProcess($proc.ProcessName, 'Stop process before Silverlight uninstall')) {
      try {
        Write-Log "Stopping process: $($proc.ProcessName) PID=$($proc.Id)"
        Stop-Process -Id $proc.Id -Force -ErrorAction Stop
      }
      catch {
        Write-Log "WARNING: Failed to stop process $($proc.ProcessName) PID=$($proc.Id): $($_.Exception.Message)"
      }
    }
    else {
      Write-Log "WhatIf/Confirm prevented stopping process '$($proc.ProcessName)'"
    }
  }
}

# ---------------- residual cleanup ----------------
function Get-SilverlightResidualPaths {
  $items = @()

  $folders = @()
  if ($env:ProgramFiles) { $folders += (Join-Path $env:ProgramFiles 'Microsoft Silverlight') }
  $pf86 = ${env:ProgramFiles(x86)}
  if ($pf86) { $folders += (Join-Path $pf86 'Microsoft Silverlight') }

  foreach ($folder in ($folders | Sort-Object -Unique)) {
    if ($folder -and (Test-Path $folder)) {
      $items += [pscustomobject]@{ Type = 'Folder'; Path = $folder }
    }
  }

  $regPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Silverlight',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Silverlight'
  )

  foreach ($rp in $regPaths) {
    if (Test-Path $rp) {
      $items += [pscustomobject]@{ Type = 'Registry'; Path = $rp }
    }
  }

  return $items
}

function Remove-SilverlightResidualPath {
  param([Parameter(Mandatory)]$Residual)

  if (-not (Test-Path $Residual.Path)) { return $true }

  if (-not $PSCmdlet.ShouldProcess($Residual.Path, "Remove Silverlight residual $($Residual.Type)")) {
    Write-Log "WhatIf/Confirm prevented residual cleanup of '$($Residual.Path)'"
    return $false
  }

  try {
    Remove-Item -Path $Residual.Path -Recurse -Force -ErrorAction Stop
    Write-Log "Removed residual $($Residual.Type): $($Residual.Path)"
    return $true
  }
  catch {
    Write-Log "WARNING: Failed to remove residual $($Residual.Type) '$($Residual.Path)': $($_.Exception.Message)"
    return $false
  }
}

# ---------------- MAIN ----------------
try {
  if (-not (Test-IsAdmin)) {
    throw 'Run elevated as Administrator or Datto/SYSTEM.'
  }

  Write-Log ("Starting Silverlight remediation. ReportOnly={0} RemoveSDK={1} CloseProcesses={2} AggressiveCleanup={3} MatchAnyPublisher={4} FailIfRemaining={5}" -f `
    $ReportOnly, $RemoveSDK, $CloseProcesses, $AggressiveCleanup, $MatchAnyPublisher, $FailIfRemaining)

  $found = @(Get-SilverlightArpEntries -RemoveSDK:$RemoveSDK -MatchAnyPublisher:$MatchAnyPublisher)
  Write-Log "Silverlight ARP entries found: $($found.Count)"
  foreach ($item in $found) {
    Write-Log "  Found: $(Format-ArpEntry $item)"
  }

  $residualBefore = @(Get-SilverlightResidualPaths)
  Write-Log "Silverlight residual paths found before cleanup: $($residualBefore.Count)"
  foreach ($r in $residualBefore) {
    Write-Log "  Residual $($r.Type): $($r.Path)"
  }

  if ($found.Count -eq 0 -and $residualBefore.Count -eq 0) {
    Write-Log 'Microsoft Silverlight was not detected. No action required.'
    exit 0
  }

  if ($ReportOnly) {
    Write-Log 'ReportOnly: no changes made.'
    if ($found.Count -gt 0) {
      Write-Log 'ReportOnly: would silently uninstall the Silverlight ARP entries listed above.'
    }
    if ($AggressiveCleanup -and $residualBefore.Count -gt 0) {
      Write-Log 'ReportOnly: would remove residual Silverlight folders/registry keys listed above.'
    }
    exit 0
  }

  if ($CloseProcesses) {
    Close-SilverlightRelatedProcesses
  }
  else {
    Write-Log 'CloseProcesses=false. Browser/Silverlight-related processes will not be force closed.'
  }

  $rebootRequired = $false
  $uninstallFailures = 0

  foreach ($item in $found) {
    $exitCode = Uninstall-ArpEntry -Entry $item
    if ($exitCode -eq 3010) { $rebootRequired = $true }
    elseif ($null -eq $exitCode) { $uninstallFailures++ }
    elseif (@(0,1605,1614) -notcontains $exitCode) { $uninstallFailures++ }
  }

  Start-Sleep -Seconds 2

  if ($AggressiveCleanup) {
    $residuals = @(Get-SilverlightResidualPaths)
    if ($residuals.Count -eq 0) {
      Write-Log 'No residual Silverlight folders/registry keys found for cleanup.'
    }
    else {
      Write-Log "AggressiveCleanup=true. Residual items to remove: $($residuals.Count)"
      foreach ($r in $residuals) {
        [void](Remove-SilverlightResidualPath -Residual $r)
      }
    }
  }
  else {
    Write-Log 'AggressiveCleanup=false. Residual folders/registry keys will be logged but not removed.'
  }

  Start-Sleep -Seconds 2

  $after = @(Get-SilverlightArpEntries -RemoveSDK:$RemoveSDK -MatchAnyPublisher:$MatchAnyPublisher)
  $residualAfter = @(Get-SilverlightResidualPaths)

  Write-Log "Post-check: Silverlight ARP entries remaining: $($after.Count)"
  foreach ($item in $after) {
    Write-Log "  Remaining ARP: $(Format-ArpEntry $item)"
  }

  Write-Log "Post-check: Silverlight residual paths remaining: $($residualAfter.Count)"
  foreach ($r in $residualAfter) {
    Write-Log "  Remaining residual $($r.Type): $($r.Path)"
  }

  if ($rebootRequired) {
    Write-Log 'At least one uninstall returned 3010. A reboot is required to finish cleanup, but this script does not force restart.'
  }

  if ($uninstallFailures -gt 0) {
    Write-Log "WARNING: One or more uninstall attempts failed or had no uninstall string. Count=$uninstallFailures"
  }

  if ($after.Count -gt 0 -and $FailIfRemaining -and -not $rebootRequired) {
    Write-Log 'ERROR: Microsoft Silverlight still appears in Add/Remove Programs after remediation.'
    exit 3
  }

  Write-Log 'Done.'
  exit 0
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)"
  exit 3
}
