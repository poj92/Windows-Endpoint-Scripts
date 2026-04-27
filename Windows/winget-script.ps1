#Requires -Version 5.1
[CmdletBinding()]
param(
  [switch]$ReportOnly,
  [switch]$Install_Winget_if_Not_Avaialble,
  [string]$LogPath = "$env:ProgramData\Datto\Logs\Winget-UpgradeAll.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$script:ActionLines = @()
$script:ErrorLines  = @()

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

if (-not $PSBoundParameters.ContainsKey('ReportOnly')) {
  $ReportOnly = Get-EnvBool 'ReportOnly' $false
}

if (-not $PSBoundParameters.ContainsKey('Install_Winget_if_Not_Avaialble')) {
  $Install_Winget_if_Not_Avaialble = Get-EnvBool 'Install_Winget_if_Not_Avaialble' $false
}

if (-not $PSBoundParameters.ContainsKey('LogPath')) {
  $lp = Get-Env 'LogPath'
  if ($lp) { $LogPath = $lp }
}

# ---------------- Logging / output ----------------
function Ensure-LogFile {
  try {
    $dir = Split-Path -Parent $LogPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
  } catch {}
}

function Write-Log([string]$Message) {
  Ensure-LogFile
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] {1}" -f $ts, $Message
  try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch {}
}

function Add-Action([string]$Text) {
  $script:ActionLines += $Text
  Write-Log $Text
}

function Add-ErrorText([string]$Text) {
  $script:ErrorLines += $Text
  Write-Log ("ERROR: {0}" -f $Text)
}

function Out-Result {
  param(
    [string]$Status,
    [string]$Summary,
    [int]$ExitCode
  )

  Write-Host ("STATUS={0}" -f $Status)
  Write-Host ("SUMMARY={0}" -f $Summary)
  Write-Host ("LOG={0}" -f $LogPath)

  foreach ($line in $script:ActionLines) {
    Write-Host $line
  }

  foreach ($line in $script:ErrorLines) {
    Write-Host ("ERROR: {0}" -f $line)
  }

  exit $ExitCode
}

function Write-BlockToLog {
  param(
    [string]$Title,
    [string]$Content
  )

  Write-Log ("----- BEGIN {0} -----" -f $Title)
  if ([string]::IsNullOrWhiteSpace($Content)) {
    Write-Log "[empty]"
  } else {
    foreach ($line in ($Content -split "`r?`n")) {
      Write-Log $line
    }
  }
  Write-Log ("----- END {0} -----" -f $Title)
}

# ---------------- Winget helpers ----------------
function Get-WingetPath {
  $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  $localAppDataPath = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
  if (Test-Path -LiteralPath $localAppDataPath) {
    return $localAppDataPath
  }

  $windowsApps = Join-Path $env:ProgramFiles 'WindowsApps'
  if (Test-Path -LiteralPath $windowsApps) {
    $candidate = Get-ChildItem -Path $windowsApps -Filter 'Microsoft.DesktopAppInstaller_*__8wekyb3d8bbwe' -Directory -ErrorAction SilentlyContinue |
      Sort-Object Name -Descending |
      Select-Object -First 1

    if ($candidate) {
      $wingetExe = Join-Path $candidate.FullName 'winget.exe'
      if (Test-Path -LiteralPath $wingetExe) {
        return $wingetExe
      }
    }
  }

  return $null
}

function Join-ArgumentList {
  param([string[]]$Parts)

  $safeParts = foreach ($p in $Parts) {
    if ($null -eq $p) { continue }
    if ($p -match '\s') {
      '"{0}"' -f ($p -replace '"','\"')
    } else {
      $p
    }
  }

  return ($safeParts -join ' ')
}

function Invoke-LoggedProcess {
  param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [Parameter(Mandatory=$true)][string[]]$Arguments,
    [Parameter(Mandatory=$true)][string]$StepName,
    [int]$TimeoutSec = 7200
  )

  $tempDir = Join-Path $env:ProgramData 'Datto\Temp\WingetUpgradeAll'
  if (-not (Test-Path -LiteralPath $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
  }

  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
  $stdoutPath = Join-Path $tempDir ("{0}-{1}-stdout.log" -f $StepName, $stamp)
  $stderrPath = Join-Path $tempDir ("{0}-{1}-stderr.log" -f $StepName, $stamp)

  $argString = Join-ArgumentList -Parts $Arguments
  Add-Action ("START {0} exe=[{1}] args=[{2}]" -f $StepName, $Exe, $argString)

  $proc = Start-Process -FilePath $Exe -ArgumentList $argString -PassThru -NoNewWindow `
    -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while (-not $proc.HasExited) {
    Start-Sleep -Seconds 2
    $proc.Refresh()

    if ($sw.Elapsed.TotalSeconds -ge $TimeoutSec) {
      try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch {}
      Add-ErrorText ("{0} timed out after {1} seconds" -f $StepName, $TimeoutSec)
      return [pscustomobject]@{
        ExitCode = -1
        StdOut   = if (Test-Path $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue } else { '' }
        StdErr   = if (Test-Path $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue } else { '' }
        TimedOut = $true
      }
    }
  }

  $result = [pscustomobject]@{
    ExitCode = [int]$proc.ExitCode
    StdOut   = if (Test-Path $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue } else { '' }
    StdErr   = if (Test-Path $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue } else { '' }
    TimedOut = $false
  }

  Add-Action ("END {0} exitcode={1}" -f $StepName, $result.ExitCode)

  try { Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue } catch {}
  try { Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue } catch {}

  return $result
}

function Install-WingetIfNeeded {
  $existing = Get-WingetPath
  if ($existing) {
    Add-Action ("winget already available at [{0}]" -f $existing)
    return $existing
  }

  if (-not [bool]$Install_Winget_if_Not_Avaialble) {
    Add-ErrorText "winget.exe was not found and Install_Winget_if_Not_Avaialble is False"
    return $null
  }

  Add-Action "winget.exe not found; attempting registration/bootstrap"

  # First try registering App Installer if it is already present
  try {
    Add-Action "Attempting App Installer registration"
    Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
    Start-Sleep -Seconds 5
  } catch {
    Add-Action ("App Installer registration attempt did not complete cleanly: {0}" -f $_.Exception.Message)
  }

  $existing = Get-WingetPath
  if ($existing) {
    Add-Action ("winget became available after registration at [{0}]" -f $existing)
    return $existing
  }

  # Fallback bootstrap path via Microsoft.WinGet.Client
  try {
    Add-Action "Installing Microsoft.WinGet.Client module"
    $progressPreference = 'SilentlyContinue'
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery -Scope AllUsers | Out-Null

    Add-Action "Running Repair-WinGetPackageManager -AllUsers"
    Repair-WinGetPackageManager -AllUsers | Out-Null
    Start-Sleep -Seconds 10
  } catch {
    Add-ErrorText ("Winget bootstrap failed: {0}" -f $_.Exception.Message)
  }

  $existing = Get-WingetPath
  if ($existing) {
    Add-Action ("winget became available after bootstrap at [{0}]" -f $existing)
    return $existing
  }

  Add-ErrorText "winget.exe is still not available after bootstrap attempts"
  return $null
}

# ---------------- Main ----------------
try {
  Add-Action "SCRIPT_VERSION=2026-04-27-WINGET-UPGRADE-ALL-v2"
  Add-Action ("Resolved variables: ReportOnly=[{0}] Install_Winget_if_Not_Avaialble=[{1}] LogPath=[{2}]" -f `
    [bool]$ReportOnly, [bool]$Install_Winget_if_Not_Avaialble, $LogPath)

  $wingetPath = Install-WingetIfNeeded
  if ([string]::IsNullOrWhiteSpace($wingetPath)) {
    Out-Result -Status 'Failed' -Summary 'winget.exe was not found or could not be installed.' -ExitCode 1
  }

  $versionResult = Invoke-LoggedProcess -Exe $wingetPath -Arguments @('--version') -StepName 'WingetVersion' -TimeoutSec 60
  Write-BlockToLog -Title 'WingetVersion STDOUT' -Content $versionResult.StdOut
  Write-BlockToLog -Title 'WingetVersion STDERR' -Content $versionResult.StdErr

  $sourceReset = Invoke-LoggedProcess -Exe $wingetPath -Arguments @('source','reset','--force','--accept-source-agreements','--disable-interactivity') -StepName 'SourceReset' -TimeoutSec 600
  Write-BlockToLog -Title 'SourceReset STDOUT' -Content $sourceReset.StdOut
  Write-BlockToLog -Title 'SourceReset STDERR' -Content $sourceReset.StdErr

  $sourceUpdate = Invoke-LoggedProcess -Exe $wingetPath -Arguments @('source','update','--accept-source-agreements','--disable-interactivity') -StepName 'SourceUpdate' -TimeoutSec 1200
  Write-BlockToLog -Title 'SourceUpdate STDOUT' -Content $sourceUpdate.StdOut
  Write-BlockToLog -Title 'SourceUpdate STDERR' -Content $sourceUpdate.StdErr

  $listArgs = @('list','--upgrade-available','--accept-source-agreements','--disable-interactivity')
  $preScan = Invoke-LoggedProcess -Exe $wingetPath -Arguments $listArgs -StepName 'PreScan' -TimeoutSec 1800
  Write-BlockToLog -Title 'PreScan STDOUT' -Content $preScan.StdOut
  Write-BlockToLog -Title 'PreScan STDERR' -Content $preScan.StdErr

  if ($preScan.TimedOut) {
    Out-Result -Status 'CompletedWithErrors' -Summary 'winget pre-scan timed out.' -ExitCode 1
  }

  if ([bool]$ReportOnly) {
    Out-Result -Status 'ReportOnly' -Summary 'Report-only mode complete. Review the log for upgradeable packages.' -ExitCode 0
  }

  $wingetNativeLog = Join-Path $env:ProgramData ("Datto\Logs\winget-upgrade-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
  $upgradeArgs = @(
    'upgrade',
    '--all',
    '--silent',
    '--accept-package-agreements',
    '--accept-source-agreements',
    '--disable-interactivity',
    '--nowarn',
    '--log', $wingetNativeLog
  )

  Add-Action ("Winget native log path [{0}]" -f $wingetNativeLog)

  $upgradeResult = Invoke-LoggedProcess -Exe $wingetPath -Arguments $upgradeArgs -StepName 'UpgradeAll' -TimeoutSec 7200
  Write-BlockToLog -Title 'UpgradeAll STDOUT' -Content $upgradeResult.StdOut
  Write-BlockToLog -Title 'UpgradeAll STDERR' -Content $upgradeResult.StdErr

  $postScan = Invoke-LoggedProcess -Exe $wingetPath -Arguments $listArgs -StepName 'PostScan' -TimeoutSec 1800
  Write-BlockToLog -Title 'PostScan STDOUT' -Content $postScan.StdOut
  Write-BlockToLog -Title 'PostScan STDERR' -Content $postScan.StdErr

  if ($upgradeResult.TimedOut) {
    Out-Result -Status 'CompletedWithErrors' -Summary 'winget upgrade timed out. Review the log.' -ExitCode 1
  }

  if ($upgradeResult.ExitCode -ne 0) {
    Out-Result -Status 'CompletedWithErrors' -Summary ("winget upgrade exited with code {0}. Review the log." -f $upgradeResult.ExitCode) -ExitCode 1
  }

  Out-Result -Status 'Success' -Summary 'winget upgrade completed. Review the log for details and post-scan results.' -ExitCode 0
}
catch {
  Add-ErrorText ("Fatal exception: {0}" -f $_.Exception.Message)
  Out-Result -Status 'Failed' -Summary 'Script failed unexpectedly.' -ExitCode 1
}
