<#
 Winget Script to auto upgrade packages where possible
 V1.3
 Author: Peter James
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
  [switch]$ReportOnly,
  [switch]$Install_Winget_if_Not_Avaialble,
  [string]$ExcludedPackages,
  [string]$LogPath = "$env:ProgramData\Datto\Logs\Winget-UpgradeAll.log"
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$script:ActionLines = @()
$script:ErrorLines  = @()

# ---------------- Datto env helpers ----------------
function Get-Env([string]$Name) {
  try { return [Environment]::GetEnvironmentVariable($Name, "Process") } catch { return $null }
}

function Get-EnvFirst {
  param([string[]]$Names)

  foreach ($name in $Names) {
    $value = Get-Env $name
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      return $value
    }
  }

  return $null
}

function Get-EnvBool([string]$Name, [bool]$Default = $false) {
  $v = Get-Env $Name
  if ($null -eq $v -or $v -eq "") { return $Default }

  switch (($v.ToString()).Trim().ToLowerInvariant()) {
    "1"     { return $true }
    "true"  { return $true }
    "yes"   { return $true }
    "y"     { return $true }
    "on"    { return $true }
    "0"     { return $false }
    "false" { return $false }
    "no"    { return $false }
    "n"     { return $false }
    "off"   { return $false }
    default { return $Default }
  }
}

if (-not $PSBoundParameters.ContainsKey("ReportOnly")) {
  $ReportOnly = Get-EnvBool "ReportOnly" $false
}

if (-not $PSBoundParameters.ContainsKey("Install_Winget_if_Not_Avaialble")) {
  $Install_Winget_if_Not_Avaialble = Get-EnvBool "Install_Winget_if_Not_Avaialble" $false
}

if (-not $PSBoundParameters.ContainsKey("ExcludedPackages")) {
  $ExcludedPackages = Get-EnvFirst @(
    "Excluded packaages",
    "Excluded packages",
    "ExcludedPackages",
    "Excluded_Packages"
  )
}

if (-not $PSBoundParameters.ContainsKey("LogPath")) {
  $lp = Get-Env "LogPath"
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
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
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

function Add-ActionLines {
  param([string[]]$Lines)

  foreach ($line in $Lines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    Add-Action $line
  }
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

# ---------------- Text / parsing helpers ----------------
function Read-TextFileSmart {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) { return "" }

  try {
    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
    return [System.IO.File]::ReadAllText($Path, $utf8Strict)
  } catch {}

  try {
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
  } catch {}

  try {
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Unicode)
  } catch {}

  try {
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
  } catch {}

  return ""
}

function Normalize-MojibakeLine {
  param([string]$Line)

  if ($null -eq $Line) { return $Line }

  $line = $Line
  $line = $line.Replace("â?¦", "...")
  $line = $line.Replace("Â", "")
  return $line
}

function Test-IsWingetNoticeLine {
  param([string]$Line)

  if ([string]::IsNullOrWhiteSpace($Line)) { return $false }

  $text = $Line.Trim()

  if ($text -match "^The source requires") { return $true }
  if ($text -match "current machine.*geographic region") { return $true }
  if ($text -match "two-letter geographic region") { return $true }
  if ($text -match "function properly.*US") { return $true }
  if ($text -match "Do you agree to all the source agreements") { return $true }
  if ($text -match "Terms of Transaction") { return $true }
  if ($text -match "The source .* requires that you view") { return $true }
  if ($text -match "Privacy Statement") { return $true }
  if ($text -match "Microsoft Store source requires") { return $true }

  return $false
}

function Test-IsProgressLine {
  param([string]$Line)

  if ([string]::IsNullOrWhiteSpace($Line)) { return $false }

  $trimmed = $Line.Trim()

  if ($trimmed.Length -eq 1) {
    $backslash = [string][char]92
    if (@("-", "/", "|", $backslash) -contains $trimmed) {
      return $true
    }
  }

  if ($trimmed -match "\b\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)\b") {
    return $true
  }

  return $false
}

function Test-IsSeparatorNoiseLine {
  param([string]$Line)

  if ([string]::IsNullOrWhiteSpace($Line)) { return $false }

  $compact = ($Line -replace "\s", "")
  if ([string]::IsNullOrWhiteSpace($compact)) { return $false }

  if ($compact -match "^-{5,}$") { return $false }

  if ($compact.Length -gt 5 -and $compact -notmatch "[A-Za-z0-9]") {
    return $true
  }

  return $false
}

function Get-CleanWingetUpgradeLines {
  param([string]$Content)

  if ([string]::IsNullOrWhiteSpace($Content)) {
    return @("[empty]")
  }

  $lines = $Content -split "`r?`n"
  $result = @()

  foreach ($raw in $lines) {
    if ($null -eq $raw) { continue }

    $line = Normalize-MojibakeLine ($raw.TrimEnd())
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    if (Test-IsWingetNoticeLine -Line $line) { continue }
    if (Test-IsProgressLine -Line $line) { continue }
    if (Test-IsSeparatorNoiseLine -Line $line) { continue }

    $result += $line
  }

  if ($result.Count -eq 0) {
    return @("[empty]")
  }

  return $result
}

function Limit-Text {
  param(
    [string]$Text,
    [int]$MaxLength
  )

  if ($null -eq $Text) { return "" }
  if ($Text.Length -le $MaxLength) { return $Text }
  if ($MaxLength -le 3) { return $Text.Substring(0, $MaxLength) }
  return ($Text.Substring(0, $MaxLength - 3) + "...")
}

function Get-UpgradeItemsFromWingetOutput {
  param([string]$Content)

  $cleanLines = Get-CleanWingetUpgradeLines -Content $Content
  $items = @()

  foreach ($line in $cleanLines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    $text = $line.Trim()
    if ($text -eq "[empty]") { continue }
    if (Test-IsWingetNoticeLine -Line $text) { continue }

    if ($text -match "^\s*Name\s+Id\s+Version\s+Available\s+Source\s*$") { continue }
    if ($text -match "^-{5,}$") { continue }
    if ($text -match "^\s*\d+\s+upgrades?\s+available\.?\s*$") { continue }
    if ($text -match "No installed package found") { continue }
    if ($text -match "No available upgrade found") { continue }
    if ($text -match "No applicable update found") { continue }

    $parts = $text -split "\s+"
    if ($parts.Count -lt 5) { continue }

    $source    = $parts[$parts.Count - 1]
    $available = $parts[$parts.Count - 2]
    $version   = $parts[$parts.Count - 3]
    $id        = $parts[$parts.Count - 4]

    if ($parts.Count -gt 4) {
      $nameParts = $parts[0..($parts.Count - 5)]
      $name = ($nameParts -join " ")
    } else {
      $name = ""
    }

    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if ([string]::IsNullOrWhiteSpace($id)) { continue }
    if ([string]::IsNullOrWhiteSpace($version)) { continue }
    if ([string]::IsNullOrWhiteSpace($available)) { continue }
    if ([string]::IsNullOrWhiteSpace($source)) { continue }

    if ($source -notmatch "^(winget|msstore)$") { continue }
    if ($id -eq "Id" -or $version -eq "Version" -or $available -eq "Available" -or $source -eq "Source") { continue }

    $already = @($items | Where-Object { $_.Id -ieq $id })
    if ($already.Count -gt 0) { continue }

    $items += [pscustomobject]@{
      Name      = $name
      Id        = $id
      Version   = $version
      Available = $available
      Source    = $source
    }
  }

  return @($items)
}

# ---------------- Exclusion helpers ----------------
function Get-ExcludedPackageTokens {
  param([string]$RawText)

  if ([string]::IsNullOrWhiteSpace($RawText)) {
    return @()
  }

  $tokens = @()

  foreach ($part in ($RawText -split ",")) {
    $token = $part.Trim()
    $token = $token.Trim('"')
    $token = $token.Trim("'")

    if (-not [string]::IsNullOrWhiteSpace($token)) {
      $tokens += $token
    }
  }

  return @($tokens)
}

function Test-IsPackageExcluded {
  param(
    $Item,
    [string[]]$Tokens
  )

  if ($null -eq $Item) { return $false }
  if ($null -eq $Tokens -or $Tokens.Count -eq 0) { return $false }

  foreach ($token in $Tokens) {
    if ([string]::IsNullOrWhiteSpace($token)) { continue }

    if ($token.Contains("*") -or $token.Contains("?")) {
      if ($Item.Id -like $token) { return $true }
      if ($Item.Name -like $token) { return $true }
      continue
    }

    if ($Item.Id -ieq $token) { return $true }
    if ($Item.Name -ieq $token) { return $true }

    # Convenience matching for friendly names such as "Chrome" or "Office".
    if ($token.Length -ge 3 -and $Item.Name -like ("*" + $token + "*")) {
      return $true
    }
  }

  return $false
}

function Split-ExcludedUpgradeItems {
  param(
    [object[]]$Items,
    [string[]]$ExcludeTokens
  )

  $actionable = @()
  $excluded = @()

  foreach ($item in $Items) {
    if (Test-IsPackageExcluded -Item $item -Tokens $ExcludeTokens) {
      $excluded += $item
    } else {
      $actionable += $item
    }
  }

  return [pscustomobject]@{
    Actionable = @($actionable)
    Excluded   = @($excluded)
  }
}

function Format-UpgradeItems {
  param(
    [object[]]$Items,
    [string]$EmptyMessage = "No upgradeable packages detected."
  )

  if ($null -eq $Items -or $Items.Count -eq 0) {
    return @($EmptyMessage)
  }

  $nameWidth = 40
  $idWidth = 42
  $versionWidth = 18
  $availableWidth = 18
  $sourceWidth = 10

  $fmt = "{0,-$nameWidth} {1,-$idWidth} {2,-$versionWidth} {3,-$availableWidth} {4,-$sourceWidth}"

  $lines = @()
  $lines += ($fmt -f "Name", "Id", "Version", "Available", "Source")
  $lines += ($fmt -f "----", "--", "-------", "---------", "------")

  foreach ($item in $Items) {
    $lines += ($fmt -f `
      (Limit-Text $item.Name $nameWidth), `
      (Limit-Text $item.Id $idWidth), `
      (Limit-Text $item.Version $versionWidth), `
      (Limit-Text $item.Available $availableWidth), `
      (Limit-Text $item.Source $sourceWidth))
  }

  $lines += ("Total packages: {0}" -f $Items.Count)

  return $lines
}

function Test-ItemStillUpgradeable {
  param(
    [string]$Id,
    [object[]]$RemainingItems
  )

  if ([string]::IsNullOrWhiteSpace($Id)) { return $false }

  foreach ($item in $RemainingItems) {
    if ($item.Id -ieq $Id) {
      return $true
    }
  }

  return $false
}

function Get-SafeStepName {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return "Package"
  }

  $safe = $Text -replace "[^A-Za-z0-9_.-]", "_"
  if ($safe.Length -gt 60) {
    $safe = $safe.Substring(0, 60)
  }

  return $safe
}

# ---------------- Winget helpers ----------------
function Get-WingetPath {
  $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  $localAppDataPath = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\winget.exe"
  if (Test-Path -LiteralPath $localAppDataPath) {
    return $localAppDataPath
  }

  $windowsApps = Join-Path $env:ProgramFiles "WindowsApps"
  if (Test-Path -LiteralPath $windowsApps) {
    $candidate = Get-ChildItem -Path $windowsApps -Filter "Microsoft.DesktopAppInstaller_*__8wekyb3d8bbwe" -Directory -ErrorAction SilentlyContinue |
      Sort-Object Name -Descending |
      Select-Object -First 1

    if ($candidate) {
      $wingetExe = Join-Path $candidate.FullName "winget.exe"
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
    if ($p -match "\s") {
      '"{0}"' -f ($p -replace '"', '\"')
    } else {
      $p
    }
  }

  return ($safeParts -join " ")
}

function Invoke-LoggedProcess {
  param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [Parameter(Mandatory=$true)][string[]]$Arguments,
    [Parameter(Mandatory=$true)][string]$StepName,
    [int]$TimeoutSec = 7200
  )

  $tempDir = Join-Path $env:ProgramData "Datto\Temp\WingetUpgradeAll"
  if (-not (Test-Path -LiteralPath $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
  }

  $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
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
        StdOut   = Read-TextFileSmart -Path $stdoutPath
        StdErr   = Read-TextFileSmart -Path $stderrPath
        TimedOut = $true
      }
    }
  }

  $result = [pscustomobject]@{
    ExitCode = [int]$proc.ExitCode
    StdOut   = Read-TextFileSmart -Path $stdoutPath
    StdErr   = Read-TextFileSmart -Path $stderrPath
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

  try {
    Add-Action "Installing Microsoft.WinGet.Client module"
    $progressPreference = "SilentlyContinue"
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
  Add-Action "SCRIPT_VERSION=2026-04-29-WINGET-PER-PACKAGE-UPGRADE-v3-EXCLUSIONS"
  Add-Action ("Resolved variables: ReportOnly=[{0}] Install_Winget_if_Not_Avaialble=[{1}] ExcludedPackagesRaw=[{2}] LogPath=[{3}]" -f `
    [bool]$ReportOnly, [bool]$Install_Winget_if_Not_Avaialble, $ExcludedPackages, $LogPath)

  $excludeTokens = @(Get-ExcludedPackageTokens -RawText $ExcludedPackages)
  if ($excludeTokens.Count -gt 0) {
    Add-Action ("Excluded package tokens: {0}" -f ($excludeTokens -join ", "))
  } else {
    Add-Action "Excluded package tokens: [none]"
  }

  $wingetPath = Install-WingetIfNeeded
  if ([string]::IsNullOrWhiteSpace($wingetPath)) {
    Out-Result -Status "Failed" -Summary "winget.exe was not found or could not be installed." -ExitCode 1
  }

  $versionResult = Invoke-LoggedProcess -Exe $wingetPath -Arguments @("--version") -StepName "WingetVersion" -TimeoutSec 60
  Write-BlockToLog -Title "WingetVersion STDOUT" -Content $versionResult.StdOut
  Write-BlockToLog -Title "WingetVersion STDERR" -Content $versionResult.StdErr

  $sourceReset = Invoke-LoggedProcess -Exe $wingetPath -Arguments @("source","reset","--force","--accept-source-agreements","--disable-interactivity") -StepName "SourceReset" -TimeoutSec 600
  Write-BlockToLog -Title "SourceReset STDOUT" -Content $sourceReset.StdOut
  Write-BlockToLog -Title "SourceReset STDERR" -Content $sourceReset.StdErr

  $sourceUpdate = Invoke-LoggedProcess -Exe $wingetPath -Arguments @("source","update","--accept-source-agreements","--disable-interactivity") -StepName "SourceUpdate" -TimeoutSec 1200
  Write-BlockToLog -Title "SourceUpdate STDOUT" -Content $sourceUpdate.StdOut
  Write-BlockToLog -Title "SourceUpdate STDERR" -Content $sourceUpdate.StdErr

  $listArgs = @("list","--upgrade-available","--accept-source-agreements","--disable-interactivity")
  $preScan = Invoke-LoggedProcess -Exe $wingetPath -Arguments $listArgs -StepName "PreScan" -TimeoutSec 1800

  $combinedPreScan = @(
    $preScan.StdOut
    $preScan.StdErr
  ) -join "`r`n"

  Write-BlockToLog -Title "PreScan COMBINED INPUT" -Content $combinedPreScan
  Write-BlockToLog -Title "PreScan STDOUT" -Content $preScan.StdOut
  Write-BlockToLog -Title "PreScan STDERR" -Content $preScan.StdErr

  if ($preScan.TimedOut) {
    Out-Result -Status "CompletedWithErrors" -Summary "winget pre-scan timed out." -ExitCode 1
  }

  $allUpgradeItems = @(Get-UpgradeItemsFromWingetOutput -Content $combinedPreScan)
  $splitItems = Split-ExcludedUpgradeItems -Items $allUpgradeItems -ExcludeTokens $excludeTokens
  $upgradeItems = @($splitItems.Actionable)
  $excludedItems = @($splitItems.Excluded)

  Add-Action "All upgradeable packages detected by winget:"
  Add-Action "BEGIN ALL UPGRADEABLE PACKAGE LIST"
  Add-ActionLines -Lines (Format-UpgradeItems -Items $allUpgradeItems)
  Add-Action "END ALL UPGRADEABLE PACKAGE LIST"

  Add-Action "Excluded packages that will NOT be upgraded:"
  Add-Action "BEGIN EXCLUDED PACKAGE LIST"
  Add-ActionLines -Lines (Format-UpgradeItems -Items $excludedItems -EmptyMessage "No excluded upgradeable packages matched.")
  Add-Action "END EXCLUDED PACKAGE LIST"

  Add-Action "Actionable packages that may be upgraded:"
  Add-Action "BEGIN ACTIONABLE PACKAGE LIST"
  Add-ActionLines -Lines (Format-UpgradeItems -Items $upgradeItems -EmptyMessage "No actionable upgradeable packages detected.")
  Add-Action "END ACTIONABLE PACKAGE LIST"

  if ([bool]$ReportOnly) {
    Out-Result -Status "ReportOnly" -Summary ("Report-only complete. Actionable={0} Excluded={1} TotalDetected={2}." -f $upgradeItems.Count, $excludedItems.Count, $allUpgradeItems.Count) -ExitCode 0
  }

  if ($upgradeItems.Count -eq 0) {
    Out-Result -Status "Success" -Summary ("No actionable upgradeable packages detected. Excluded={0} TotalDetected={1}." -f $excludedItems.Count, $allUpgradeItems.Count) -ExitCode 0
  }

  $packageResults = @()

  foreach ($item in $upgradeItems) {
    $safeStepName = Get-SafeStepName -Text $item.Id
    $wingetNativeLog = Join-Path $env:ProgramData ("Datto\Logs\winget-upgrade-{0}-{1}.log" -f $safeStepName, (Get-Date -Format "yyyyMMdd-HHmmss"))

    $upgradeArgs = @(
      "upgrade",
      "--id", $item.Id,
      "--exact",
      "--silent",
      "--accept-package-agreements",
      "--accept-source-agreements",
      "--disable-interactivity",
      "--nowarn",
      "--log", $wingetNativeLog
    )

    if (-not [string]::IsNullOrWhiteSpace($item.Source)) {
      $upgradeArgs += "--source"
      $upgradeArgs += $item.Source
    }

    Add-Action ("PACKAGE START Name=[{0}] Id=[{1}] Current=[{2}] Available=[{3}] Source=[{4}]" -f `
      $item.Name, $item.Id, $item.Version, $item.Available, $item.Source)
    Add-Action ("PACKAGE WINGET LOG Id=[{0}] Log=[{1}]" -f $item.Id, $wingetNativeLog)

    $result = Invoke-LoggedProcess -Exe $wingetPath -Arguments $upgradeArgs -StepName ("Upgrade_" + $safeStepName) -TimeoutSec 3600

    Write-BlockToLog -Title ("Upgrade {0} STDOUT" -f $item.Id) -Content $result.StdOut
    Write-BlockToLog -Title ("Upgrade {0} STDERR" -f $item.Id) -Content $result.StdErr

    if ($result.TimedOut) {
      Add-ErrorText ("PACKAGE COMMAND TIMED OUT Name=[{0}] Id=[{1}]" -f $item.Name, $item.Id)
    }
    elseif ($result.ExitCode -eq 0) {
      Add-Action ("PACKAGE COMMAND SUCCESS Name=[{0}] Id=[{1}] ExitCode=[{2}]" -f $item.Name, $item.Id, $result.ExitCode)
    }
    else {
      Add-ErrorText ("PACKAGE COMMAND FAILED Name=[{0}] Id=[{1}] ExitCode=[{2}]" -f $item.Name, $item.Id, $result.ExitCode)
    }

    $packageResults += [pscustomobject]@{
      Name      = $item.Name
      Id        = $item.Id
      Version   = $item.Version
      Available = $item.Available
      Source    = $item.Source
      ExitCode  = $result.ExitCode
      TimedOut  = $result.TimedOut
      LogPath   = $wingetNativeLog
    }
  }

  $postScan = Invoke-LoggedProcess -Exe $wingetPath -Arguments $listArgs -StepName "PostScan" -TimeoutSec 1800

  $combinedPostScan = @(
    $postScan.StdOut
    $postScan.StdErr
  ) -join "`r`n"

  Write-BlockToLog -Title "PostScan COMBINED INPUT" -Content $combinedPostScan
  Write-BlockToLog -Title "PostScan STDOUT" -Content $postScan.StdOut
  Write-BlockToLog -Title "PostScan STDERR" -Content $postScan.StdErr

  $remainingAllItems = @(Get-UpgradeItemsFromWingetOutput -Content $combinedPostScan)
  $remainingSplitItems = Split-ExcludedUpgradeItems -Items $remainingAllItems -ExcludeTokens $excludeTokens
  $remainingItems = @($remainingSplitItems.Actionable)
  $remainingExcludedItems = @($remainingSplitItems.Excluded)

  Add-Action "Post-upgrade remaining actionable upgradeable packages:"
  Add-Action "BEGIN REMAINING ACTIONABLE PACKAGE LIST"
  Add-ActionLines -Lines (Format-UpgradeItems -Items $remainingItems -EmptyMessage "No remaining actionable upgradeable packages detected.")
  Add-Action "END REMAINING ACTIONABLE PACKAGE LIST"

  Add-Action "Post-upgrade remaining excluded upgradeable packages:"
  Add-Action "BEGIN REMAINING EXCLUDED PACKAGE LIST"
  Add-ActionLines -Lines (Format-UpgradeItems -Items $remainingExcludedItems -EmptyMessage "No remaining excluded upgradeable packages detected.")
  Add-Action "END REMAINING EXCLUDED PACKAGE LIST"

  $failedCount = 0
  $stillUpgradeableCount = 0
  $successCount = 0

  Add-Action "Per-package final result summary:"
  Add-Action "BEGIN PACKAGE RESULT SUMMARY"

  foreach ($pkg in $packageResults) {
    $stillUpgradeable = Test-ItemStillUpgradeable -Id $pkg.Id -RemainingItems $remainingItems

    if ($pkg.TimedOut) {
      $failedCount++
      Add-Action ("RESULT=FAILED_TIMEOUT Name=[{0}] Id=[{1}] From=[{2}] To=[{3}] ExitCode=[{4}]" -f `
        $pkg.Name, $pkg.Id, $pkg.Version, $pkg.Available, $pkg.ExitCode)
      continue
    }

    if ($pkg.ExitCode -ne 0) {
      $failedCount++
      Add-Action ("RESULT=FAILED_EXITCODE Name=[{0}] Id=[{1}] From=[{2}] To=[{3}] ExitCode=[{4}]" -f `
        $pkg.Name, $pkg.Id, $pkg.Version, $pkg.Available, $pkg.ExitCode)
      continue
    }

    if ($stillUpgradeable) {
      $stillUpgradeableCount++
      Add-Action ("RESULT=STILL_UPGRADEABLE Name=[{0}] Id=[{1}] From=[{2}] To=[{3}] ExitCode=[{4}]" -f `
        $pkg.Name, $pkg.Id, $pkg.Version, $pkg.Available, $pkg.ExitCode)
      continue
    }

    $successCount++
    Add-Action ("RESULT=SUCCESS Name=[{0}] Id=[{1}] From=[{2}] To=[{3}] ExitCode=[{4}]" -f `
      $pkg.Name, $pkg.Id, $pkg.Version, $pkg.Available, $pkg.ExitCode)
  }

  Add-Action "END PACKAGE RESULT SUMMARY"

  $summary = "Attempted=$($packageResults.Count) Success=$successCount Failed=$failedCount StillUpgradeable=$stillUpgradeableCount Excluded=$($excludedItems.Count) RemainingActionable=$($remainingItems.Count) RemainingExcluded=$($remainingExcludedItems.Count)"

  if ($failedCount -gt 0 -or $stillUpgradeableCount -gt 0) {
    Out-Result -Status "CompletedWithErrors" -Summary $summary -ExitCode 1
  }

  Out-Result -Status "Success" -Summary $summary -ExitCode 0
}
catch {
  Add-ErrorText ("Fatal exception: {0}" -f $_.Exception.Message)
  Out-Result -Status "Failed" -Summary "Script failed unexpectedly." -ExitCode 1
}