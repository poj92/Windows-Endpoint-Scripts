#Requires -Version 5.1
<#
Browser-PerBrowser-Nudge.ps1

Per-browser logic:
- For EACH installed browser, determine whether it was opened within the last N hours.
- If a specific browser was not opened within the window, launch it and close it after a delay.
- Does NOT rely solely on Prefetch (which can be bumped by background activity).

Signals used (in this order):
1) If browser is currently running (StartTime within lookback) -> counts
2) User profile activity files (best proxy for real user usage)
3) Process creation event logs (Sysmon ID 1, else Security 4688 if enabled)
4) Prefetch last write time (least reliable; can be background)

Exit codes:
  0 = all installed browsers were opened within lookback (no action)
  1 = at least one browser was nudged (opened then closed)
  2 = ReportOnly and at least one browser would be nudged
  3 = no supported browsers installed
  4 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$LookbackHours = 24,
  [int]$OpenSeconds = 120,
  [string]$Url = 'about:blank',
  [switch]$ReportOnly,

  # Preference order for processing (also affects "which gets nudged first")
  [ValidateSet('Edge','Chrome','Firefox','Brave','Opera')]
  [string[]]$Preference = @('Edge','Chrome','Firefox','Brave','Opera')
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Get-ConsoleUserName {
  try {
    $u = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
    if ($u) { return $u }
  } catch { }
  return $null
}

function Get-ProfilePathsForUser {
  param([Parameter(Mandatory)][string]$UserName)

  try {
    $sid = (New-Object System.Security.Principal.NTAccount($UserName)).
      Translate([System.Security.Principal.SecurityIdentifier]).Value
    $key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
    $p = (Get-ItemProperty -Path $key -ErrorAction Stop).ProfileImagePath
    if (-not $p -or -not (Test-Path $p)) { return $null }

    return [pscustomobject]@{
      UserName       = $UserName
      ProfilePath    = $p
      LocalAppData   = Join-Path $p 'AppData\Local'
      RoamingAppData = Join-Path $p 'AppData\Roaming'
    }
  } catch {
    return $null
  }
}

function Get-AppPathExe {
  param([Parameter(Mandatory)][string]$ExeName)

  $subKeys = @(
    "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
    "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
  )

  foreach ($root in @([Microsoft.Win32.Registry]::LocalMachine, [Microsoft.Win32.Registry]::CurrentUser)) {
    foreach ($sub in $subKeys) {
      try {
        $k = $root.OpenSubKey($sub)
        if ($k) {
          $val = $k.GetValue('')
          $k.Close()
          if ($val -and (Test-Path $val)) { return $val }
        }
      } catch { }
    }
  }
  return $null
}

function Get-PrefetchLastRun {
  param([Parameter(Mandatory)][string]$ExeBaseName)

  $prefetchDir = Join-Path $env:SystemRoot 'Prefetch'
  if (-not (Test-Path $prefetchDir)) { return $null }

  $pattern = ('{0}.EXE-*.pf' -f $ExeBaseName.ToUpperInvariant())
  $files = Get-ChildItem -Path $prefetchDir -Filter $pattern -ErrorAction SilentlyContinue
  if (-not $files) { return $null }

  ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
}

function Get-LatestWriteTime {
  param([Parameter(Mandatory)][string[]]$PathsOrWildcards)

  $latest = $null
  foreach ($p in $PathsOrWildcards) {
    try {
      $items = Get-ChildItem -Path $p -ErrorAction SilentlyContinue
      foreach ($i in $items) {
        if (-not $latest -or $i.LastWriteTime -gt $latest) { $latest = $i.LastWriteTime }
      }
    } catch { }
  }
  return $latest
}

function Get-ProcessStartWithinLookback {
  param(
    [Parameter(Mandatory)][string[]]$ProcNames,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  $latest = $null
  foreach ($pn in $ProcNames) {
    $procs = Get-Process -Name $pn -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
      try {
        if ($p.StartTime -ge $Cutoff) {
          if (-not $latest -or $p.StartTime -gt $latest) { $latest = $p.StartTime }
        }
      } catch { }
    }
  }
  return $latest
}

function Get-ProcessCreateTimesFromEventLogs {
  param(
    [Parameter(Mandatory)][string[]]$ExeLeafNames,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  $result = @{}  # leaf -> datetime

  # Sysmon (best, if present)
  try {
    $sysmonLog = Get-WinEvent -ListLog 'Microsoft-Windows-Sysmon/Operational' -ErrorAction SilentlyContinue
    if ($sysmonLog) {
      $events = Get-WinEvent -FilterHashtable @{ LogName='Microsoft-Windows-Sysmon/Operational'; Id=1; StartTime=$Cutoff } -ErrorAction SilentlyContinue
      foreach ($e in $events) {
        try {
          $xml = [xml]$e.ToXml()
          $img = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'Image' } | Select-Object -First 1).'#text'
          if (-not $img) { continue }
          $leaf = [System.IO.Path]::GetFileName($img).ToLowerInvariant()
          if ($ExeLeafNames -notcontains $leaf) { continue }
          if (-not $result.ContainsKey($leaf) -or $e.TimeCreated -gt $result[$leaf]) {
            $result[$leaf] = $e.TimeCreated
          }
        } catch { }
      }
      return $result
    }
  } catch { }

  # Security 4688 (if enabled)
  try {
    $events = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4688; StartTime=$Cutoff } -ErrorAction SilentlyContinue
    foreach ($e in $events) {
      $msg = $e.Message
      if (-not $msg) { continue }
      foreach ($leaf in $ExeLeafNames) {
        if ($msg -match [regex]::Escape($leaf)) {
          if (-not $result.ContainsKey($leaf) -or $e.TimeCreated -gt $result[$leaf]) {
            $result[$leaf] = $e.TimeCreated
          }
        }
      }
    }
  } catch { }

  return $result
}

function Find-InstalledBrowsers {
  param([string[]]$Names)

  $defs = @(
    @{
      Name='Edge'; ProcNames=@('msedge'); ExeLeaf='msedge.exe'; PrefetchBases=@('msedge');
      Candidates=@("$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                   "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe");
      ProfileKind='Chromium'; ProfileRootRel=@('Microsoft\Edge\User Data')
    },
    @{
      Name='Chrome'; ProcNames=@('chrome'); ExeLeaf='chrome.exe'; PrefetchBases=@('chrome');
      Candidates=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                   "$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe");
      ProfileKind='Chromium'; ProfileRootRel=@('Google\Chrome\User Data')
    },
    @{
      Name='Brave'; ProcNames=@('brave'); ExeLeaf='brave.exe'; PrefetchBases=@('brave');
      Candidates=@("$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                   "$env:ProgramFiles (x86)\BraveSoftware\Brave-Browser\Application\brave.exe");
      ProfileKind='Chromium'; ProfileRootRel=@('BraveSoftware\Brave-Browser\User Data')
    },
    @{
      Name='Firefox'; ProcNames=@('firefox'); ExeLeaf='firefox.exe'; PrefetchBases=@('firefox');
      Candidates=@("$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                   "$env:ProgramFiles (x86)\Mozilla Firefox\firefox.exe");
      ProfileKind='Firefox'; ProfileRootRel=@('Mozilla\Firefox')
    },
    @{
      Name='Opera'; ProcNames=@('opera','launcher'); ExeLeaf='launcher.exe'; PrefetchBases=@('opera','launcher');
      Candidates=@("$env:ProgramFiles\Opera\launcher.exe",
                   "$env:ProgramFiles (x86)\Opera\launcher.exe");
      ProfileKind='Opera'; ProfileRootRel=@('Opera Software\Opera Stable')
    }
  )

  $installed = @()

  foreach ($name in $Names) {
    $b = $defs | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if (-not $b) { continue }

    $exePath = $null
    foreach ($c in $b.Candidates) { if ($c -and (Test-Path $c)) { $exePath = $c; break } }
    if (-not $exePath) { $exePath = Get-AppPathExe -ExeName $b.ExeLeaf }

    if ($exePath) {
      $installed += [pscustomobject]@{
        Name=$b.Name; ExePath=$exePath; ExeLeaf=$b.ExeLeaf; ProcNames=$b.ProcNames; PrefetchBases=$b.PrefetchBases;
        ProfileKind=$b.ProfileKind; ProfileRootRel=$b.ProfileRootRel
      }
    }
  }

  # keep preference order
  $ordered = @()
  foreach ($p in $Names) {
    $hit = $installed | Where-Object { $_.Name -eq $p } | Select-Object -First 1
    if ($hit) { $ordered += $hit }
  }
  $ordered
}

function Get-ProfileActivityTime {
  param(
    [Parameter(Mandatory)]$Browser,
    [Parameter(Mandatory)]$UserPaths
  )

  if (-not $UserPaths) { return $null }

  if ($Browser.ProfileKind -eq 'Chromium') {
    $root = Join-Path $UserPaths.LocalAppData ($Browser.ProfileRootRel[0])
    if (-not (Test-Path $root)) { return $null }

    # Look across Default + Profile* for files that change with real usage.
    $paths = @(
      (Join-Path $root 'Default\History'),
      (Join-Path $root 'Default\Current Session'),
      (Join-Path $root 'Default\Current Tabs'),
      (Join-Path $root 'Default\Last Session'),
      (Join-Path $root 'Default\Last Tabs'),
      (Join-Path $root 'Default\Preferences'),
      (Join-Path $root 'Profile *\History'),
      (Join-Path $root 'Profile *\Current Session'),
      (Join-Path $root 'Profile *\Current Tabs'),
      (Join-Path $root 'Profile *\Last Session'),
      (Join-Path $root 'Profile *\Last Tabs'),
      (Join-Path $root 'Profile *\Preferences')
    )
    return Get-LatestWriteTime -PathsOrWildcards $paths
  }

  if ($Browser.ProfileKind -eq 'Firefox') {
    $base = Join-Path $UserPaths.RoamingAppData ($Browser.ProfileRootRel[0])
    if (-not (Test-Path $base)) { return $null }

    # Try profiles.ini to find profile dirs; else wildcard
    $paths = @(
      (Join-Path $base 'profiles.ini'),
      (Join-Path $base 'Profiles\*\places.sqlite'),
      (Join-Path $base 'Profiles\*\sessionstore.jsonlz4'),
      (Join-Path $base 'Profiles\*\prefs.js')
    )
    return Get-LatestWriteTime -PathsOrWildcards $paths
  }

  if ($Browser.ProfileKind -eq 'Opera') {
    # Opera Stable profile is usually under Roaming
    $base = Join-Path $UserPaths.RoamingAppData ($Browser.ProfileRootRel[0])
    if (-not (Test-Path $base)) { return $null }

    $paths = @(
      (Join-Path $base 'History'),
      (Join-Path $base 'Current Session'),
      (Join-Path $base 'Current Tabs'),
      (Join-Path $base 'Last Session'),
      (Join-Path $base 'Last Tabs'),
      (Join-Path $base 'Preferences')
    )
    return Get-LatestWriteTime -PathsOrWildcards $paths
  }

  return $null
}

function Get-BrowserLastOpened {
  param(
    [Parameter(Mandatory)]$Browser,
    [Parameter(Mandatory)][datetime]$Cutoff,
    [Parameter(Mandatory)]$EventTimesByLeaf,
    $UserPaths
  )

  # 1) If currently running within window
  $running = Get-ProcessStartWithinLookback -ProcNames $Browser.ProcNames -Cutoff $Cutoff
  if ($running) {
    return [pscustomobject]@{ Time=$running; Source='RunningProcess' }
  }

  # 2) Profile activity (best user signal)
  $profileTime = Get-ProfileActivityTime -Browser $Browser -UserPaths $UserPaths
  if ($profileTime) {
    return [pscustomobject]@{ Time=$profileTime; Source='ProfileActivity' }
  }

  # 3) Event logs
  $leaf = $Browser.ExeLeaf.ToLowerInvariant()
  if ($EventTimesByLeaf.ContainsKey($leaf)) {
    return [pscustomobject]@{ Time=$EventTimesByLeaf[$leaf]; Source='EventLog' }
  }

  # 4) Prefetch (least reliable)
  $prefTimes = @()
  foreach ($base in $Browser.PrefetchBases) {
    $t = $null
    try { $t = Get-PrefetchLastRun -ExeBaseName $base } catch { $t = $null }
    if ($t) { $prefTimes += $t }
  }
  if ($prefTimes.Count -gt 0) {
    $pt = ($prefTimes | Sort-Object -Descending | Select-Object -First 1)
    return [pscustomobject]@{ Time=$pt; Source='Prefetch' }
  }

  return [pscustomobject]@{ Time=$null; Source='None' }
}

function Close-NewProcesses {
  param(
    [Parameter(Mandatory)][string[]]$ProcNames,
    [Parameter(Mandatory)][hashtable]$ExistingPidsByName,
    [Parameter(Mandatory)][datetime]$LaunchTime
  )

  $newProcs = @()

  foreach ($pn in $ProcNames) {
    $existing = @()
    if ($ExistingPidsByName.ContainsKey($pn)) { $existing = $ExistingPidsByName[$pn] }

    $procs = Get-Process -Name $pn -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
      if ($existing -contains $p.Id) { continue }

      $ok = $true
      try {
        if ($p.StartTime -lt $LaunchTime.AddSeconds(-5)) { $ok = $false }
      } catch { $ok = $true }

      if ($ok) { $newProcs += $p }
    }
  }

  if (-not $newProcs -or $newProcs.Count -eq 0) { return }

  foreach ($p in $newProcs) {
    try { if ($p.MainWindowHandle -ne 0) { [void]$p.CloseMainWindow() } } catch { }
  }

  Start-Sleep -Seconds 10

  foreach ($p in $newProcs) {
    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
  }
}

function Get-LaunchArgs {
  param([Parameter(Mandatory)][string]$BrowserName, [Parameter(Mandatory)][string]$Url)

  # Use flags that force a visible window where possible
  if ($BrowserName -in @('Edge','Chrome','Brave')) { return "--new-window $Url" }
  if ($BrowserName -eq 'Firefox') { return "-new-window $Url" }
  if ($BrowserName -eq 'Opera') { return $Url }
  return $Url
}

# ---------------- main ----------------
try {
  $cutoff = (Get-Date).AddHours(-[double]$LookbackHours)

  $installed = Find-InstalledBrowsers -Names $Preference
  if (-not $installed -or $installed.Count -eq 0) {
    Write-Host "No supported browsers found installed on this machine."
    exit 3
  }

  $consoleUser = Get-ConsoleUserName
  $userPaths = $null
  if ($consoleUser) { $userPaths = Get-ProfilePathsForUser -UserName $consoleUser }

  Write-Host "Browser-PerBrowser-Nudge"
  Write-Host ("Lookback window: last {0} hours (since {1})" -f $LookbackHours, $cutoff)
  if ($userPaths) {
    Write-Host ("Console user: {0}  Profile: {1}" -f $userPaths.UserName, $userPaths.ProfilePath)
  } else {
    Write-Host "Console user profile: not available (SYSTEM/no interactive user)."
  }

  Write-Host ""
  Write-Host "Installed browsers detected:"
  foreach ($b in $installed) {
    Write-Host ("- {0} => {1}" -f $b.Name, $b.ExePath)
  }

  # Build event time map once (per run)
  $exeLeafs = @($installed | ForEach-Object { $_.ExeLeaf.ToLowerInvariant() }) | Sort-Object -Unique
  $eventMap = Get-ProcessCreateTimesFromEventLogs -ExeLeafNames $exeLeafs -Cutoff $cutoff

  # Evaluate each browser independently
  $status = @()
  foreach ($b in $installed) {
    $info = Get-BrowserLastOpened -Browser $b -Cutoff $cutoff -EventTimesByLeaf $eventMap -UserPaths $userPaths
    $within = [bool]($info.Time -and $info.Time -ge $cutoff)

    $status += [pscustomobject]@{
      Browser=$b.Name
      ExePath=$b.ExePath
      LastOpened=$info.Time
      Source=$info.Source
      OpenedWithinLookback=$within
      NeedsAction = -not $within
      ProcNames = $b.ProcNames
    }
  }

  Write-Host ""
  Write-Host "Per-browser last-open status:"
  foreach ($s in $status) {
    $t = '<unknown>'
    if ($s.LastOpened) { $t = $s.LastOpened.ToString('yyyy-MM-dd HH:mm:ss') }
    Write-Host ("{0}: LastOpened={1} Source={2} OpenedWithinLookback={3} NeedsAction={4}" -f $s.Browser, $t, $s.Source, $s.OpenedWithinLookback, $s.NeedsAction)
  }

  $toNudge = @($status | Where-Object { $_.NeedsAction })
  if (-not $toNudge -or $toNudge.Count -eq 0) {
    Write-Host ""
    Write-Host "Result: All installed browsers were opened within the lookback window. No action taken."
    exit 0
  }

  Write-Host ""
  Write-Host "Result: The following browsers were NOT opened within the lookback window and will be nudged:"
  foreach ($x in $toNudge) { Write-Host ("- {0}" -f $x.Browser) }

  if ($ReportOnly) {
    Write-Host ("ReportOnly: Would launch each above browser for {0} seconds then close." -f $OpenSeconds)
    exit 2
  }

  $didAny = $false

  foreach ($x in $toNudge) {
    $browser = $installed | Where-Object { $_.Name -eq $x.Browser } | Select-Object -First 1
    if (-not $browser) { continue }

    $existingPids = @{}
    foreach ($pn in $browser.ProcNames) {
      $existingPids[$pn] = @(Get-Process -Name $pn -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    }

    $launchTime = Get-Date
    $args = Get-LaunchArgs -BrowserName $browser.Name -Url $Url

    Write-Host ""
    Write-Host ("Launching {0} for {1} seconds..." -f $browser.Name, $OpenSeconds)

    if ($PSCmdlet.ShouldProcess($browser.Name, "Launch and close after delay")) {
      Start-Process -FilePath $browser.ExePath -ArgumentList $args | Out-Null
      Start-Sleep -Seconds $OpenSeconds
      Write-Host ("Closing newly-started {0} processes..." -f $browser.Name)
      Close-NewProcesses -ProcNames $browser.ProcNames -ExistingPidsByName $existingPids -LaunchTime $launchTime
      $didAny = $true
    }
  }

  Write-Host ""
  Write-Host "Done."
  if ($didAny) { exit 1 } else { exit 0 }
}
catch {
  Write-Error $_
  exit 4
}