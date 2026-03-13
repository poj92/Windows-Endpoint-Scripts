#Requires -Version 5.1
<#
Browser-Force-Open-Close.ps1 (Per-browser)

For each installed browser:
- Determine if it was opened within the last N hours
- If not, launch it and close it after a delay

Signals (in order):
1) Running process started within lookback
2) User profile activity timestamps (best proxy for real user usage) if a user profile can be determined
3) Prefetch last write time (least reliable; can be bumped by background activity)

Exit codes:
  0 = all installed browsers opened within lookback (no action)
  1 = at least one browser was nudged (opened then closed)
  2 = ReportOnly and at least one browser would be nudged
  3 = no supported browsers installed
  4 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$LookbackHours = 24,
  [int]$OpenSeconds   = 120,
  [string]$Url        = 'about:blank',
  [switch]$ReportOnly,

  # Order matters: used for install detection and nudging order
  [string[]]$Preference = @('Edge','Chrome','Firefox','Brave','Opera')
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

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

function Get-ConsoleUserName {
  try {
    $u = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
    if ($u) { return $u }
  } catch { }
  return $null
}

function Get-UserProfilePathsFromUsername {
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

function Get-BestAvailableUserProfilePaths {
  # If we can’t identify the console user (SYSTEM context), pick the “most recently used” profile.
  # Best-effort: choose the ProfileImagePath whose NTUSER.DAT was most recently written.
  $profileListRoot = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
  if (-not (Test-Path $profileListRoot)) { return $null }

  $best = $null

  Get-ChildItem $profileListRoot -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      $p = (Get-ItemProperty $_.PSPath -ErrorAction Stop).ProfileImagePath
      if (-not $p) { return }
      if (-not (Test-Path $p)) { return }

      # Skip built-in / service profiles
      $leaf = Split-Path $p -Leaf
      if ($leaf -in @('Public','Default','Default User','All Users')) { return }
      if ($p -match '\\Windows\\System32\\') { return }

      $ntuser = Join-Path $p 'NTUSER.DAT'
      if (-not (Test-Path $ntuser)) { return }

      $t = (Get-Item $ntuser -ErrorAction Stop).LastWriteTime

      if (-not $best -or $t -gt $best.Time) {
        $best = [pscustomobject]@{ ProfilePath=$p; Time=$t }
      }
    } catch { }
  }

  if (-not $best) { return $null }

  return [pscustomobject]@{
    UserName       = '<best-effort>'
    ProfilePath    = $best.ProfilePath
    LocalAppData   = Join-Path $best.ProfilePath 'AppData\Local'
    RoamingAppData = Join-Path $best.ProfilePath 'AppData\Roaming'
  }
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

function Get-ProfileActivityTime {
  param(
    [Parameter(Mandatory)]$Browser,
    $UserPaths   # <-- NOT mandatory; can be $null safely
  )

  if (-not $UserPaths) { return $null }

  if ($Browser.ProfileKind -eq 'Chromium') {
    $root = Join-Path $UserPaths.LocalAppData $Browser.ProfileRootRel
    if (-not (Test-Path $root)) { return $null }

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
    $base = Join-Path $UserPaths.RoamingAppData $Browser.ProfileRootRel
    if (-not (Test-Path $base)) { return $null }

    $paths = @(
      (Join-Path $base 'profiles.ini'),
      (Join-Path $base 'Profiles\*\places.sqlite'),
      (Join-Path $base 'Profiles\*\sessionstore.jsonlz4'),
      (Join-Path $base 'Profiles\*\prefs.js')
    )
    return Get-LatestWriteTime -PathsOrWildcards $paths
  }

  if ($Browser.ProfileKind -eq 'Opera') {
    $base = Join-Path $UserPaths.RoamingAppData $Browser.ProfileRootRel
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

function Find-InstalledBrowsers {
  param([string[]]$Names)

  $defs = @(
    @{
      Name='Edge';    ProcNames=@('msedge');  ExeLeaf='msedge.exe';  PrefetchBases=@('msedge');
      Candidates=@("$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                   "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe");
      ProfileKind='Chromium'; ProfileRootRel='Microsoft\Edge\User Data'
    },
    @{
      Name='Chrome';  ProcNames=@('chrome');  ExeLeaf='chrome.exe';  PrefetchBases=@('chrome');
      Candidates=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                   "$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe");
      ProfileKind='Chromium'; ProfileRootRel='Google\Chrome\User Data'
    },
    @{
      Name='Brave';   ProcNames=@('brave');   ExeLeaf='brave.exe';   PrefetchBases=@('brave');
      Candidates=@("$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                   "$env:ProgramFiles (x86)\BraveSoftware\Brave-Browser\Application\brave.exe");
      ProfileKind='Chromium'; ProfileRootRel='BraveSoftware\Brave-Browser\User Data'
    },
    @{
      Name='Firefox'; ProcNames=@('firefox'); ExeLeaf='firefox.exe'; PrefetchBases=@('firefox');
      Candidates=@("$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                   "$env:ProgramFiles (x86)\Mozilla Firefox\firefox.exe");
      ProfileKind='Firefox'; ProfileRootRel='Mozilla\Firefox'
    },
    @{
      Name='Opera';   ProcNames=@('opera','launcher'); ExeLeaf='launcher.exe'; PrefetchBases=@('opera','launcher');
      Candidates=@("$env:ProgramFiles\Opera\launcher.exe",
                   "$env:ProgramFiles (x86)\Opera\launcher.exe");
      ProfileKind='Opera'; ProfileRootRel='Opera Software\Opera Stable'
    }
  )

  $installed = @()

  foreach ($n in $Names) {
    $d = $defs | Where-Object { $_.Name -eq $n } | Select-Object -First 1
    if (-not $d) { continue }

    $exePath = $null
    foreach ($c in $d.Candidates) { if ($c -and (Test-Path $c)) { $exePath = $c; break } }
    if (-not $exePath) { $exePath = Get-AppPathExe -ExeName $d.ExeLeaf }

    if ($exePath) {
      $installed += [pscustomobject]@{
        Name=$d.Name; ExePath=$exePath; ExeLeaf=$d.ExeLeaf;
        ProcNames=$d.ProcNames; PrefetchBases=$d.PrefetchBases;
        ProfileKind=$d.ProfileKind; ProfileRootRel=$d.ProfileRootRel
      }
    }
  }

  return $installed
}

function Get-LaunchArgs {
  param([Parameter(Mandatory)][string]$BrowserName, [Parameter(Mandatory)][string]$Url)
  if ($BrowserName -in @('Edge','Chrome','Brave')) { return "--new-window $Url" }
  if ($BrowserName -eq 'Firefox') { return "-new-window $Url" }
  return $Url
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

# ---------------- main ----------------
try {
  $cutoff = (Get-Date).AddHours(-[double]$LookbackHours)

  $installed = Find-InstalledBrowsers -Names $Preference
  if (-not $installed -or $installed.Count -eq 0) {
    Write-Host "No supported browsers found installed on this machine."
    exit 3
  }

  # Determine user profile paths (console user if present; else best-effort most recently used)
  $consoleUser = Get-ConsoleUserName
  $userPaths = $null
  if ($consoleUser) { $userPaths = Get-UserProfilePathsFromUsername -UserName $consoleUser }
  if (-not $userPaths) { $userPaths = Get-BestAvailableUserProfilePaths }

  Write-Host "Browser-PerBrowser-Nudge"
  Write-Host ("Lookback window: last {0} hours (since {1})" -f $LookbackHours, $cutoff)

  if ($userPaths) {
    Write-Host ("User profile for activity checks: {0}" -f $userPaths.ProfilePath)
  } else {
    Write-Host "User profile for activity checks: <none available>; profile activity signal skipped."
  }

  Write-Host ""
  Write-Host "Installed browsers detected:"
  foreach ($b in $installed) { Write-Host ("- {0} => {1}" -f $b.Name, $b.ExePath) }

  # Per-browser evaluation
  $status = @()
  foreach ($b in $installed) {
    $runningTime = Get-ProcessStartWithinLookback -ProcNames $b.ProcNames -Cutoff $cutoff
    if ($runningTime) {
      $last = $runningTime
      $src  = 'RunningProcess'
    } else {
      $profileTime = Get-ProfileActivityTime -Browser $b -UserPaths $userPaths
      if ($profileTime) {
        $last = $profileTime
        $src  = 'ProfileActivity'
      } else {
        # last resort
        $prefTimes = @()
        foreach ($base in $b.PrefetchBases) {
          $t = $null
          try { $t = Get-PrefetchLastRun -ExeBaseName $base } catch { $t = $null }
          if ($t) { $prefTimes += $t }
        }
        if ($prefTimes.Count -gt 0) {
          $last = ($prefTimes | Sort-Object -Descending | Select-Object -First 1)
          $src  = 'Prefetch'
        } else {
          $last = $null
          $src  = 'None'
        }
      }
    }

    $within = [bool]($last -and $last -ge $cutoff)

    $status += [pscustomobject]@{
      Browser=$b.Name
      ExePath=$b.ExePath
      ProcNames=$b.ProcNames
      LastOpened=$last
      Source=$src
      OpenedWithinLookback=$within
      NeedsAction = -not $within
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
  Write-Host "Result: Browsers needing action (not opened within window):"
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