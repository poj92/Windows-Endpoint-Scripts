#Requires -Version 5.1
<#
Browser-Force-Open-Close.ps1

Checks whether ANY installed browser has been opened within the last N hours.
If not, launches one installed browser and closes only newly started processes after a delay.

Installed browser detection (SYSTEM-friendly):
- Known machine paths (Program Files / Program Files (x86))
- App Paths registry (HKLM + HKLM\WOW6432Node + HKCU)

Last-open detection:
- Prefetch last write time for the browser EXE name (best signal)
- Fallback: currently running browser processes with StartTime within lookback

Exit codes:
  0 = at least one installed browser opened within lookback (no action)
  1 = none opened; script launched and then closed a browser
  2 = ReportOnly and launch would have occurred
  3 = no supported browser installed
  4 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$LookbackHours = 24,
  [int]$OpenSeconds = 120,
  [string]$Url = 'about:blank',
  [switch]$ReportOnly,
  [string[]]$Preference = @('Edge','Chrome','Firefox','Brave','Opera')
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Get-AppPathExe {
  param([Parameter(Mandatory)][string]$ExeName)

  # Read default value from App Paths using .NET registry APIs (works across PS versions)
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

function Get-ProcessLastRunFallback {
  param(
    [Parameter(Mandatory)][string[]]$ProcNames,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  $times = @()
  foreach ($pn in $ProcNames) {
    $procs = Get-Process -Name $pn -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
      try {
        if ($p.StartTime -ge $Cutoff) { $times += $p.StartTime }
      } catch { }
    }
  }

  if (-not $times -or $times.Count -eq 0) { return $null }
  ($times | Sort-Object -Descending | Select-Object -First 1)
}

function Find-InstalledBrowsers {
  # Define supported browsers (machine-wide paths + app paths)
  $defs = @(
    @{
      Name='Edge';    ProcNames=@('msedge');  AppExe='msedge.exe';  PrefetchBases=@('msedge');
      Candidates=@("$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                   "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe")
    },
    @{
      Name='Chrome';  ProcNames=@('chrome');  AppExe='chrome.exe';  PrefetchBases=@('chrome');
      Candidates=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                   "$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe")
    },
    @{
      Name='Firefox'; ProcNames=@('firefox'); AppExe='firefox.exe'; PrefetchBases=@('firefox');
      Candidates=@("$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                   "$env:ProgramFiles (x86)\Mozilla Firefox\firefox.exe")
    },
    @{
      Name='Brave';   ProcNames=@('brave');   AppExe='brave.exe';   PrefetchBases=@('brave');
      Candidates=@("$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                   "$env:ProgramFiles (x86)\BraveSoftware\Brave-Browser\Application\brave.exe")
    },
    @{
      # Opera typically launches via launcher.exe but spawns opera.exe
      Name='Opera';   ProcNames=@('opera','launcher'); AppExe='launcher.exe'; PrefetchBases=@('opera','launcher');
      Candidates=@("$env:ProgramFiles\Opera\launcher.exe",
                   "$env:ProgramFiles (x86)\Opera\launcher.exe")
    }
  )

  $installed = @()

  foreach ($b in $defs) {
    $exePath = $null

    foreach ($c in $b.Candidates) {
      if ($c -and (Test-Path $c)) { $exePath = $c; break }
    }

    if (-not $exePath) {
      $exePath = Get-AppPathExe -ExeName $b.AppExe
    }

    if ($exePath) {
      $installed += [pscustomobject]@{
        Name = $b.Name
        ExePath = $exePath
        ProcNames = $b.ProcNames
        PrefetchBases = $b.PrefetchBases
      }
    }
  }

  # Order by preference
  $ordered = @()
  foreach ($p in $Preference) {
    $hit = $installed | Where-Object { $_.Name -eq $p } | Select-Object -First 1
    if ($hit) { $ordered += $hit }
  }
  foreach ($x in $installed) {
    if (-not ($ordered | Where-Object { $_.Name -eq $x.Name })) { $ordered += $x }
  }

  return $ordered
}

function Get-LastOpenStatus {
  param(
    [Parameter(Mandatory)]$InstalledBrowsers,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  $status = @()

  foreach ($b in $InstalledBrowsers) {
    $prefetchTimes = @()
    foreach ($base in $b.PrefetchBases) {
      $t = $null
      try { $t = Get-PrefetchLastRun -ExeBaseName $base } catch { $t = $null }
      if ($t) { $prefetchTimes += $t }
    }

    $prefetchTime = $null
    if ($prefetchTimes.Count -gt 0) {
      $prefetchTime = ($prefetchTimes | Sort-Object -Descending | Select-Object -First 1)
    }

    $fallbackTime = $null
    if (-not $prefetchTime) {
      $fallbackTime = Get-ProcessLastRunFallback -ProcNames $b.ProcNames -Cutoff $Cutoff
    }

    $last = $prefetchTime
    $source = 'Prefetch'
    if (-not $last) { $last = $fallbackTime; $source = 'ProcessStartTime' }
    if (-not $last) { $source = 'None' }

    $status += [pscustomobject]@{
      Browser = $b.Name
      ExePath = $b.ExePath
      ProcNames = $b.ProcNames
      LastOpened = $last
      Source = $source
      OpenedWithinLookback = [bool]($last -and $last -ge $Cutoff)
    }
  }

  return $status
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
      } catch {
        # If StartTime not accessible, still treat as potentially new (best-effort)
        $ok = $true
      }

      if ($ok) { $newProcs += $p }
    }
  }

  if (-not $newProcs -or $newProcs.Count -eq 0) { return }

  # Graceful close first
  foreach ($p in $newProcs) {
    try {
      if ($p.MainWindowHandle -ne 0) { [void]$p.CloseMainWindow() }
    } catch { }
  }

  Start-Sleep -Seconds 10

  # Force close leftovers
  foreach ($p in $newProcs) {
    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
  }
}

# ---------------- main ----------------
try {
  $cutoff = (Get-Date).AddHours(-[double]$LookbackHours)

  $installed = Find-InstalledBrowsers
  if (-not $installed -or $installed.Count -eq 0) {
    Write-Host "No supported browsers found installed on this machine."
    exit 3
  }

  Write-Host ("Lookback window: last {0} hours (since {1})" -f $LookbackHours, $cutoff)
  Write-Host "Installed browsers detected:"
  foreach ($b in $installed) {
    Write-Host ("- {0} => {1}" -f $b.Name, $b.ExePath)
  }

  $status = Get-LastOpenStatus -InstalledBrowsers $installed -Cutoff $cutoff

  Write-Host ""
  Write-Host "Last-open status:"
  foreach ($s in $status) {
    $t = '<unknown>'
    if ($s.LastOpened) { $t = $s.LastOpened.ToString('yyyy-MM-dd HH:mm:ss') }
    Write-Host ("{0}: LastOpened={1} Source={2} OpenedWithinLookback={3}" -f $s.Browser, $t, $s.Source, $s.OpenedWithinLookback)
  }

  $openedRecently = $false
  foreach ($s in $status) {
    if ($s.OpenedWithinLookback) { $openedRecently = $true; break }
  }

  if ($openedRecently) {
    Write-Host ""
    Write-Host "Result: At least one installed browser was opened within the lookback window. No action taken."
    exit 0
  }

  Write-Host ""
  Write-Host "Result: No installed browser was detected as opened within the lookback window."

  $toLaunch = $installed | Select-Object -First 1
  if ($ReportOnly) {
    Write-Host ("ReportOnly: Would launch {0} for {1} seconds then close." -f $toLaunch.Name, $OpenSeconds)
    exit 2
  }

  # Record existing PIDs per process name so we only close what we start
  $existingPids = @{}
  foreach ($pn in $toLaunch.ProcNames) {
    $existingPids[$pn] = @(Get-Process -Name $pn -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
  }

  $launchTime = Get-Date
  Write-Host ("Launching {0} to {1}..." -f $toLaunch.Name, $Url)

  if ($PSCmdlet.ShouldProcess($toLaunch.Name, "Launch and close after delay")) {
    Start-Process -FilePath $toLaunch.ExePath -ArgumentList $Url | Out-Null
    Start-Sleep -Seconds $OpenSeconds
    Write-Host "Closing newly-started browser processes..."
    Close-NewProcesses -ProcNames $toLaunch.ProcNames -ExistingPidsByName $existingPids -LaunchTime $launchTime
  }

  Write-Host "Done."
  exit 1
}
catch {
  Write-Error $_
  exit 4
}