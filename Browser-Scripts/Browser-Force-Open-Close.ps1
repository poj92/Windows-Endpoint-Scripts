#Requires -Version 5.1
<#
Browser-Force-Open-Close.ps1 (SAFE PER-BROWSER)

Goal:
- For EACH installed browser:
  - If the browser is currently running (any instance): DO NOT TOUCH IT (no launch/no close).
  - Else, decide if it was opened within the last N hours.
  - If not opened, launch it, keep open for OpenSeconds, then close ONLY the instance we started.

Signals (when NOT running):
- Prefetch last-write time for the browser exe (best-effort; can be bumped by background launches)

Safety:
- Uses unique temp profile directories and closes only processes whose CommandLine contains that profile path.
- Never touches browsers that were already running when the script started.

Exit codes:
  0 = no action needed (or all candidates already running / within window)
  1 = at least one browser was launched+closed by this script
  2 = ReportOnly and at least one browser would be launched
  3 = no supported browsers installed
  4 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$LookbackHours = 24,
  [int]$OpenSeconds   = 120,
  [string]$Url        = 'about:blank',
  [switch]$ReportOnly,
  [string[]]$Preference = @('Edge','Chrome','Firefox','Brave','Opera')
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Write-Log([string]$Msg) {
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Write-Host ("[{0}] {1}" -f $ts, $Msg)
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

function Find-InstalledBrowsers {
  param([string[]]$Names)

  $defs = @(
    @{
      Name='Edge'
      Type='Chromium'
      ExeLeaf='msedge.exe'
      ProcNames=@('msedge')
      Candidates=@(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe"
      )
      PrefetchBases=@('msedge')
      CloseNames=@('msedge.exe')
    },
    @{
      Name='Chrome'
      Type='Chromium'
      ExeLeaf='chrome.exe'
      ProcNames=@('chrome')
      Candidates=@(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe"
      )
      PrefetchBases=@('chrome')
      CloseNames=@('chrome.exe')
    },
    @{
      Name='Brave'
      Type='Chromium'
      ExeLeaf='brave.exe'
      ProcNames=@('brave')
      Candidates=@(
        "$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
        "$env:ProgramFiles (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
      )
      PrefetchBases=@('brave')
      CloseNames=@('brave.exe')
    },
    @{
      Name='Firefox'
      Type='Firefox'
      ExeLeaf='firefox.exe'
      ProcNames=@('firefox')
      Candidates=@(
        "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
        "$env:ProgramFiles (x86)\Mozilla Firefox\firefox.exe"
      )
      PrefetchBases=@('firefox')
      CloseNames=@('firefox.exe')
    },
    @{
      Name='Opera'
      Type='Chromium'
      ExeLeaf='launcher.exe'
      ProcNames=@('opera','launcher')
      Candidates=@(
        "$env:ProgramFiles\Opera\launcher.exe",
        "$env:ProgramFiles (x86)\Opera\launcher.exe"
      )
      PrefetchBases=@('opera','launcher')
      # Opera may spawn opera.exe and/or launcher.exe; close both if they include our marker.
      CloseNames=@('opera.exe','launcher.exe')
    }
  )

  $installed = @()
  foreach ($n in $Names) {
    $d = $defs | Where-Object { $_.Name -eq $n } | Select-Object -First 1
    if (-not $d) { continue }

    $exePath = $null
    foreach ($c in $d.Candidates) {
      if ($c -and (Test-Path $c)) { $exePath = $c; break }
    }
    if (-not $exePath) { $exePath = Get-AppPathExe -ExeName $d.ExeLeaf }

    if ($exePath) {
      $installed += [pscustomobject]@{
        Name=$d.Name
        Type=$d.Type
        ExePath=$exePath
        ExeLeaf=$d.ExeLeaf
        ProcNames=$d.ProcNames
        PrefetchBases=$d.PrefetchBases
        CloseNames=$d.CloseNames
      }
    }
  }

  return $installed
}

function Is-BrowserRunningNow {
  param([Parameter(Mandatory)][string[]]$ProcNames)

  foreach ($pn in $ProcNames) {
    $p = Get-Process -Name $pn -ErrorAction SilentlyContinue
    if ($p) { return $true }
  }
  return $false
}

function Get-LastOpenedWhenNotRunning {
  param(
    [Parameter(Mandatory)]$Browser
  )

  # Only best-effort Prefetch when not running (still can be bumped by background activity).
  $prefTimes = @()
  foreach ($base in $Browser.PrefetchBases) {
    $t = $null
    try { $t = Get-PrefetchLastRun -ExeBaseName $base } catch { $t = $null }
    if ($t) { $prefTimes += $t }
  }

  if ($prefTimes.Count -gt 0) {
    return [pscustomobject]@{ Time=($prefTimes | Sort-Object -Descending | Select-Object -First 1); Source='Prefetch' }
  }

  return [pscustomobject]@{ Time=$null; Source='None' }
}

function New-UniqueProfileDir {
  param([Parameter(Mandatory)][string]$BrowserName)

  $base = Join-Path $env:ProgramData 'BrowserNudge'
  $dir1 = Join-Path $base $BrowserName
  $dir2 = Join-Path $dir1 ([guid]::NewGuid().ToString('N'))

  if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
  if (-not (Test-Path $dir1)) { New-Item -ItemType Directory -Path $dir1 -Force | Out-Null }
  New-Item -ItemType Directory -Path $dir2 -Force | Out-Null

  return $dir2
}

function Get-LaunchArgs {
  param(
    [Parameter(Mandatory)]$Browser,
    [Parameter(Mandatory)][string]$ProfileDir,
    [Parameter(Mandatory)][string]$Url
  )

  if ($Browser.Type -eq 'Chromium') {
    # Use isolated profile to avoid touching user profile and to identify our processes later.
    return @(
      "--user-data-dir=`"$ProfileDir`"",
      "--no-first-run",
      "--no-default-browser-check",
      "--new-window",
      $Url
    ) -join ' '
  }

  if ($Browser.Type -eq 'Firefox') {
    # -no-remote keeps it separate from any existing instance (but we skip if running anyway).
    return @(
      "-no-remote",
      "-profile", "`"$ProfileDir`"",
      "-new-window",
      $Url
    ) -join ' '
  }

  return $Url
}

function Stop-OnlyOurBrowserProcesses {
  param(
    [Parameter(Mandatory)][string[]]$ProcessExeNames,
    [Parameter(Mandatory)][string]$MarkerText
  )

  # MarkerText is the unique profile directory; match on command line.
  foreach ($exe in $ProcessExeNames) {
    try {
      $filter = "Name='$exe'"
      $procs = Get-CimInstance Win32_Process -Filter $filter -ErrorAction SilentlyContinue
      foreach ($p in $procs) {
        $cmd = $p.CommandLine
        if ($cmd -and ($cmd -like "*$MarkerText*")) {
          try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
          } catch { }
        }
      }
    } catch { }
  }
}

# ---------------- main ----------------
try {
  Write-Host "Browser-Force-Open-Close"
  $cutoff = (Get-Date).AddHours(-[double]$LookbackHours)

  $installed = Find-InstalledBrowsers -Names $Preference
  if (-not $installed -or $installed.Count -eq 0) {
    Write-Host "No supported browsers found installed on this machine."
    exit 3
  }

  Write-Host ("Lookback window: last {0} hours (since {1})" -f $LookbackHours, $cutoff)
  Write-Host ""
  Write-Host "Installed browsers detected:"
  foreach ($b in $installed) { Write-Host ("- {0} => {1}" -f $b.Name, $b.ExePath) }

  # Evaluate each browser independently
  $status = @()
  foreach ($b in $installed) {
    $runningNow = Is-BrowserRunningNow -ProcNames $b.ProcNames

    if ($runningNow) {
      $status += [pscustomobject]@{
        Browser=$b.Name
        RunningNow=$true
        LastOpened=$null
        Source='RunningNow'
        OpenedWithinLookback=$true
        NeedsAction=$false
      }
      continue
    }

    $info = Get-LastOpenedWhenNotRunning -Browser $b
    $within = [bool]($info.Time -and $info.Time -ge $cutoff)

    $status += [pscustomobject]@{
      Browser=$b.Name
      RunningNow=$false
      LastOpened=$info.Time
      Source=$info.Source
      OpenedWithinLookback=$within
      NeedsAction = -not $within
    }
  }

  Write-Host ""
  Write-Host "Per-browser status (running browsers are never touched):"
  foreach ($s in $status) {
    $t = '<unknown>'
    if ($s.LastOpened) { $t = $s.LastOpened.ToString('yyyy-MM-dd HH:mm:ss') }
    Write-Host ("{0}: RunningNow={1} LastOpened={2} Source={3} OpenedWithinLookback={4} NeedsAction={5}" -f `
      $s.Browser, $s.RunningNow, $t, $s.Source, $s.OpenedWithinLookback, $s.NeedsAction)
  }

  $toNudge = @()
  foreach ($s in $status) {
    if (-not $s.RunningNow -and $s.NeedsAction) { $toNudge += $s.Browser }
  }

  if (-not $toNudge -or $toNudge.Count -eq 0) {
    Write-Host ""
    Write-Host "Result: No browsers require action (or they are currently running)."
    exit 0
  }

  Write-Host ""
  Write-Host "Result: Browsers needing action (NOT running and NOT opened within window):"
  foreach ($bn in $toNudge) { Write-Host ("- {0}" -f $bn) }

  if ($ReportOnly) {
    Write-Host ("ReportOnly: Would launch each listed browser for {0} seconds then close ONLY the instance started by this script." -f $OpenSeconds)
    exit 2
  }

  $didAny = $false

  foreach ($bn in $toNudge) {
    $browser = $installed | Where-Object { $_.Name -eq $bn } | Select-Object -First 1
    if (-not $browser) { continue }

    # Safety re-check: if it started running since evaluation, do not touch.
    if (Is-BrowserRunningNow -ProcNames $browser.ProcNames) {
      Write-Host ("Skipping {0}: browser is now running." -f $browser.Name)
      continue
    }

    $profileDir = New-UniqueProfileDir -BrowserName $browser.Name
    $args = Get-LaunchArgs -Browser $browser -ProfileDir $profileDir -Url $Url

    Write-Host ""
    Write-Host ("Launching {0} for {1} seconds (isolated profile)..." -f $browser.Name, $OpenSeconds)

    if ($PSCmdlet.ShouldProcess($browser.Name, "Launch and close after delay")) {
      Start-Process -FilePath $browser.ExePath -ArgumentList $args | Out-Null
      Start-Sleep -Seconds $OpenSeconds

      # Close ONLY the processes that include our unique profile dir in the command line.
      Write-Host ("Closing only the instance started by this script for {0}..." -f $browser.Name)
      Stop-OnlyOurBrowserProcesses -ProcessExeNames $browser.CloseNames -MarkerText $profileDir

      # Cleanup profile dir (best-effort)
      try { Remove-Item -Path $profileDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }

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