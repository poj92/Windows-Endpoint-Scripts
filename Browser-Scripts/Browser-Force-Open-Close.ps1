#Requires -Version 5.1
<#
Browser-Force-Open-Close.ps1

Checks whether ANY installed browser has been opened within the last N hours.
If not, launches an installed browser and closes only the newly started processes after a delay.

Installed browser detection:
- Known paths
- App Paths registry
- Uninstall registry entries (DisplayName / DisplayIcon / InstallLocation)

"Opened within lookback" detection:
- Prefetch last-write time for browser EXE name (best signal when available)
- Fallback: currently running browser processes whose StartTime is within lookback (best-effort)

Exit codes:
  0 = at least one installed browser opened within lookback (no action)
  1 = none opened; script launched and then closed a browser
  2 = ReportOnly and launch would have occurred
  3 = no supported browser installed (nothing to do)
  4 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$LookbackHours = 24,
  [int]$OpenSeconds = 120,
  [string]$Url = 'about:blank',
  [switch]$ReportOnly,

  [ValidateSet('Edge','Chrome','Firefox','Brave','Opera')]
  [string[]]$Preference = @('Edge','Chrome','Firefox','Brave','Opera')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-BrowserMeta {
  param([Parameter(Mandatory)][string]$Name)
  switch ($Name) {
    'Edge'    { @{ Name='Edge';    Proc='msedge';  Exe='msedge.exe'  ; Known=@("$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe",
                                                                               "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") } }
    'Chrome'  { @{ Name='Chrome';  Proc='chrome';  Exe='chrome.exe'  ; Known=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                                                                               "$env:ProgramFiles (x86)\Google\Chrome\Application\chrome.exe") } }
    'Firefox' { @{ Name='Firefox'; Proc='firefox'; Exe='firefox.exe' ; Known=@("$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                                                                               "$env:ProgramFiles (x86)\Mozilla Firefox\firefox.exe") } }
    'Brave'   { @{ Name='Brave';   Proc='brave';   Exe='brave.exe'   ; Known=@("$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe",
                                                                               "$env:ProgramFiles (x86)\BraveSoftware\Brave-Browser\Application\brave.exe") } }
    'Opera'   { @{ Name='Opera';   Proc='opera';   Exe='launcher.exe'; Known=@("$env:ProgramFiles\Opera\launcher.exe",
                                                                               "$env:ProgramFiles (x86)\Opera\launcher.exe",
                                                                               "$env:LOCALAPPDATA\Programs\Opera\launcher.exe") } }
    default { throw "Unknown browser name: $Name" }
  }
}

function Get-AppPathsExe {
  param([Parameter(Mandatory)][string]$ExeName)

  $keys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
  )

  foreach ($k in $keys) {
    try {
      if (Test-Path $k) {
        $val = (Get-ItemProperty -Path $k -ErrorAction Stop).'(default)'
        if ($val -and (Test-Path $val)) { return $val }
      }
    } catch { }
  }
  return $null
}

function Get-UninstallEntries {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  foreach ($p in $paths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
      Where-Object {
        # Defensive: only require DisplayName; ignore SystemComponent if missing
        $_.PSObject.Properties.Match('DisplayName').Count -gt 0 -and
        $_.DisplayName -and
        (($_.PSObject.Properties.Match('SystemComponent').Count -eq 0) -or ($_.SystemComponent -ne 1))
      } |
      ForEach-Object {
        [pscustomobject]@{
          DisplayName     = $_.DisplayName
          DisplayIcon     = (if ($_.PSObject.Properties.Match('DisplayIcon').Count -gt 0) { $_.DisplayIcon } else { $null })
          InstallLocation = (if ($_.PSObject.Properties.Match('InstallLocation').Count -gt 0) { $_.InstallLocation } else { $null })
        }
      }
  }
}

function Clean-DisplayIconPath {
  param([string]$DisplayIcon)
  if (-not $DisplayIcon) { return $null }
  $s = $DisplayIcon.Trim()

  if ($s.StartsWith('"') -and $s.EndsWith('"')) {
    $s = $s.Trim('"')
  }

  # Strip icon index after comma
  $s = ($s -split ',')[0].Trim()

  if ($s -and (Test-Path $s)) { return $s }
  return $null
}

function Find-InstalledBrowsers {
  param([string[]]$Names)

  $uninstall = @(Get-UninstallEntries)
  $installed = @()

  foreach ($name in $Names) {
    $m = Get-BrowserMeta -Name $name

    # 1) Known paths
    $exePath = $null
    foreach ($p in $m.Known) {
      if ($p -and (Test-Path $p)) { $exePath = $p; break }
    }

    # 2) App Paths
    if (-not $exePath) {
      $exePath = Get-AppPathsExe -ExeName $m.Exe
    }

    # 3) Uninstall entries heuristics
    if (-not $exePath) {
      $hits = $uninstall | Where-Object { $_.DisplayName -match $m.Name }
      foreach ($h in $hits) {
        $icon = Clean-DisplayIconPath -DisplayIcon $h.DisplayIcon
        if ($icon) { $exePath = $icon; break }
        if ($h.InstallLocation) {
          $candidate = Join-Path $h.InstallLocation $m.Exe
          if (Test-Path $candidate) { $exePath = $candidate; break }
        }
      }
    }

    if ($exePath) {
      $installed += [pscustomobject]@{
        Name     = $m.Name
        ProcName = $m.Proc
        ExePath  = $exePath
      }
    }
  }

  $installed | Sort-Object Name -Unique
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
    [Parameter(Mandatory)][string]$ProcName,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  $procs = Get-Process -Name $ProcName -ErrorAction SilentlyContinue
  if (-not $procs) { return $null }

  $times = foreach ($p in $procs) {
    try {
      if ($p.StartTime -ge $Cutoff) { $p.StartTime }
    } catch { }
  }

  if (-not $times) { return $null }
  ($times | Sort-Object -Descending | Select-Object -First 1)
}

function Get-LastOpenStatus {
  param(
    [Parameter(Mandatory)]$InstalledBrowsers,
    [Parameter(Mandatory)][datetime]$Cutoff
  )

  foreach ($b in $InstalledBrowsers) {
    $prefetchTime = $null
    try { $prefetchTime = Get-PrefetchLastRun -ExeBaseName $b.ProcName } catch { $prefetchTime = $null }

    $fallbackTime = $null
    if (-not $prefetchTime) {
      $fallbackTime = Get-ProcessLastRunFallback -ProcName $b.ProcName -Cutoff $Cutoff
    }

    $last = if ($prefetchTime) { $prefetchTime } else { $fallbackTime }

    [pscustomobject]@{
      Browser    = $b.Name
      ProcName   = $b.ProcName
      ExePath    = $b.ExePath
      LastOpened = $last
      Source     = if ($prefetchTime) { 'Prefetch' } elseif ($fallbackTime) { 'ProcessStartTime' } else { 'None' }
      OpenedWithinLookback = [bool]($last -and $last -ge $Cutoff)
    }
  }
}

function Start-Browser {
  param([Parameter(Mandatory)][string]$ExePath, [Parameter(Mandatory)][string]$Url)
  Start-Process -FilePath $ExePath -ArgumentList $Url -PassThru
}

function Close-NewBrowserProcesses {
  param(
    [Parameter(Mandatory)][string]$ProcessName,
    [Parameter(Mandatory)][int[]]$ExistingPids,
    [Parameter(Mandatory)][datetime]$LaunchTime
  )

  $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
  if (-not $procs) { return }

  $newProcs = foreach ($p in $procs) {
    if ($ExistingPids -contains $p.Id) { continue }

    $ok = $true
    try {
      if ($p.StartTime -lt $LaunchTime.AddSeconds(-5)) { $ok = $false }
    } catch {
      $ok = $true
    }

    if ($ok) { $p }
  }

  if (-not $newProcs) { return }

  foreach ($p in $newProcs) {
    try {
      if ($p.MainWindowHandle -ne 0) { [void]$p.CloseMainWindow() }
    } catch { }
  }

  Start-Sleep -Seconds 10

  foreach ($p in $newProcs) {
    try {
      if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    } catch { }
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

  Write-Host ("Lookback window: last {0} hours (since {1})" -f $LookbackHours, $cutoff)
  Write-Host "Installed browsers detected:"
  foreach ($b in $installed) { Write-Host ("- {0} ({1}) => {2}" -f $b.Name, $b.ProcName, $b.ExePath) }

  $status = @(Get-LastOpenStatus -InstalledBrowsers $installed -Cutoff $cutoff)

  Write-Host ""
  Write-Host "Last-open status:"
  foreach ($s in $status) {
    $t = if ($s.LastOpened) { $s.LastOpened.ToString('yyyy-MM-dd HH:mm:ss') } else { '<unknown>' }
    Write-Host ("{0}: LastOpened={1} Source={2} OpenedWithinLookback={3}" -f $s.Browser, $t, $s.Source, $s.OpenedWithinLookback)
  }

  $openedRecently = [bool]($status | Where-Object { $_.OpenedWithinLookback } | Select-Object -First 1)
  if ($openedRecently) {
    Write-Host ""
    Write-Host "Result: At least one installed browser was opened within the lookback window. No action taken."
    exit 0
  }

  Write-Host ""
  Write-Host "Result: No installed browser was detected as opened within the lookback window."

  # Choose what to launch: first installed in preference order
  $toLaunch = $null
  foreach ($pref in $Preference) {
    $hit = $installed | Where-Object { $_.Name -eq $pref } | Select-Object -First 1
    if ($hit) { $toLaunch = $hit; break }
  }
  if (-not $toLaunch) { $toLaunch = $installed | Select-Object -First 1 }

  if ($ReportOnly) {
    Write-Host ("ReportOnly: Would launch {0} ({1}) for {2} seconds then close." -f $toLaunch.Name, $toLaunch.ExePath, $OpenSeconds)
    exit 2
  }

  $existing = @(Get-Process -Name $toLaunch.ProcName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
  $launchTime = Get-Date

  Write-Host ("Launching {0} to {1}..." -f $toLaunch.Name, $Url)

  if ($PSCmdlet.ShouldProcess($toLaunch.Name, "Launch and close after delay")) {
    [void](Start-Browser -ExePath $toLaunch.ExePath -Url $Url)
    Start-Sleep -Seconds $OpenSeconds
    Write-Host ("Closing newly-started {0} processes..." -f $toLaunch.ProcName)
    Close-NewBrowserProcesses -ProcessName $toLaunch.ProcName -ExistingPids $existing -LaunchTime $launchTime
  }

  Write-Host "Done."
  exit 1
}
catch {
  Write-Error $_
  exit 4
}