#Requires -Version 5.1
<#
- Installs all available updates that do NOT require reboot (RebootBehavior=NeverReboots)
- If reboot is required (pending reboot OR updates requiring reboot OR install reports reboot):
  Show an interactive user prompt:
    - reboot in X minutes (countdown)
    - Postpone 30m / 1h / 2h
    - Reboot now
- Works under SYSTEM/RMM by launching UI in the logged-in user's session (InteractiveToken scheduled task)

Exit codes:
  0 = no reboot needed; updates installed or none available
  1 = reboot prompt launched/scheduled
  2 = updates found but none installed (e.g., only reboot-requiring updates and IncludeRebootUpdates not set)
  3 = error
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [int]$CountdownMinutes = 10,
  [int[]]$PostponeOptionsMinutes = @(30, 60, 120),
  [switch]$IncludeRebootUpdates,
  [switch]$ReportOnly,

  [string]$UiTitle = "A security message from Nexus Open Systems Ltd",
  [string]$Reason  = "Windows updates require a restart to finish installing.",

  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\WindowsUpdate\WindowsUpdateReboot.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Path (Split-Path -Parent $LogPath) -Force | Out-Null

function Write-Log([string]$Message) {
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] {1}" -f $ts, $Message
  Write-Host $line
  try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PendingReboot {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach ($p in $paths) { if (Test-Path $p) { return $true } }

  try {
    $sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($sess -and $sess.PendingFileRenameOperations) { return $true }
  } catch { }

  return $false
}

function Get-ActiveSessionPresent {
  try {
    $out = & quser 2>$null
    if ($out) {
      foreach ($l in $out) {
        if ($l -match '\sActive\s') { return $true }
      }
    }
  } catch { }
  return $false
}

function Send-UserMessage([string]$Text) {
  try { & msg.exe * $Text | Out-Null } catch { }
}

function Ensure-ScheduledTasksModule {
  try { Import-Module ScheduledTasks -ErrorAction Stop; return $true } catch { return $false }
}

function New-RebootPromptHelperScript {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes,
    [string]$UiTitle,
    [string]$Reason,
    [string]$Dir
  )

  $helper = Join-Path $Dir 'RebootPromptUI.ps1'
  $optsCsv = ($PostponeOptionsMinutes | ForEach-Object { [int]$_ }) -join ','

  $content = @"
param(
  [int]`$CountdownMinutes = $CountdownMinutes,
  [string]`$PostponeCsv = '$optsCsv',
  [string]`$UiTitle = @'
$UiTitle
'@,
  [string]`$Reason = @'
$Reason
'@
)

Set-StrictMode -Off
`$ErrorActionPreference = 'Stop'

# Make WinForms reliable
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Suppress shutdown /a output completely
function Abort-ShutdownSilently {
  try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { }
}

# Schedule reboot at deadline via ScheduledTasks (no Windows /t toast)
function Set-RebootDeadline([datetime]`$When) {
  Abort-ShutdownSilently

  try { Import-Module ScheduledTasks -ErrorAction Stop } catch { return `$false }

  try {
    `$action    = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument ('/r /t 0 /c "' + `$Reason + '"')
    `$trigger   = New-ScheduledTaskTrigger -Once -At `$When
    `$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    `$settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    Register-ScheduledTask -TaskName 'Nexus_Reboot_Deadline' -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force | Out-Null
    return `$true
  } catch {
    return `$false
  }
}

# Parse postpone options
`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) {
  foreach (`$x in (`$PostponeCsv -split ',')) { try { `$PostponeOptionsMinutes += [int]`$x } catch { } }
}
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) { `$PostponeOptionsMinutes = @(30,60,120) }

# IMPORTANT: use script scope so postpone updates the live countdown
`$script:deadline = (Get-Date).AddMinutes([double]`$CountdownMinutes)
[void](Set-RebootDeadline -When `$script:deadline)

`$form = New-Object System.Windows.Forms.Form
`$form.Text = `$UiTitle
`$form.Size = New-Object System.Drawing.Size(720, 300)
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true
`$form.Add_FormClosing({ if (`$_.CloseReason -eq 'UserClosing') { `$_.Cancel = `$true } })

`$label1 = New-Object System.Windows.Forms.Label
`$label1.AutoSize = `$true
`$label1.MaximumSize = New-Object System.Drawing.Size(680, 0)
`$label1.Location = New-Object System.Drawing.Point(18, 18)
`$label1.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$label1.Text = `$Reason
`$form.Controls.Add(`$label1)

`$label2 = New-Object System.Windows.Forms.Label
`$label2.AutoSize = `$true
`$label2.Location = New-Object System.Drawing.Point(18, 70)
`$label2.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
`$form.Controls.Add(`$label2)

`$label3 = New-Object System.Windows.Forms.Label
`$label3.AutoSize = `$true
`$label3.Location = New-Object System.Drawing.Point(18, 105)
`$label3.Font = New-Object System.Drawing.Font('Segoe UI', 9)
`$label3.Text = ''
`$form.Controls.Add(`$label3)

function Update-Countdown {
  `$remain = `$script:deadline - (Get-Date)
  if (`$remain.TotalSeconds -le 0) {
    `$label2.Text = 'Rebooting now...'
    & shutdown.exe /r /t 0 /c "`$Reason" | Out-Null
    Start-Sleep -Seconds 1
    `$form.Close()
    return
  }
  `$mins = [int][Math]::Floor(`$remain.TotalMinutes)
  `$secs = [int]`$remain.Seconds
  `$label2.Text = ('Time remaining: {0:00}:{1:00}' -f `$mins, `$secs)
}

`$timer = New-Object System.Windows.Forms.Timer
`$timer.Interval = 1000
`$timer.Add_Tick({ Update-Countdown })
`$timer.Start()
Update-Countdown

`$combo = New-Object System.Windows.Forms.ComboBox
`$combo.DropDownStyle = 'DropDownList'
`$combo.Location = New-Object System.Drawing.Point(18, 130)
`$combo.Size = New-Object System.Drawing.Size(280, 28)

foreach (`$m in `$PostponeOptionsMinutes) {
  if (`$m -eq 30)      { [void]`$combo.Items.Add('Postpone 30 minutes') }
  elseif (`$m -eq 60)  { [void]`$combo.Items.Add('Postpone 1 hour') }
  elseif (`$m -eq 120) { [void]`$combo.Items.Add('Postpone 2 hours') }
  else                { [void]`$combo.Items.Add(("Postpone {0} minutes" -f `$m)) }
}
`$combo.SelectedIndex = 0
`$form.Controls.Add(`$combo)

function Get-SelectedMinutes {
  `$idx = `$combo.SelectedIndex
  if (`$idx -lt 0) { return `$PostponeOptionsMinutes[0] }
  return `$PostponeOptionsMinutes[`$idx]
}

function Set-Postpone([int]`$Minutes) {
  `$script:deadline = (Get-Date).AddMinutes([double]`$Minutes)

  if (Set-RebootDeadline -When `$script:deadline) {
    `$label3.Text = ('Postponed until {0}' -f `$script:deadline.ToString('HH:mm'))
  } else {
    `$label3.Text = 'Postpone failed (could not schedule reboot).'
  }
  Update-Countdown
}

`$btnPostpone = New-Object System.Windows.Forms.Button
`$btnPostpone.Text = 'Postpone'
`$btnPostpone.Size = New-Object System.Drawing.Size(140, 36)
`$btnPostpone.Location = New-Object System.Drawing.Point(315, 128)
`$btnPostpone.Add_Click({ Set-Postpone -Minutes (Get-SelectedMinutes) })
`$form.Controls.Add(`$btnPostpone)

`$btnNow = New-Object System.Windows.Forms.Button
`$btnNow.Text = 'Reboot now'
`$btnNow.Size = New-Object System.Drawing.Size(140, 36)
`$btnNow.Location = New-Object System.Drawing.Point(470, 128)
`$btnNow.Add_Click({
  Abort-ShutdownSilently
  & shutdown.exe /r /t 0 /c "`$Reason" | Out-Null
  `$form.Close()
})
`$form.Controls.Add(`$btnNow)

[void]`$form.ShowDialog()
"@

  Set-Content -Path $helper -Value $content -Encoding UTF8 -Force
  return $helper
}

function Start-RebootPrompt {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes,
    [string]$UiTitle,
    [string]$Reason
  )

  $dir = Split-Path -Parent $LogPath
  $helper = New-RebootPromptHelperScript -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason -Dir $dir

  # If no active session, no interactive UI is possible
  if (-not (Get-ActiveSessionPresent)) {
    Write-Log "No active user session detected; falling back to msg.exe + scheduled reboot deadline."
    Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
    if (Ensure-ScheduledTasksModule) {
      $when = (Get-Date).AddMinutes([double]$CountdownMinutes)
      try {
        $action    = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument ('/r /t 0 /c "' + $Reason + '"')
        $trigger   = New-ScheduledTaskTrigger -Once -At $when
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName 'Nexus_Reboot_Deadline' -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
      } catch { }
    }
    return
  }

  # Create a wrapper .cmd so schtasks /TR quoting is stable
  $wrapper = Join-Path (Split-Path -Parent $helper) 'RunRebootPrompt.cmd'
  $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $cmd = "@echo off`r`n`"$ps`" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$helper`" -CountdownMinutes $CountdownMinutes`r`n"
  Set-Content -Path $wrapper -Value $cmd -Encoding ASCII -Force

  $taskName = "Nexus_RebootPromptUI"

  # Locale-safe date for schtasks
  $dt = Get-Date
  $sd = $dt.ToString((Get-Culture).DateTimeFormat.ShortDatePattern) -replace '[\.\-]', '/'
  $st = $dt.AddMinutes(1).ToString('HH:mm')

  Write-Log "Launching UI as SYSTEM (interactive) via schtasks /IT..."

  # Clean old task name if present
  try { & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null } catch { }

  $createOut = & schtasks.exe /Create /TN $taskName /TR "`"$wrapper`"" /SC ONCE /ST $st /SD $sd /RU SYSTEM /RL HIGHEST /IT /F 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Log "schtasks /Create failed; falling back to msg.exe."
    Write-Log ($createOut -join ' ')
    Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
    return
  }

  $runOut = & schtasks.exe /Run /TN $taskName 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Log "schtasks /Run failed; falling back to msg.exe."
    Write-Log ($runOut -join ' ')
    Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
    return
  }

  # Optional: delete the launcher task after starting (helper remains running)
  try { & schtasks.exe /Delete /TN $taskName /F 2>$null | Out-Null } catch { }
}

function Get-AvailableUpdates {
  $session  = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $criteria = "IsInstalled=0 and IsHidden=0"
  $result   = $searcher.Search($criteria)

  $list = @()
  for ($i=0; $i -lt $result.Updates.Count; $i++) {
    $u = $result.Updates.Item($i)
    try { if ($u.EulaAccepted -eq $false) { $u.AcceptEula() } } catch { }

    # 0=NeverReboots, 1=AlwaysRequiresReboot, 2=CanRequestReboot
    $rb = 2
    try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

    $list += [pscustomobject]@{ Update=$u; Title=[string]$u.Title; RebootBehavior=$rb }
  }

  [pscustomobject]@{ Session=$session; Updates=$list }
}

function Install-Updates {
  param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][object[]]$UpdatesToInstall)

  if (-not $UpdatesToInstall -or $UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount=0; RebootRequired=$false }
  }

  $coll = New-Object -ComObject Microsoft.Update.UpdateColl
  foreach ($x in $UpdatesToInstall) { [void]$coll.Add($x.Update) }

  $downloader = $Session.CreateUpdateDownloader()
  $downloader.Updates = $coll
  Write-Log ("Downloading {0} update(s)..." -f $coll.Count)
  [void]$downloader.Download()

  $installer = $Session.CreateUpdateInstaller()
  $installer.Updates = $coll
  Write-Log ("Installing {0} update(s)..." -f $coll.Count)
  $res = $installer.Install()

  [pscustomobject]@{ InstalledCount=[int]$coll.Count; RebootRequired=[bool]$res.RebootRequired }
}

# ---------------- main ----------------
try {
  Write-Host "Windows-Update-Force-Apply"

  if (-not (Test-IsAdmin)) { throw "Run this script elevated (Administrator / SYSTEM)." }

  Write-Log "Scanning for Windows updates..."
  $beforePendingReboot = Test-PendingReboot
  if ($beforePendingReboot) { Write-Log "System already indicates a pending reboot." }

  $scan = Get-AvailableUpdates
  $updates = $scan.Updates

  if (-not $updates -or $updates.Count -eq 0) {
    Write-Log "No available updates found."
    if ($beforePendingReboot -and -not $ReportOnly) {
      Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason
      exit 1
    }
    exit 0
  }

  Write-Log ("Found {0} available update(s)." -f $updates.Count)

  $never    = @($updates | Where-Object { $_.RebootBehavior -eq 0 })
  $mayReboot = @($updates | Where-Object { $_.RebootBehavior -ne 0 })

  Write-Log ("Updates that should not require reboot: {0}" -f $never.Count)
  Write-Log ("Updates that may/require reboot: {0}" -f $mayReboot.Count)

  if ($ReportOnly) {
    Write-Log "ReportOnly: No installs."
    if ($beforePendingReboot -or $mayReboot.Count -gt 0) { exit 2 }
    exit 0
  }

  $rebootFromInstall = $false

  if ($never.Count -gt 0) {
    $r1 = Install-Updates -Session $scan.Session -UpdatesToInstall $never
    $rebootFromInstall = $rebootFromInstall -or $r1.RebootRequired
    Write-Log ("Installed {0} non-reboot update(s). RebootRequired={1}" -f $r1.InstalledCount, $r1.RebootRequired)
  }

  if ($IncludeRebootUpdates -and $mayReboot.Count -gt 0) {
    $r2 = Install-Updates -Session $scan.Session -UpdatesToInstall $mayReboot
    $rebootFromInstall = $rebootFromInstall -or $r2.RebootRequired
    Write-Log ("Installed {0} reboot update(s). RebootRequired={1}" -f $r2.InstalledCount, $r2.RebootRequired)
  }

  $afterPendingReboot = Test-PendingReboot
  $needsReboot = $beforePendingReboot -or $afterPendingReboot -or $rebootFromInstall -or ($mayReboot.Count -gt 0)

  if ($needsReboot) {
    Write-Log "Reboot required/pending. Prompting user."
    Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason
    exit 1
  }

  Write-Log "Finished. No reboot required."
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}