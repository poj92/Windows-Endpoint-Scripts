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

function Send-UserMessage([string]$Text) {
  try { & msg.exe * $Text | Out-Null } catch { }
}

function Get-ActiveUserFromExplorer {
  try {
    $explorers = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $explorers) {
      try {
        $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction SilentlyContinue
        if ($o -and $o.User) {
          $user = if ($o.Domain) { ($o.Domain + "\" + $o.User) } else { ($env:COMPUTERNAME + "\" + $o.User) }
          return [pscustomobject]@{ User=$user; SessionId=[int]$p.SessionId }
        }
      } catch { }
    }
  } catch { }
  return $null
}

function Get-ActiveUser {
  try {
    $csUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
    if ($csUser) {
      $ex = Get-ActiveUserFromExplorer
      if ($ex -and $ex.User -and ($ex.User.ToLowerInvariant() -eq $csUser.ToLowerInvariant())) {
        return $ex
      }
      return [pscustomobject]@{ User=$csUser; SessionId=$null }
    }
  } catch { }
  return Get-ActiveUserFromExplorer
}

function Ensure-ScheduledTasksModule {
  try { Import-Module ScheduledTasks -ErrorAction Stop; return $true } catch { return $false }
}

function Set-RebootDeadlineTask {
  param(
    [Parameter(Mandatory)][datetime]$When,
    [Parameter(Mandatory)][string]$Reason
  )

  # Cancel any previous shutdown timer (prevents Windows /t notifications from earlier runs)
  try { & shutdown.exe /a | Out-Null } catch { }

  if (-not (Ensure-ScheduledTasksModule)) { return $false }
  if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) { return $false }

  $taskName  = "Nexus_Reboot_Deadline"
  $action    = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument ('/r /t 0 /c "' + $Reason + '"')
  $trigger   = New-ScheduledTaskTrigger -Once -At $When
  $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
  $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

  try {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    return $true
  } catch {
    Write-Log ("WARNING: Failed to schedule reboot deadline task: {0}" -f $_.Exception.Message)
    return $false
  }
}

function New-RebootPromptHelperScript {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes,
    [string]$UiTitle,
    [string]$Reason
  )

  $dir = Split-Path -Parent $LogPath
  $helper = Join-Path $dir 'RebootPromptUI.ps1'
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

`$ErrorActionPreference = 'Stop'

# Force a deadline reboot task (avoids Windows /t "about to be logged off" toast)
Import-Module ScheduledTasks -ErrorAction SilentlyContinue
`$TaskName = 'Nexus_Reboot_Deadline_UI'

function Set-RebootDeadline([datetime]`$When) {
  try { & shutdown.exe /a | Out-Null } catch { }

  try {
    `$action  = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument ('/r /t 0 /c "' + `$Reason + '"')
    `$trigger = New-ScheduledTaskTrigger -Once -At `$When
    `$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    `$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName `$TaskName -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force | Out-Null
    return `$true
  } catch { return `$false }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) {
  foreach (`$x in (`$PostponeCsv -split ',')) { try { `$PostponeOptionsMinutes += [int]`$x } catch { } }
}
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) { `$PostponeOptionsMinutes = @(30,60,120) }

`$deadline = (Get-Date).AddMinutes([double]`$CountdownMinutes)
[void](Set-RebootDeadline -When `$deadline)

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

function Update-Countdown {
  `$remain = `$deadline - (Get-Date)
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
  `$deadline = (Get-Date).AddMinutes([double]`$Minutes)
  [void](Set-RebootDeadline -When `$deadline)
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
  try { & shutdown.exe /a | Out-Null } catch { }
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

  $helper = New-RebootPromptHelperScript -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason
  $runningAsSystem = ($env:USERNAME -eq 'SYSTEM')

  # If user-run interactive: launch hidden + STA so UI is stable and no console flashes
  if ([Environment]::UserInteractive -and -not $runningAsSystem) {
    Write-Log "Interactive user context: launching UI hidden (STA)."
    Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList @(
      "-NoProfile","-ExecutionPolicy","Bypass","-STA",
      "-File", "`"$helper`"",
      "-CountdownMinutes", "$CountdownMinutes"
    ) | Out-Null
    return
  }

  # SYSTEM context: schedule UI to run as active user (LogonType Interactive) and run it STA + hidden
  $active = Get-ActiveUser
  if ($active -and $active.User -and (Ensure-ScheduledTasksModule)) {
    $taskName = "Nexus_RebootPromptUI"
    $args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$helper`" -CountdownMinutes $CountdownMinutes"
    $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $args
    $trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(10)
    $principal = New-ScheduledTaskPrincipal -UserId $active.User -LogonType Interactive -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    Write-Log ("SYSTEM context: launching UI in session as {0} (LogonType=Interactive, STA, Hidden)." -f $active.User)

    try {
      Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
      Start-ScheduledTask -TaskName $taskName | Out-Null
      return
    } catch {
      Write-Log "WARNING: Failed to launch UI via ScheduledTasks."
      Write-Log $_.Exception.Message
    }
  }

  # Fallback: schedule reboot + msg
  Write-Log "No interactive UI available; scheduling reboot deadline and notifying with msg.exe."
  Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
  $ok = Set-RebootDeadlineTask -When ((Get-Date).AddMinutes([double]$CountdownMinutes)) -Reason $Reason
  if (-not $ok) {
    & shutdown.exe /r /t ($CountdownMinutes*60) /c ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.") | Out-Null
  }
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

    $rb = 2
    try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

    $list += [pscustomobject]@{ Update=$u; Title=[string]$u.Title; RebootBehavior=$rb }
  }

  [pscustomobject]@{ Session=$session; Updates=$list }
}

function Install-Updates {
  param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][object[]]$UpdatesToInstall)

  if (-not $UpdatesToInstall -or $UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount=0; RebootRequired=$false; ResultCode=0 }
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

  [pscustomobject]@{
    InstalledCount = [int]$coll.Count
    RebootRequired = [bool]$res.RebootRequired
    ResultCode     = [int]$res.ResultCode
  }
}

# ---------------- main ----------------
try {
  if (-not (Test-IsAdmin)) { throw "Run this script elevated (Administrator)." }

  Write-Log "Scanning for Windows updates..."
  $beforePendingReboot = Test-PendingReboot
  if ($beforePendingReboot) { Write-Log "System already indicates a pending reboot." }

  $scan = Get-AvailableUpdates
  $updates = $scan.Updates

  if (-not $updates -or $updates.Count -eq 0) {
    Write-Log "No available updates found."
    if ($beforePendingReboot) {
      if ($ReportOnly) { Write-Log "ReportOnly: Would prompt reboot."; exit 1 }
      Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason
      exit 1
    }
    exit 0
  }

  Write-Log ("Found {0} available update(s)." -f $updates.Count)

  $never     = @($updates | Where-Object { $_.RebootBehavior -eq 0 })
  $mayReboot  = @($updates | Where-Object { $_.RebootBehavior -ne 0 })

  Write-Log ("Updates that should not require reboot: {0}" -f $never.Count)
  Write-Log ("Updates that may/require reboot: {0}" -f $mayReboot.Count)

  if ($ReportOnly) {
    Write-Log "ReportOnly: No installs."
    if ($mayReboot.Count -gt 0 -or $beforePendingReboot) { Write-Log "ReportOnly: Would prompt reboot."; exit 1 }
    exit 0
  }

  $didInstall = $false
  $rebootFromInstall = $false

  if ($never.Count -gt 0) {
    $r1 = Install-Updates -Session $scan.Session -UpdatesToInstall $never
    $didInstall = ($r1.InstalledCount -gt 0)
    $rebootFromInstall = $rebootFromInstall -or $r1.RebootRequired
    Write-Log ("Installed {0} non-reboot update(s). RebootRequired={1}" -f $r1.InstalledCount, $r1.RebootRequired)
  }

  if ($IncludeRebootUpdates -and $mayReboot.Count -gt 0) {
    $r2 = Install-Updates -Session $scan.Session -UpdatesToInstall $mayReboot
    $didInstall = $didInstall -or ($r2.InstalledCount -gt 0)
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

  if (-not $didInstall -and $updates.Count -gt 0) {
    Write-Log "Updates found but none installed (likely only reboot-requiring updates)."
    exit 2
  }

  Write-Log "Finished. No reboot required."
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}