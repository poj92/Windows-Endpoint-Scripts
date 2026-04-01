#Requires -Version 5.1
<#
Windows Update + Consent-only Reboot + Auto Post-boot Update + Auto Re-prompt

Behavior:
- NEVER reboots automatically. Reboot only occurs when user clicks "Reboot now".
- "Postpone" schedules a reminder that shows the same UI later (no reboot).
- UI appears at user logon ONLY if reboot is actually required.
- Post-boot worker re-runs Windows Update automatically after a user-initiated reboot.
- If another reboot becomes required after post-boot updates, the user is prompted again at logon.
- Logs user clicks, UI shown, ONLOGON decisions, report-only output, and POSTBOOT actions.

Exit codes:
  0 = no reboot needed; updates installed or none available
  1 = reboot required/pending; prompt launched or scheduled (no forced reboot)
  2 = report-only or updates found but none installed under rules
  3 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [int]$CountdownMinutes = 10,
  [int[]]$PostponeOptionsMinutes = @(30, 60, 120),

  [switch]$IncludeRebootUpdates,
  [bool]$EnsureLatestCumulativeUpdate = $true,
  [switch]$ReportOnly,

  [int]$PostRebootMaxPasses = 4,

  [string]$UiTitle = "A security message from Nexus Open Systems Ltd",
  [string]$Reason  = "Windows updates require a restart to finish installing. Please plug your computer into power if it's not already, save your work, and restart as soon as possible to ensure your system is secure and up to date.",

  [string]$BaseDir = "$env:ProgramData\NexusOpenSystems\WindowsUpdate",
  [string]$LogPath = "$env:ProgramData\NexusOpenSystems\WindowsUpdate\WindowsUpdateReboot.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $LogPath) -Force | Out-Null

$StatePath          = Join-Path $BaseDir "WU-State.json"
$UiHelperPath       = Join-Path $BaseDir "RebootPromptUI.ps1"
$PostBootPath       = Join-Path $BaseDir "WU-PostBootWorker.ps1"

$TaskPostBoot       = "Nexus_WU_PostBootWorker"
$TaskOnLogon        = "Nexus_WU_PromptOnLogon"
$TaskImmediate      = "Nexus_WU_PromptNow"

function Write-Log([string]$Message) {
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] {1}" -f $ts, $Message
  Write-Host $line
  try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PendingReboot {
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach ($p in $paths) {
    if (Test-Path $p) { return $true }
  }

  try {
    $sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
      -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
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

function Save-State([hashtable]$obj) {
  try {
    ($obj | ConvertTo-Json -Depth 6) | Set-Content -Path $StatePath -Encoding UTF8 -Force
  } catch { }
}

function Remove-Task([string]$Name) {
  try { Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
  try { schtasks.exe /Delete /TN $Name /F *> $null } catch { }
}

function Abort-AnyShutdown {
  try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { }
}

function Import-ScheduledTasksModule {
  try {
    Import-Module ScheduledTasks -ErrorAction Stop
    return $true
  } catch {
    Write-Log "ScheduledTasks module not available: $($_.Exception.Message)"
    return $false
  }
}

function Remove-LogonPromptTaskIfNotNeeded {
  if (-not (Test-PendingReboot)) {
    Remove-Task $TaskOnLogon
    Write-Log "Removed ONLOGON prompt task (reboot not required)."
  }
}

# ---------------- Windows Update COM scan/install ----------------
function Get-AvailableUpdates {
  $session  = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
  $result   = $searcher.Search($criteria)

  $list = @()
  if ($result -and $result.Updates) {
    for ($i = 0; $i -lt $result.Updates.Count; $i++) {
      $u = $result.Updates.Item($i)
      try { if ($u.EulaAccepted -eq $false) { $u.AcceptEula() } } catch { }

      $rb = 2
      try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

      $title = [string]$u.Title
      $isSSU = ($title -match 'Servicing Stack Update')
      $isLCU = ($title -match 'Cumulative Update') -and ($title -match 'Windows')

      $list += [pscustomobject]@{
        Update         = $u
        Title          = $title
        RebootBehavior = $rb
        IsSSU          = $isSSU
        IsLCU          = $isLCU
      }
    }
  }

  return [pscustomobject]@{
    Session = $session
    Updates = $list
  }
}

function Install-Updates {
  param(
    [Parameter(Mandatory)]$Session,
    [Parameter(Mandatory)][object[]]$UpdatesToInstall
  )

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
  }
}

function Run-WindowsUpdatePasses {
  param(
    [int]$MaxPasses,
    [switch]$IncludeRebootUpdates,
    [bool]$EnsureLatestCumulativeUpdate
  )

  $pass = 0
  $rebootFromInstall = $false
  $installedTotal = 0
  $foundAny = $false
  $eligibleFound = $false

  while ($pass -lt $MaxPasses) {
    $pass++
    $scan = Get-AvailableUpdates
    $updates = $scan.Updates

    if (-not $updates -or $updates.Count -eq 0) {
      Write-Log ("WU pass {0}/{1}: No available updates." -f $pass, $MaxPasses)
      break
    }

    $foundAny = $true
    Write-Log ("WU pass {0}/{1}: Found {2} update(s)." -f $pass, $MaxPasses, $updates.Count)

    $ssu    = @($updates | Where-Object { $_.IsSSU })
    $lcu    = @($updates | Where-Object { $_.IsLCU })
    $others = @($updates | Where-Object { -not $_.IsSSU -and -not $_.IsLCU })

    $toInstall = @()
    if ($EnsureLatestCumulativeUpdate) {
      $toInstall += $ssu
      $toInstall += $lcu
      if ($IncludeRebootUpdates) { $toInstall += $others }
      else { $toInstall += @($others | Where-Object { $_.RebootBehavior -eq 0 }) }
    } else {
      if ($IncludeRebootUpdates) { $toInstall = $updates }
      else { $toInstall = @($updates | Where-Object { $_.RebootBehavior -eq 0 }) }
    }

    $toInstall = @($toInstall | Select-Object -Unique)

    if (-not $toInstall -or $toInstall.Count -eq 0) {
      Write-Log ("WU pass {0}/{1}: Updates found but none eligible under rules." -f $pass, $MaxPasses)
      break
    }

    $eligibleFound = $true
    $r = Install-Updates -Session $scan.Session -UpdatesToInstall $toInstall
    $installedTotal += $r.InstalledCount
    $rebootFromInstall = $rebootFromInstall -or $r.RebootRequired

    Write-Log ("WU pass {0}/{1}: Installed {2}. RebootRequired={3}" -f $pass, $MaxPasses, $r.InstalledCount, $r.RebootRequired)

    if ($rebootFromInstall) { break }
  }

  return [pscustomobject]@{
    FoundAny           = $foundAny
    EligibleFound      = $eligibleFound
    InstalledTotal     = $installedTotal
    RebootFromInstall  = $rebootFromInstall
  }
}

# ---------------- UI helper ----------------
function Write-UiHelper {
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
'@,
  [string]`$LogPath = @'
$LogPath
'@
)

Set-StrictMode -Off
`$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Write-LogLocal([string]`$msg) {
  try {
    `$ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path `$LogPath -Value ("[{0}] {1}" -f `$ts, `$msg) -Encoding UTF8
  } catch { }
}

function Test-PendingReboot {
  `$paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach (`$p in `$paths) {
    if (Test-Path `$p) { return `$true }
  }
  try {
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

function Import-ScheduledTasksModuleLocal {
  try {
    Import-Module ScheduledTasks -ErrorAction Stop
    return `$true
  } catch {
    Write-LogLocal ("UI: ScheduledTasks module unavailable: {0}" -f `$_.Exception.Message)
    return `$false
  }
}

function Get-ReminderTaskName {
  `$safeUser = (`$env:USERNAME -replace '[^A-Za-z0-9_.-]', '_')
  if ([string]::IsNullOrWhiteSpace(`$safeUser)) { `$safeUser = 'UnknownUser' }
  return ("Nexus_WU_RebootReminder_{0}" -f `$safeUser)
}

function Schedule-Reminder([datetime]`$When) {
  if (-not (Import-ScheduledTasksModuleLocal)) { return `$false }

  try {
    `$taskName = Get-ReminderTaskName
    try { Unregister-ScheduledTask -TaskName `$taskName -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null } catch { }

    `$psExe  = "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    `$self   = "`$PSCommandPath"
    `$args   = "-NoProfile -ExecutionPolicy Bypass -STA -File `"`$self`" -CountdownMinutes `$CountdownMinutes -PostponeCsv `"`$PostponeCsv`""

    `$action    = New-ScheduledTaskAction -Execute `$psExe -Argument `$args
    `$trigger   = New-ScheduledTaskTrigger -Once -At `$When
    `$principal = New-ScheduledTaskPrincipal -UserId `$env:USERNAME -LogonType InteractiveToken -RunLevel Limited
    `$settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    Register-ScheduledTask -TaskName `$taskName -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force | Out-Null
    return `$true
  } catch {
    Write-LogLocal ("UI: failed to schedule reminder: {0}" -f `$_.Exception.Message)
    return `$false
  }
}

if (-not (Test-PendingReboot)) {
  Write-LogLocal "UI: reboot not required; exiting."
  exit 0
}

`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) {
  foreach (`$x in (`$PostponeCsv -split ',')) {
    try { `$PostponeOptionsMinutes += [int]`$x } catch { }
  }
}
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) {
  `$PostponeOptionsMinutes = @(30,60,120)
}

Write-LogLocal "UI_SHOWN"

`$script:AllowClose = `$false
`$script:deadlineInfoOnly = (Get-Date).AddMinutes([double]`$CountdownMinutes)

`$form = New-Object System.Windows.Forms.Form
`$form.Text = `$UiTitle
`$form.Size = New-Object System.Drawing.Size(780, 340)
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true
`$form.MaximizeBox = `$false
`$form.MinimizeBox = `$false
`$form.FormBorderStyle = 'FixedDialog'
`$form.Add_FormClosing({
  if (-not `$script:AllowClose -and `$_.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
    `$_.Cancel = `$true
  }
})

`$label1 = New-Object System.Windows.Forms.Label
`$label1.AutoSize = `$true
`$label1.MaximumSize = New-Object System.Drawing.Size(740, 0)
`$label1.Location = New-Object System.Drawing.Point(18, 18)
`$label1.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$label1.Text = `$Reason
`$form.Controls.Add(`$label1)

`$label2 = New-Object System.Windows.Forms.Label
`$label2.AutoSize = `$true
`$label2.Location = New-Object System.Drawing.Point(18, 88)
`$label2.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
`$form.Controls.Add(`$label2)

`$label3 = New-Object System.Windows.Forms.Label
`$label3.AutoSize = `$true
`$label3.Location = New-Object System.Drawing.Point(18, 122)
`$label3.Font = New-Object System.Drawing.Font('Segoe UI', 9)
`$label3.Text = 'No restart will occur unless you choose Reboot now.'
`$form.Controls.Add(`$label3)

function Update-Countdown {
  `$remain = `$script:deadlineInfoOnly - (Get-Date)
  if (`$remain.TotalSeconds -le 0) {
    `$label2.Text = 'Countdown ended: please choose Reboot now or Postpone.'
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
`$combo.Location = New-Object System.Drawing.Point(18, 165)
`$combo.Size = New-Object System.Drawing.Size(340, 28)
foreach (`$m in `$PostponeOptionsMinutes) {
  if (`$m -eq 30)      { [void]`$combo.Items.Add('Postpone 30 minutes') }
  elseif (`$m -eq 60)  { [void]`$combo.Items.Add('Postpone 1 hour') }
  elseif (`$m -eq 120) { [void]`$combo.Items.Add('Postpone 2 hours') }
  else                 { [void]`$combo.Items.Add(("Postpone {0} minutes" -f `$m)) }
}
`$combo.SelectedIndex = 0
`$form.Controls.Add(`$combo)

function Get-SelectedMinutes {
  if (`$combo.SelectedIndex -lt 0) { return `$PostponeOptionsMinutes[0] }
  return `$PostponeOptionsMinutes[`$combo.SelectedIndex]
}

`$btnPostpone = New-Object System.Windows.Forms.Button
`$btnPostpone.Text = 'Postpone'
`$btnPostpone.Size = New-Object System.Drawing.Size(170, 42)
`$btnPostpone.Location = New-Object System.Drawing.Point(380, 160)
`$btnPostpone.Add_Click({
  `$mins = Get-SelectedMinutes
  `$when = (Get-Date).AddMinutes([double]`$mins)
  if (Schedule-Reminder -When `$when) {
    Write-LogLocal ("USER_ACTION: Postpone {0} minutes (re-prompt at {1})" -f `$mins, `$when.ToString('yyyy-MM-dd HH:mm:ss'))
    `$label3.Text = ('Okay — we will remind you again at {0}.' -f `$when.ToString('HH:mm'))
    Start-Sleep -Milliseconds 800
    `$script:AllowClose = `$true
    `$form.Close()
  } else {
    Write-LogLocal ("USER_ACTION: Postpone FAILED ({0} minutes)" -f `$mins)
    `$label3.Text = 'Postpone failed (could not schedule reminder).'
  }
})
`$form.Controls.Add(`$btnPostpone)

`$btnNow = New-Object System.Windows.Forms.Button
`$btnNow.Text = 'Reboot now'
`$btnNow.Size = New-Object System.Drawing.Size(170, 42)
`$btnNow.Location = New-Object System.Drawing.Point(570, 160)
`$btnNow.Add_Click({
  Write-LogLocal "USER_ACTION: Reboot now"
  `$script:AllowClose = `$true
  try { & shutdown.exe /r /t 0 /c "`$Reason" | Out-Null } catch { }
  `$form.Close()
})
`$form.Controls.Add(`$btnNow)

[void]`$form.ShowDialog()
"@

  Set-Content -Path $UiHelperPath -Value $content -Encoding UTF8 -Force
}

# ---------------- User-context prompt tasks ----------------
function Register-LogonPromptTask {
  if (-not (Import-ScheduledTasksModule)) { throw "ScheduledTasks module is required." }

  Remove-Task $TaskOnLogon

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $arg   = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$UiHelperPath`""

  $action    = New-ScheduledTaskAction -Execute $psExe -Argument $arg
  $trigger   = New-ScheduledTaskTrigger -AtLogOn
  $principal = New-ScheduledTaskPrincipal -GroupId "Users" -RunLevel Limited
  $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

  Register-ScheduledTask -TaskName $TaskOnLogon -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Log "Registered ONLOGON prompt task (runs in user context; UI self-checks pending reboot)."
}

function Launch-UI-IfRebootRequiredNow {
  if (-not (Test-PendingReboot)) {
    Write-Log "Immediate UI not launched: reboot not required."
    return
  }

  if (-not (Get-ActiveSessionPresent)) {
    Write-Log "Immediate UI not launched: no active user session. ONLOGON task will handle later."
    return
  }

  if (-not (Import-ScheduledTasksModule)) {
    Write-Log "Immediate UI not launched: ScheduledTasks module unavailable."
    return
  }

  try {
    Remove-Task $TaskImmediate

    $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $arg   = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$UiHelperPath`""

    $action    = New-ScheduledTaskAction -Execute $psExe -Argument $arg
    $trigger   = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
    $principal = New-ScheduledTaskPrincipal -GroupId "Users" -RunLevel Limited
    $settings  = New-ScheduledTaskSettingsSet -DeleteExpiredTaskAfter (New-TimeSpan -Hours 2) `
                                              -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
                                              -MultipleInstances IgnoreNew `
                                              -AllowStartIfOnBatteries `
                                              -DontStopIfGoingOnBatteries

    Register-ScheduledTask -TaskName $TaskImmediate -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
    Start-ScheduledTask -TaskName $TaskImmediate
    Write-Log "Attempted immediate UI launch via user-context scheduled task."
  } catch {
    Write-Log ("Immediate UI launch failed: {0}" -f $_.Exception.Message)
  }
}

# ---------------- Post-boot worker ----------------
function Write-PostBootWorker {
  $content = @"
param(
  [string]`$LogPath,
  [int]`$MaxPasses = $PostRebootMaxPasses,
  [switch]`$IncludeRebootUpdates,
  [bool]`$EnsureLatestCumulativeUpdate = $EnsureLatestCumulativeUpdate,
  [string]`$TaskOnLogon = '$TaskOnLogon',
  [string]`$TaskPostBoot = '$TaskPostBoot'
)

Set-StrictMode -Off
`$ErrorActionPreference = 'Stop'

function Write-Log([string]`$Message) {
  `$ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Add-Content -Path `$LogPath -Value ("[{0}] {1}" -f `$ts, `$Message) -Encoding UTF8
}

function Test-PendingReboot {
  `$paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach (`$p in `$paths) {
    if (Test-Path `$p) { return `$true }
  }
  try {
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

function Remove-TaskLocal([string]`$Name) {
  try { Unregister-ScheduledTask -TaskName `$Name -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null } catch { }
  try { schtasks.exe /Delete /TN `$Name /F *> `$null } catch { }
}

function Get-AvailableUpdates {
  `$session  = New-Object -ComObject Microsoft.Update.Session
  `$searcher = `$session.CreateUpdateSearcher()
  `$criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
  `$result   = `$searcher.Search(`$criteria)

  `$list = @()
  if (`$result -and `$result.Updates) {
    for (`$i = 0; `$i -lt `$result.Updates.Count; `$i++) {
      `$u = `$result.Updates.Item(`$i)
      try { if (`$u.EulaAccepted -eq `$false) { `$u.AcceptEula() } } catch { }

      `$rb = 2
      try { `$rb = [int]`$u.InstallationBehavior.RebootBehavior } catch { `$rb = 2 }

      `$title = [string]`$u.Title
      `$isSSU = (`$title -match 'Servicing Stack Update')
      `$isLCU = (`$title -match 'Cumulative Update') -and (`$title -match 'Windows')

      `$list += [pscustomobject]@{
        Update         = `$u
        Title          = `$title
        RebootBehavior = `$rb
        IsSSU          = `$isSSU
        IsLCU          = `$isLCU
      }
    }
  }

  return [pscustomobject]@{
    Session = `$session
    Updates = `$list
  }
}

function Install-Updates([object]`$Session, [object[]]`$UpdatesToInstall) {
  if (-not `$UpdatesToInstall -or `$UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount = 0; RebootRequired = `$false }
  }

  `$coll = New-Object -ComObject Microsoft.Update.UpdateColl
  foreach (`$x in `$UpdatesToInstall) { [void]`$coll.Add(`$x.Update) }

  `$downloader = `$Session.CreateUpdateDownloader()
  `$downloader.Updates = `$coll
  Write-Log ("POSTBOOT: Downloading {0} update(s)..." -f `$coll.Count)
  [void]`$downloader.Download()

  `$installer = `$Session.CreateUpdateInstaller()
  `$installer.Updates = `$coll
  Write-Log ("POSTBOOT: Installing {0} update(s)..." -f `$coll.Count)
  `$res = `$installer.Install()

  return [pscustomobject]@{
    InstalledCount = [int]`$coll.Count
    RebootRequired = [bool]`$res.RebootRequired
  }
}

Write-Log "POSTBOOT: Starting Windows Update worker."
`$pass = 0
`$rebootFromInstall = `$false
`$installedTotal = 0

while (`$pass -lt `$MaxPasses) {
  `$pass++
  `$scan = Get-AvailableUpdates
  `$updates = `$scan.Updates

  if (-not `$updates -or `$updates.Count -eq 0) {
    Write-Log ("POSTBOOT: Pass {0}/{1}: No available updates." -f `$pass, `$MaxPasses)
    break
  }

  Write-Log ("POSTBOOT: Pass {0}/{1}: Found {2} update(s)." -f `$pass, `$MaxPasses, `$updates.Count)

  `$ssu    = @(`$updates | Where-Object { `$_.IsSSU })
  `$lcu    = @(`$updates | Where-Object { `$_.IsLCU })
  `$others = @(`$updates | Where-Object { -not `$_.IsSSU -and -not `$_.IsLCU })

  `$toInstall = @()
  if (`$EnsureLatestCumulativeUpdate) {
    `$toInstall += `$ssu
    `$toInstall += `$lcu
    if (`$IncludeRebootUpdates) { `$toInstall += `$others }
    else { `$toInstall += @(`$others | Where-Object { `$_.RebootBehavior -eq 0 }) }
  } else {
    if (`$IncludeRebootUpdates) { `$toInstall = `$updates }
    else { `$toInstall = @(`$updates | Where-Object { `$_.RebootBehavior -eq 0 }) }
  }

  `$toInstall = @(`$toInstall | Select-Object -Unique)
  if (-not `$toInstall -or `$toInstall.Count -eq 0) {
    Write-Log ("POSTBOOT: Pass {0}/{1}: Updates found but none eligible under rules." -f `$pass, `$MaxPasses)
    break
  }

  `$r = Install-Updates -Session `$scan.Session -UpdatesToInstall `$toInstall
  `$installedTotal += `$r.InstalledCount
  `$rebootFromInstall = `$rebootFromInstall -or `$r.RebootRequired

  Write-Log ("POSTBOOT: Pass {0}/{1}: Installed {2}. RebootRequired={3}" -f `$pass, `$MaxPasses, `$r.InstalledCount, `$r.RebootRequired)

  if (`$rebootFromInstall) { break }
}

`$pending = Test-PendingReboot
Write-Log ("POSTBOOT: InstalledTotal={0} RebootFromInstall={1} PendingReboot={2}" -f `$installedTotal, `$rebootFromInstall, `$pending)

if (`$pending -or `$rebootFromInstall) {
  Write-Log "POSTBOOT: Reboot required. Keeping ONLOGON prompt task."
} else {
  Remove-TaskLocal `$TaskOnLogon
  Write-Log "POSTBOOT: System is clean. Removed ONLOGON prompt task."
}

Remove-TaskLocal `$TaskPostBoot
Write-Log "POSTBOOT: Worker task removed."
"@

  Set-Content -Path $PostBootPath -Value $content -Encoding UTF8 -Force
}

function Register-PostBootWorkerTask {
  if (-not (Import-ScheduledTasksModule)) { throw "ScheduledTasks module is required." }

  Write-PostBootWorker
  Remove-Task $TaskPostBoot

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $arg   = "-NoProfile -ExecutionPolicy Bypass -File `"$PostBootPath`" -LogPath `"$LogPath`" -MaxPasses $PostRebootMaxPasses -TaskOnLogon `"$TaskOnLogon`" -TaskPostBoot `"$TaskPostBoot`""
  if ($IncludeRebootUpdates) { $arg += " -IncludeRebootUpdates" }
  if ($EnsureLatestCumulativeUpdate) { $arg += " -EnsureLatestCumulativeUpdate `$$true" }

  $action    = New-ScheduledTaskAction -Execute $psExe -Argument $arg
  $trigger   = New-ScheduledTaskTrigger -AtStartup
  $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
  $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

  Register-ScheduledTask -TaskName $TaskPostBoot -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Log "Registered post-boot Windows Update worker task (runs at startup as SYSTEM)."
}

# ---------------- MAIN ----------------
try {
  Write-Host "Windows Update (Consent-only + auto re-run + user-context prompt; prompt only if reboot required)"

  if (-not (Test-IsAdmin)) {
    throw "Run this script elevated (Administrator / SYSTEM)."
  }

  Abort-AnyShutdown
  Write-UiHelper

  if ($ReportOnly) {
    Write-Log "ReportOnly=True. Scanning only."
    $beforePending = Test-PendingReboot
    if ($beforePending) { Write-Log "ReportOnly: System already indicates a pending reboot." }

    $result = Run-WindowsUpdatePasses -MaxPasses 1 -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate
    $afterPending = Test-PendingReboot
    $needsReboot = $beforePending -or $afterPending -or $result.RebootFromInstall

    Write-Log ("ReportOnly: FoundAny={0} EligibleFound={1} InstalledTotal={2} NeedsReboot={3}" -f `
      $result.FoundAny, $result.EligibleFound, $result.InstalledTotal, $needsReboot)

    exit 2
  }

  Register-LogonPromptTask

  Write-Log "Pre-reboot phase: scanning/installing updates..."
  $beforePending = Test-PendingReboot
  if ($beforePending) { Write-Log "System already indicates a pending reboot." }

  $result = Run-WindowsUpdatePasses -MaxPasses 3 -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate

  $afterPending = Test-PendingReboot
  $needsReboot = $beforePending -or $afterPending -or $result.RebootFromInstall

  Write-Log ("Pre-reboot phase complete. FoundAny={0} EligibleFound={1} InstalledTotal={2} NeedsReboot={3}" -f `
    $result.FoundAny, $result.EligibleFound, $result.InstalledTotal, $needsReboot)

  if ($needsReboot) {
    Save-State @{
      CreatedAt                    = (Get-Date).ToString("s")
      IncludeRebootUpdates         = [bool]$IncludeRebootUpdates
      EnsureLatestCumulativeUpdate = [bool]$EnsureLatestCumulativeUpdate
      PostRebootMaxPasses          = [int]$PostRebootMaxPasses
    }

    Register-PostBootWorkerTask
    Launch-UI-IfRebootRequiredNow
    exit 1
  }

  try {
    if (Test-Path $StatePath) {
      Remove-Item $StatePath -Force -ErrorAction SilentlyContinue
    }
  } catch { }

  Remove-LogonPromptTaskIfNotNeeded
  Remove-Task $TaskImmediate

  Write-Log "Finished. No reboot required and no pending updates detected."
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}