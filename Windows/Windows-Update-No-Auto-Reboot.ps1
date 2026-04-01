#Requires -Version 5.1
<#
Windows Update + Consent-only Reboot + Auto Post-boot Update + Auto Re-prompt

Requirements satisfied:
- NEVER reboots automatically. Reboot only occurs when user clicks "Reboot now".
- "Postpone" schedules a REMINDER that shows the same UI again later (no reboot).
- UI auto-appears on next logon ONLY if a reboot is actually required (pending reboot signals).
- Post-boot worker re-runs Windows Update automatically after a user-initiated reboot, to keep applying updates.
- If another reboot becomes required after post-boot updates, ONLOGON prompt remains so user is asked again.
- Logs user clicks (USER_ACTION: ...), UI shown, ONLOGON decisions, and POSTBOOT actions.

Exit codes:
  0 = no reboot needed; updates installed or none available
  1 = reboot required/pending; prompt launched/attempted (no forced reboot)
  2 = report-only or updates found but none installed under rules
  3 = error
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [int]$CountdownMinutes = 10,
  [int[]]$PostponeOptionsMinutes = @(30, 60, 120),

  [switch]$IncludeRebootUpdates,
  [switch]$EnsureLatestCumulativeUpdate = $true,
  [switch]$ReportOnly,

  # Post-reboot automation
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

$StatePath       = Join-Path $BaseDir "WU-State.json"
$UiHelperPath    = Join-Path $BaseDir "RebootPromptUI.ps1"
$PostBootPath    = Join-Path $BaseDir "WU-PostBootWorker.ps1"
$LogonRunnerPath = Join-Path $BaseDir "WU-PromptOnLogon.ps1"

$TaskPostBoot    = "Nexus_WU_PostBootWorker"
$TaskReminder    = "Nexus_WU_RebootReminder"   # postpone reminder (re-prompt later)
$TaskOnLogon     = "Nexus_WU_PromptOnLogon"    # re-prompt at next logon if reboot required

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
    $sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
      -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($sess -and $sess.PendingFileRenameOperations) { return $true }
  } catch { }
  return $false
}

function Get-ActiveSessionPresent {
  try {
    $out = & quser 2>$null
    if ($out) { foreach ($l in $out) { if ($l -match '\sActive\s') { return $true } } }
  } catch { }
  return $false
}

function Save-State([hashtable]$obj) {
  ($obj | ConvertTo-Json -Depth 6) | Set-Content -Path $StatePath -Encoding UTF8 -Force
}
function Load-State {
  if (-not (Test-Path $StatePath)) { return $null }
  try { return (Get-Content $StatePath -Raw | ConvertFrom-Json) } catch { return $null }
}

function Remove-Task([string]$Name) {
  try { schtasks.exe /Delete /TN $Name /F *> $null } catch { }
}

# Keep this conservative: do NOT delete OnLogon prompt if a reboot is required.
function Remove-LogonPromptTaskIfNotNeeded {
  if (-not (Test-PendingReboot)) {
    Remove-Task $TaskOnLogon
    Write-Log "Removed ONLOGON prompt task (reboot not required)."
  }
}

# Abort any in-progress shutdown started by other tooling
function Abort-AnyShutdown {
  try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { }
}

# ---------------- Windows Update COM scan/install ----------------
function Get-AvailableUpdates {
  $session  = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
  $result   = $searcher.Search($criteria)

  $list = @()
  if ($result -and $result.Updates) {
    for ($i=0; $i -lt $result.Updates.Count; $i++) {
      $u = $result.Updates.Item($i)
      try { if ($u.EulaAccepted -eq $false) { $u.AcceptEula() } } catch { }

      $rb = 2
      try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

      $title = [string]$u.Title
      $isSSU = ($title -match 'Servicing Stack Update')
      $isLCU = ($title -match 'Cumulative Update') -and ($title -match 'Windows')

      $list += [pscustomobject]@{ Update=$u; Title=$title; RebootBehavior=$rb; IsSSU=$isSSU; IsLCU=$isLCU }
    }
  }
  return [pscustomobject]@{ Session=$session; Updates=$list }
}

function Install-Updates {
  param([Parameter(Mandatory)]$Session,[Parameter(Mandatory)][object[]]$UpdatesToInstall)

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
  param([int]$MaxPasses,[switch]$IncludeRebootUpdates,[switch]$EnsureLatestCumulativeUpdate)

  $pass = 0
  $rebootFromInstall = $false
  $installedTotal = 0
  $foundAny = $false

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

    $ssu = @($updates | Where-Object { $_.IsSSU })
    $lcu = @($updates | Where-Object { $_.IsLCU })
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

    $r = Install-Updates -Session $scan.Session -UpdatesToInstall $toInstall
    $installedTotal += $r.InstalledCount
    $rebootFromInstall = $rebootFromInstall -or $r.RebootRequired

    Write-Log ("WU pass {0}/{1}: Installed {2}. RebootRequired={3}" -f $pass, $MaxPasses, $r.InstalledCount, $r.RebootRequired)
    if ($rebootFromInstall) { break }
  }

  return [pscustomobject]@{
    FoundAny = $foundAny
    InstalledTotal = $installedTotal
    RebootFromInstall = $rebootFromInstall
  }
}

# ---------------- UI helper (consent-only; postpone = re-prompt later; reminder only shows if reboot required) ----------------
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
'@,
  [string]`$TaskReminder = '$TaskReminder'
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
  foreach (`$p in `$paths) { if (Test-Path `$p) { return `$true } }
  try {
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

function Schedule-Reminder([datetime]`$When) {
  try { Import-Module ScheduledTasks -ErrorAction Stop } catch { return `$false }
  try {
    try { Unregister-ScheduledTask -TaskName `$TaskReminder -Confirm:`$false -ErrorAction SilentlyContinue } catch { }

    `$psExe = "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    `$self = "`$PSCommandPath"
    `$args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"`$self`" -CountdownMinutes `$CountdownMinutes -PostponeCsv `"`$PostponeCsv`""
    `$action = New-ScheduledTaskAction -Execute `$psExe -Argument `$args
    `$trigger = New-ScheduledTaskTrigger -Once -At `$When
    `$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    `$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName `$TaskReminder -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force | Out-Null
    return `$true
  } catch { return `$false }
}

# PATCH: If this UI is running due to a reminder, ONLY show if reboot is required.
# (If not required, exit and remove reminder task.)
if (-not (Test-PendingReboot)) {
  Write-LogLocal "REMINDER/ONDEMAND: reboot not required; exiting and deleting reminder task if present."
  try { Unregister-ScheduledTask -TaskName `$TaskReminder -Confirm:`$false -ErrorAction SilentlyContinue } catch { }
  exit 0
}

`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) { foreach (`$x in (`$PostponeCsv -split ',')) { try { `$PostponeOptionsMinutes += [int]`$x } catch { } } }
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) { `$PostponeOptionsMinutes = @(30,60,120) }

Write-LogLocal "UI_SHOWN"

`$script:deadlineInfoOnly = (Get-Date).AddMinutes([double]`$CountdownMinutes)

`$form = New-Object System.Windows.Forms.Form
`$form.Text = `$UiTitle
`$form.Size = New-Object System.Drawing.Size(780, 340)
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true
`$form.Add_FormClosing({ if (`$_.CloseReason -eq 'UserClosing') { `$_.Cancel = `$true } })

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
  if (`$remain.TotalSeconds -le 0) { `$label2.Text = 'Countdown ended: please choose Reboot now or Postpone.'; return }
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
  else                { [void]`$combo.Items.Add(("Postpone {0} minutes" -f `$m)) }
}
`$combo.SelectedIndex = 0
`$form.Controls.Add(`$combo)

function Get-SelectedMinutes { if (`$combo.SelectedIndex -lt 0) { return `$PostponeOptionsMinutes[0] } return `$PostponeOptionsMinutes[`$combo.SelectedIndex] }

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
  & shutdown.exe /r /t 0 /c "`$Reason" | Out-Null
  `$form.Close()
})
`$form.Controls.Add(`$btnNow)

[void]`$form.ShowDialog()
"@
  Set-Content -Path $UiHelperPath -Value $content -Encoding UTF8 -Force
}

# ---------------- ONLOGON Runner (ONLY prompt if reboot required) ----------------
function Write-LogonRunner {
  $content = @"
param(
  [string]`$LogPath,
  [string]`$UiHelperPath,
  [string]`$TaskOnLogon
)

Set-StrictMode -Off
`$ErrorActionPreference = 'SilentlyContinue'

function Write-Log([string]`$Message) {
  try {
    `$ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path `$LogPath -Value ("[{0}] {1}" -f `$ts, `$Message) -Encoding UTF8
  } catch { }
}

function Test-PendingReboot {
  `$paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
  )
  foreach (`$p in `$paths) { if (Test-Path `$p) { return `$true } }
  try {
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

# PATCH: only show prompt if reboot is required
`$pending = Test-PendingReboot
if (-not `$pending) {
  Write-Log "ONLOGON: reboot not required; deleting logon prompt task."
  try { schtasks.exe /Delete /TN `$TaskOnLogon /F *> `$null } catch { }
  exit 0
}

Write-Log "ONLOGON: reboot IS required; launching UI."
`$psExe = "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
Start-Process -FilePath `$psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"`$UiHelperPath`"" -WindowStyle Hidden | Out-Null
"@
  Set-Content -Path $LogonRunnerPath -Value $content -Encoding UTF8 -Force
}

function Register-LogonPromptTask {
  Write-LogonRunner
  Remove-Task $TaskOnLogon

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $arg = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$LogonRunnerPath`" -LogPath `"$LogPath`" -UiHelperPath `"$UiHelperPath`" -TaskOnLogon `"$TaskOnLogon`""

  # /IT makes it interactive at logon; /RU SYSTEM ensures permissions
  $out = & schtasks.exe /Create /TN $TaskOnLogon /SC ONLOGON /RU SYSTEM /RL HIGHEST /IT /TR "`"$psExe`" $arg" /F 2>&1
  if ($LASTEXITCODE -ne 0) { throw "Failed to create ONLOGON prompt task: $($out -join ' ')" }
  Write-Log "Registered ONLOGON prompt task (will only show UI if reboot is required)."
}

# ---------------- Post-boot worker: rerun Windows Update until clean, and ensure ONLOGON prompt if reboot required ----------------
function Write-PostBootWorker {
  $content = @"
param(
  [string]`$LogPath,
  [int]`$MaxPasses = $PostRebootMaxPasses,
  [switch]`$IncludeRebootUpdates,
  [switch]`$EnsureLatestCumulativeUpdate,
  [string]`$TaskOnLogon,
  [string]`$UiHelperPath,
  [string]`$LogonRunnerPath
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
  foreach (`$p in `$paths) { if (Test-Path `$p) { return `$true } }
  try {
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

function Get-AvailableUpdates {
  `$session  = New-Object -ComObject Microsoft.Update.Session
  `$searcher = `$session.CreateUpdateSearcher()
  `$criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
  `$result   = `$searcher.Search(`$criteria)

  `$list = @()
  if (`$result -and `$result.Updates) {
    for (`$i=0; `$i -lt `$result.Updates.Count; `$i++) {
      `$u = `$result.Updates.Item(`$i)
      try { if (`$u.EulaAccepted -eq `$false) { `$u.AcceptEula() } } catch { }
      `$rb = 2
      try { `$rb = [int]`$u.InstallationBehavior.RebootBehavior } catch { `$rb = 2 }
      `$title = [string]`$u.Title
      `$isSSU = (`$title -match 'Servicing Stack Update')
      `$isLCU = (`$title -match 'Cumulative Update') -and (`$title -match 'Windows')
      `$list += [pscustomobject]@{ Update=`$u; Title=`$title; RebootBehavior=`$rb; IsSSU=`$isSSU; IsLCU=`$isLCU }
    }
  }
  return [pscustomobject]@{ Session=`$session; Updates=`$list }
}

function Install-Updates([object]`$Session,[object[]]`$UpdatesToInstall) {
  if (-not `$UpdatesToInstall -or `$UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount=0; RebootRequired=`$false }
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

  return [pscustomobject]@{ InstalledCount=[int]`$coll.Count; RebootRequired=[bool]`$res.RebootRequired }
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

  `$ssu = @(`$updates | Where-Object { `$_.IsSSU })
  `$lcu = @(`$updates | Where-Object { `$_.IsLCU })
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

# If reboot is required after postboot work, ensure ONLOGON prompt task exists (prompt-only when reboot required).
if (`$pending -or `$rebootFromInstall) {
  Write-Log "POSTBOOT: Reboot required. ONLOGON task should prompt user at logon."
} else {
  # Clean: remove ONLOGON prompt if exists
  try { schtasks.exe /Delete /TN `$TaskOnLogon /F *> `$null } catch { }
  Write-Log "POSTBOOT: System is clean. Removed ONLOGON prompt task."
}

# Self-delete postboot task (will be recreated when needed)
try { schtasks.exe /Delete /TN '$TaskPostBoot' /F *> `$null } catch { }
"@
  Set-Content -Path $PostBootPath -Value $content -Encoding UTF8 -Force
}

function Register-PostBootWorkerTask {
  Write-PostBootWorker
  Remove-Task $TaskPostBoot

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $arg  = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PostBootPath`" -LogPath `"$LogPath`" -MaxPasses $PostRebootMaxPasses -TaskOnLogon `"$TaskOnLogon`" -UiHelperPath `"$UiHelperPath`" -LogonRunnerPath `"$LogonRunnerPath`""
  if ($IncludeRebootUpdates) { $arg += " -IncludeRebootUpdates" }
  if ($EnsureLatestCumulativeUpdate) { $arg += " -EnsureLatestCumulativeUpdate" }

  $out = & schtasks.exe /Create /TN $TaskPostBoot /SC ONSTART /RU SYSTEM /RL HIGHEST /TR "`"$psExe`" $arg" /F 2>&1
  if ($LASTEXITCODE -ne 0) { throw "Failed to create post-boot worker task: $($out -join ' ')" }
  Write-Log "Registered post-boot Windows Update worker task (runs at startup)."
}

function Launch-UI-IfRebootRequiredNow {
  if (-not (Test-PendingReboot)) {
    Write-Log "UI not launched now: reboot not required."
    return
  }
  if (-not (Get-ActiveSessionPresent)) {
    Write-Log "UI not launched now: no active session. ONLOGON task will handle when user logs on."
    return
  }
  # Run UI helper in the currently active session (this script runs as SYSTEM; easiest is to run as scheduled onlogon too,
  # but for immediate display we just start powershell - it will still show in session if already interactive context.
  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  Start-Process -FilePath $psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$UiHelperPath`"" -WindowStyle Hidden | Out-Null
  Write-Log "Launched UI immediately (reboot required)."
}

# ---------------- MAIN ----------------
try {
  Write-Host "Windows Update (Consent-only + auto re-run + auto prompt on logon; prompt only if reboot required)"
  if (-not (Test-IsAdmin)) { throw "Run this script elevated (Administrator / SYSTEM)." }

  if ($ReportOnly) {
    Write-Log "ReportOnly=True. Exiting."
    exit 2
  }

  Abort-AnyShutdown

  # Ensure helper + logon runner exist
  Write-UiHelper
  Register-LogonPromptTask

  Write-Log "Pre-reboot phase: scanning/installing updates..."
  $beforePending = Test-PendingReboot
  if ($beforePending) { Write-Log "System already indicates a pending reboot." }

  $result = Run-WindowsUpdatePasses -MaxPasses 3 -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate

  $afterPending = Test-PendingReboot
  $needsReboot = $beforePending -or $afterPending -or $result.RebootFromInstall
  Write-Log ("Pre-reboot phase complete. FoundAny={0} InstalledTotal={1} NeedsReboot={2}" -f $result.FoundAny, $result.InstalledTotal, $needsReboot)

  if ($needsReboot) {
    # Register post-boot update worker so that after user-initiated reboot we continue applying updates.
    Save-State @{
      CreatedAt = (Get-Date).ToString("s")
      IncludeRebootUpdates = [bool]$IncludeRebootUpdates
      EnsureLatestCumulativeUpdate = [bool]$EnsureLatestCumulativeUpdate
      PostRebootMaxPasses = [int]$PostRebootMaxPasses
    }
    Register-PostBootWorkerTask

    # UI shows now only if reboot required (it is) and active session exists; otherwise ONLOGON runner handles.
    Launch-UI-IfRebootRequiredNow

    exit 1
  }

  # Clean: remove state and ONLOGON prompt task
  try { if (Test-Path $StatePath) { Remove-Item $StatePath -Force -ErrorAction SilentlyContinue } } catch { }
  Remove-LogonPromptTaskIfNotNeeded
  Write-Log "Finished. No reboot required and no pending updates detected."
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  exit 3
}