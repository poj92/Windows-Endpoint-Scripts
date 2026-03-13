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

  # Message heading/title requested
  [string]$UiTitle = "A security message from Nexus Open Systems Ltd",

  [string]$Reason = "Windows updates require a restart to finish installing."
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Write-Log([string]$Message) {
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Write-Host ("[{0}] {1}" -f $ts, $Message)
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
  foreach ($p in $paths) { if (Test-Path $p) { return $true } }

  try {
    $sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($sess -and $sess.PendingFileRenameOperations) { return $true }
  } catch { }

  return $false
}

function Get-ActiveSessionInfo {
  # Returns @{ User = 'DOMAIN\User'; SessionId = 1 } or $null
  try {
    $out = & quser 2>$null
    if (-not $out) { return $null }

    foreach ($line in $out) {
      # Typical:
      # >username         sessionname      1  Active      ...
      # username          console          1  Active      ...
      $l = $line.Trim()
      if ($l -match '^(>?)(\S+)\s+(\S+)\s+(\d+)\s+Active\s') {
        $user = $Matches[2]
        $sid  = [int]$Matches[4]
        # Try to expand to DOMAIN\User via Win32_ComputerSystem.UserName if available
        try {
          $csUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
          if ($csUser -and ($csUser -match "\\$user$")) {
            return @{ User = $csUser; SessionId = $sid }
          }
        } catch { }
        return @{ User = $user; SessionId = $sid }
      }
    }
  } catch { }

  # Fallback: explorer.exe owner in any session
  try {
    $explorers = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $explorers) {
      try {
        $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction SilentlyContinue
        if ($o -and $o.User) {
          $user = if ($o.Domain) { ($o.Domain + "\" + $o.User) } else { $o.User }
          return @{ User = $user; SessionId = [int]$p.SessionId }
        }
      } catch { }
    }
  } catch { }

  return $null
}

function Send-UserMessage([string]$Text) {
  try { & msg.exe * $Text | Out-Null } catch { }
}

function New-RebootPromptHelperScript {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes,
    [string]$UiTitle,
    [string]$Reason
  )

  $dir = Join-Path $env:ProgramData 'RebootPrompt'
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Schedule-Reboot([int]`$Seconds, [string]`$Msg) {
  try { & shutdown.exe /a | Out-Null } catch { }
  & shutdown.exe /r /t `$Seconds /c `$Msg | Out-Null
}

# Parse postpone options
`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) {
  foreach (`$x in (`$PostponeCsv -split ',')) { try { `$PostponeOptionsMinutes += [int]`$x } catch { } }
}
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) { `$PostponeOptionsMinutes = @(30,60,120) }

# Schedule initial reboot so it still happens if user ignores the UI
`$deadline = (Get-Date).AddMinutes([double]`$CountdownMinutes)
Schedule-Reboot -Seconds ([int](`$CountdownMinutes*60)) -Msg ("`$Reason This computer will reboot in `$CountdownMinutes minute(s).")

`$form = New-Object System.Windows.Forms.Form
`$form.Text = `$UiTitle
`$form.Size = New-Object System.Drawing.Size(640, 260)
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true

`$label1 = New-Object System.Windows.Forms.Label
`$label1.AutoSize = `$true
`$label1.MaximumSize = New-Object System.Drawing.Size(600, 0)
`$label1.Location = New-Object System.Drawing.Point(18, 18)
`$label1.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$label1.Text = `$Reason + ' This computer will reboot in $CountdownMinutes minutes.'
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
    Schedule-Reboot -Seconds 0 -Msg ("`$Reason Rebooting now.")
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

# Dropdown for postpone options
`$combo = New-Object System.Windows.Forms.ComboBox
`$combo.DropDownStyle = 'DropDownList'
`$combo.Location = New-Object System.Drawing.Point(18, 120)
`$combo.Size = New-Object System.Drawing.Size(220, 28)

foreach (`$m in `$PostponeOptionsMinutes) {
  if (`$m -eq 30)      { [void]`$combo.Items.Add('Postpone 30 minutes') }
  elseif (`$m -eq 60)  { [void]`$combo.Items.Add('Postpone 1 hour') }
  elseif (`$m -eq 120) { [void]`$combo.Items.Add('Postpone 2 hours') }
  else                { [void]`$combo.Items.Add(("Postpone {0} minutes" -f `$m)) }
}
`$combo.SelectedIndex = 0
`$form.Controls.Add(`$combo)

function Get-SelectedPostponeMinutes {
  `$idx = `$combo.SelectedIndex
  if (`$idx -lt 0) { return `$PostponeOptionsMinutes[0] }
  return `$PostponeOptionsMinutes[`$idx]
}

function Set-Postpone([int]`$Minutes) {
  `$deadline = (Get-Date).AddMinutes([double]`$Minutes)
  Schedule-Reboot -Seconds ([int](`$Minutes*60)) -Msg ("`$Reason This computer will reboot in `$Minutes minute(s).")
  Update-Countdown
}

`$btnPostpone = New-Object System.Windows.Forms.Button
`$btnPostpone.Text = 'Postpone'
`$btnPostpone.Size = New-Object System.Drawing.Size(120, 34)
`$btnPostpone.Location = New-Object System.Drawing.Point(255, 118)
`$btnPostpone.Add_Click({
  `$m = Get-SelectedPostponeMinutes
  Set-Postpone -Minutes `$m
})
`$form.Controls.Add(`$btnPostpone)

`$btnNow = New-Object System.Windows.Forms.Button
`$btnNow.Text = 'Reboot now'
`$btnNow.Size = New-Object System.Drawing.Size(120, 34)
`$btnNow.Location = New-Object System.Drawing.Point(390, 118)
`$btnNow.Add_Click({
  `$deadline = Get-Date
  Schedule-Reboot -Seconds 0 -Msg ("`$Reason Rebooting now.")
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

  # If already running in an interactive user context, show directly.
  if ([Environment]::UserInteractive -and -not $runningAsSystem) {
    Write-Log "Interactive user context detected: launching reboot prompt UI directly."
    Start-Process -FilePath "powershell.exe" -ArgumentList @(
      "-NoProfile","-ExecutionPolicy","Bypass","-File", "`"$helper`"",
      "-CountdownMinutes", "$CountdownMinutes"
    ) | Out-Null
    return
  }

  # SYSTEM context: launch in active user's session via InteractiveToken scheduled task
  $sess = Get-ActiveSessionInfo
  if ($sess -and $sess.User) {
    try { Import-Module ScheduledTasks -ErrorAction Stop } catch { }

    if (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue) {
      $taskName = "RebootPromptUI"
      $args = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File `"$helper`" -CountdownMinutes $CountdownMinutes"

      $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $args
      $trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(15)
      $principal = New-ScheduledTaskPrincipal -UserId $sess.User -LogonType InteractiveToken -RunLevel Highest
      $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

      Write-Log ("Launching reboot prompt UI in user session as {0} (InteractiveToken)." -f $sess.User)

      try {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName | Out-Null
        return
      } catch {
        Write-Log "InteractiveToken task launch failed; falling back to msg.exe + shutdown timer."
        Write-Log $_.Exception.Message
      }
    }
  }

  # Fallback: msg.exe (simple) + scheduled reboot timer
  Write-Log "Falling back to msg.exe notification + shutdown timer (no interactive UI available)."
  Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
  & shutdown.exe /r /t ($CountdownMinutes*60) /c ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.") | Out-Null
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

    $list += [pscustomobject]@{
      Update = $u
      Title  = [string]$u.Title
      RebootBehavior = $rb   # 0 NeverReboots, 1 AlwaysRequiresReboot, 2 CanRequestReboot
    }
  }

  return [pscustomobject]@{ Session=$session; Updates=$list }
}

function Install-Updates {
  param(
    [Parameter(Mandatory)]$Session,
    [Parameter(Mandatory)][object[]]$UpdatesToInstall
  )

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

  return [pscustomobject]@{
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
      if ($ReportOnly) { Write-Log "ReportOnly: Would prompt user to reboot."; exit 1 }
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
    Write-Log "ReportOnly: No updates will be installed."
    if ($mayReboot.Count -gt 0 -or $beforePendingReboot) { Write-Log "ReportOnly: Reboot prompt would be shown."; exit 1 }
    exit 0
  }

  $didInstall = $false
  $rebootFromInstall = $false

  if ($never.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("Windows Update", "Install non-reboot updates")) {
      $r1 = Install-Updates -Session $scan.Session -UpdatesToInstall $never
      $didInstall = ($r1.InstalledCount -gt 0)
      $rebootFromInstall = $rebootFromInstall -or $r1.RebootRequired
      Write-Log ("Installed {0} non-reboot update(s). RebootRequired={1}" -f $r1.InstalledCount, $r1.RebootRequired)
    }
  }

  if ($IncludeRebootUpdates -and $mayReboot.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("Windows Update", "Install reboot-requiring updates")) {
      $r2 = Install-Updates -Session $scan.Session -UpdatesToInstall $mayReboot
      $didInstall = $didInstall -or ($r2.InstalledCount -gt 0)
      $rebootFromInstall = $rebootFromInstall -or $r2.RebootRequired
      Write-Log ("Installed {0} reboot-requiring update(s). RebootRequired={1}" -f $r2.InstalledCount, $r2.RebootRequired)
    }
  }

  $afterPendingReboot = Test-PendingReboot
  $needsReboot = $beforePendingReboot -or $afterPendingReboot -or $rebootFromInstall -or ($mayReboot.Count -gt 0)

  if ($needsReboot) {
    Write-Log "Reboot is required (or reboot-requiring updates are pending). Prompting user with countdown/options."
    Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes -UiTitle $UiTitle -Reason $Reason
    exit 1
  }

  if (-not $didInstall -and $updates.Count -gt 0) {
    Write-Log "Updates were found but none were installed (likely only reboot-requiring updates)."
    exit 2
  }

  Write-Log "Finished. No reboot required."
  exit 0
}
catch {
  Write-Error $_
  exit 3
}