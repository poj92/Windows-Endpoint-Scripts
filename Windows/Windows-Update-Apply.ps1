#Requires -Version 5.1
<#
.SYNOPSIS
  Scans for available Windows updates, installs them, and prompts the user to reboot if needed.

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [int]$CountdownMinutes = 10,
  [int[]]$PostponeOptionsMinutes = @(30, 60, 120),
  [switch]$IncludeRebootUpdates,
  [switch]$ReportOnly
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Tls12ForPs5 {
  if ($PSVersionTable.PSVersion.Major -lt 6) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  }
}

function Write-Log {
  param([string]$Message)
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Write-Host ("[{0}] {1}" -f $ts, $Message)
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

function Send-UserMessage {
  param([string]$Text)
  try { & msg.exe * $Text | Out-Null } catch { }
}

function New-RebootPromptArtifacts {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes
  )

  $dir = Join-Path $env:ProgramData 'RebootPrompt'
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

  $helperPs1 = Join-Path $dir 'RebootPromptUI.ps1'
  $wrapperCmd = Join-Path $dir 'RunRebootPrompt.cmd'

  $optsCsv = ($PostponeOptionsMinutes | ForEach-Object { [int]$_ }) -join ','

  $ps1 = @"
param(
  [int]`$CountdownMinutes = $CountdownMinutes,
  [string]`$PostponeCsv = '$optsCsv',
  [string]`$Reason = 'Windows updates require a restart to finish installing.'
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) {
  foreach (`$x in (`$PostponeCsv -split ',')) {
    try { `$PostponeOptionsMinutes += [int]`$x } catch { }
  }
}
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) {
  `$PostponeOptionsMinutes = @(30,60,120)
}

function Schedule-Reboot([int]`$Seconds, [string]`$Msg) {
  try { & shutdown.exe /a | Out-Null } catch { }
  & shutdown.exe /r /t `$Seconds /c `$Msg | Out-Null
}

`$deadline = (Get-Date).AddMinutes([double]`$CountdownMinutes)
Schedule-Reboot -Seconds ([int](`$CountdownMinutes*60)) -Msg "`$Reason Computer will reboot in `$CountdownMinutes minute(s)."

`$form = New-Object System.Windows.Forms.Form
`$form.Text = 'Restart required'
`$form.Size = New-Object System.Drawing.Size(560, 230)
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true

`$form.Add_FormClosing({
  if (`$_.CloseReason -eq 'UserClosing') { `$_.Cancel = `$true }
})

`$label1 = New-Object System.Windows.Forms.Label
`$label1.AutoSize = `$true
`$label1.Location = New-Object System.Drawing.Point(18, 18)
`$label1.Font = New-Object System.Drawing.Font('Segoe UI', 10)
`$label1.Text = `$Reason
`$form.Controls.Add(`$label1)

`$label2 = New-Object System.Windows.Forms.Label
`$label2.AutoSize = `$true
`$label2.Location = New-Object System.Drawing.Point(18, 55)
`$label2.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
`$form.Controls.Add(`$label2)

function Update-Countdown {
  `$remain = `$deadline - (Get-Date)
  if (`$remain.TotalSeconds -le 0) {
    `$label2.Text = 'Rebooting now...'
    Schedule-Reboot -Seconds 0 -Msg "`$Reason Rebooting now."
    Start-Sleep -Seconds 1
    `$form.Close()
    return
  }
  `$mins = [int][Math]::Floor(`$remain.TotalMinutes)
  `$secs = [int]`$remain.Seconds
  `$label2.Text = ('Computer will reboot in {0:00}:{1:00}' -f `$mins, `$secs)
}

`$timer = New-Object System.Windows.Forms.Timer
`$timer.Interval = 1000
`$timer.Add_Tick({ Update-Countdown })
`$timer.Start()
Update-Countdown

function Set-Postpone([int]`$Minutes) {
  `$deadline = (Get-Date).AddMinutes([double]`$Minutes)
  Schedule-Reboot -Seconds ([int](`$Minutes*60)) -Msg "`$Reason Computer will reboot in `$Minutes minute(s)."
  Update-Countdown
}

`$btnNow = New-Object System.Windows.Forms.Button
`$btnNow.Text = 'Reboot now'
`$btnNow.Size = New-Object System.Drawing.Size(110, 32)
`$btnNow.Location = New-Object System.Drawing.Point(18, 130)
`$btnNow.Add_Click({
  `$deadline = Get-Date
  Schedule-Reboot -Seconds 0 -Msg "`$Reason Rebooting now."
  `$form.Close()
})
`$form.Controls.Add(`$btnNow)

`$x = 150
foreach (`$m in `$PostponeOptionsMinutes) {
  `$b = New-Object System.Windows.Forms.Button
  if (`$m -eq 30) { `$b.Text = 'Postpone 30m' }
  elseif (`$m -eq 60) { `$b.Text = 'Postpone 1h' }
  elseif (`$m -eq 120) { `$b.Text = 'Postpone 2h' }
  else { `$b.Text = ("Postpone {0}m" -f `$m) }

  `$b.Size = New-Object System.Drawing.Size(115, 32)
  `$b.Location = New-Object System.Drawing.Point(`$x, 130)
  `$b.Add_Click({ Set-Postpone -Minutes `$m })
  `$form.Controls.Add(`$b)
  `$x += 125
}

[void]`$form.ShowDialog()
"@

  Set-Content -Path $helperPs1 -Value $ps1 -Encoding UTF8 -Force

  # Wrapper CMD avoids schtasks /TR quoting issues
  $cmd = "@echo off`r`n" +
         "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$helperPs1`" -CountdownMinutes $CountdownMinutes -PostponeCsv `"$optsCsv`"`r`n"
  Set-Content -Path $wrapperCmd -Value $cmd -Encoding ASCII -Force

  return [pscustomobject]@{ HelperPs1=$helperPs1; WrapperCmd=$wrapperCmd }
}

function Start-RebootPrompt {
  param(
    [int]$CountdownMinutes,
    [int[]]$PostponeOptionsMinutes
  )

  $art = New-RebootPromptArtifacts -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes
  $hasSession = Get-ActiveSessionPresent

  if ($hasSession) {
    $tn = "RebootPrompt_" + ([guid]::NewGuid().ToString('N'))

    # schedule a minute or two in the future; NO /SD (avoids locale date issues)
    $start = (Get-Date).AddMinutes(2)
    $st = $start.ToString('HH:mm')

    Write-Log "Launching reboot prompt UI via scheduled task (interactive)."
    $createArgs = @(
      '/Create','/TN', $tn,
      '/TR', "`"$($art.WrapperCmd)`"",
      '/SC','ONCE',
      '/ST', $st,
      '/RL','HIGHEST',
      '/RU','SYSTEM',
      '/IT',
      '/F',
      '/Z'
    )

    $c = & schtasks.exe @createArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
      Write-Log "schtasks /Create failed. Falling back to msg.exe + shutdown timer."
      Write-Log ($c -join ' ')
      Send-UserMessage ("Windows updates require a restart. This computer will reboot in {0} minutes." -f $CountdownMinutes)
      & shutdown.exe /r /t ($CountdownMinutes*60) /c "Windows updates require a restart. This computer will reboot in $CountdownMinutes minutes." | Out-Null
      return
    }

    $r = & schtasks.exe /Run /TN $tn 2>&1
    if ($LASTEXITCODE -ne 0) {
      Write-Log "schtasks /Run failed. Falling back to msg.exe + shutdown timer."
      Write-Log ($r -join ' ')
      Send-UserMessage ("Windows updates require a restart. This computer will reboot in {0} minutes." -f $CountdownMinutes)
      & shutdown.exe /r /t ($CountdownMinutes*60) /c "Windows updates require a restart. This computer will reboot in $CountdownMinutes minutes." | Out-Null
      return
    }
  }
  else {
    Write-Log "No active user session detected. Scheduling reboot and notifying via msg.exe."
    Send-UserMessage ("Windows updates require a restart. This computer will reboot in {0} minutes." -f $CountdownMinutes)
    & shutdown.exe /r /t ($CountdownMinutes*60) /c "Windows updates require a restart. This computer will reboot in $CountdownMinutes minutes." | Out-Null
  }
}

function Get-AvailableUpdates {
  $session = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $criteria = "IsInstalled=0 and IsHidden=0"
  $result = $searcher.Search($criteria)

  $list = @()
  for ($i=0; $i -lt $result.Updates.Count; $i++) {
    $u = $result.Updates.Item($i)
    try { if ($u.EulaAccepted -eq $false) { $u.AcceptEula() } } catch { }

    $rb = 2
    try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

    $list += [pscustomobject]@{
      Update = $u
      Title  = [string]$u.Title
      RebootBehavior = $rb # 0 NeverReboots, 1 AlwaysRequiresReboot, 2 CanRequestReboot
    }
  }

  return [pscustomobject]@{ Session = $session; Updates = $list }
}

function Install-Updates {
  param(
    [Parameter(Mandatory)]$Session,
    [Parameter(Mandatory)][object[]]$UpdatesToInstall
  )

  if (-not $UpdatesToInstall -or $UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount = 0; RebootRequired = $false; ResultCode = 0 }
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
      Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes
      exit 1
    }
    exit 0
  }

  Write-Log ("Found {0} available update(s)." -f $updates.Count)

  $never = @($updates | Where-Object { $_.RebootBehavior -eq 0 })
  $mayReboot = @($updates | Where-Object { $_.RebootBehavior -ne 0 })

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
    Start-RebootPrompt -CountdownMinutes $CountdownMinutes -PostponeOptionsMinutes $PostponeOptionsMinutes
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