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

  # Ensure latest offered cumulative update is installed (SSU/LCU prioritized)
  [switch]$EnsureLatestCumulativeUpdate = $true,

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

function Clear-NexusRebootTasks {
  $tasks = @('Nexus_RebootPromptUI','Nexus_Reboot_Deadline','Nexus_Reboot_Deadline_UI')
  foreach ($t in ($tasks | Sort-Object -Unique)) {
    try { schtasks.exe /Delete /TN $t /F *> $null } catch { }
  }
  try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { }
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
    if ($out) { foreach ($l in $out) { if ($l -match '\sActive\s') { return $true } } }
  } catch { }
  return $false
}

function Send-UserMessage([string]$Text) {
  try { & msg.exe * $Text | Out-Null } catch { }
}

function Get-OsBuildInfo {
  try {
    $cv = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
    $ubr = $cv.UBR
    $build = $cv.CurrentBuildNumber
    $disp = $cv.DisplayVersion
    if (-not $disp) { $disp = $cv.ReleaseId }
    return "Product=$($cv.ProductName) Version=$disp Build=$build UBR=$ubr"
  } catch {
    return $null
  }
}

function Get-MostRecentInstalledUpdate {
  try {
    $hf = Get-HotFix -ErrorAction Stop |
      Where-Object { $_.HotFixID } |
      Sort-Object @{
        Expression = { try { [datetime]$_.InstalledOn } catch { [datetime]::MinValue } }
      } -Descending |
      Select-Object -First 1

    if ($hf) {
      $dt = $null
      try { $dt = [datetime]$hf.InstalledOn } catch { $dt = $null }
      $dtText = if ($dt) { $dt.ToString('yyyy-MM-dd') } else { "<unknown>" }
      return "HotFixID=$($hf.HotFixID) InstalledOn=$dtText Description=$($hf.Description)"
    }
  } catch { }
  return $null
}

# ---------- Reliable UI launch as SYSTEM in active session ----------
$script:LauncherLoaded = $false
function Ensure-SystemSessionLauncher {
  if ($script:LauncherLoaded) { return $true }
  try {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class NexusSessionLauncher
{
  private const uint TOKEN_ASSIGN_PRIMARY = 0x0001;
  private const uint TOKEN_DUPLICATE = 0x0002;
  private const uint TOKEN_QUERY = 0x0008;
  private const uint TOKEN_ADJUST_DEFAULT = 0x0080;
  private const uint TOKEN_ADJUST_SESSIONID = 0x0100;
  private const uint MAXIMUM_ALLOWED = 0x02000000;

  private const int SecurityImpersonation = 2;
  private const int TokenPrimary = 1;
  private const int TokenSessionId = 12;

  private const uint CREATE_UNICODE_ENVIRONMENT = 0x00000400;

  private enum WTS_CONNECTSTATE_CLASS
  {
    WTSActive, WTSConnected, WTSConnectQuery, WTSShadow, WTSDisconnected, WTSIdle,
    WTSListen, WTSReset, WTSDown, WTSInit
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct WTS_SESSION_INFO
  {
    public int SessionID;
    [MarshalAs(UnmanagedType.LPStr)]
    public string pWinStationName;
    public WTS_CONNECTSTATE_CLASS State;
  }

  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  private struct STARTUPINFO
  {
    public int cb;
    public string lpReserved;
    public string lpDesktop;
    public string lpTitle;
    public int dwX;
    public int dwY;
    public int dwXSize;
    public int dwYSize;
    public int dwXCountChars;
    public int dwYCountChars;
    public int dwFillAttribute;
    public int dwFlags;
    public short wShowWindow;
    public short cbReserved2;
    public IntPtr lpReserved2;
    public IntPtr hStdInput;
    public IntPtr hStdOutput;
    public IntPtr hStdError;
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct PROCESS_INFORMATION
  {
    public IntPtr hProcess;
    public IntPtr hThread;
    public int dwProcessId;
    public int dwThreadId;
  }

  [DllImport("kernel32.dll")]
  private static extern IntPtr GetCurrentProcess();

  [DllImport("advapi32.dll", SetLastError=true)]
  private static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);

  [DllImport("advapi32.dll", SetLastError=true)]
  private static extern bool DuplicateTokenEx(
    IntPtr hExistingToken,
    uint dwDesiredAccess,
    IntPtr lpTokenAttributes,
    int ImpersonationLevel,
    int TokenType,
    out IntPtr phNewToken);

  [DllImport("advapi32.dll", SetLastError=true)]
  private static extern bool SetTokenInformation(IntPtr TokenHandle, int TokenInformationClass, ref int TokenInformation, int TokenInformationLength);

  [DllImport("userenv.dll", SetLastError=true)]
  private static extern bool CreateEnvironmentBlock(out IntPtr lpEnvironment, IntPtr hToken, bool bInherit);

  [DllImport("userenv.dll", SetLastError=true)]
  private static extern bool DestroyEnvironmentBlock(IntPtr lpEnvironment);

  [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  private static extern bool CreateProcessAsUser(
    IntPtr hToken,
    string lpApplicationName,
    string lpCommandLine,
    IntPtr lpProcessAttributes,
    IntPtr lpThreadAttributes,
    bool bInheritHandles,
    uint dwCreationFlags,
    IntPtr lpEnvironment,
    string lpCurrentDirectory,
    ref STARTUPINFO lpStartupInfo,
    out PROCESS_INFORMATION lpProcessInformation);

  [DllImport("kernel32.dll", SetLastError=true)]
  private static extern bool CloseHandle(IntPtr hObject);

  [DllImport("Wtsapi32.dll", SetLastError=true)]
  private static extern bool WTSEnumerateSessions(IntPtr hServer, int Reserved, int Version, out IntPtr ppSessionInfo, out int pCount);

  [DllImport("Wtsapi32.dll")]
  private static extern void WTSFreeMemory(IntPtr pMemory);

  private static int GetActiveSessionId()
  {
    IntPtr pInfo;
    int count;
    if (WTSEnumerateSessions(IntPtr.Zero, 0, 1, out pInfo, out count))
    {
      int dataSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
      long current = (long)pInfo;
      for (int i=0; i<count; i++)
      {
        WTS_SESSION_INFO si = (WTS_SESSION_INFO)Marshal.PtrToStructure(new IntPtr(current), typeof(WTS_SESSION_INFO));
        if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive)
        {
          WTSFreeMemory(pInfo);
          return si.SessionID;
        }
        current += dataSize;
      }
      WTSFreeMemory(pInfo);
    }
    return -1;
  }

  public static bool StartAsSystemInActiveSession(string appPath, string fullCmdLine, out int lastError)
  {
    lastError = 0;
    int sessionId = GetActiveSessionId();
    if (sessionId < 0)
    {
      lastError = 0x57; // invalid parameter-ish
      return false;
    }

    IntPtr hToken;
    if (!OpenProcessToken(GetCurrentProcess(),
      TOKEN_ASSIGN_PRIMARY | TOKEN_DUPLICATE | TOKEN_QUERY | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID,
      out hToken))
    {
      lastError = Marshal.GetLastWin32Error();
      return false;
    }

    IntPtr hDup;
    bool ok = DuplicateTokenEx(hToken, MAXIMUM_ALLOWED, IntPtr.Zero, SecurityImpersonation, TokenPrimary, out hDup);
    CloseHandle(hToken);
    if (!ok)
    {
      lastError = Marshal.GetLastWin32Error();
      return false;
    }

    ok = SetTokenInformation(hDup, TokenSessionId, ref sessionId, sizeof(int));
    if (!ok)
    {
      lastError = Marshal.GetLastWin32Error();
      CloseHandle(hDup);
      return false;
    }

    IntPtr env = IntPtr.Zero;
    CreateEnvironmentBlock(out env, hDup, false);

    STARTUPINFO si = new STARTUPINFO();
    si.cb = Marshal.SizeOf(si);
    si.lpDesktop = "winsta0\\default";

    PROCESS_INFORMATION pi;
    ok = CreateProcessAsUser(
      hDup,
      appPath,
      fullCmdLine,
      IntPtr.Zero, IntPtr.Zero,
      false,
      CREATE_UNICODE_ENVIRONMENT,
      env,
      null,
      ref si,
      out pi
    );

    if (env != IntPtr.Zero) DestroyEnvironmentBlock(env);
    CloseHandle(hDup);

    if (!ok)
    {
      lastError = Marshal.GetLastWin32Error();
      return false;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return true;
  }
}
"@ -ErrorAction Stop | Out-Null

    $script:LauncherLoaded = $true
    return $true
  } catch {
    $script:LauncherLoaded = $false
    return $false
  }
}

# ---------- Helper UI script creation ----------
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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Abort-ShutdownSilently { try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { } }

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
  } catch { return `$false }
}

`$PostponeOptionsMinutes = @()
if (`$PostponeCsv) { foreach (`$x in (`$PostponeCsv -split ',')) { try { `$PostponeOptionsMinutes += [int]`$x } catch { } } }
if (-not `$PostponeOptionsMinutes -or `$PostponeOptionsMinutes.Count -eq 0) { `$PostponeOptionsMinutes = @(30,60,120) }

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

function Get-SelectedMinutes { if (`$combo.SelectedIndex -lt 0) { return `$PostponeOptionsMinutes[0] } return `$PostponeOptionsMinutes[`$combo.SelectedIndex] }

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

  if (-not (Get-ActiveSessionPresent)) {
    Write-Log "No active user session detected; showing msg.exe and scheduling Windows countdown reboot."
    Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
    & shutdown.exe /r /t ($CountdownMinutes*60) /c $Reason | Out-Null
    return
  }

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $cmdLine = "`"$psExe`" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$helper`" -CountdownMinutes $CountdownMinutes"

  # Primary: CreateProcessAsUser (SYSTEM token) into active session
  if (Ensure-SystemSessionLauncher) {
    $err = 0
    $ok = [NexusSessionLauncher]::StartAsSystemInActiveSession($psExe, $cmdLine, [ref]$err)
    if ($ok) {
      Write-Log "Reboot UI launched via SessionLauncher (SYSTEM in active session)."
      return
    } else {
      Write-Log ("WARNING: SessionLauncher failed. Win32Error={0}" -f $err)
    }
  } else {
    Write-Log "WARNING: SessionLauncher not available (Add-Type blocked)."
  }

  # Secondary: schtasks /IT fallback
  $wrapper = Join-Path (Split-Path -Parent $helper) 'RunRebootPrompt.cmd'
  $cmd = "@echo off`r`n$cmdLine`r`n"
  Set-Content -Path $wrapper -Value $cmd -Encoding ASCII -Force

  $taskName = "Nexus_RebootPromptUI"
  $sd = (Get-Date).ToString('MM/dd/yyyy')
  $st = (Get-Date).AddMinutes(1).ToString('HH:mm')

  Write-Log "Fallback: launching UI via schtasks /IT..."
  try { schtasks.exe /Delete /TN $taskName /F *> $null } catch { }

  $createOut = & schtasks.exe /Create /TN $taskName /TR "`"$wrapper`"" /SC ONCE /ST $st /SD $sd /RU SYSTEM /RL HIGHEST /IT /F 2>&1
  if ($LASTEXITCODE -eq 0) {
    $runOut = & schtasks.exe /Run /TN $taskName 2>&1
    if ($LASTEXITCODE -eq 0) {
      try { schtasks.exe /Delete /TN $taskName /F *> $null } catch { }
      Write-Log "Reboot UI launched via schtasks /IT."
      return
    }
    Write-Log ("WARNING: schtasks /Run failed: {0}" -f ($runOut -join ' '))
  } else {
    Write-Log ("WARNING: schtasks /Create failed: {0}" -f ($createOut -join ' '))
  }

  # Final fallback to prevent silent reboot: Windows built-in countdown
  Write-Log "FINAL FALLBACK: showing msg.exe + Windows shutdown countdown (ensures user sees notification)."
  Send-UserMessage ($Reason + " This computer will reboot in " + $CountdownMinutes + " minutes.")
  & shutdown.exe /r /t ($CountdownMinutes*60) /c $Reason | Out-Null
}

function Get-AvailableUpdates {
  try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
    $result   = $searcher.Search($criteria)

    if (-not $result -or -not $result.Updates) {
      return [pscustomobject]@{ Session=$session; Updates=@() }
    }

    $list = @()
    for ($i=0; $i -lt $result.Updates.Count; $i++) {
      $u = $result.Updates.Item($i)
      try { if ($u.EulaAccepted -eq $false) { $u.AcceptEula() } } catch { }

      $rb = 2
      try { $rb = [int]$u.InstallationBehavior.RebootBehavior } catch { $rb = 2 }

      $title = [string]$u.Title
      $isSSU = ($title -match 'Servicing Stack Update')
      $isLCU = ($title -match 'Cumulative Update') -and ($title -match 'Windows')

      $list += [pscustomobject]@{
        Update = $u
        Title  = $title
        RebootBehavior = $rb
        IsSSU = $isSSU
        IsLCU = $isLCU
      }
    }

    return [pscustomobject]@{ Session=$session; Updates=$list }
  }
  catch {
    Write-Log ("Windows Update COM scan failed: {0}" -f $_.Exception.Message)
    return [pscustomobject]@{ Session=$null; Updates=@() }
  }
}

function Install-Updates {
  param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)][object[]]$UpdatesToInstall)

  if (-not $Session) { throw "Windows Update session is null (scan failed earlier)." }
  if (-not $UpdatesToInstall -or $UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{ InstalledCount=0; RebootRequired=$false; Titles=@(); Kbs=@(); LcuKbs=@(); LcuKbsUnique=@() }
  }

  $titles = @($UpdatesToInstall | ForEach-Object { $_.Title })

  $kbs = @()
  foreach ($t in $titles) {
    $m = [regex]::Match($t, '(KB\d{6,8})')
    if ($m.Success) { $kbs += $m.Value }
  }
  $kbs = $kbs | Sort-Object -Unique

  $lcuTitles = @(
    $UpdatesToInstall | Where-Object {
      ($_.PSObject.Properties.Name -contains 'IsLCU' -and $_.IsLCU) -or
      (($_.Title -match 'Cumulative Update') -and ($_.Title -match 'Windows'))
    } | ForEach-Object { $_.Title }
  )

  $lcuKbs = @()
  foreach ($t in $lcuTitles) {
    $m = [regex]::Match($t, '(KB\d{6,8})')
    if ($m.Success) { $lcuKbs += $m.Value }
  }
  $lcuKbsOrdered = $lcuKbs
  $lcuKbsUnique  = $lcuKbs | Sort-Object -Unique

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
    Titles         = $titles
    Kbs            = $kbs
    LcuKbs         = $lcuKbsOrdered
    LcuKbsUnique   = $lcuKbsUnique
  }
}

# ---------------- main ----------------
try {
  Write-Host "Windows-Update-Force-Apply"
  if (-not (Test-IsAdmin)) { throw "Run this script elevated (Administrator / SYSTEM)." }

  Clear-NexusRebootTasks

  $os = Get-OsBuildInfo
  if ($os) { Write-Log ("OS build (pre-run): {0}" -f $os) }

  $last = Get-MostRecentInstalledUpdate
  if ($last) { Write-Log ("Most recent installed update (pre-run): {0}" -f $last) }

  Write-Log "Scanning for Windows updates..."
  $beforePendingReboot = Test-PendingReboot
  if ($beforePendingReboot) { Write-Log "System already indicates a pending reboot." }

  $maxPasses = 3
  $pass = 0
  $anyInstalledKbs = @()
  $installedLcuKbsOrdered = @()
  $rebootFromInstall = $false

  while ($pass -lt $maxPasses) {
    $pass++
    $scan = Get-AvailableUpdates
    $updates = $scan.Updates

    if (-not $updates -or $updates.Count -eq 0) {
      Write-Log "No available updates found."
      break
    }

    Write-Log ("Pass {0}/{1}: Found {2} available update(s)." -f $pass, $maxPasses, $updates.Count)

    $ssu = @($updates | Where-Object { $_.IsSSU })
    $lcu = @($updates | Where-Object { $_.IsLCU })

    $toInstall = @()

    if ($EnsureLatestCumulativeUpdate) {
      $toInstall += $ssu
      $toInstall += $lcu

      $others = @($updates | Where-Object { -not $_.IsSSU -and -not $_.IsLCU })
      if ($IncludeRebootUpdates) {
        $toInstall += $others
      } else {
        $toInstall += @($others | Where-Object { $_.RebootBehavior -eq 0 })
      }
    } else {
      if ($IncludeRebootUpdates) {
        $toInstall = $updates
      } else {
        $toInstall = @($updates | Where-Object { $_.RebootBehavior -eq 0 })
      }
    }

    $toInstall = @($toInstall | Select-Object -Unique)
    if (-not $toInstall -or $toInstall.Count -eq 0) {
      Write-Log "No eligible updates to install in this pass."
      break
    }

    if ($ReportOnly) {
      Write-Log ("ReportOnly: would install {0} update(s) this pass." -f $toInstall.Count)
      exit 2
    }

    $r = Install-Updates -Session $scan.Session -UpdatesToInstall $toInstall
    $rebootFromInstall = $rebootFromInstall -or $r.RebootRequired

    if ($r.Kbs) { $anyInstalledKbs += $r.Kbs }
    if ($r.LcuKbs -and $r.LcuKbs.Count -gt 0) { $installedLcuKbsOrdered += $r.LcuKbs }

    Write-Log ("Installed {0} update(s) this pass. RebootRequired={1}" -f $r.InstalledCount, $r.RebootRequired)
    if ($r.Kbs.Count -gt 0) { Write-Log ("KBs installed (this pass): {0}" -f ($r.Kbs -join ', ')) }
    if ($r.LcuKbsUnique -and $r.LcuKbsUnique.Count -gt 0) { Write-Log ("LCU KBs installed (this pass): {0}" -f ($r.LcuKbsUnique -join ', ')) }

    if ($rebootFromInstall) { break }
  }

  $afterPendingReboot = Test-PendingReboot
  $needsReboot = $beforePendingReboot -or $afterPendingReboot -or $rebootFromInstall

  $anyInstalledKbs = $anyInstalledKbs | Sort-Object -Unique
  if ($anyInstalledKbs.Count -gt 0) {
    Write-Log ("KBs installed this run (all passes): {0}" -f ($anyInstalledKbs -join ', '))
  }

  if ($installedLcuKbsOrdered.Count -gt 0) {
    $latestLcu = $installedLcuKbsOrdered[-1]
    Write-Log ("Latest LCU installed this run: {0}" -f $latestLcu)
    $allLcuUnique = $installedLcuKbsOrdered | Sort-Object -Unique
    Write-Log ("LCU KBs installed this run (all): {0}" -f ($allLcuUnique -join ', '))
  } else {
    Write-Log "Latest LCU installed this run: <none>"
  }

  $os2 = Get-OsBuildInfo
  if ($os2) { Write-Log ("OS build (post-run): {0}" -f $os2) }

  $last2 = Get-MostRecentInstalledUpdate
  if ($last2) { Write-Log ("Most recent installed update (post-run): {0}" -f $last2) }

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