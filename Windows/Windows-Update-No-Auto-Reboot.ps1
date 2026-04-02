#Requires -Version 5.1
<#
Windows Update + Consent-only Reboot + Auto Post-boot Update + Auto Re-prompt

Behavior:
- NEVER reboots automatically. Reboot only occurs when user clicks "Reboot now".
- "Postpone" schedules a reminder that shows the same UI later (no reboot).
- UI appears at user logon ONLY if reboot is actually required.
- Post-boot worker re-runs Windows Update automatically after a user-initiated reboot.
- If another reboot becomes required after post-boot updates, the user is prompted again at logon.
- Uses the same proven notification path as the working script:
  1) If already running as the interactive user, launch UI directly
  2) If running as SYSTEM, try CreateProcessAsUser into the active session
  3) Fallback to schtasks /IT
  4) Final fallback is msg.exe only (never reboots)

Exit codes:
  0 = no reboot needed; updates installed or none available
  1 = reboot required/pending; prompt launched or scheduled (no forced reboot)
  2 = report-only OR updates found but none were installed under rules
  3 = error
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [ValidateSet('Normal','PostBoot','PromptOnly')]
  [string]$Mode = 'Normal',

  [int]$CountdownMinutes = 10,
  [int[]]$PostponeOptionsMinutes = @(30, 60, 120),
  [string]$PostponeOptionsCsv,

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

if ($PostponeOptionsCsv) {
  $parsed = @()
  foreach ($x in ($PostponeOptionsCsv -split ',')) {
    try { $parsed += [int]$x } catch { }
  }
  if ($parsed.Count -gt 0) { $PostponeOptionsMinutes = $parsed }
}

New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $LogPath) -Force | Out-Null

$StatePath           = Join-Path $BaseDir "WU-State.json"
$UiHelperPath        = Join-Path $BaseDir "RebootPromptUI.ps1"
$UiWrapperPath       = Join-Path $BaseDir "RunRebootPrompt.cmd"
$ReminderWrapperPath = Join-Path $BaseDir "RunRebootReminder.cmd"
$MainScriptCopyPath  = Join-Path $BaseDir "Windows-Update-No-Auto-Reboot.ps1"

$TaskPostBoot        = "Nexus_WU_PostBootWorker"
$TaskReminder        = "Nexus_WU_RebootReminder"
$TaskOnLogon         = "Nexus_WU_PromptOnLogon"
$TaskImmediate       = "Nexus_WU_PromptNow"

function Write-Log {
  param(
    [string]$Message,
    [ValidateSet('INFO','WARN','ERROR','DEBUG')]
    [string]$Level = 'INFO'
  )

  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $line = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
  Write-Host $line
  try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Write-ExceptionLog {
  param(
    [string]$Prefix,
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  if ($null -eq $ErrorRecord) {
    Write-Log ("{0}: <no error record supplied>" -f $Prefix) 'ERROR'
    return
  }

  $msg  = $ErrorRecord.Exception.Message
  $type = $ErrorRecord.Exception.GetType().FullName
  $at   = $ErrorRecord.InvocationInfo.PositionMessage

  Write-Log ("{0}: ExceptionType={1}; Message={2}" -f $Prefix, $type, $msg) 'ERROR'
  if ($at) {
    Write-Log ("{0}: Location={1}" -f $Prefix, ($at -replace "`r|`n",' ')) 'DEBUG'
  }
}

function Ensure-BaseDirPermissions {
  try {
    $out = & icacls.exe $BaseDir /grant "*S-1-5-32-545:(OI)(CI)M" /T /C 2>&1
    Write-Log ("ACL update output: {0}" -f (($out | ForEach-Object { $_.ToString().Trim() }) -join ' | ')) 'DEBUG'
  } catch {
    Write-ExceptionLog -Prefix 'Ensure-BaseDirPermissions failed' -ErrorRecord $_
  }
}

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-CurrentIdentityName {
  try {
    return [Security.Principal.WindowsIdentity]::GetCurrent().Name
  } catch {
    if ($env:USERDOMAIN -and $env:USERNAME) {
      return ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    }
    return $env:USERNAME
  }
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
    return ("Product={0} Version={1} Build={2} UBR={3}" -f $cv.ProductName, $disp, $build, $ubr)
  } catch {
    return $null
  }
}

function Convert-ToSafeDateTime {
  param([object]$Value)

  if ($null -eq $Value) { return $null }
  if ($Value -is [datetime]) { return $Value }

  $s = [string]$Value
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }

  $parsed = $null
  $styles = [Globalization.DateTimeStyles]::AllowWhiteSpaces
  $formats = @(
    'M/d/yyyy',
    'MM/dd/yyyy',
    'd/M/yyyy',
    'dd/MM/yyyy',
    'yyyy-MM-dd',
    'yyyy/MM/dd',
    'ddd MMM dd HH:mm:ss yyyy'
  )

  foreach ($fmt in $formats) {
    if ([datetime]::TryParseExact($s, $fmt, $null, $styles, [ref]$parsed)) {
      return $parsed
    }
  }

  if ([datetime]::TryParse($s, [ref]$parsed)) {
    return $parsed
  }

  return $null
}

function Get-LatestInstalledUpdate {
  try {
    $hotfixes = Get-CimInstance Win32_QuickFixEngineering -ErrorAction Stop | ForEach-Object {
      $dt = Convert-ToSafeDateTime $_.InstalledOn
      [pscustomobject]@{
        HotFixID     = $_.HotFixID
        Description  = $_.Description
        InstalledOn  = $dt
        RawInstalled = $_.InstalledOn
      }
    }

    $latest = $hotfixes |
      Where-Object { $_.InstalledOn } |
      Sort-Object InstalledOn -Descending |
      Select-Object -First 1

    if ($latest) { return $latest }
  } catch {
    Write-Log ("Failed to query latest installed update: {0}" -f $_.Exception.Message) 'WARN'
  }

  return $null
}

function Write-LatestInstalledUpdateLog([string]$Prefix = 'LATEST_INSTALLED_UPDATE') {
  $latest = Get-LatestInstalledUpdate
  if ($latest) {
    Write-Log ("{0}: HotFixID={1}; Description={2}; InstalledOn={3}; RawInstalled='{4}'" -f `
      $Prefix,
      $latest.HotFixID,
      $latest.Description,
      $latest.InstalledOn.ToString('yyyy-MM-dd'),
      $latest.RawInstalled)
  } else {
    Write-Log ("{0}: not found" -f $Prefix) 'WARN'
  }
}

function Save-State([hashtable]$obj) {
  try {
    ($obj | ConvertTo-Json -Depth 6) | Set-Content -Path $StatePath -Encoding UTF8 -Force
  } catch { }
}

function Remove-Task([string]$Name) {
  try { schtasks.exe /Delete /TN $Name /F *> $null } catch { }
}

function Remove-UiTasks {
  foreach ($t in @($TaskReminder, $TaskImmediate)) {
    Remove-Task $t
  }
}

function Remove-LogonPromptTaskIfNotNeeded {
  if (-not (Test-PendingReboot)) {
    Remove-Task $TaskOnLogon
    Write-Log "Removed ONLOGON prompt task (reboot not required)."
  }
}

function Abort-AnyShutdown {
  try { & cmd.exe /c "shutdown.exe /a >nul 2>&1" | Out-Null } catch { }
}

function Write-TaskDiagnostics {
  param(
    [Parameter(Mandatory)][string]$TaskName,
    [string]$Prefix = 'TASK'
  )

  try { Import-Module ScheduledTasks -ErrorAction Stop } catch {
    Write-Log ("{0}: ScheduledTasks module unavailable; skipping diagnostics for '{1}'." -f $Prefix, $TaskName) 'WARN'
    return
  }

  try {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
  } catch {
    Write-Log ("{0}: Task '{1}' not found." -f $Prefix, $TaskName) 'WARN'
    return
  }

  $info = $null
  try { $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction Stop } catch { }

  $actionText = @()
  foreach ($a in @($task.Actions)) {
    $actionText += ("Execute='{0}' Arguments='{1}'" -f $a.Execute, $a.Arguments)
  }

  $triggerText = @()
  foreach ($t in @($task.Triggers)) {
    $triggerType = $null
    try { $triggerType = $t.CimClass.CimClassName } catch { }
    $triggerText += ("Type='{0}' StartBoundary='{1}' EndBoundary='{2}' Enabled='{3}'" -f `
      $triggerType, $t.StartBoundary, $t.EndBoundary, $t.Enabled)
  }

  Write-Log ("{0}: Name='{1}'; State='{2}'; PrincipalUserId='{3}'; PrincipalGroupId='{4}'; LogonType='{5}'; RunLevel='{6}'" -f `
    $Prefix, $TaskName, $task.State, $task.Principal.UserId, $task.Principal.GroupId, $task.Principal.LogonType, $task.Principal.RunLevel) 'DEBUG'

  if ($info) {
    Write-Log ("{0}: LastRunTime='{1}'; LastTaskResult='{2}'; NextRunTime='{3}'; NumberOfMissedRuns='{4}'" -f `
      $Prefix, $info.LastRunTime, $info.LastTaskResult, $info.NextRunTime, $info.NumberOfMissedRuns) 'DEBUG'
  }

  if ($actionText.Count -gt 0) {
    Write-Log ("{0}: Actions={1}" -f $Prefix, ($actionText -join ' | ')) 'DEBUG'
  }

  if ($triggerText.Count -gt 0) {
    Write-Log ("{0}: Triggers={1}" -f $Prefix, ($triggerText -join ' | ')) 'DEBUG'
  }
}

function Ensure-MainScriptCopy {
  try {
    if (-not $PSCommandPath) { return $MainScriptCopyPath }
    if ($PSCommandPath -ne $MainScriptCopyPath) {
      Copy-Item -Path $PSCommandPath -Destination $MainScriptCopyPath -Force
      Write-Log ("Copied main script to '{0}'." -f $MainScriptCopyPath) 'DEBUG'
    }
  } catch {
    Write-ExceptionLog -Prefix 'Ensure-MainScriptCopy failed' -ErrorRecord $_
  }
  return $MainScriptCopyPath
}

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
      lastError = 0x57;
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
    Write-ExceptionLog -Prefix 'Ensure-SystemSessionLauncher failed' -ErrorRecord $_
    return $false
  }
}

function Write-UiHelperScript {
  $optsCsv = ($PostponeOptionsMinutes | ForEach-Object { [int]$_ }) -join ','
  $mainScriptPathForReminder = $MainScriptCopyPath
  $reminderWrapperPathForHelper = $ReminderWrapperPath
  $baseDirForHelper = $BaseDir

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
  [string]`$TaskReminder = '$TaskReminder',
  [string]`$MainScriptPath = @'
$mainScriptPathForReminder
'@,
  [string]`$ReminderWrapperPath = @'
$reminderWrapperPathForHelper
'@,
  [string]`$BaseDir = @'
$baseDirForHelper
'@
)

Set-StrictMode -Off
`$ErrorActionPreference = 'Stop'

function Write-LogLocal {
  param(
    [string]`$Message,
    [ValidateSet('INFO','WARN','ERROR','DEBUG')]
    [string]`$Level = 'INFO'
  )
  try {
    `$ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path `$LogPath -Value ("[{0}] [{1}] {2}" -f `$ts, `$Level, `$Message) -Encoding UTF8
  } catch { }
}

function Write-ExceptionLocal {
  param(
    [string]`$Prefix,
    [System.Management.Automation.ErrorRecord]`$ErrorRecord
  )

  if (`$null -eq `$ErrorRecord) {
    Write-LogLocal ("{0}: <no error record supplied>" -f `$Prefix) 'ERROR'
    return
  }

  Write-LogLocal ("{0}: ExceptionType={1}; Message={2}" -f `
    `$Prefix,
    `$ErrorRecord.Exception.GetType().FullName,
    `$ErrorRecord.Exception.Message) 'ERROR'

  if (`$ErrorRecord.InvocationInfo -and `$ErrorRecord.InvocationInfo.PositionMessage) {
    Write-LogLocal ("{0}: Location={1}" -f `$Prefix, (`$ErrorRecord.InvocationInfo.PositionMessage -replace "`r|`n",' ')) 'DEBUG'
  }
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
    `$sess = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
      -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if (`$sess -and `$sess.PendingFileRenameOperations) { return `$true }
  } catch { }
  return `$false
}

function Remove-ReminderTask {
  try { Unregister-ScheduledTask -TaskName `$TaskReminder -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null } catch { }
  try { schtasks.exe /Delete /TN `$TaskReminder /F *> `$null } catch { }
}

function Write-ReminderTaskDiagnostics {
  param([string]`$TaskName)

  try {
    Import-Module ScheduledTasks -ErrorAction Stop
    `$task = Get-ScheduledTask -TaskName `$TaskName -ErrorAction Stop
    `$info = Get-ScheduledTaskInfo -TaskName `$TaskName -ErrorAction Stop

    `$actionText = @()
    foreach (`$a in @(`$task.Actions)) {
      `$actionText += ("Execute='{0}' Arguments='{1}'" -f `$a.Execute, `$a.Arguments)
    }

    Write-LogLocal ("UI_REMINDER_TASK: Name='{0}'; State='{1}'; PrincipalUserId='{2}'; LogonType='{3}'; RunLevel='{4}'" -f `
      `$TaskName, `$task.State, `$task.Principal.UserId, `$task.Principal.LogonType, `$task.Principal.RunLevel) 'DEBUG'
    Write-LogLocal ("UI_REMINDER_TASK: LastRunTime='{0}'; LastTaskResult='{1}'; NextRunTime='{2}'" -f `
      `$info.LastRunTime, `$info.LastTaskResult, `$info.NextRunTime) 'DEBUG'

    if (`$actionText.Count -gt 0) {
      Write-LogLocal ("UI_REMINDER_TASK: Actions={0}" -f (`$actionText -join ' | ')) 'DEBUG'
    }
  } catch {
    Write-ExceptionLocal -Prefix 'UI_REMINDER_TASK_DIAGNOSTICS_ERROR' -ErrorRecord `$_
  }
}

function Schedule-Reminder([datetime]`$When) {
  try {
    Import-Module ScheduledTasks -ErrorAction Stop

    `$psExe = "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    `$cmdLine = ('"{0}" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}" -Mode PromptOnly -CountdownMinutes {2} -PostponeOptionsCsv "{3}" -BaseDir "{4}" -LogPath "{5}"' -f `
      `$psExe, `$MainScriptPath, `$CountdownMinutes, `$PostponeCsv, `$BaseDir, `$LogPath)

    New-Item -ItemType Directory -Path `$BaseDir -Force | Out-Null
    Set-Content -Path `$ReminderWrapperPath -Value ("@echo off`r`n" + `$cmdLine + "`r`n") -Encoding ASCII -Force
    Write-LogLocal ("UI: reminder wrapper written to '{0}' with command '{1}'" -f `$ReminderWrapperPath, `$cmdLine) 'DEBUG'

    Remove-ReminderTask

    `$action = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\cmd.exe" -Argument "/c `"`$ReminderWrapperPath`""
    `$trigger = New-ScheduledTaskTrigger -Once -At `$When
    `$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    `$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    Register-ScheduledTask -TaskName `$TaskReminder -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force | Out-Null

    Write-LogLocal ("UI: reminder task registered for {0}" -f `$When.ToString('yyyy-MM-dd HH:mm:ss')) 'INFO'
    Write-ReminderTaskDiagnostics -TaskName `$TaskReminder
    return `$true
  } catch {
    Write-ExceptionLocal -Prefix 'UI_REMINDER_SCHEDULE_ERROR' -ErrorRecord `$_
    return `$false
  }
}

try {
  Write-LogLocal ("UI_START: Identity='{0}'; UserInteractive='{1}'; SessionId='{2}'; PID='{3}'; Machine='{4}'" -f `
    ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
    [Environment]::UserInteractive,
    [System.Diagnostics.Process]::GetCurrentProcess().SessionId,
    `$PID,
    `$env:COMPUTERNAME)

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  Write-LogLocal "UI_ADDTYPE_OK"

  if (-not (Test-PendingReboot)) {
    Write-LogLocal "UI: reboot not required; exiting."
    Remove-ReminderTask
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

  Write-LogLocal ("UI_SHOWN: CountdownMinutes='{0}'; Options='{1}'" -f `$CountdownMinutes, (`$PostponeOptionsMinutes -join ','))

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
      Write-LogLocal ("USER_ACTION: Postpone FAILED ({0} minutes)" -f `$mins) 'ERROR'
      `$label3.Text = 'Postpone failed (could not schedule reminder). Please try again or choose Reboot now.'
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
    try { & shutdown.exe /r /t 0 /c "`$Reason" | Out-Null } catch {
      Write-ExceptionLocal -Prefix 'UI_SHUTDOWN_ERROR' -ErrorRecord `$_
    }
    `$form.Close()
  })
  `$form.Controls.Add(`$btnNow)

  [void]`$form.ShowDialog()
}
catch {
  Write-ExceptionLocal -Prefix 'UI_HELPER_ERROR' -ErrorRecord `$_
  exit 1
}
"@

  Set-Content -Path $UiHelperPath -Value $content -Encoding UTF8 -Force
  Write-Log ("Wrote UI helper script to '{0}'." -f $UiHelperPath) 'DEBUG'
}

function Get-UiCommandLine {
  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  return ('"{0}" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{1}" -CountdownMinutes {2}' -f `
    $psExe, $UiHelperPath, $CountdownMinutes)
}

function Write-UiWrapperFile {
  $cmdLine = Get-UiCommandLine
  Set-Content -Path $UiWrapperPath -Value ("@echo off`r`n" + $cmdLine + "`r`n") -Encoding ASCII -Force
  Write-Log ("Wrote UI wrapper file to '{0}'." -f $UiWrapperPath) 'DEBUG'
}

function Launch-UiInActiveSessionNow {
  if (-not (Test-PendingReboot)) {
    Write-Log "Immediate UI not launched: reboot not required."
    return
  }

  if (-not (Get-ActiveSessionPresent)) {
    Write-Log "Immediate UI not launched: no active user session. ONLOGON task will handle later." 'WARN'
    return
  }

  $currentIdentity = Get-CurrentIdentityName
  $isSystem = $currentIdentity -eq 'NT AUTHORITY\SYSTEM'
  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $cmdLine = Get-UiCommandLine

  Write-Log ("Immediate UI launch attempt. Identity='{0}'; CommandLine={1}" -f $currentIdentity, $cmdLine) 'DEBUG'

  if (-not $isSystem) {
    try {
      Start-Process -FilePath $psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$UiHelperPath`" -CountdownMinutes $CountdownMinutes" -WindowStyle Hidden
      Write-Log "Immediate UI launched directly because script is already running in user context."
      return
    } catch {
      Write-ExceptionLog -Prefix 'Immediate UI direct launch failed' -ErrorRecord $_
    }
  }

  if ($isSystem -and (Ensure-SystemSessionLauncher)) {
    try {
      $err = 0
      $ok = [NexusSessionLauncher]::StartAsSystemInActiveSession($psExe, $cmdLine, [ref]$err)
      if ($ok) {
        Write-Log "Immediate UI launched via SessionLauncher."
        return
      } else {
        Write-Log ("Immediate UI SessionLauncher failed. Win32Error={0}" -f $err) 'WARN'
      }
    } catch {
      Write-ExceptionLog -Prefix 'Immediate UI SessionLauncher failed' -ErrorRecord $_
    }
  }

  try {
    Write-UiWrapperFile
    $start = (Get-Date).AddMinutes(3)
    $sd = $start.ToString('MM/dd/yyyy')
    $st = $start.ToString('HH:mm')

    Remove-Task $TaskImmediate
    $createOut = & schtasks.exe /Create /TN $TaskImmediate /TR "`"$UiWrapperPath`"" /SC ONCE /SD $sd /ST $st /RU $currentIdentity /RL LIMITED /IT /F 2>&1
    Write-Log ("IMMEDIATE schtasks /Create output: {0}" -f (($createOut | ForEach-Object { $_.ToString().Trim() }) -join ' | ')) 'DEBUG'

    if ($LASTEXITCODE -eq 0) {
      Write-TaskDiagnostics -TaskName $TaskImmediate -Prefix 'IMMEDIATE_TASK_BEFORE_RUN'
      $runOut = & schtasks.exe /Run /TN $TaskImmediate 2>&1
      Write-Log ("IMMEDIATE schtasks /Run output: {0}" -f (($runOut | ForEach-Object { $_.ToString().Trim() }) -join ' | ')) 'DEBUG'

      if ($LASTEXITCODE -eq 0) {
        Write-Log ("Attempted immediate UI launch via schtasks.exe /Run as '{0}'." -f $currentIdentity)
        Start-Sleep -Seconds 2
        Write-TaskDiagnostics -TaskName $TaskImmediate -Prefix 'IMMEDIATE_TASK_AFTER_RUN'
        return
      }
    }
  } catch {
    Write-ExceptionLog -Prefix 'Immediate UI schtasks fallback failed' -ErrorRecord $_
  }

  Send-UserMessage "Windows updates require a restart. Please save your work and reboot when ready."
  Write-Log "Immediate UI final fallback was msg.exe only (no reboot)." 'WARN'
}

function Register-LogonPromptTask {
  Remove-Task $TaskOnLogon
  Write-UiWrapperFile

  $currentIdentity = Get-CurrentIdentityName
  $isSystem = $currentIdentity -eq 'NT AUTHORITY\SYSTEM'

  if ($isSystem) {
    $out = & schtasks.exe /Create /TN $TaskOnLogon /SC ONLOGON /RU SYSTEM /RL HIGHEST /IT /TR "`"$UiWrapperPath`"" /F 2>&1
    Write-Log ("ONLOGON schtasks /Create output: {0}" -f (($out | ForEach-Object { $_.ToString().Trim() }) -join ' | ')) 'DEBUG'
    if ($LASTEXITCODE -ne 0) {
      throw ("Failed to create ONLOGON prompt task as SYSTEM: {0}" -f ($out -join ' '))
    }
    Write-Log "Registered ONLOGON prompt task as SYSTEM (/IT)."
  }
  else {
    $out = & schtasks.exe /Create /TN $TaskOnLogon /SC ONLOGON /RU $currentIdentity /RL LIMITED /IT /TR "`"$UiWrapperPath`"" /F 2>&1
    Write-Log ("ONLOGON schtasks /Create output: {0}" -f (($out | ForEach-Object { $_.ToString().Trim() }) -join ' | ')) 'DEBUG'
    if ($LASTEXITCODE -ne 0) {
      throw ("Failed to create ONLOGON prompt task as current user: {0}" -f ($out -join ' '))
    }
    Write-Log ("Registered ONLOGON prompt task as current user '{0}'." -f $currentIdentity)
  }

  Write-TaskDiagnostics -TaskName $TaskOnLogon -Prefix 'ONLOGON_TASK'
}

function Get-AvailableUpdates {
  $session  = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
  Write-Log ("Searching Windows Update with criteria: {0}" -f $criteria) 'DEBUG'
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

function Write-AvailableUpdatesLog {
  param(
    [Parameter(Mandatory)][object[]]$Updates,
    [string]$Prefix = 'UPDATES'
  )

  if (-not $Updates -or $Updates.Count -eq 0) {
    Write-Log ("{0}: No available updates." -f $Prefix)
    return
  }

  Write-Log ("{0}: Found {1} available update(s)." -f $Prefix, $Updates.Count)
  foreach ($u in $Updates) {
    Write-Log ("{0}: Title='{1}'; RebootBehavior={2}; IsSSU={3}; IsLCU={4}" -f `
      $Prefix, $u.Title, $u.RebootBehavior, $u.IsSSU, $u.IsLCU)
  }
}

function Select-EligibleUpdates {
  param(
    [Parameter(Mandatory)][object[]]$Updates,
    [switch]$IncludeRebootUpdates,
    [bool]$EnsureLatestCumulativeUpdate
  )

  $ssu    = @($Updates | Where-Object { $_.IsSSU })
  $lcu    = @($Updates | Where-Object { $_.IsLCU })
  $others = @($Updates | Where-Object { -not $_.IsSSU -and -not $_.IsLCU })

  $toInstall = @()
  if ($EnsureLatestCumulativeUpdate) {
    $toInstall += $ssu
    $toInstall += $lcu
    if ($IncludeRebootUpdates) { $toInstall += $others }
    else { $toInstall += @($others | Where-Object { $_.RebootBehavior -eq 0 }) }
  } else {
    if ($IncludeRebootUpdates) { $toInstall = $Updates }
    else { $toInstall = @($Updates | Where-Object { $_.RebootBehavior -eq 0 }) }
  }

  return @($toInstall | Select-Object -Unique)
}

function Install-Updates {
  param(
    [Parameter(Mandatory)]$Session,
    [Parameter(Mandatory)][object[]]$UpdatesToInstall
  )

  if (-not $Session) { throw "Windows Update session is null." }
  if (-not $UpdatesToInstall -or $UpdatesToInstall.Count -eq 0) {
    return [pscustomobject]@{
      InstalledCount = 0
      RebootRequired = $false
    }
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
    $updates = @($scan.Updates)

    if (-not $updates -or $updates.Count -eq 0) {
      Write-Log ("WU pass {0}/{1}: No available updates." -f $pass, $MaxPasses)
      break
    }

    $foundAny = $true
    Write-AvailableUpdatesLog -Updates $updates -Prefix ("WU pass {0}/{1}" -f $pass, $MaxPasses)

    $toInstall = Select-EligibleUpdates -Updates $updates -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate
    if (-not $toInstall -or $toInstall.Count -eq 0) {
      Write-Log ("WU pass {0}/{1}: Updates found but none eligible under rules." -f $pass, $MaxPasses)
      break
    }

    $eligibleFound = $true
    Write-Log ("WU pass {0}/{1}: Installing eligible update(s): {2}" -f `
      $pass, $MaxPasses, (($toInstall | ForEach-Object { $_.Title }) -join ' | '))

    $r = Install-Updates -Session $scan.Session -UpdatesToInstall $toInstall
    $installedTotal += $r.InstalledCount
    $rebootFromInstall = $rebootFromInstall -or $r.RebootRequired

    Write-Log ("WU pass {0}/{1}: Installed {2}. RebootRequired={3}" -f $pass, $MaxPasses, $r.InstalledCount, $r.RebootRequired)
    Write-LatestInstalledUpdateLog -Prefix ("WU_PASS_{0}_LATEST_INSTALLED_UPDATE" -f $pass)

    if ($rebootFromInstall) { break }
  }

  return [pscustomobject]@{
    FoundAny          = $foundAny
    EligibleFound     = $eligibleFound
    InstalledTotal    = $installedTotal
    RebootFromInstall = $rebootFromInstall
  }
}

function Register-PostBootWorkerTask {
  try { Import-Module ScheduledTasks -ErrorAction Stop } catch { throw "ScheduledTasks module is required." }

  Remove-Task $TaskPostBoot
  $scriptToRun = Ensure-MainScriptCopy

  $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
  $arg = ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Mode PostBoot -CountdownMinutes {1} -PostponeOptionsCsv "{2}" -BaseDir "{3}" -LogPath "{4}" -PostRebootMaxPasses {5} -EnsureLatestCumulativeUpdate:{6}' -f `
    $scriptToRun, $CountdownMinutes, ($PostponeOptionsMinutes -join ','), $BaseDir, $LogPath, $PostRebootMaxPasses, $EnsureLatestCumulativeUpdate.ToString().ToLower())

  if ($IncludeRebootUpdates) { $arg += " -IncludeRebootUpdates" }

  $action    = New-ScheduledTaskAction -Execute $psExe -Argument $arg
  $trigger   = New-ScheduledTaskTrigger -AtStartup
  $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
  $settings  = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

  Register-ScheduledTask -TaskName $TaskPostBoot -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Log "Registered post-boot Windows Update worker task (runs at startup as SYSTEM)."
  Write-TaskDiagnostics -TaskName $TaskPostBoot -Prefix 'POSTBOOT_TASK'
}

function Run-PostBootWorker {
  Write-Log ("POSTBOOT: Starting Windows Update worker. Identity='{0}'; PSVersion='{1}'" -f `
    (Get-CurrentIdentityName), $PSVersionTable.PSVersion)

  $pass = 0
  $rebootFromInstall = $false
  $installedTotal = 0

  while ($pass -lt $PostRebootMaxPasses) {
    $pass++
    $scan = Get-AvailableUpdates
    $updates = @($scan.Updates)

    if (-not $updates -or $updates.Count -eq 0) {
      Write-Log ("POSTBOOT: Pass {0}/{1}: No available updates." -f $pass, $PostRebootMaxPasses)
      break
    }

    Write-AvailableUpdatesLog -Updates $updates -Prefix ("POSTBOOT: Pass {0}/{1}" -f $pass, $PostRebootMaxPasses)

    $toInstall = Select-EligibleUpdates -Updates $updates -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate
    if (-not $toInstall -or $toInstall.Count -eq 0) {
      Write-Log ("POSTBOOT: Pass {0}/{1}: Updates found but none eligible under rules." -f $pass, $PostRebootMaxPasses)
      break
    }

    Write-Log ("POSTBOOT: Pass {0}/{1}: Installing eligible update(s): {2}" -f `
      $pass, $PostRebootMaxPasses, (($toInstall | ForEach-Object { $_.Title }) -join ' | '))

    $r = Install-Updates -Session $scan.Session -UpdatesToInstall $toInstall
    $installedTotal += $r.InstalledCount
    $rebootFromInstall = $rebootFromInstall -or $r.RebootRequired

    Write-Log ("POSTBOOT: Pass {0}/{1}: Installed {2}. RebootRequired={3}" -f $pass, $PostRebootMaxPasses, $r.InstalledCount, $r.RebootRequired)
    Write-LatestInstalledUpdateLog -Prefix ("POSTBOOT_PASS_{0}_LATEST_INSTALLED_UPDATE" -f $pass)

    if ($rebootFromInstall) { break }
  }

  $pending = Test-PendingReboot
  Write-Log ("POSTBOOT: InstalledTotal={0} RebootFromInstall={1} PendingReboot={2}" -f $installedTotal, $rebootFromInstall, $pending)
  Write-LatestInstalledUpdateLog -Prefix 'POSTBOOT_LATEST_INSTALLED_UPDATE'

  if ($pending -or $rebootFromInstall) {
    Write-Log "POSTBOOT: Reboot required. Keeping ONLOGON prompt task."
  } else {
    Remove-Task $TaskOnLogon
    Write-Log "POSTBOOT: System is clean. Removed ONLOGON prompt task."
  }

  Remove-Task $TaskPostBoot
  Write-Log "POSTBOOT: Worker task removed."
}

try {
  Ensure-BaseDirPermissions

  Write-Host "Windows Update (Consent-only + auto re-run + proven active-session UI launch; prompt only if reboot required)"
  Write-Log ("Script start: Mode='{0}'; Identity='{1}'; Computer='{2}'; PSVersion='{3}'; IncludeRebootUpdates='{4}'; EnsureLatestCumulativeUpdate='{5}'; ReportOnly='{6}'; CountdownMinutes='{7}'" -f `
    $Mode, (Get-CurrentIdentityName), $env:COMPUTERNAME, $PSVersionTable.PSVersion, [bool]$IncludeRebootUpdates, $EnsureLatestCumulativeUpdate, [bool]$ReportOnly, $CountdownMinutes)

  $os = Get-OsBuildInfo
  if ($os) { Write-Log ("OS_BUILD_INFO: {0}" -f $os) }

  if (-not (Test-IsAdmin)) {
    throw "Run this script elevated (Administrator / SYSTEM)."
  }

  Abort-AnyShutdown
  Ensure-MainScriptCopy | Out-Null
  Write-UiHelperScript

  if ($Mode -eq 'PromptOnly') {
    Write-Log "PromptOnly mode: checking reboot state and relaunching UI if needed."
    if (Test-PendingReboot) {
      Launch-UiInActiveSessionNow
      exit 1
    } else {
      Write-Log "PromptOnly mode: reboot not required; nothing to show."
      Remove-Task $TaskReminder
      exit 0
    }
  }

  if ($Mode -eq 'PostBoot') {
    Run-PostBootWorker
    exit 0
  }

  if ($ReportOnly) {
    Write-Log "ReportOnly=True. Scan-only mode; no installs will be attempted."
    $beforePending = Test-PendingReboot
    if ($beforePending) { Write-Log "ReportOnly: System already indicates a pending reboot." }

    $scan = Get-AvailableUpdates
    $updates = @($scan.Updates)
    Write-AvailableUpdatesLog -Updates $updates -Prefix 'REPORTONLY'

    $eligible = Select-EligibleUpdates -Updates $updates -IncludeRebootUpdates:$IncludeRebootUpdates -EnsureLatestCumulativeUpdate:$EnsureLatestCumulativeUpdate
    if ($eligible.Count -gt 0) {
      Write-Log ("ReportOnly: Eligible updates count={0}" -f $eligible.Count)
      foreach ($u in $eligible) {
        Write-Log ("ReportOnly: Eligible update: Title='{0}'; RebootBehavior={1}" -f $u.Title, $u.RebootBehavior)
      }
    } else {
      Write-Log "ReportOnly: No eligible updates under current rules."
    }

    $afterPending = Test-PendingReboot
    $needsReboot = $beforePending -or $afterPending

    Write-Log ("ReportOnly: FoundAny={0} EligibleFound={1} NeedsReboot={2}" -f `
      [bool]($updates.Count -gt 0), [bool]($eligible.Count -gt 0), $needsReboot)
    Write-LatestInstalledUpdateLog -Prefix 'REPORTONLY_LATEST_INSTALLED_UPDATE'
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
  Write-LatestInstalledUpdateLog -Prefix 'PREBOOT_LATEST_INSTALLED_UPDATE'

  if ($needsReboot) {
    Save-State @{
      CreatedAt                    = (Get-Date).ToString("s")
      IncludeRebootUpdates         = [bool]$IncludeRebootUpdates
      EnsureLatestCumulativeUpdate = [bool]$EnsureLatestCumulativeUpdate
      PostRebootMaxPasses          = [int]$PostRebootMaxPasses
    }
    Write-Log ("Saved state file to '{0}'." -f $StatePath) 'DEBUG'

    Register-PostBootWorkerTask
    Launch-UiInActiveSessionNow
    exit 1
  }

  try {
    if (Test-Path $StatePath) {
      Remove-Item $StatePath -Force -ErrorAction SilentlyContinue
      Write-Log ("Removed state file '{0}'." -f $StatePath) 'DEBUG'
    }
  } catch { }

  Remove-UiTasks
  Remove-LogonPromptTaskIfNotNeeded

  if ($result.FoundAny -and -not $result.InstalledTotal) {
    Write-Log "Finished. Updates were found but none were installed under current rules."
    exit 2
  }

  Write-Log "Finished. No reboot required and no pending updates detected."
  exit 0
}
catch {
  Write-ExceptionLog -Prefix 'Unhandled script error' -ErrorRecord $_
  exit 3
}