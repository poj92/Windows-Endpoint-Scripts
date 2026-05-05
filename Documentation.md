# Windows Endpoint Scripts - Documentation

Complete reference guide for all PowerShell scripts in this repository. These scripts are optimized for deployment through Datto RMM and other endpoint management platforms.

Author:             Peter James
Email:              peter.james@nexusos.co.uk
Document Version:   v1.2

---

## Table of Contents

1. [Browser Scripts](#browser-scripts)
2. [Java Scripts](#java-scripts)
3. [Windows Update Scripts](#windows-update-scripts)
4. [ConnectSecure](#connectsecure)
5. [Dell Peripheral Manager](#dell-peripheral-manager)
6. [Office Application](#office-application)
7. [Python](#python)
8. [VC++ Redistributables](#vc-redistributables)
9. [Zoom](#zoom)
10. [Winget](#winget)
11. [Collection Utilities](#collection-utilities)

---

## Browser Scripts

### Browser-Update-Detection.ps1

**Purpose**: Detects when browsers have pending updates that require a restart to apply. Tracks browser usage state and generates update notifications.

**Execution Context**: Must run as SYSTEM

**Key Features**:
- Detects Chrome, Firefox, and Edge sessions
- Tracks browser usage history with JSON-based state tracking
- Identifies browsers running long enough to warrant restart
- Logs all detection activity for auditing

**Configuration Paths**:
- Base Directory: `C:\ProgramData\Datto\BrowserUpdateCheck`
- Log File: `Detection.log`
- Tracking File: `BrowserUsageTracking.json`
- Reload Queue: `ReloadQueue.json`

**Parameters** (Configuration Variables):
- `$BrowserReloadThresholdHours` (default: 24) - Hours of runtime before restart recommended
- `$ApiTimeoutSeconds` (default: 20) - API call timeout

**Exit Codes**:
- Returns output to stdout for RMM parsing

**Typical RMM Setup**: Set as detection script paired with Browser-Apply-Update-Notify-Reload.ps1

---

### Browser-Apply-Update-Notify-Reload.ps1

**Purpose**: Forces graceful browser restart for pending updates. Shows user notification with countdown, handles postpone requests, and enforces browser reload when updates are pending.

**Execution Context**: Must run as logged-in user context (use Datto InteractiveToken)

**Key Features**:
- Graceful browser restart with user notification
- Countdown timer (default 300 seconds)
- Postpone options for users
- Scheduled task cleanup on completion
- Comprehensive logging for troubleshooting

**RMM Component Variables**:
- `ScheduledTaskName` - Optional scheduled task identifier

**Configuration Variables**:
- `$WarningTimeSeconds` (default: 300) - Countdown until forced restart
- `$CompanyName` (default: "Nexus Open Systems Ltd") - Displayed in notifications
- `$RemediationScriptPath` - Path to remediation script

**Exit Codes**: Not specified (returns success/failure via output)

**Typical RMM Setup**: Deploy as remediation script triggered by Browser-Update-Detection.ps1 detection

---

### Browser-Force-Idle-Open-Close.ps1

**Purpose**: Safely launches browsers that haven't been used within a lookback window, then closes them after a specified duration. Useful for triggering update checks without user intervention.

**Execution Context**: Can run as SYSTEM or user context

**Key Features**:
- Per-browser safety checks (never touches already-running browsers)
- Uses Prefetch timestamps to detect recent usage
- Creates temporary profile directories for clean launches
- Only closes processes it started (via profile path matching)
- Supports multiple browser preferences

**Parameters**:
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `LookbackHours` | int | 24 | Hours to look back in usage history |
| `OpenSeconds` | int | 120 | Seconds to keep browser open |
| `Url` | string | 'about:blank' | URL to open |
| `ReportOnly` | switch | False | Report without making changes |
| `Preference` | string[] | @('Edge','Chrome','Firefox','Brave','Opera') | Browser launch order preference |

**Exit Codes**:
- `0` - No action needed (all browsers running or within window)
- `1` - At least one browser launched and closed successfully
- `2` - ReportOnly mode: would have launched browsers
- `3` - No supported browsers installed
- `4` - Error occurred

---

### Browser-Force-Update-And-Reload.ps1

**Purpose**: Forces Google Chrome and Microsoft Edge to check for and apply updates immediately. Prompts users for reload if required, with auto-reload after countdown expires.

**Execution Context**: Recommended as SYSTEM (uses InteractiveToken for user-facing prompts)

**Key Features**:
- Force immediate update check for Chrome and Edge
- User-friendly reload prompt with countdown
- Automatic reload if user doesn't respond
- Multiple update triggers for stability across builds
- Preserves browser sessions via session restore
- One-time scheduled task for user prompts

**RMM Component Variables**:
- `PromptOnly` - Show prompt without browser update
- `ScheduledTaskName` - Custom scheduled task identifier
- `BasePath` - Working directory (default: `C:\ProgramData\Datto\BrowserForceUpdate`)
- `CompanyName` - Displayed in notifications
- `CountdownSeconds` (default: 300) - Reload countdown timer
- `UpdateWaitSeconds` (default: 90) - Max wait for update check
- `PollIntervalSeconds` (default: 5) - Check interval
- `GracefulCloseWaitSeconds` (default: 10) - Grace period for close
- `RelaunchDelaySeconds` (default: 2) - Delay before relaunch
- `PromptOnlyWhenBrowserIsRunning` (default: true) - Only prompt if running
- `ForceKillRemainingProcesses` (default: true) - Force kill unresponsive

**Exit Codes**: Returns success/failure via script invocation

---

## Java Scripts

### Java-JRE-Update.ps1

**Purpose**: Manages Java JRE (Temurin or Oracle) installations. Detects, reports, installs, or upgrades JRE for specified major versions with optional cleanup.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Supports Java 8, 11, 17, 21 (major families)
- Automatic version detection via winget
- Optional MSI fallback for environments without winget
- Selective removal of older versions in same family
- JAVA_HOME and PATH cleanup
- Comprehensive version normalization

**RMM Component Variables**:

| Parameter | Type | Default | Allowed Values | Description |
|-----------|------|---------|-----------------|-------------|
| `Vendor` | string | Temurin | Temurin, Oracle | Java distribution |
| `TargetFamily` | int | 17 | 8, 11, 17, 21 | Java major version |
| `ReportOnly` | switch | False | - | Check without making changes |
| `RemoveOlder` | switch | False | - | Remove older versions in family |
| `Cleanup` | switch | False | - | Clean stale env vars and PATH |
| `Force` | switch | False | - | Force update even if current |
| `UseMsiFallback` | switch | True | - | Fall back to MSI if winget unavailable |
| `LogPath` | string | $env:ProgramData\JavaUpdate\JavaJRE-Update.log | - | Log file location |

**Exit Codes**:
- `0` - Up-to-date or no action required
- `1` - Performed an update or install
- `2` - ReportOnly mode: update would be required
- `3` - Update required but cannot proceed (winget missing, MSI fallback disabled)

**Example RMM Command**: `.\Java\Java-JRE-Update.ps1 -TargetFamily 17 -RemoveOlder -Cleanup`

---

### Java-SDK-Update.ps1

**Purpose**: Manages Java JDK (Temurin or Oracle) installations. Detects, reports, installs, or upgrades JDK for specified major versions with optional cleanup.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Same versioning engine as JRE script
- Supports Java 8, 11, 17, 21
- Auto-detection via winget
- MSI fallback support (Temurin only)
- Selective removal of older versions
- Environment cleanup capabilities

**RMM Component Variables**: Identical to Java-JRE-Update.ps1

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Vendor` | string | Temurin | Java distribution |
| `TargetFamily` | int | 21 | Java major version (8/11/17/21) |
| `ReportOnly` | switch | False | Check without changes |
| `RemoveOlder` | switch | False | Remove older versions in family |
| `Cleanup` | switch | False | Clean stale env vars and PATH |
| `Force` | switch | False | Force update when current |
| `UseMsiFallback` | switch | True | Fall back to MSI without winget |
| `LogPath` | string | $env:ProgramData\JavaUpdate\JavaJDK-Update.log | Log file location |

**Exit Codes**: Same as Java-JRE-Update.ps1

**Note**: Default TargetFamily is 21 (JDK) vs 17 (JRE)

---

## Windows Update Scripts

### Windows-Update-No-Auto-Reboot.ps1

**Purpose**: Installs Windows updates with full user control. Never reboots automatically; user must explicitly click "Reboot now". Supports postponement, auto post-boot updates, and re-prompting if additional reboots needed.

**Execution Context**: Can run as SYSTEM (creates tasks for user session interaction)

**Key Features**:
- User consent-only reboot strategy
- Postpone options (30m, 1h, 2h configurable)
- Post-boot worker auto-runs updates after user reboot
- Re-prompting if another reboot becomes required
- Interactive UI via scheduled task or fallback to msg.exe
- Comprehensive reboot detection (registry, WMI, updates COM)

**RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Mode` | string | Normal | Execution mode: Normal, PostBoot, PromptOnly |
| `CountdownMinutes` | int | 10 | Reboot countdown duration |
| `PostponeOptionsMinutes` | int[] | @(30, 60, 120) | Postpone duration options |
| `PostponeOptionsCsv` | string | - | CSV override for postpone options |
| `IncludeRebootUpdates` | switch | False | Include reboot-requiring updates |
| `EnsureLatestCumulativeUpdate` | bool | True | Prioritize SSU/LCU updates |
| `ReportOnly` | switch | False | Report without installing |
| `PostRebootMaxPasses` | int | 4 | Post-boot update attempts |
| `UiTitle` | string | "A security message from Nexus Open Systems Ltd" | Notification title |
| `Reason` | string | (security message) | Notification message |
| `BaseDir` | string | $env:ProgramData\NexusOpenSystems\WindowsUpdate | Working directory |
| `LogPath` | string | .../WindowsUpdateReboot.log | Log file location |

**Scheduled Tasks Created**:
- `Nexus_WU_PostBootWorker` - Runs updates after user reboot
- `Nexus_WU_RebootReminder` - Reminder timeout scheduler
- `Nexus_WU_PromptOnLogon` - Show prompt at user logon
- `Nexus_WU_PromptNow` - Show prompt immediately

**Exit Codes**:
- `0` - No reboot needed; updates installed or none available
- `1` - Reboot required/pending; prompt launched or scheduled (no forced reboot)
- `2` - Report-only OR updates found but none installed under rules
- `3` - Error

**Typical RMM Setup**: Schedule weekly or bi-weekly; use Mode=Normal for standard operation, Mode=PostBoot for automatic post-reboot updates

---

### Recent Windows folder changes

- Added `Windows-Update-No-Auto-Reboot-Ensure-LatestBuild.ps1`: a thin wrapper around `Windows-Update-No-Auto-Reboot.ps1` that enforces `EnsureLatestCumulativeUpdate` (SSU/LCU prioritization) while preserving the same header, parameters, logging, and exit-code patterns for RMM usage.
- Standardized headers/parameters/logging across the Windows folder to match the `Windows-Update-No-Auto-Reboot.ps1` reference pattern. Files reviewed and confirmed or updated:
  - `Windows-Update-No-Auto-Reboot.ps1` (reference, unchanged)
  - `Windows-Update-Apply-Auto-Reboot.ps1` (follows reference pattern)
  - `Windows-Update-No-Postpone.ps1` (follows reference pattern)
  - `Windows-Update-No-Auto-Reboot-Ensure-LatestBuild.ps1` (new wrapper added)
  - `winget-script.ps1` (package-management style; uses consistent Datto env helpers and logging)

These changes ensure consistent RMM variable handling, `Write-Log` usage, and explicit exit-code documentation across Windows update scripts. If you want, I can further unify helper function names (e.g., exact `Write-Log` signature) across all files.


### Windows-Update-Apply-Auto-Reboot.ps1

**Purpose**: Installs all available Windows updates and automatically reboots if required (with countdown). No postpone option; reboot is mandatory after countdown expires.

**Execution Context**: Can run as SYSTEM (InteractiveToken for user UI)

**Key Features**:
- Installs reboot-free updates immediately
- Auto-reboot for required updates with countdown
- User-friendly notification UI
- Supports no-reboot updates and reboot-requiring updates
- Session isolation (respects already-running-browsers rule)
- Automatic cumulative update prioritization

**RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `CountdownMinutes` | int | 10 | Reboot countdown timer |
| `IncludeRebootUpdates` | switch | False | Include reboot-requiring updates |
| `EnsureLatestCumulativeUpdate` | bool | True | Prioritize SSU/LCU |
| `ReportOnly` | switch | False | Report mode |
| `UiTitle` | string | "A security message from Nexus Open Systems Ltd" | Notification title |
| `Reason` | string | (security message) | Notification message |
| `LogPath` | string | $env:ProgramData\NexusOpenSystems\WindowsUpdate\WindowsUpdateReboot.log | Log file location |

**Exit Codes**:
- `0` - No reboot needed; updates installed
- `1` - Reboot prompt launched/scheduled
- `2` - Updates found but none installed
- `3` - Error

---

### Windows-Update-No-Postpone.ps1

**Purpose**: Installs Windows updates with countdown reboot UI but NO postpone option. "Restart now" or wait for countdown; reboot is inevitable.

**Execution Context**: Can run as SYSTEM

**Key Features**:
- Forces reboot decision (now or countdown)
- No postpone options available
- Installs non-reboot updates first
- Interactive countdown with restart button
- Fallback to msg.exe + Windows shutdown countdown
- All reboot precondition checking

**RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `CountdownMinutes` | int | 10 | Reboot countdown duration |
| `IncludeRebootUpdates` | switch | False | Include reboot-requiring updates |
| `EnsureLatestCumulativeUpdate` | bool | True | Prioritize SSU/LCU |
| `ReportOnly` | switch | False | Report mode |
| `UiTitle` | string | "A security message from Nexus Open Systems Ltd" | Notification title |
| `Reason` | string | (security message) | Notification message |
| `LogPath` | string | $env:ProgramData\NexusOpenSystems\WindowsUpdate\WindowsUpdateReboot.log | Log file location |

**Exit Codes**: Same as Windows-Update-Apply-Auto-Reboot.ps1

**Typical RMM Setup**: Use in maintenance windows where immediate reboot is acceptable

---

## ConnectSecure

### ConnectSecure-Agent-Install.ps1

**Purpose**: Installs or updates ConnectSecure (CyberCNS) security agent on Windows endpoints.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Checks for existing agent installation
- Downloads latest agent from ConnectSecure API
- Installs with provided credentials
- Validates successful installation

**Datto RMM Component Variables** (Required):

Create these EXACT variable names in Datto component Variable editor:

| Variable Name | Type | Example | Description |
|---------------|------|---------|-------------|
| `CompanyID` | String | ABC123 | Organization tenant ID |
| `TenantID` | String | TID456 | Tenant identifier |
| `Secret` | String | secret_key_value | Authentication secret |

**Configuration**:
- Agent Path Check: `C:\Program Files (x86)\CyberCNSAgent\cybercnsagent.exe`
- API Endpoint: `https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows`
- Download Location: `%TEMP%\cybercnsagent.exe`
- TLS Protocol: 1.2 (minimum)

**Exit Codes**:
- `0` - Agent already installed or installation succeeded
- Non-zero - Installation failed (check downloaded log)

**Typical RMM Setup**: Deploy once per endpoint or include in onboarding workflow

---

### ConnectSecure-Agent-Removal.ps1

**Purpose**: Uninstalls the ConnectSecure (CyberCNS) agent from a Windows endpoint.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Silent uninstall via agent executable
- Requires an uninstallation key provided via Datto environment variable

**Datto RMM Component Variables**:

| Variable Name | Type | Description |
|---------------|------|-------------|
| `cyberKey` | String | Uninstallation key required by the agent uninstall command |

**Command**:
Runs from the agent install folder and invokes the agent uninstall switch:

```
cd "C:\Program Files (x86)\CyberCNSAgent"
.\cybercnsagent.exe -r -y $env:cyberKey
```


## Dell Peripheral Manager

### DDPM-Fix.ps1

**Purpose**: Detects and remediates vulnerable Dell Display and Peripheral Manager (DDPM) installations. Enforces minimum safe version and updates outdated installations.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Automatic Dell device detection (skips non-Dell systems)
- Identifies vulnerable DDPM/DPM versions
- Removes old versions and verifies uninstall
- Downloads and installs specified DDPM version
- SHA-256 validation for downloaded packages
- Comprehensive exit codes for Detect and Remediate modes

**Datto RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Mode` | string | Remediate | Execution mode: Detect or Remediate |
| `MinimumSafeVersion` | string | 1.7.6 | Minimum safe DDPM version |
| `SkipNonDell` | bool | true | Skip non-Dell devices |
| `LatestDDPMUrl` | string | (empty) | Download URL for DDPM installer |
| `LatestDDPMVersion` | string | (empty) | DDPM version to install |
| `LatestDDPMSha256` | string | (empty) | SHA-256 hash for verification |
| `LatestDDPMFileName` | string | (empty) | Downloaded file name |
| `LatestDDPMSilentArgs` | string | /S | Installation arguments |
| `LogToFile` | bool | true | Enable file logging |
| `LogPath` | string | C:\ProgramData\DattoRMM\Logs\DellPeripheral-DDPM-Compliance.log | Log file location |
| `WorkingDirectory` | string | C:\ProgramData\DattoRMM\Packages\DellDDPM | Working directory |
| `ForceTls12` | switch | False | Force TLS 1.2 |

**Exit Codes**:
- `0` - Compliant or remediation succeeded
- `1` - Non-compliant detected (Detect mode)
- `2` - Fatal script error
- `3` - Skipped non-Dell device
- `4` - Below-minimum found, uninstall verification failed
- `5` - Older version removed, DDPM package metadata missing
- `6` - DDPM downloaded/installed but not detected afterwards

**Example RMM Setup**:

```powershell
# Detect mode
.\DDPM-Fix.ps1 -Mode Detect -MinimumSafeVersion 1.7.6 -SkipNonDell $true

# Remediate mode
.\DDPM-Fix.ps1 -Mode Remediate `
  -MinimumSafeVersion 1.7.6 `
  -LatestDDPMUrl "https://dl.dell.com/..." `
  -LatestDDPMVersion "2.2.1.16" `
  -LatestDDPMSha256 "ae6e965495ad54b78fab64a580f143cfa7c73e82ee5a473a1a925a6ac5c203a7" `
  -LatestDDPMFileName "DDPM-Setup_2.2.1.16.exe"
```

---

## Office Application

### Office-Force-Update.ps1

**Purpose**: Triggers immediate updates for Click-to-Run Microsoft 365 Apps / Office 365 on Windows endpoints.

**Execution Context**: Can run as SYSTEM or user context

**Key Features**:
- Detects Click-to-Run Office installations
- Triggers Office Automatic Updates 2.0 task
- Safe execution (doesn't force-close apps by default)
- Optional app closure during maintenance windows
- Registry-based configuration tracking

**RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ForceCloseOfficeApps` | bool | false | Close Office apps before update (use only in maintenance window) |

**Monitored Office Processes**:
- WINWORD (Word)
- EXCEL (Excel)
- POWERPNT (PowerPoint)
- OUTLOOK (Outlook)
- ONENOTE (OneNote)
- MSACCESS (Access)
- VISIO (Visio)
- LYNC (Skype for Business)
- MSPUB (Publisher)

**Registry Locations Checked**:
- HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
- HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Updates

**Scheduled Task**:
- Task Name: `\Microsoft\Office\Office Automatic Updates 2.0`

**Exit Codes**:
- `0` - Update triggered successfully
- `1` - Click-to-Run Office not found
- `2` - Office update scheduled task not found
- `3` - Failed to trigger scheduled task

---

## Python

### Python-Update.ps1

**Purpose**: Reports and manages Python installations. Detects current version and initiates updates to latest stable release when below minimum version.

**Execution Context**: Recommended as Administrator (for reliable version detection/installation)

**Key Features**:
- Silent version reporting
- Automatic update detection
- Latest stable version installation
- Comprehensive action logging
- Datto RMM variable support

**Datto RMM Component Variables** (Optional):

| Variable Name | Type | Default | Description |
|---------------|------|---------|-------------|
| `MinimumVersion` | String | (auto-detect) | Minimum required Python version |
| `UpdateMethod` | String | Auto | Update installation method |
| `LogPath` | String | $env:ProgramData\Datto\Logs\ | Log file directory |

**Exit Codes**: Not explicitly defined (returns success/failure)

**Note**: Script requires further development for complete integration. Check implementation for current capabilities.

---

## VC++ Redistributables

### VC++-Install.ps1

**Purpose**: Manages Visual C++ redistributable installations. Detects installed versions, enforces minimum version requirements, and updates older installations selectively by architecture.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Architecture-aware installation (x86 and/or x64)
- Minimum version enforcement
- Selective uninstall of below-minimum versions
- Multiple pass detections with removal verification
- MSI force option for problematic systems
- Automatic cleanup of downloaded installers

**Datto RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `TargetUrlX64` | string | (empty) | x64 VC++ installer URL |
| `TargetUrlX86` | string | (empty) | x86 VC++ installer URL |
| `MinKeepVersion` | string | (required) | Minimum version to retain |
| `ReportOnly` | switch | False | Report-only mode |
| `IncludeX64` | switch | False | Include x64 architecture |
| `IncludeX86` | switch | False | Include x86 architecture |
| `ForceMSI` | switch | False | Force MSI installation method |
| `LogPath` | string | $env:ProgramData\NexusOpenSystems\VCRedist\VCRedistUpdate.log | Log file location |

**Rules**:
1. Detects installed VC++ entries from Add/Remove Programs
2. If no entries below MinKeepVersion: Do nothing
3. If any below MinKeepVersion:
   - Install URLs only for affected architectures
   - Uninstall all below-minimum entries
   - Re-scan and remove remaining old entries
4. Clean up downloaded installers

**Exit Codes**:
- `0` - No changes needed
- `1` - Changes made successfully
- `2` - Report-only mode: changes would be made
- `3` - Error occurred

**Example RMM Setup**:

```powershell
.\VC++-Install.ps1 `
  -TargetUrlX64 "https://cdn.visualstudio.com/..." `
  -TargetUrlX86 "https://cdn.visualstudio.com/..." `
  -MinKeepVersion "14.30" `
  -IncludeX64 `
  -IncludeX86
```

---

## Zoom

### Zoom-Cleanup.ps1

**Purpose**: Manages Zoom client installations. Detects current version and removes outdated installations while preserving minimum required version.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Version detection and comparison
- Selective removal of outdated versions
- Optional process termination before cleanup
- Optional residual file cleanup
- Report-only mode for verification
- Operation timeout handling

**Datto RMM Component Variables** (Create in Datto component Variable editor):

| Variable Name | Type | Default | Description |
|---------------|------|---------|-------------|
| `MinimumVersionToKeep` | String | 6.3.0.0 | Minimum Zoom version to retain |
| `ReportOnly` | Boolean | true | Don't make changes, report only |
| `StopZoomProcesses` | Boolean | true | Terminate Zoom processes before cleanup |
| `CleanupResiduals` | Boolean | true | Remove stale Zoom files/config |
| `OperationTimeoutSec` | String/Number | 600 | Operation timeout in seconds |

**Environment Variables** (Read from Datto):
- Reads variables from `$env:VariableName`
- Falls back to hardcoded defaults if not set

**Configuration**:
- Log Path: `C:\ProgramData\Datto\Logs\Zoom-Cleanup.log`
- Detection Registry: Windows Add/Remove Programs (HKLM, HKCU)

**Exit Codes**: Not explicitly defined (returns success/failure)

**Example RMM Setup**:

```
MinimumVersionToKeep = 6.4.0.0
ReportOnly = false
StopZoomProcesses = true
CleanupResiduals = true
OperationTimeoutSec = 300
```

---

## Winget

### winget-script.ps1

**Purpose**: Upgrades all installed packages via Windows Package Manager (winget) to their latest available versions. Supports automatic winget installation if missing.

**Execution Context**: Can run as SYSTEM or user (must have package access)

**Key Features**:
- Auto-detects installed packages
- Individual package upgrade with success/failure tracking
- Optional automatic winget installation
- Comprehensive upgrade logging with timestamps
- Report-only mode for planned upgrades
- Datto RMM environment variable support

**Datto RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ReportOnly` | bool | false | Report without upgrading |
| `Install_Winget_if_Not_Available` | bool | false | Auto-install winget if missing |
| `LogPath` | string | $env:ProgramData\Datto\Logs\Winget-UpgradeAll.log | Log file location |

**Environment Variable Overrides** (from Datto):
- `ReportOnly` (env) - Boolean override
- `Install_Winget_if_Not_Avaialble` (env) - Boolean override
- `LogPath` (env) - Path override

**Configuration**:
- Temp Directory: `C:\ProgramData\Datto\Temp\WingetUpgradeAll`
- Individual Package Logs: `C:\ProgramData\Datto\Logs\winget-upgrade-{StepName}-{DateTime}.log`
- Supports custom RMM helper functions for environment variable handling

**Exit Codes**: Not explicitly defined (returns success/failure)

**Example RMM Setup**:

```
ReportOnly = false              (check first without upgrading)
Install_Winget_if_Not_Avaialble = true  (auto-download winget if needed)
```

**Notes**:
- Author: Peter James
- Version: 1.3
- Designed for Datto RMM deployment

---

## Collection Utilities

### Collect-Install-Logs.ps1

**Purpose**: Retrieves and displays recent application installation and uninstallation events from Windows Application logs. Useful for auditing deployment and software management activities.

**Execution Context**: Can run as standard user (read-only event log access)

**Key Features**:
- Configurable historical lookback window (hours)
- Filters for specific MSI installation/uninstallation event IDs
- User-friendly action classification
- Chronologically sorted output
- Silent execution (no errors if no events found)

**Datto RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `StartTime` | int | 24 | Hours to look back in event log |

**Monitored Event IDs**:
- `11707` - MSI Install successful
- `11708` - MSI Install failed
- `11724` - MSI Uninstall successful
- `11725` - MSI Uninstall failed
- `1033` - Product installed
- `1034` - Product removed

**Output Columns**:
- TimeCreated - Event timestamp
- Id - Event ID
- ProviderName - Always "MsiInstaller"
- MachineName - Computer name
- Action - Translated action (Install/Uninstall/etc)
- Message - Full event message

**Exit Codes**:
- `0` - Success (events found or none available)

**Example RMM Setup**:

```powershell
.\Collect-Install-Logs.ps1   # Uses default 24 hours
# Output is formatted for Datto RMM event reporting
```

**Typical Use Case**: Discovery script for installation/deployment compliance auditing

---
