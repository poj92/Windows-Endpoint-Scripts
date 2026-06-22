# Windows Endpoint Scripts — concise reference

Purpose: short guide to the repository's PowerShell remediation and detection scripts, designed for Datto RMM or similar endpoint managers.

Author: Peter James — peter.james@nexusos.co.uk
Version: v1.3

Contents (major groups)
- Browser
- Java (JRE/JDK)
- .NET runtime/SDK
- Windows Update
- ConnectSecure
- Dell (DDPM)
- Office
- Python
- VC++ redistributables
- Zoom
- winget helper
- Security cleanup
- Collection utilities

Notes on style
- Each script includes: purpose, recommended execution context (SYSTEM, Administrator, or user), key Datto RMM variables, and exit codes.

One-line summaries
- Browser: detection, notification, and remediation helpers for Chrome/Edge/Firefox. Detection runs as SYSTEM; user-facing reloads use InteractiveToken.
- Java: `Java-JRE-Update.ps1` and `Java-SDK-Update.ps1` manage Temurin/Oracle JRE/JDK (8/11/17/21). Administrator; prefer `winget` with MSI fallback. Use `-ReportOnly` to test.
- .NET: `Net-Runtime-Install-Latest-*NoRestart.ps1` enforces minimum versions, removes old ARP entries, and cleans folders. Administrator; use `-MinKeepVersion` and `-ReportOnly`.
- Windows Update: three patterns — no-auto-reboot (user prompts), auto-reboot (maintenance windows), and no-postpone (forced reboot after countdown). SYSTEM or scheduled-task setups are typical.
- ConnectSecure: `ConnectSecure-Agent-Install.ps1` and `ConnectSecure-Agent-Removal.ps1` install/remove the CyberCNS agent. Administrator; requires `CompanyID`, `TenantID`, `Secret` (install) and `cyberKey` (uninstall).
- Dell (DDPM): `DDPM-Fix.ps1` detects/remediates vulnerable Dell DDPM installs. Run Detect first, then Remediate with installer URL and SHA256 when needed.
- Office: `Office-Force-Update.ps1` triggers Click-to-Run updates. Can run as SYSTEM or user; `ForceCloseOfficeApps` is optional.
- Python: `Python-Update.ps1` reports and may update Python installs; requires Administrator for installs and is marked for further integration.
- VC++: `VC++-Install.ps1` enforces minimum redistributables per-architecture. Administrator; supports `-ReportOnly`.
- Zoom: `Zoom-Cleanup.ps1` removes outdated Zoom clients while preserving a minimum version. Administrator; supports `-ReportOnly`.
- winget: `winget-script.ps1` upgrades installed packages via Windows Package Manager. Supports `-ReportOnly` and optional winget auto-install.
- Security cleanup: `Remove-Microsoft-Silverlight.ps1` removes Silverlight and optional SDKs. Administrator; supports `-AggressiveCleanup` and `-ReportOnly`.
- Collection utilities: `Collect-Install-Logs.ps1` gathers MSI install/uninstall events for auditing and runs as non-admin (read-only event log access).

How to use
- Test first: run scripts with `-ReportOnly` where available.
- Respect execution context: follow each script's header (SYSTEM vs Administrator vs user).
- Logs: most scripts write to `%ProgramData%`; many expose a `LogPath` variable to override.

Examples
- Java update:

  .\Java\Java-JRE-Update.ps1 -TargetFamily 17 -RemoveOlder -Cleanup

- .NET report-only:

  .\.Net\Net-Runtime-Install-Latest-Including-Host-Force-Removal-NoRestart.ps1 -MinKeepVersion "8.0.11" -ReportOnly

Notes
- Windows update scripts were standardized to share headers/parameters/logging and explicit exit-code documentation. Use the `No-Auto-Reboot` script for interactive, user-consent workflows and the `Apply-Auto-Reboot` / `No-Postpone` scripts for maintenance windows where a reboot is acceptable.

If you'd like, I can:
- Produce a single-line index file mapping each script to its path.
- Expand any section into a short table of parameters and key exit codes.
- Trim this further to only include specific categories (e.g., Browser + Windows Update).

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

## Security Cleanup

### Remove-Microsoft-Silverlight.ps1

**Purpose**: Detects and removes Microsoft Silverlight from Windows endpoints, including optional Silverlight SDK components and residual folders or registry keys.

**Execution Context**: Must run with Administrator privileges

**Key Features**:
- Uses Add/Remove Programs inventory instead of `Win32_Product`
- Supports report-only inventory mode
- Optionally closes browser and Silverlight-related processes before uninstall
- Removes residual Silverlight folders and registry keys when enabled
- Can broaden matching to any Silverlight publisher entry when required
- Logs reboot-required uninstall results without forcing a restart

**Datto RMM Component Variables**:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ReportOnly` | switch | False | Inventory and log only |
| `RemoveSDK` | switch | True | Remove Silverlight SDK/developer components as well |
| `CloseProcesses` | switch | False | Close browsers and Silverlight-related processes before uninstall |
| `AggressiveCleanup` | switch | True | Remove leftover folders and registry keys after uninstall |
| `MatchAnyPublisher` | switch | False | Match any Silverlight ARP entry, not only Microsoft entries |
| `FailIfRemaining` | switch | True | Exit 3 if Silverlight still appears after remediation |
| `LogPath` | string | $env:ProgramData\NexusOpenSystems\Silverlight\Remove-Silverlight.log | Log file location |

**Exit Codes**:
- `0` - Success, no action required, report-only mode, or remediation completed
- `3` - Error or Silverlight still present after remediation when `FailIfRemaining` is enabled

**Example RMM Setup**:

```powershell
.\Remove-Microsoft-Silverlight.ps1 -CloseProcesses -AggressiveCleanup
```

**Typical Use Case**: Decommission Silverlight from managed endpoints as part of security hardening

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
