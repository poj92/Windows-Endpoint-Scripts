#Requires -Version 5.1
<#
Wrapper: Windows-Update-No-Auto-Reboot (Ensure Latest Build)

Purpose:
- Thin wrapper that invokes `Windows-Update-No-Auto-Reboot.ps1` while forcing
	`EnsureLatestCumulativeUpdate` (SSU/LCU prioritization). Keeps the same
	execution/parameter/logging pattern as the primary script so component
	variables and RMM usage remain consistent.

Execution Context: Can run as SYSTEM (uses scheduled task to show UI in active session)

Behavior:
- Forwards parameters to the main `Windows-Update-No-Auto-Reboot.ps1` script
- Ensures `EnsureLatestCumulativeUpdate` is enabled

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[ValidateSet('Normal','PostBoot','PromptOnly')]
	[string]$Mode = 'Normal',

	[int]$CountdownMinutes = 10,
	[int[]]$PostponeOptionsMinutes = @(30, 60, 120),

	[switch]$IncludeRebootUpdates,

	# Enforced by this wrapper
	[switch]$EnsureLatestCumulativeUpdate = $true,

	[switch]$ReportOnly,

	[int]$PostRebootMaxPasses = 4,

	[string]$UiTitle = "A security message from Nexus Open Systems Ltd",
	[string]$Reason  = "Windows updates require a restart to finish installing. Please plug your computer into power if it's not already, save your work, and restart as soon as possible to ensure your system is secure and up to date.",

	[string]$BaseDir = "$env:ProgramData\NexusOpenSystems\WindowsUpdate",
	[string]$LogPath = "$env:ProgramData\NexusOpenSystems\WindowsUpdate\WindowsUpdateReboot.log"
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

function Write-Log([string]$Message) {
	$ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
	$line = "[{0}] {1}" -f $ts, $Message
	Write-Host $line
	try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

# Ensure main script exists next to this wrapper
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$mainScript = Join-Path $scriptDir 'Windows-Update-No-Auto-Reboot.ps1'

if (-not (Test-Path $mainScript)) {
	Write-Log "ERROR: Main script not found at $mainScript"
	Write-Host "Main script not found: $mainScript"
	exit 3
}

# Build splatted parameter set and force EnsureLatestCumulativeUpdate
$params = @{
	Mode = $Mode
	CountdownMinutes = $CountdownMinutes
	PostponeOptionsMinutes = $PostponeOptionsMinutes
	IncludeRebootUpdates = $IncludeRebootUpdates
	EnsureLatestCumulativeUpdate = $true
	ReportOnly = $ReportOnly
	PostRebootMaxPasses = $PostRebootMaxPasses
	UiTitle = $UiTitle
	Reason = $Reason
	BaseDir = $BaseDir
	LogPath = $LogPath
}

Write-Log "Invoking $mainScript with EnsureLatestCumulativeUpdate=$($params.EnsureLatestCumulativeUpdate)"
& $mainScript @params
$rc = $LASTEXITCODE
Write-Log "Main script exited with code $rc"
exit $rc

