# Windows-Endpoint-Scripts

This repository contains PowerShell scripts for Windows endpoint management, optimized for deployment through Datto RMM and other endpoint management platforms.

## Scripts

### Browser-Scripts

- `Browser-Update-Detection.ps1` - Detects Chrome, Firefox, and Edge sessions that have been running long enough to warrant a restart and tracks usage state.
- `Browser-Force-Restart.ps1` - Forces a browser restart workflow for pending updates, including queue handling, logging, and scheduled task cleanup.

### Java

- `Java-JRE-Update.ps1` - Detects, reports, installs, or upgrades the target Java JRE family using winget, with optional cleanup and old-version removal.
- `Java-SDK-Update.ps1` - Detects, reports, installs, or upgrades the target Java JDK family using winget, with optional cleanup and old-version removal.

## Usage

Run scripts in an elevated PowerShell session on the target Windows device.

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force

.\Browser-Scripts\Browser-Update-Detection.ps1
.\Browser-Scripts\Browser-Force-Restart.ps1

.\Java\Java-JRE-Update.ps1 -TargetFamily 17
.\Java\Java-SDK-Update.ps1 -TargetFamily 21
```

Useful Java script options:

- `-ReportOnly` checks whether an update is needed without making changes.
- `-RemoveOlder` removes older installs in the same major family.
- `-Cleanup` removes stale Java environment variables and PATH entries.

## Requirements

- Windows device with PowerShell 5.1+
- Administrative privileges
- `winget` for Java update/install workflows
