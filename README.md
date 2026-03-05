# Windows-Endpoint-Scripts

This repository contains PowerShell scripts for Windows endpoint management, optimized for deployment through Datto RMM and other endpoint management platforms.

## Scripts

### Browser Update Management

#### Check-BrowserUpdates.ps1
Checks for available browser updates for Google Chrome, Mozilla Firefox, and Microsoft Edge. Provides detailed reporting without forcing updates.

**See**: [Check Browser Updates Documentation](./README.md#check-browser-updates)

#### Force-BrowserUpdates.ps1
Forces immediate browser updates with user notification and deferral options. Closes and restarts browsers to apply pending updates.

**See**: [Force Updates Deployment Guide](./FORCE_RESTART_DEPLOYMENT.md) | [Datto Deployment Guide](./DATTO_DEPLOYMENT.md)

**Features**:
- Immediate update application
- User deferral option
- Pre/post-update notifications
- Comprehensive logging

#### Force-BrowserReload-PendingUpdates.ps1
Monitors browser usage and automatically reloads browsers that haven't been opened in 72 hours and have pending updates. Provides a 5-minute countdown warning.

**See**: [Browser Reload Deployment Guide](./BROWSER_RELOAD_DEPLOYMENT.md)

**Features**:
- 72-hour inactivity threshold
- Pending update detection
- 5-minute countdown warning
- User cancellation option
- Usage tracking via JSON

## Quick Start

### Using with Datto RMM

1. Create a new PowerShell component
2. Copy script contents
3. Set execution policy: `Set-ExecutionPolicy Bypass -Scope Process -Force`
4. Configure schedule (see deployment guides)
5. Deploy to target sites/devices

### Manual Execution

```powershell
# Run with administrative privileges
Set-ExecutionPolicy Bypass -Scope Process -Force
.\scriptname.ps1
```

## Deployment Guides

- **[BROWSER_RELOAD_DEPLOYMENT.md](./BROWSER_RELOAD_DEPLOYMENT.md)** - Browser reload for pending updates
- **[DATTO_DEPLOYMENT.md](./DATTO_DEPLOYMENT.md)** - General Datto RMM deployment instructions
- **[FORCE_RESTART_DEPLOYMENT.md](./FORCE_RESTART_DEPLOYMENT.md)** - Force browser updates deployment

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Administrative privileges
- Supported browsers: Chrome, Firefox, Edge

## Logging

All scripts maintain comprehensive logging:
- **Location**: `C:\ProgramData\Datto\BrowserUpdateCheck\`
- **Windows Event Log**: Custom source in "Datto RMM" log
- **Console Output**: For RMM monitoring

## Support

For issues or questions, please review the relevant deployment guide or check the script logs at `C:\ProgramData\Datto\BrowserUpdateCheck\`.
