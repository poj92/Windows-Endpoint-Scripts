# Browser Reload for Pending Updates - Deployment Guide

## Overview

The `Force-BrowserReload-PendingUpdates.ps1` script monitors browser usage and automatically reloads browsers when:
- The browser has not been opened in **72 hours**
- The browser has a **pending update** waiting to be applied
- Users receive a **5-minute countdown warning** before automatic reload

This script is designed for scheduled deployment through Datto RMM or other endpoint management systems.

## Features

### Intelligent Update Detection
- **Chrome**: Checks GoogleUpdate registry keys, detects `new_chrome.exe`, and monitors version folders
- **Firefox**: Examines `active-update.xml` and updated folders
- **Edge**: Checks EdgeUpdate registry and detects `new_msedge.exe`

### Usage Tracking
- Maintains JSON tracking file at `C:\ProgramData\Datto\BrowserUpdateCheck\BrowserUsageTracking.json`
- Updates timestamps when browsers are detected running
- Tracks last opened time for each browser independently

### User-Friendly Notifications
- **Countdown Dialog**: Shows graphical 5-minute countdown timer
- **Cancellation Option**: Users can cancel the reload from the dialog
- **SYSTEM Context Support**: Works when run as SYSTEM via scheduled tasks
- **Session Detection**: Automatically finds active user sessions

### Comprehensive Logging
- File logging: `C:\ProgramData\Datto\BrowserUpdateCheck\ForceBrowserReload.log`
- Windows Event Log: Custom source "ForceBrowserReload" in "Datto RMM" log
- Console output for Datto monitoring

## How It Works

### Workflow

1. **Update Usage Tracking**
   - Script checks if browsers are currently running
   - Updates last opened timestamp for running browsers

2. **Check Inactivity**
   - Loads usage tracking JSON
   - Calculates hours since last browser use
   - Skips browsers opened within last 72 hours

3. **Detect Pending Updates**
   - Checks registry keys for update flags
   - Looks for staged update files (new_chrome.exe, new_msedge.exe)
   - Examines update XML files (Firefox)

4. **Show Countdown Warning**
   - Creates graphical countdown dialog via scheduled task
   - Displays 5-minute timer with cancel option
   - Waits for user response or timeout

5. **Reload Browser**
   - Forcibly closes browser processes
   - Restarts browser in minimized mode
   - Updates tracking timestamp to prevent re-reload

## Deployment

### Datto RMM Component

#### Recommended Schedule
- **Frequency**: Daily at off-peak hours (e.g., 7:00 AM)
- **Alternative**: Multiple times per day (e.g., 7 AM, 1 PM, 7 PM)

#### Component Configuration

**Name**: Browser Reload - Pending Updates (72hr Inactive)

**Component Type**: PowerShell Script

**Script Content**: Copy entire contents of `Force-BrowserReload-PendingUpdates.ps1`

**Execution Policy**: 
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
```

**Run As**: SYSTEM (recommended for cross-session compatibility)

**Timeout**: 600 seconds (10 minutes)

### Manual Deployment

#### One-Time Execution
```powershell
# Run with administrative privileges
Set-ExecutionPolicy Bypass -Scope Process -Force
.\Force-BrowserReload-PendingUpdates.ps1
```

#### Scheduled Task (Windows Task Scheduler)
```powershell
# Create scheduled task to run daily at 7:00 AM
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-ExecutionPolicy Bypass -File "C:\Scripts\Force-BrowserReload-PendingUpdates.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00AM
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName "Browser Reload - Pending Updates" -Action $action -Trigger $trigger -Principal $principal -Settings $settings
```

## Configuration Options

### Adjustable Parameters

Edit these variables at the top of the script to customize behavior:

```powershell
# Inactivity threshold (default: 72 hours)
$InactivityThresholdHours = 72

# Warning countdown time (default: 300 seconds / 5 minutes)
$WarningTimeSeconds = 300

# Log file location
$LogPath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = "$LogPath\ForceBrowserReload.log"
```

### Examples

**Reduce inactivity threshold to 48 hours:**
```powershell
$InactivityThresholdHours = 48
```

**Extend warning time to 10 minutes:**
```powershell
$WarningTimeSeconds = 600
```

**Reduce warning time to 2 minutes:**
```powershell
$WarningTimeSeconds = 120
```

## Monitoring & Verification

### Check Logs

**View recent log entries:**
```powershell
Get-Content "C:\ProgramData\Datto\BrowserUpdateCheck\ForceBrowserReload.log" -Tail 50
```

**Filter for specific browser:**
```powershell
Get-Content "C:\ProgramData\Datto\BrowserUpdateCheck\ForceBrowserReload.log" | Select-String "Chrome"
```

### Check Usage Tracking

**View tracking data:**
```powershell
Get-Content "C:\ProgramData\Datto\BrowserUpdateCheck\BrowserUsageTracking.json" | ConvertFrom-Json
```

### Check Windows Event Log

```powershell
Get-EventLog -LogName "Datto RMM" -Source "ForceBrowserReload" -Newest 20
```

### Verify Pending Updates

**Chrome:**
```powershell
# Check Chrome version
Get-ItemProperty "HKLM:\SOFTWARE\Google\Chrome\BLBeacon" -Name version

# Check for pending update
Test-Path "C:\Program Files\Google\Chrome\Application\new_chrome.exe"
```

**Firefox:**
```powershell
# Check for update staging folder
Test-Path "C:\Program Files\Mozilla Firefox\updated"
```

**Edge:**
```powershell
# Check Edge version
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon" -Name version

# Check for pending update
Test-Path "C:\Program Files\Microsoft\Edge\Application\new_msedge.exe"
```

## Best Practices

### Deployment Strategy

1. **Pilot Group**: Deploy to small test group first (10-20 endpoints)
2. **Monitor Results**: Review logs for 1-2 weeks
3. **Adjust Timing**: Optimize schedule based on user feedback
4. **Full Rollout**: Deploy to all endpoints

### Combining with Force-BrowserUpdates.ps1

For comprehensive browser update management:

1. **Force-BrowserReload-PendingUpdates.ps1**: Run daily to catch inactive browsers
2. **Force-BrowserUpdates.ps1**: Run weekly/monthly to force updates for active browsers

Recommended schedule:
- **Reload Script**: Daily at 7:00 AM
- **Force Update Script**: Monthly on patch Tuesday + 7 days

### User Communication

Consider notifying users about the automated browser maintenance:

- Send email explaining the 72-hour inactivity policy
- Document the 5-minute warning and cancellation option
- Provide instructions for monitoring browser update status

## Troubleshooting

### Script Doesn't Detect Browsers

**Issue**: Browsers shown as not installed

**Solution**: 
- Verify browser installation paths
- Check for 32-bit vs 64-bit installation locations
- Review log file for path detection errors

### Tracking File Issues

**Issue**: Browsers always shown as inactive

**Solution**:
```powershell
# Delete and recreate tracking file
Remove-Item "C:\ProgramData\Datto\BrowserUpdateCheck\BrowserUsageTracking.json" -Force
# Run script again to recreate
```

### Countdown Dialog Not Appearing

**Issue**: Users don't see countdown warning

**Possible Causes**:
- Script not running as SYSTEM
- No active user session
- User session not detected by `query user`

**Solution**:
```powershell
# Test session detection
query user

# Verify script is running as SYSTEM
whoami
```

### Pending Updates Not Detected

**Issue**: Updates exist but aren't detected

**Solution**:
- Manually verify update status using commands in Monitoring section
- Check registry paths for your browser version
- Review detection logic for your specific browser configuration

### Browser Doesn't Restart After Reload

**Issue**: Browser closes but doesn't reopen

**Possible Causes**:
- Incorrect executable path
- Browser installation corrupted
- Insufficient permissions

**Solution**:
- Test browser launch manually
- Verify browser installation integrity
- Check script execution privileges

## Exit Codes

- **0**: Success - Script completed normally
- **Non-zero**: Error occurred (check logs)

## Security Considerations

- Script runs with SYSTEM privileges for cross-session compatibility
- Temporary files created in `%TEMP%` are automatically cleaned up
- Scheduled tasks created for user notifications are removed after use
- No sensitive data stored in tracking or log files

## Additional Notes

### Inactive Browser Definition

A browser is considered "inactive" if:
- It has not been opened in the past 72 hours, OR
- No usage timestamp exists in the tracking file (first run)

### Update Application

Browser updates are applied when:
- Browser processes are terminated
- Browser restarts and detects staged update
- Update files are moved to active location

The script doesn't download updates - it only triggers the application of updates that browsers have already downloaded.

## Support

For issues or questions:
1. Check log files first
2. Review this documentation
3. Test individual functions in PowerShell ISE
4. Contact your IT administrator or Datto RMM support

## Version History

- **v1.0**: Initial release with Chrome, Firefox, and Edge support
