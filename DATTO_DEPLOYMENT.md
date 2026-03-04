# Browser Update Check - Datto RMM Deployment Guide

This guide explains how to deploy the Browser Update Check scripts in Datto RMM.

## Files Included

- **Check-BrowserUpdates.ps1** - Main script that checks for pending browser updates
- **Setup-BrowserUpdateTask.ps1** - Preparation/validation script (optional, for testing)

## Features

✓ Checks Google Chrome, Firefox, and Microsoft Edge for pending updates  
✓ Displays pop-up notifications to logged-in users  
✓ Logs results to Windows Event Log  
✓ Logs results to local log file  
✓ Proper exit codes for Datto alerting  
✓ Runs as SYSTEM for comprehensive coverage  

## Quick Start

### 1. Prepare Environment (Optional)

Run the setup script on the test machine first:

```powershell
.\Setup-BrowserUpdateTask.ps1
```

This validates the main script and creates necessary directories.

### 2. Create Custom Component in Datto RMM

#### In Datto RMM Dashboard:

1. Navigate to: **Automation** → **Custom Components** → **Create New**

2. Configure Component Details:
   - **Name:** Browser Update Check
   - **Description:** Checks for pending updates on Chrome, Firefox, and Edge
   - **Script Type:** PowerShell
   - **Execution Context:** SYSTEM (important)
   - **Timeout:** 60 seconds

3. **Paste Script Content:**
   - Copy the entire contents of `Check-BrowserUpdates.ps1`
   - Paste into the script field in Datto

4. **Set Schedule:**
   - **Frequency:** Hourly, Every 6 Hours, Daily, or Weekly (based on your preference)
   - **Run on:** All devices or specific device groups

5. **Configure Alerts:**
   - Exit Code 0 = No action (all up to date) - No alert
   - Exit Code 1 = Updates pending - Alert/Notify

6. **Save Component**

### 3. Deploy to Devices

Once created, you can:
- Deploy immediately to all devices
- Deploy to specific device groups
- Schedule deployment for a specific time

## Monitoring Results

### Event Log
- **Log Name:** Datto RMM
- **Source:** BrowserUpdateCheck
- **View in Event Viewer:** Applications and Services Logs → Datto RMM

### Log Files
- **Location:** `C:\ProgramData\Datto\BrowserUpdateCheck\BrowserUpdateCheck.log`
- Text file with timestamps for each check
- Readable from Datto file browse feature

### Datto Dashboard
- Component execution history shows exit codes
- Set up alerts for exit code 1 (updates detected)
- Schedule notifications to your team

## Exit Codes

| Code | Status | Meaning |
|------|--------|---------|
| 0 | Success | All browsers are up to date |
| 1 | Alert | One or more browsers have pending updates |

## Customization

### Change Log Location
In `Check-BrowserUpdates.ps1`, modify:
```powershell
$LogPath = "C:\ProgramData\Datto\BrowserUpdateCheck"
```

### Change Event Log Name
Modify:
```powershell
$EventLogName = "Datto RMM"
$EventLogSource = "BrowserUpdateCheck"
```

### Add More Browsers
Edit the test functions (`Test-ChromeUpdate`, `Test-FirefoxUpdate`, `Test-EdgeUpdate`) or add new ones.

## Troubleshooting

### Script Not Running
- Verify execution context is set to SYSTEM in Datto
- Check Windows PowerShell execution policy: `Get-ExecutionPolicy`
- Run test on single device first

### No Pop-up Notifications Appearing
- Pop-ups only show to active user sessions
- Use `msg.exe` command from Windows at terminal if testing in RDP

### Missing Log Files
- Verify `C:\ProgramData\Datto\` directory exists
- Check SYSTEM account has write permissions
- Review Datto component execution logs for errors

## Advanced: Separate Scheduled Task Deployment

If you prefer to use Windows Scheduled Tasks instead of Datto scheduling:

1. Run `Setup-BrowserUpdateTask.ps1` as Administrator
2. This creates a scheduled task that:
   - Runs every 6 hours automatically
   - Also runs when users log on
   - Executes as SYSTEM
   - Same logging to Event Log and files

To manage the task:
```powershell
# View task
Get-ScheduledTask -TaskName "BrowserUpdateCheck"

# Run manually
Start-ScheduledTask -TaskName "BrowserUpdateCheck"

# Remove task
Unregister-ScheduledTask -TaskName "BrowserUpdateCheck" -Confirm:$false
```

## Support Notes

- Script handles both 32-bit and 64-bit installations
- Works on Windows Server 2016+, Windows 10/11
- No external dependencies required
- Safe to run multiple times on same system

## Version History

**v1.0** - Initial Datto RMM optimized release
- Event Log + File logging
- Proper exit codes
- User session notifications
