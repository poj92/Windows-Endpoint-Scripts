# Force Browser Restart on Update - Datto RMM Deployment Guide

This script automatically restarts browsers when BOTH conditions are met:
1. Updates are pending
2. Browser has not been used for 72+ hours

## Features

✓ Checks for pending browser updates  
✓ Detects last execution time from process history  
✓ Only restarts if both conditions met (safety measure)  
✓ Shows 5-minute countdown before restart  
✓ Users can postpone restart (24-hour grace period)  
✓ Preserves browser tabs on restart (Chrome/Firefox/Edge auto-restore)  
✓ Logs to Windows Event Log and file  
✓ Proper exit codes for Datto alerting  
✓ Runs as SYSTEM for comprehensive coverage  

## How It Works

### Detection Logic

1. **Check for Updates**: Detects pending updates on Chrome, Firefox, Edge
2. **Check Last Run Time**: Checks:
   - Running process creation time
   - Recent files accessed by browser
   - Windows Security event logs
3. **Conditions for Restart**:
   - Update pending = YES
   - Last used > 72 hours ago = YES
   - User postponed within 24 hours = NO
   → If all true, trigger restart sequence

### Postponement System

- Users can postpone restart notifications
- Postponement lasts 24 hours per browser
- After 24 hours, notification appears again if conditions still met
- Stored in registry: `HKCU:\Software\Datto\BrowserUpdates`

### Restart Behavior

**Timer Sequence:**
1. Initial notification with 5-minute countdown
2. Displays countdown at each minute mark
3. After 5 minutes, browser is force-closed
4. Browser automatically restarts
5. Tabs are restored (built-in browser feature)

## Datto RMM Deployment

### 1. Create Custom Component in Datto RMM

1. Navigate to: **Automation** → **Custom Components** → **Create New**

2. Configure Component Details:
   - **Name:** Force Browser Restart on Update
   - **Description:** Restarts browsers with pending updates if idle 72+ hours
   - **Script Type:** PowerShell
   - **Execution Context:** SYSTEM (required)
   - **Timeout:** 600 seconds (10 minutes - allows full countdown + restart)

3. **Paste Script Content:**
   - Copy entire contents of `Force-BrowserRestartOnUpdate.ps1`
   - Paste into the script field in Datto

4. **Set Schedule:**
   - **Frequency:** Daily or Every 12 Hours (check twice daily)
   - **Running Time:** Off-peak hours (early morning recommended)
   - **Run on:** All devices or specific device groups

5. **Configure Alerts:**
   - Exit Code 0 = No action needed (all browsers up to date or recently used)
   - Exit Code 1 = Restart performed (update + 72+ hour condition met)

6. **Save Component**

### 2. Optional: Configure Advanced Settings

```powershell
# Change the idle threshold (default 72 hours)
$HoursSinceLastUse = 72

# Change countdown timer (default 5 minutes)
$CountdownMinutes = 5

# Change postponement window (default 24 hours)
# Edit in Get-PostponeStatus function
```

### 3. Deploy to Devices

Once created, deploy to:
- All endpoints
- Specific device groups
- Schedule for off-peak hours

## User Experience

### Notification Sequence

**Minute 5 remaining:**
```
BROWSER UPDATES REQUIRED

Pending updates detected for: Chrome, Firefox

These browsers have not been restarted in over 72 hours.

Your browsers will automatically restart in 5 minute(s) to install the updates.

Save your work now.
```

**Minute 4, 3, 2, 1 remaining:**
- Log entry shows countdown progression
- User has time to save work

**After countdown expires:**
- Browser is force-closed
- Browser restarts automatically
- Tabs are restored

## Monitoring Results

### Event Log
- **Log Name:** Datto RMM
- **Source:** BrowserRestartCheck
- **View in Event Viewer:** Applications and Services Logs → Datto RMM

### Log Files
- **Location:** `C:\ProgramData\Datto\BrowserRestartCheck\BrowserRestartCheck.log`
- Contains timeline of checks and actions taken

### Datto Dashboard
- Component execution history
- Exit code tracking
- Output messages show which browsers were restarted

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | No restart needed (updates not pending, or browser recently used, or user postponed) |
| 1 | Restart performed (both conditions met and countdown executed) |

## Troubleshooting

### Notification Not Appearing

**Issue:** Users don't see countdown notification

**Solutions:**
- Verify script running as SYSTEM in Datto
- Check Windows Event Log for errors
- Test on single device first
- Ensure users are actively logged in (notifications go to active sessions)

### Browser Not Restarting

**Issue:** Countdown completes but browser doesn't restart

**Solutions:**
- Check if browser executable path exists (verify installation)
- Check `BrowserRestartCheck.log` for error messages
- Verify browser is not running as elevated/protected process
- Manual test: `taskkill /IM chrome.exe /F` then `start chrome.exe`

### Postponement Not Working

**Issue:** Postponement system not registering

**Solutions:**
- Registry path may not exist: `HKCU:\Software\Datto\BrowserUpdates`
- Script may not have permission to create registry entries
- Check registry manually: `reg query "HKCU\Software\Datto\BrowserUpdates"`

### Last Run Time Not Accurate

**Issue:** Script thinks browser is idle when it was recently used

**Solutions:**
- Check Security event log: `Get-WinEvent -LogName Security | Where EventID -eq 4688 | head -20`
- Recent items may be limited by Windows policy
- Consider also checking browser history directory

## Advanced Configuration

### Change Update Detection

Edit the `Test-*Update` functions to add more sophisticated detection:

```powershell
# Example: Check specific registry keys
$regPath = "HKCU:\Software\Google\Chrome\Update"
$updateCheckTime = Get-ItemProperty -Path $regPath -Name "UpdateCheckTime"
```

### Change Restart Delay

Modify the countdown loop iteration:

```powershell
# Default is 60 seconds per minute
# To check more frequently (every 30 seconds):
Start-Sleep -Seconds 30
```

### Change Postponement Duration

In `Get-PostponeStatus` function:

```powershell
# Change from 24 hours to 48 hours:
if ($hoursSincePostpone -lt 48) {
    return $true
}
```

## FAQ

**Q: Will this restart a browser if the user is currently using it?**
A: Yes. The script shows a 5-minute countdown to save work. If browser is still running at end of countdown, it's force-closed and restarted.

**Q: What happens to open tabs?**
A: All major browsers (Chrome, Firefox, Edge) have built-in session recovery. When the browser restarts, it automatically reopens previously open tabs.

**Q: Can users disable this?**
A: Only through postponement (24-hour window). For permanent disable, the script would need to be removed from Datto or the device excluded from component deployment.

**Q: How often should this run?**
A: Recommended daily or every 12 hours. Running too frequently wastes resources; running too infrequently means longer delays for updates.

**Q: What if the browser is frozen or unresponsive?**
A: The script uses force-close (`Stop-Process -Force`), which terminates frozen processes. On restart, browser session recovery kicks in.

## Deployment Checklist

- [ ] Script copied to all endpoints (or deployed via Datto)
- [ ] Custom component created in Datto
- [ ] Execution context set to SYSTEM
- [ ] Schedule configured (daily/every 12 hours)
- [ ] Tested on pilot devices
- [ ] Users informed about expected behavior
- [ ] Monitoring configured for exit codes
- [ ] Event Log alerts setup in Datto

## Version History

**v1.0** - Initial Datto RMM optimized release
- Process execution time detection
- 5-minute countdown with user options
- Postponement system (24-hour window)
- Tab preservation on restart
- Event Log + File logging
- Proper exit codes for alerting
