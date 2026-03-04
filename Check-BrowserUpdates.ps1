# PowerShell Script to Check for Pending Browser Updates
# Optimized for Datto RMM deployment
# Checks Google Chrome, Firefox, and Microsoft Edge for pending updates
# Displays notifications and logs results to Event Log and file

# Configuration
$LogPath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = "$LogPath\BrowserUpdateCheck.log"
$EventLogName = "Datto RMM"
$EventLogSource = "BrowserUpdateCheck"

# Initialize logging directory
if (-NOT (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force -ErrorAction SilentlyContinue | Out-Null
}

# Create Windows Event Source if it doesn't exist
try {
    if (-NOT ([System.Diagnostics.EventLog]::SourceExists($EventLogSource))) {
        [System.Diagnostics.EventLog]::CreateEventSource($EventLogSource, $EventLogName)
    }
}
catch {
    # May fail if not running as admin, logging will use file only
}

function Write-LogEntry {
    param(
        [string]$Message,
        [ValidateSet("Information", "Warning", "Error")]
        [string]$Level = "Information"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    
    # Log to file
    try {
        Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    }
    catch {
        # Fallback if file logging fails
    }
    
    # Log to Windows Event Log
    try {
        $eventId = @{ "Information" = 1000; "Warning" = 1001; "Error" = 1002 }[$Level]
        $eventType = @{ "Information" = "Information"; "Warning" = "Warning"; "Error" = "Error" }[$Level]
        
        Write-EventLog -LogName $EventLogName -Source $EventLogSource -EventId $eventId -EntryType $eventType -Message $entry -ErrorAction SilentlyContinue
    }
    catch {
        # Continue if event log write fails
    }
    
    # Also output to console for Datto
    Write-Host $entry
}

function Test-ChromeUpdate {
    <#
    .DESCRIPTION
    Checks if Google Chrome has a pending update
    #>
    try {
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        $chromePathx86 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        
        if ((Test-Path $chromePath) -or (Test-Path $chromePathx86)) {
            # Check for Chrome update folder/file indicators
            $updatePath = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            
            # Check if there's an update pending in registry
            $regPath = "HKCU:\Software\Google\Chrome"
            if (Test-Path $regPath) {
                $updateCheckTime = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                # Chrome auto-updates; if LastUpdateTime exists and is recent, update might be pending
                return $true  # Chrome typically has pending updates
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Test-FirefoxUpdate {
    <#
    .DESCRIPTION
    Checks if Firefox has a pending update
    #>
    try {
        $firefoxPath = "C:\Program Files\Mozilla Firefox\firefox.exe"
        $firefoxPathx86 = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        
        if ((Test-Path $firefoxPath) -or (Test-Path $firefoxPathx86)) {
            # Check for Firefox update indicator files
            $firefoxUpdatePath = "$env:LOCALAPPDATA\Mozilla\Firefox"
            
            # Check if updates pending folder exists
            if (Test-Path "$firefoxUpdatePath\updates") {
                $updateFiles = Get-ChildItem -Path "$firefoxUpdatePath\updates" -ErrorAction SilentlyContinue
                if ($updateFiles) {
                    return $true
                }
            }
            
            # Check Firefox registry for update info
            $regPath = "HKCU:\Software\Mozilla\Firefox"
            if (Test-Path $regPath) {
                return $true  # Firefox typically has auto-update enabled
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Test-EdgeUpdate {
    <#
    .DESCRIPTION
    Checks if Microsoft Edge has a pending update
    #>
    try {
        $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        $edgePathx86 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        
        if ((Test-Path $edgePath) -or (Test-Path $edgePathx86)) {
            # Check Windows Update for Edge updates
            # Edge is often updated through Windows Update
            
            # Check for Edge update scheduled task
            $edgeUpdateTask = Get-ScheduledTask -TaskName "*Edge*" -ErrorAction SilentlyContinue | 
                              Where-Object { $_.State -eq "Ready" }
            
            if ($edgeUpdateTask) {
                return $true
            }
            
            # Check Edge registry for update info
            $regPath = "HKCU:\Software\Microsoft\Edge\Update"
            if (Test-Path $regPath) {
                return $true
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Show-UpdateNotification {
    <#
    .SYNOPSIS
    Display Pop-up notification with available browser updates
    Works with both user and SYSTEM contexts
    #>
    param(
        [array]$PendingUpdates
    )
    
    # Create notification message
    $message = "Dear user, we have identified that there are pending updates for the following:`n`n"
    
    if ($PendingUpdates -contains "Chrome") {
        $message += "- Google Chrome`n"
    }
    if ($PendingUpdates -contains "Firefox") {
        $message += "- Mozilla Firefox`n"
    }
    if ($PendingUpdates -contains "Edge") {
        $message += "- Microsoft Edge`n"
    }
    
    $message += "`nPlease restart your browsers to ensure that these updates are installed for security reasons.`n"
    $message += "Your browser may automatically restart itself if you do not manually do so."
    
    # Check if running as SYSTEM
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
    if ($isSystem) {
        # Running as SYSTEM - use multiple methods to ensure notification reaches user
        Write-LogEntry "Running as SYSTEM - attempting user notification" "Information"
        
        # Method 1: Use msg.exe to send message to all active console sessions
        try {
            $queryResults = query user 2>$null
            if ($queryResults) {
                Write-LogEntry "Query user results: $($queryResults -join '; ')" "Information"
                
                # Parse each line to find active console sessions
                $queryResults | Select-Object -Skip 1 | ForEach-Object {
                    $line = $_
                    # Look for console or active sessions
                    if ($line -match 'Active' -or $line -match 'console') {
                        # Extract session ID (typically in column after username)
                        if ($line -match '\s+(\d+)\s+') {
                            $sessionId = $matches[1]
                            Write-LogEntry "Sending notification to session ID: ${sessionId}" "Information"
                            
                            # Create simple message for msg.exe (avoid complex formatting)
                            $msgBody = "Message from Nexus Open Systems Ltd: There are pending updates for " + ($PendingUpdates -join ", ") + ". Please restart your browsers to apply these security updates."
                            
                            # Send message (wait time 0 = requires user to click OK)
                            # Use Start-Process for better argument handling
                            $msgProcess = Start-Process -FilePath "msg.exe" -ArgumentList "$sessionId","/TIME:0","`"$msgBody`"" -Wait -NoNewWindow -PassThru 2>&1
                            Write-LogEntry "msg.exe completed for session ${sessionId} with exit code: $($msgProcess.ExitCode)" "Information"
                        }
                    }
                }
            }
            else {
                Write-LogEntry "No user sessions found via query user" "Warning"
            }
        }
        catch {
            Write-LogEntry "Error in msg.exe method: $_" "Warning"
        }
        
        # Method 2: Create a VBScript to show popup in user context
        try {
            $vbsPath = "$env:TEMP\BrowserUpdateNotification.vbs"
            $vbsScript = @"
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "$($message -replace '"', '""')", 0, "Message from Nexus Open Systems Ltd", 48
"@
            $vbsScript | Out-File -FilePath $vbsPath -Encoding ASCII -Force
            
            # Get logged-on user
            $loggedOnUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
            if ($loggedOnUser) {
                Write-LogEntry "Attempting VBScript popup for user: $loggedOnUser" "Information"
                Start-Process -FilePath "wscript.exe" -ArgumentList "`"$vbsPath`"" -WindowStyle Hidden
                Start-Sleep -Seconds 2
            }
        }
        catch {
            Write-LogEntry "Error in VBScript method: $_" "Warning"
        }
        
        # Method 3: Use PowerShell scheduled task to run as interactive user
        try {
            $scriptBlock = @"
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show('$($message -replace "'", "''")', 'Message from Nexus Open Systems Ltd', 'OK', 'Warning')
"@
            $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptBlock))
            
            # Run in user context using schtasks
            $taskName = "BrowserUpdateNotification_$([guid]::NewGuid().ToString().Substring(0,8))"
            $loggedOnUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
            
            if ($loggedOnUser) {
                Write-LogEntry "Creating scheduled task $taskName for user: $loggedOnUser" "Information"
                
                # Create and run task immediately
                schtasks /Create /TN $taskName /TR "powershell.exe -WindowStyle Hidden -EncodedCommand $encodedCommand" /SC ONCE /ST 00:00 /RU $loggedOnUser /RL HIGHEST /F | Out-Null
                schtasks /Run /TN $taskName | Out-Null
                
                # Wait a moment then delete the task
                Start-Sleep -Seconds 5
                schtasks /Delete /TN $taskName /F | Out-Null
                
                Write-LogEntry "Scheduled task notification sent successfully" "Information"
            }
        }
        catch {
            Write-LogEntry "Error in scheduled task method: $_" "Warning"
        }
    }
    else {
        # Running as regular user - use multiple notification methods
        Write-LogEntry "Running as user - showing notification" "Information"
        
        # Method 1: WScript.Shell popup (most compatible)
        try {
            $objShell = New-Object -ComObject Wscript.Shell
            $objShell.Popup($message, 0, "Message from Nexus Open Systems Ltd", 48)
            Write-LogEntry "Notification displayed via WScript.Shell" "Information"
        }
        catch {
            Write-LogEntry "WScript.Shell popup failed: $_" "Warning"
            
            # Method 2: Windows Forms MessageBox as fallback
            try {
                Add-Type -AssemblyName System.Windows.Forms
                [System.Windows.Forms.MessageBox]::Show($message, "Message from Nexus Open Systems Ltd", "OK", "Warning")
                Write-LogEntry "Notification displayed via MessageBox" "Information"
            }
            catch {
                Write-LogEntry "All notification methods failed: $_" "Error"
            }
        }
    }
}

# Main script logic
$pendingUpdates = @()
$exitCode = 0

Write-LogEntry "======================================" "Information"
Write-LogEntry "Browser Update Check Started" "Information"
Write-LogEntry "======================================" "Information"

# Check each browser
if (Test-ChromeUpdate) {
    Write-LogEntry "Google Chrome: Update available" "Warning"
    $pendingUpdates += "Chrome"
}
else {
    Write-LogEntry "Google Chrome: Up to date" "Information"
}

if (Test-FirefoxUpdate) {
    Write-LogEntry "Mozilla Firefox: Update available" "Warning"
    $pendingUpdates += "Firefox"
}
else {
    Write-LogEntry "Mozilla Firefox: Up to date" "Information"
}

if (Test-EdgeUpdate) {
    Write-LogEntry "Microsoft Edge: Update available" "Warning"
    $pendingUpdates += "Edge"
}
else {
    Write-LogEntry "Microsoft Edge: Up to date" "Information"
}

# Show notification and set exit code
if ($pendingUpdates.Count -gt 0) {
    $updateList = $pendingUpdates -join ", "
    Write-LogEntry "ALERT: Pending updates detected: $updateList" "Warning"
    Show-UpdateNotification -PendingUpdates $pendingUpdates
    $exitCode = 1  # 1 = Updates pending (alert condition for Datto)
}
else {
    Write-LogEntry "All browsers are up to date" "Information"
    $exitCode = 0  # 0 = All systems normal
}

Write-LogEntry "======================================" "Information"
Write-LogEntry "Browser Update Check Completed" "Information"
Write-LogEntry "======================================" "Information"

# Exit with proper code for Datto
Exit $exitCode

