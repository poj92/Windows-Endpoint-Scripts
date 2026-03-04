# PowerShell Script to Force Browser Updates
# Optimized for Datto RMM deployment
# Forces updates for Google Chrome, Firefox, and Microsoft Edge
# Notifies users before and after applying updates

# Configuration
$LogPath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = "$LogPath\ForceBrowserUpdates.log"
$EventLogName = "Datto RMM"
$EventLogSource = "ForceBrowserUpdates"

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
        $eventId = @{ "Information" = 2000; "Warning" = 2001; "Error" = 2002 }[$Level]
        $eventType = @{ "Information" = "Information"; "Warning" = "Warning"; "Error" = "Error" }[$Level]
        
        Write-EventLog -LogName $EventLogName -Source $EventLogSource -EventId $eventId -EntryType $eventType -Message $entry -ErrorAction SilentlyContinue
    }
    catch {
        # Continue if event log write fails
    }
    
    # Also output to console for Datto
    Write-Host $entry
}

function Show-PreUpdateNotification {
    <#
    .SYNOPSIS
    Notify user that browser updates are about to be applied
    #>
    param(
        [array]$BrowsersToUpdate
    )
    
    # Check if running as SYSTEM
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
    if ($isSystem) {
        Write-LogEntry "Running as SYSTEM - notifying user of pending updates" "Information"
        
        try {
        function Request-UpdateDeferral {
            <#
            .SYNOPSIS
            Ask user if they want to defer the update
            Returns $true if user wants to defer, $false to proceed with update
            #>
    
            # Check if running as SYSTEM
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
            if ($isSystem) {
                Write-LogEntry "Running as SYSTEM - requesting user input for update deferral" "Information"
        
                try {
                    # Get active user sessions
                    $queryResults = query user 2>$null
                    if ($queryResults) {
                        # Parse to get logged-on user
                        $queryResults | Select-Object -Skip 1 | ForEach-Object {
                            $line = $_
                            if ($line -match 'Active' -or $line -match 'console') {
                                if ($line -match '\s+(\d+)\s+') {
                                    $sessionId = $matches[1]
                            
                                    # Create a marker file location
                                    $markerFile = "$env:TEMP\UpdateDeferred_$([guid]::NewGuid().ToString().Substring(0,8)).txt"
                            
                                    # Create PowerShell script that shows MessageBox with Yes/No buttons
                                    $psScriptPath = "$env:TEMP\DeferUpdatePrompt_$([guid]::NewGuid().ToString().Substring(0,8)).ps1"
                            
                                    $scriptContent = @"
        Add-Type -AssemblyName System.Windows.Forms
        `$result = [System.Windows.Forms.MessageBox]::Show(
            'Browser updates are ready to be applied. Do you want to defer this update to a later time?`n`nClick YES to defer or NO to apply updates now.',
            'Browser Updates Available',
            'YesNo',
            'Question'
        )
        if (`$result -eq 'Yes') {
            'deferred' | Out-File -FilePath '$markerFile' -Force
            Exit 1
        }
        else {
            Exit 0
        }
        "@
                                    $scriptContent | Out-File -FilePath $psScriptPath -Encoding UTF8 -Force
                            
                                    # Create scheduled task to run script in user context
                                    $taskName = "DeferUpdatePrompt_$([guid]::NewGuid().ToString().Substring(0,8))"
                            
                                    Write-LogEntry "Showing deferral prompt to user" "Information"
                            
                                    # Create and execute task
                                    $createCmd = "schtasks /Create /TN `"$taskName`" /TR `"powershell.exe -ExecutionPolicy Bypass -File `'$psScriptPath`'`" /SC ONCE /ST 00:00 /RU `"*`" /RL HIGHEST /IT /F"
                                    cmd.exe /c $createCmd 2>&1 | Out-Null
                            
                                    # Run task
                                    & cmd.exe /c "schtasks /Run /TN `"$taskName`"" 2>&1 | Out-Null
                            
                                    # Wait for user response
                                    Start-Sleep -Seconds 8
                            
                                    # Check if deferral marker file was created
                                    $deferred = $false
                                    if (Test-Path $markerFile) {
                                        Write-LogEntry "User chose to defer updates" "Information"
                                        $deferred = $true
                                        Remove-Item -Path $markerFile -Force -ErrorAction SilentlyContinue
                                    }
                                    else {
                                        Write-LogEntry "User chose to proceed with updates" "Information"
                                    }
                            
                                    # Clean up task
                                    $deleteCmd = "schtasks /Delete /TN `"$taskName`" /F"
                                    cmd.exe /c $deleteCmd 2>&1 | Out-Null
                            
                                    # Clean up script file
                                    Start-Sleep -Milliseconds 500
                                    Remove-Item -Path $psScriptPath -Force -ErrorAction SilentlyContinue
                            
                                    return $deferred
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-LogEntry "Error requesting update deferral: $_" "Warning"
                    return $false
                }
            }
            else {
                # Running as regular user - show MessageBox directly
                try {
                    Add-Type -AssemblyName System.Windows.Forms
                    $result = [System.Windows.Forms.MessageBox]::Show(
                        'Browser updates are ready to be applied. Do you want to defer this update to a later time?`n`nClick YES to defer or NO to apply updates now.',
                        'Browser Updates Available',
                        'YesNo',
                        'Question'
                    )
            
                    if ($result -eq 'Yes') {
                        Write-LogEntry "User chose to defer updates" "Information"
                        return $true
                    }
                    else {
                        Write-LogEntry "User chose to proceed with updates" "Information"
                        return $false
                    }
                }
                catch {
                    Write-LogEntry "Error requesting deferral in user context: $_" "Warning"
                    return $false
                }
            }
    
            return $false
        }

            # Get active user sessions
            $queryResults = query user 2>$null
            if ($queryResults) {
                # Parse each line to find active sessions
                $queryResults | Select-Object -Skip 1 | ForEach-Object {
                    $line = $_
                    # Look for active sessions or console
                    if ($line -match 'Active' -or $line -match 'console') {
                        # Extract session ID
                        if ($line -match '\s+(\d+)\s+') {
                            $sessionId = $matches[1]
                            Write-LogEntry "Sending pre-update notification to session ID: ${sessionId}" "Information"
                            
                            # Create temp file with notification
                            $msgFile = "$env:TEMP\BrowserPreUpdateMsg_$([guid]::NewGuid().ToString().Substring(0,8)).txt"
                            
                            # Build notification message
                            $messageContent = @"
Message from Nexus Open Systems Ltd

BROWSER UPDATES IN PROGRESS

The following browsers will now be updated to the latest version:

"@
                            
                            foreach ($browser in $BrowsersToUpdate) {
                                $messageContent += "- $browser`n"
                            }
                            
                            $messageContent += "`nPlease save any open work in these browsers before proceeding.`n`nThis may take a few minutes. Your browsers will be closed and automatically reopened after the updates are complete."
                            
                            # Write message to file
                            $messageContent | Out-File -FilePath $msgFile -Encoding ASCII -Force
                            
                            # Send notification using msg.exe
                            $cmdString = "msg.exe $sessionId /TIME:0 < `"$msgFile`""
                            cmd.exe /c $cmdString 2>&1 | Out-Null
                            
                            # Clean up
                            Start-Sleep -Milliseconds 500
                            Remove-Item -Path $msgFile -Force -ErrorAction SilentlyContinue
                            
                            Write-LogEntry "Pre-update notification sent to session ${sessionId}" "Information"
                        }
                    }
                }
            }
        }
        catch {
            Write-LogEntry "Error notifying user of pending updates: $_" "Warning"
        }
    }
}

function Show-PostUpdateNotification {
    <#
    .SYNOPSIS
    Notify user that browser updates have been applied
    #>
    param(
        [array]$UpdatedBrowsers,
        [array]$FailedUpdates
    )
    
    # Check if running as SYSTEM
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
    if ($isSystem) {
        Write-LogEntry "Running as SYSTEM - notifying user of completed updates" "Information"
        
        try {
            # Get active user sessions
            $queryResults = query user 2>$null
            if ($queryResults) {
                # Parse each line to find active sessions
                $queryResults | Select-Object -Skip 1 | ForEach-Object {
                    $line = $_
                    # Look for active sessions or console
                    if ($line -match 'Active' -or $line -match 'console') {
                        # Extract session ID
                        if ($line -match '\s+(\d+)\s+') {
                            $sessionId = $matches[1]
                            
                            # Create temp file with notification
                            $msgFile = "$env:TEMP\BrowserPostUpdateMsg_$([guid]::NewGuid().ToString().Substring(0,8)).txt"
                            
                            # Build notification message
                            $messageContent = "Message from Nexus Open Systems Ltd`n`nBROWSER UPDATES COMPLETED`n`n"
                            
                            if ($UpdatedBrowsers.Count -gt 0) {
                                $messageContent += "Successfully updated:`n"
                                foreach ($browser in $UpdatedBrowsers) {
                                    $messageContent += "- $browser`n"
                                }
                            }
                            
                            if ($FailedUpdates.Count -gt 0) {
                                $messageContent += "`nUpdate attempts (may require manual restart):`n"
                                foreach ($browser in $FailedUpdates) {
                                    $messageContent += "- $browser`n"
                                }
                            }
                            
                            $messageContent += "`nYour system is now more secure with the latest browser updates installed."
                            
                            # Write message to file
                            $messageContent | Out-File -FilePath $msgFile -Encoding ASCII -Force
                            
                            # Send notification using msg.exe
                            $cmdString = "msg.exe $sessionId /TIME:0 < `"$msgFile`""
                            cmd.exe /c $cmdString 2>&1 | Out-Null
                            
                            # Clean up
                            Start-Sleep -Milliseconds 500
                            Remove-Item -Path $msgFile -Force -ErrorAction SilentlyContinue
                            
                            Write-LogEntry "Post-update notification sent to session ${sessionId}" "Information"
                        }
                    }
                }
            }
        }
        catch {
            Write-LogEntry "Error notifying user of completed updates: $_" "Warning"
        }
    }
}

function Update-ChromeNow {
    <#
    .DESCRIPTION
    Force Google Chrome to update immediately
    #>
    Write-LogEntry "Attempting to force Google Chrome update" "Information"
    
    try {
        # Close Chrome
        Write-LogEntry "Closing Google Chrome" "Information"
        Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Check for Chrome update utility
        $chromePath = "C:\Program Files\Google\Chrome\Application"
        $chromePathx86 = "C:\Program Files (x86)\Google\Chrome\Application"
        
        $updateUtility = $null
        if (Test-Path "$chromePath\chrome.exe") {
            $updateUtility = $chromePath
        }
        elseif (Test-Path "$chromePathx86\chrome.exe") {
            $updateUtility = $chromePathx86
        }
        
        if ($updateUtility) {
            # Trigger Chrome update check
            Write-LogEntry "Triggering Chrome update check" "Information"
            & "$updateUtility\chrome.exe" --offlineoff --enable-plugins 2>&1 | Out-Null
            Start-Sleep -Seconds 5
            
            # Start Chrome (it will update on startup if available)
            Write-LogEntry "Starting Google Chrome" "Information"
            Start-Process -FilePath "$updateUtility\chrome.exe" -WindowStyle Minimized -ErrorAction SilentlyContinue
            
            Write-LogEntry "Google Chrome update initiated" "Information"
            return $true
        }
        else {
            Write-LogEntry "Google Chrome installation not found" "Warning"
            return $false
        }
    }
    catch {
        Write-LogEntry "Error updating Google Chrome: $_" "Error"
        return $false
    }
}

function Update-FirefoxNow {
    <#
    .DESCRIPTION
    Force Mozilla Firefox to update immediately
    #>
    Write-LogEntry "Attempting to force Mozilla Firefox update" "Information"
    
    try {
        # Close Firefox
        Write-LogEntry "Closing Mozilla Firefox" "Information"
        Get-Process firefox -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Check for Firefox
        $firefoxPath = "C:\Program Files\Mozilla Firefox"
        $firefoxPathx86 = "C:\Program Files (x86)\Mozilla Firefox"
        
        $firefoxExe = $null
        if (Test-Path "$firefoxPath\firefox.exe") {
            $firefoxExe = "$firefoxPath\firefox.exe"
        }
        elseif (Test-Path "$firefoxPathx86\firefox.exe") {
            $firefoxExe = "$firefoxPathx86\firefox.exe"
        }
        
        if ($firefoxExe) {
            # Start Firefox (it will check for updates on startup)
            Write-LogEntry "Starting Mozilla Firefox for update check" "Information"
            Start-Process -FilePath $firefoxExe -ArgumentList "-new-instance" -WindowStyle Minimized -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 10
            
            # Close Firefox to allow update installation
            Get-Process firefox -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            Write-LogEntry "Mozilla Firefox update initiated" "Information"
            return $true
        }
        else {
            Write-LogEntry "Mozilla Firefox installation not found" "Warning"
            return $false
        }
    }
    catch {
        Write-LogEntry "Error updating Mozilla Firefox: $_" "Error"
        return $false
    }
}

function Update-EdgeNow {
    <#
    .DESCRIPTION
    Force Microsoft Edge to update immediately
    #>
    Write-LogEntry "Attempting to force Microsoft Edge update" "Information"
    
    try {
        # Close Edge
        Write-LogEntry "Closing Microsoft Edge" "Information"
        Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Check for Edge
        $edgePath = "C:\Program Files\Microsoft\Edge\Application"
        $edgePathx86 = "C:\Program Files (x86)\Microsoft\Edge\Application"
        
        $edgeExe = $null
        if (Test-Path "$edgePath\msedge.exe") {
            $edgeExe = "$edgePath\msedge.exe"
        }
        elseif (Test-Path "$edgePathx86\msedge.exe") {
            $edgeExe = "$edgePathx86\msedge.exe"
        }
        
        if ($edgeExe) {
            # Start Edge (it will check for updates on startup)
            Write-LogEntry "Starting Microsoft Edge for update check" "Information"
            Start-Process -FilePath $edgeExe -WindowStyle Minimized -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 10
            
            # Close Edge to allow update installation
            Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            Write-LogEntry "Microsoft Edge update initiated" "Information"
            return $true
        }
        else {
            Write-LogEntry "Microsoft Edge installation not found" "Warning"
            return $false
        }
    }
    catch {
        Write-LogEntry "Error updating Microsoft Edge: $_" "Error"
        return $false
    }
}

# Main script logic
$browsersToUpdate = @()
$successfulUpdates = @()
$failedUpdates = @()

Write-LogEntry "======================================" "Information"
Write-LogEntry "Force Browser Updates Started" "Information"
Write-LogEntry "======================================" "Information"

# Check which browsers are installed
Write-LogEntry "Detecting installed browsers" "Information"

if ((Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")) {
    Write-LogEntry "Google Chrome detected" "Information"
    $browsersToUpdate += "Google Chrome"
}

if ((Test-Path "C:\Program Files\Mozilla Firefox\firefox.exe") -or (Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe")) {
    Write-LogEntry "Mozilla Firefox detected" "Information"
    $browsersToUpdate += "Mozilla Firefox"
}

if ((Test-Path "C:\Program Files\Microsoft\Edge\Application\msedge.exe") -or (Test-Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")) {
    Write-LogEntry "Microsoft Edge detected" "Information"
    $browsersToUpdate += "Microsoft Edge"
}

# Notify user that updates are starting
if ($browsersToUpdate.Count -gt 0) {
    Write-LogEntry "Requesting user deferral option" "Information"
    $userDeferred = Request-UpdateDeferral
    
    if ($userDeferred) {
        Write-LogEntry "User deferred updates - exiting without applying changes" "Warning"
        Write-LogEntry "======================================" "Information"
        Write-LogEntry "Force Browser Updates Cancelled by User" "Information"
        Write-LogEntry "======================================" "Information"
        Exit 0
    }
    
    Write-LogEntry "Notifying user of pending updates" "Information"
    Show-PreUpdateNotification -BrowsersToUpdate $browsersToUpdate
    Start-Sleep -Seconds 2
    
    # Apply updates
    Write-LogEntry "Applying browser updates" "Information"
    
    if ($browsersToUpdate -contains "Google Chrome") {
        if (Update-ChromeNow) {
            $successfulUpdates += "Google Chrome"
        }
        else {
            $failedUpdates += "Google Chrome"
        }
    }
    
    if ($browsersToUpdate -contains "Mozilla Firefox") {
        if (Update-FirefoxNow) {
            $successfulUpdates += "Mozilla Firefox"
        }
        else {
            $failedUpdates += "Mozilla Firefox"
        }
    }
    
    if ($browsersToUpdate -contains "Microsoft Edge") {
        if (Update-EdgeNow) {
            $successfulUpdates += "Microsoft Edge"
        }
        else {
            $failedUpdates += "Microsoft Edge"
        }
    }
    
    # Notify user that updates are complete
    Write-LogEntry "Notifying user of completed updates" "Information"
    Show-PostUpdateNotification -UpdatedBrowsers $successfulUpdates -FailedUpdates $failedUpdates
}
else {
    Write-LogEntry "No supported browsers detected" "Warning"
}

Write-LogEntry "======================================" "Information"
Write-LogEntry "Force Browser Updates Completed" "Information"
Write-LogEntry "======================================" "Information"

# Exit with success code
Exit 0
