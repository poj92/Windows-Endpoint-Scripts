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
    Checks if Google Chrome has a pending update by comparing the executable version with available version folders
    #>
    try {
        $chromePaths = @(
            "C:\Program Files\Google\Chrome\Application",
            "C:\Program Files (x86)\Google\Chrome\Application"
        )
        
        foreach ($appPath in $chromePaths) {
            if (Test-Path $appPath) {
                $exePath = Join-Path $appPath "chrome.exe"
                
                if (Test-Path $exePath) {
                    # Get the version of the currently running executable
                    $exeVersion = (Get-Item $exePath).VersionInfo.FileVersion
                    if ([string]::IsNullOrEmpty($exeVersion)) {
                        continue
                    }
                    
                    # Get all version folders
                    $versionFolders = Get-ChildItem -Path $appPath -Directory -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }
                    
                    # Check if any version folder is newer than the running version
                    foreach ($folder in $versionFolders) {
                        try {
                            $folderVersion = [version]$folder.Name
                            $currentVersion = [version]$exeVersion
                            
                            if ($folderVersion -gt $currentVersion) {
                                # A newer version folder exists - update is pending
                                return $true
                            }
                        }
                        catch {
                            # Skip if version comparison fails
                            continue
                        }
                    }
                }
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
            # Check for actual Firefox update files in the updates directory
            $updatesPaths = @(
                "C:\Program Files\Mozilla Firefox\updates",
                "C:\Program Files (x86)\Mozilla Firefox\updates"
            )
            
            foreach ($updatesPath in $updatesPaths) {
                if (Test-Path $updatesPath) {
                    # Check for active-update.xml which indicates a downloaded/pending update
                    $activeUpdateFile = Join-Path $updatesPath "active-update.xml"
                    if (Test-Path $activeUpdateFile) {
                        return $true
                    }
                    
                    # Check for update directories with .mar files (Mozilla Archive)
                    $updateDirs = Get-ChildItem -Path $updatesPath -Directory -ErrorAction SilentlyContinue
                    foreach ($dir in $updateDirs) {
                        $marFiles = Get-ChildItem -Path $dir.FullName -Filter "*.mar" -ErrorAction SilentlyContinue
                        if ($marFiles) {
                            return $true
                        }
                    }
                }
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
    Checks if Microsoft Edge has a pending update by comparing the executable version with available version folders
    #>
    try {
        $edgePaths = @(
            "C:\Program Files\Microsoft\Edge\Application",
            "C:\Program Files (x86)\Microsoft\Edge\Application"
        )
        
        foreach ($appPath in $edgePaths) {
            if (Test-Path $appPath) {
                $exePath = Join-Path $appPath "msedge.exe"
                
                if (Test-Path $exePath) {
                    # Get the version of the currently running executable
                    $exeVersion = (Get-Item $exePath).VersionInfo.FileVersion
                    if ([string]::IsNullOrEmpty($exeVersion)) {
                        continue
                    }
                    
                    # Get all version folders
                    $versionFolders = Get-ChildItem -Path $appPath -Directory -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }
                    
                    # Check if any version folder is newer than the running version
                    foreach ($folder in $versionFolders) {
                        try {
                            $folderVersion = [version]$folder.Name
                            $currentVersion = [version]$exeVersion
                            
                            if ($folderVersion -gt $currentVersion) {
                                # A newer version folder exists - update is pending
                                return $true
                            }
                        }
                        catch {
                            # Skip if version comparison fails
                            continue
                        }
                    }
                }
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
    $message = "We have found critical vulnerabilities on your system due to the following outdated browsers:`n`n"
    
    if ($PendingUpdates -contains "Chrome") {
        $message += "- Google Chrome`n"
    }
    if ($PendingUpdates -contains "Firefox") {
        $message += "- Mozilla Firefox`n"
    }
    if ($PendingUpdates -contains "Edge") {
        $message += "- Microsoft Edge`n"
    }
    
    $message += "`nYou are advised to launch each of these browsers and apply the updates. If a browser is already open, please close the browser completely and reload it."
    
    # Check if running as SYSTEM
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
    if ($isSystem) {
        # Running as SYSTEM - use msg.exe to notify active user session
        Write-LogEntry "Running as SYSTEM - attempting user notification" "Information"
        
        try {
            # Get active user sessions
            $queryResults = query user 2>$null
            if ($queryResults) {
                Write-LogEntry "Query user results found" "Information"
                
                # Parse each line to find active sessions
                $queryResults | Select-Object -Skip 1 | ForEach-Object {
                    $line = $_
                    # Look for active sessions or console
                    if ($line -match 'Active' -or $line -match 'console') {
                        # Extract session ID
                        if ($line -match '\s+(\d+)\s+') {
                            $sessionId = $matches[1]
                            Write-LogEntry "Sending notification to session ID: ${sessionId}" "Information"
                            
                            # Create a temp file with the message (msg.exe works better with file input)
                            $msgFile = "$env:TEMP\BrowserUpdateMsg_$([guid]::NewGuid().ToString().Substring(0,8)).txt"
                            
                            # Build message content
                            $messageContent = @"
Message from Nexus Open Systems Ltd

We have found critical vulnerabilities on your system due to the following outdated browsers:

"@
                            
                            if ($PendingUpdates -contains "Chrome") {
                                $messageContent += "- Google Chrome`n"
                            }
                            if ($PendingUpdates -contains "Firefox") {
                                $messageContent += "- Mozilla Firefox`n"
                            }
                            if ($PendingUpdates -contains "Edge") {
                                $messageContent += "- Microsoft Edge`n"
                            }
                            
                            $messageContent += "`nYou are advised to launch each of these browsers and apply the updates. If a browser is already open, please close the browser completely and reload it."
                            
                            # Write message to file
                            $messageContent | Out-File -FilePath $msgFile -Encoding ASCII -Force
                            
                            # Send message using msg.exe with cmd.exe to handle file redirection
                            $cmdString = "msg.exe $sessionId /TIME:0 < `"$msgFile`""
                            cmd.exe /c $cmdString 2>&1 | Out-Null
                            
                            # Clean up the temp file
                            Start-Sleep -Milliseconds 500
                            Remove-Item -Path $msgFile -Force -ErrorAction SilentlyContinue
                            
                            Write-LogEntry "msg.exe notification sent to session ${sessionId}" "Information"
                        }
                    }
                }
            }
            else {
                Write-LogEntry "No active user sessions found" "Warning"
            }
        }
        catch {
            Write-LogEntry "Error sending notification via msg.exe: $_" "Warning"
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

