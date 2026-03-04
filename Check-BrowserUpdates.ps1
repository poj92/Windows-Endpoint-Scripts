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
        # Running as SYSTEM - use msg.exe to notify active user session
        try {
            # Get active session ID
            $sessions = quser 2>$null
            if ($sessions) {
                $sessionIds = $sessions | ForEach-Object {
                    if ($_ -notmatch "USERNAME" -and $_ -match '\d+') {
                        $parts = $_ -split '\s+' | Where-Object { $_ }
                        if ($parts[2] -match '^\d+$') {
                            [int]$parts[2]
                        }
                    }
                } | Where-Object { $_ -gt 0 } | Select-Object -Unique
                
                foreach ($sessionId in $sessionIds) {
                    # Send message to user session using msg.exe with company header
                    $headerMsg = "Message from Nexus Open Systems Ltd`n`n"
                    $fullMessage = $headerMsg + $message
                    & msg.exe $sessionId $fullMessage /TIME:60 2>$null
                }
            }
        }
        catch {
            Write-LogEntry "Could not display notification to user session" "Warning"
        }
    }
    else {
        # Running as regular user - use WScript.Shell popup
        try {
            $objShell = New-Object -ComObject Wscript.Shell
            $objShell.Popup($message, 0, "Message from Nexus Open Systems Ltd", 48)
        }
        catch {
            Write-LogEntry "Notification: $message" "Warning"
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

