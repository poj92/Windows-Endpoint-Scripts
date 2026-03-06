# PowerShell Script to Force Browser Reload for Pending Updates
# Monitors browser usage and forces reload if:
# - Browser has not been opened in 72 hours
# - Browser has been open continuously for 72 hours
# - Browser has a pending update
# - Gives user 5-minute warning before automatic reload

# Configuration
$LogPath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = "$LogPath\ForceBrowserReload.log"
$TrackingFile = "$LogPath\BrowserUsageTracking.json"
$EventLogName = "Datto RMM"
$EventLogSource = "ForceBrowserReload"
$InactivityThresholdHours = 72
$WarningTimeSeconds = 300  # 5 minutes

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
        $eventId = @{ "Information" = 3000; "Warning" = 3001; "Error" = 3002 }[$Level]
        $eventType = @{ "Information" = "Information"; "Warning" = "Warning"; "Error" = "Error" }[$Level]
        
        Write-EventLog -LogName $EventLogName -Source $EventLogSource -EventId $eventId -EntryType $eventType -Message $entry -ErrorAction SilentlyContinue
    }
    catch {
        # Continue if event log write fails
    }
    
    # Also output to console for Datto
    Write-Host $entry
}

function Get-ChromeVersion {
    <#
    .SYNOPSIS
    Get installed Chrome version from registry
    #>
    try {
        $path = "HKLM:\SOFTWARE\Google\Chrome\BLBeacon"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).version
        }
    }
    catch {
        Write-LogEntry "Error reading Chrome installed version: $_" "Warning"
    }

    return $null
}

function Get-EdgeVersion {
    <#
    .SYNOPSIS
    Get installed Edge version from registry
    #>
    try {
        $path = "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).version
        }
    }
    catch {
        Write-LogEntry "Error reading Edge installed version: $_" "Warning"
    }

    return $null
}

function Get-FirefoxVersion {
    <#
    .SYNOPSIS
    Get installed Firefox version from registry
    #>
    try {
        $path = "HKLM:\SOFTWARE\Mozilla\Mozilla Firefox"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)."CurrentVersion"
        }
    }
    catch {
        Write-LogEntry "Error reading Firefox installed version: $_" "Warning"
    }

    return $null
}

function Get-LatestChromeVersion {
    <#
    .SYNOPSIS
    Get latest stable Chrome version from Google API
    #>
    try {
        $response = Invoke-RestMethod -Uri "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions" -TimeoutSec 20 -ErrorAction Stop
        if ($response.versions -and $response.versions.Count -gt 0) {
            return $response.versions[0].version
        }
    }
    catch {
        Write-LogEntry "Unable to fetch latest Chrome version: $_" "Warning"
    }

    return $null
}

function Get-LatestEdgeVersion {
    <#
    .SYNOPSIS
    Get latest stable Edge version from Microsoft API
    #>
    try {
        $response = Invoke-RestMethod -Uri "https://edgeupdates.microsoft.com/api/products" -TimeoutSec 20 -ErrorAction Stop
        $stableRelease = $response.ProductReleases | Where-Object { $_.Product -eq "Stable" } | Select-Object -First 1

        if ($stableRelease) {
            return $stableRelease.ProductVersion
        }
    }
    catch {
        Write-LogEntry "Unable to fetch latest Edge version: $_" "Warning"
    }

    return $null
}

function Get-LatestFirefoxVersion {
    <#
    .SYNOPSIS
    Get latest stable Firefox version from Mozilla API
    #>
    try {
        $response = Invoke-RestMethod -Uri "https://product-details.mozilla.org/1.0/firefox_versions.json" -TimeoutSec 20 -ErrorAction Stop
        return $response.LATEST_FIREFOX_VERSION
    }
    catch {
        Write-LogEntry "Unable to fetch latest Firefox version: $_" "Warning"
    }

    return $null
}

function Get-BrowserVersionStatus {
    <#
    .SYNOPSIS
    Return installed/latest/up-to-date status for detected browsers
    #>
    $results = @()

    $installedChrome = Get-ChromeVersion
    if ($installedChrome) {
        $latestChrome = Get-LatestChromeVersion
        $results += [PSCustomObject]@{
            Browser = "Chrome"
            Installed = $installedChrome
            Latest = $(if ($latestChrome) { $latestChrome } else { "Unknown" })
            UpToDate = $(if ($latestChrome) { $installedChrome -eq $latestChrome } else { $null })
        }
    }

    $installedEdge = Get-EdgeVersion
    if ($installedEdge) {
        $latestEdge = Get-LatestEdgeVersion
        $results += [PSCustomObject]@{
            Browser = "Edge"
            Installed = $installedEdge
            Latest = $(if ($latestEdge) { $latestEdge } else { "Unknown" })
            UpToDate = $(if ($latestEdge) { $installedEdge -eq $latestEdge } else { $null })
        }
    }

    $installedFirefox = Get-FirefoxVersion
    if ($installedFirefox) {
        $latestFirefox = Get-LatestFirefoxVersion
        $results += [PSCustomObject]@{
            Browser = "Firefox"
            Installed = $installedFirefox
            Latest = $(if ($latestFirefox) { $latestFirefox } else { "Unknown" })
            UpToDate = $(if ($latestFirefox) { $installedFirefox -eq $latestFirefox } else { $null })
        }
    }

    return $results
}

function Write-BrowserVersionSummary {
    <#
    .SYNOPSIS
    Log browser version status summary
    #>
    param(
        [object[]]$VersionResults
    )

    if (-not $VersionResults -or $VersionResults.Count -eq 0) {
        Write-LogEntry "No supported browsers detected for version comparison" "Information"
        return
    }

    Write-LogEntry "Browser version status summary:" "Information"

    foreach ($result in $VersionResults) {
        $upToDateText = if ($null -eq $result.UpToDate) {
            "Unknown"
        }
        elseif ($result.UpToDate) {
            "Yes"
        }
        else {
            "No"
        }

        Write-LogEntry "$($result.Browser): Installed=$($result.Installed) | Latest=$($result.Latest) | UpToDate=$upToDateText" "Information"
    }
}

function Get-BrowserUsageTracking {
    <#
    .SYNOPSIS
    Load browser usage tracking data from JSON file
    #>
    if (Test-Path $TrackingFile) {
        try {
            $tracking = Get-Content -Path $TrackingFile -Raw -ErrorAction Stop | ConvertFrom-Json
            return $tracking
        }
        catch {
            Write-LogEntry "Error loading tracking file, creating new tracking data" "Warning"
            return @{
                Chrome = @{ SessionStart = $null; IsRunning = $false }
                Firefox = @{ SessionStart = $null; IsRunning = $false }
                Edge = @{ SessionStart = $null; IsRunning = $false }
            }
        }
    }
    else {
        return @{
            Chrome = @{ SessionStart = $null; IsRunning = $false }
            Firefox = @{ SessionStart = $null; IsRunning = $false }
            Edge = @{ SessionStart = $null; IsRunning = $false }
        }
    }
}

function Save-BrowserUsageTracking {
    param(
        [object]$TrackingData
    )
    
    try {
        $TrackingData | ConvertTo-Json -Depth 3 | Set-Content -Path $TrackingFile -Force -ErrorAction Stop
    }
    catch {
        Write-LogEntry "Error saving tracking file: $_" "Error"
    }
}

function Update-BrowserSessionTracking {
    <#
    .SYNOPSIS
    Track browser session start times - only updates when browser starts, not while running
    #>
    $tracking = Get-BrowserUsageTracking
    $currentTime = Get-Date
    
    # Check Chrome
    $chromeRunning = Get-Process chrome -ErrorAction SilentlyContinue
    if ($chromeRunning) {
        # Browser is running
        if (-not $tracking.Chrome.IsRunning) {
            # Browser just started - record session start
            $tracking.Chrome.SessionStart = $currentTime.ToString("o")
            $tracking.Chrome.IsRunning = $true
            Write-LogEntry "Chrome started new session - recorded session start time" "Information"
        }
        else {
            # Browser was already running - don't update timestamp
            Write-LogEntry "Chrome already running - maintaining session start time" "Information"
        }
    }
    else {
        # Browser is not running
        if ($tracking.Chrome.IsRunning) {
            # Browser just closed
            $tracking.Chrome.IsRunning = $false
            Write-LogEntry "Chrome closed - session ended" "Information"
        }
    }
    
    # Check Firefox
    $firefoxRunning = Get-Process firefox -ErrorAction SilentlyContinue
    if ($firefoxRunning) {
        if (-not $tracking.Firefox.IsRunning) {
            $tracking.Firefox.SessionStart = $currentTime.ToString("o")
            $tracking.Firefox.IsRunning = $true
            Write-LogEntry "Firefox started new session - recorded session start time" "Information"
        }
        else {
            Write-LogEntry "Firefox already running - maintaining session start time" "Information"
        }
    }
    else {
        if ($tracking.Firefox.IsRunning) {
            $tracking.Firefox.IsRunning = $false
            Write-LogEntry "Firefox closed - session ended" "Information"
        }
    }
    
    # Check Edge
    $edgeRunning = Get-Process msedge -ErrorAction SilentlyContinue
    if ($edgeRunning) {
        if (-not $tracking.Edge.IsRunning) {
            $tracking.Edge.SessionStart = $currentTime.ToString("o")
            $tracking.Edge.IsRunning = $true
            Write-LogEntry "Edge started new session - recorded session start time" "Information"
        }
        else {
            Write-LogEntry "Edge already running - maintaining session start time" "Information"
        }
    }
    else {
        if ($tracking.Edge.IsRunning) {
            $tracking.Edge.IsRunning = $false
            Write-LogEntry "Edge closed - session ended" "Information"
        }
    }
    
    Save-BrowserUsageTracking -TrackingData $tracking
}

function Test-BrowserSessionAge {
    <#
    .SYNOPSIS
    Check if browser session has been running for longer than threshold OR is not running and hasn't been started in threshold time
    Returns $true if browser needs reload (either running too long OR inactive too long)
    #>
    param(
        [string]$BrowserName,
        [object]$TrackingData,
        [int]$ThresholdHours
    )
    
    $sessionStart = $TrackingData.$BrowserName.SessionStart
    $isRunning = $TrackingData.$BrowserName.IsRunning
    
    if ($null -eq $sessionStart -or $sessionStart -eq "") {
        Write-LogEntry "$BrowserName has no session start timestamp - assuming needs reload" "Information"
        return $true
    }
    
    try {
        $sessionStartDate = [DateTime]::Parse($sessionStart)
        $hoursSinceSessionStart = (Get-Date) - $sessionStartDate
        
        if ($hoursSinceSessionStart.TotalHours -ge $ThresholdHours) {
            if ($isRunning) {
                Write-LogEntry "$BrowserName has been running continuously for $([math]::Round($hoursSinceSessionStart.TotalHours, 2)) hours" "Information"
            }
            else {
                Write-LogEntry "$BrowserName has not been started for $([math]::Round($hoursSinceSessionStart.TotalHours, 2)) hours" "Information"
            }
            return $true
        }
        else {
            if ($isRunning) {
                Write-LogEntry "$BrowserName current session started $([math]::Round($hoursSinceSessionStart.TotalHours, 2)) hours ago (still running)" "Information"
            }
            else {
                Write-LogEntry "$BrowserName last session was $([math]::Round($hoursSinceSessionStart.TotalHours, 2)) hours ago" "Information"
            }
            return $false
        }
    }
    catch {
        Write-LogEntry "Error parsing session start date for $BrowserName : $_" "Warning"
        return $true
    }
}

function Test-ChromePendingUpdate {
    <#
    .SYNOPSIS
    Check if Chrome has a pending update by examining GoogleUpdate registry keys
    #>
    try {
        # Check Chrome version registry
        $chromeVersion = $null
        $chromeRegPath = "HKLM:\SOFTWARE\Google\Chrome\BLBeacon"
        
        if (Test-Path $chromeRegPath) {
            $chromeVersion = (Get-ItemProperty -Path $chromeRegPath -ErrorAction SilentlyContinue).version
        }
        
        # Check for pending update indicators
        # Method 1: Check if GoogleUpdate has a newer version staged
        $updateRegPath = "HKLM:\SOFTWARE\Google\Update\ClientState\{8A69D345-D564-463c-AFF1-A69D9E530F96}"
        
        if (Test-Path $updateRegPath) {
            $updateProps = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue
            
            # Check for update available flag or opv (old product version) different from current version
            if ($updateProps.UpdateAvailable -eq 1) {
                Write-LogEntry "Chrome has pending update (UpdateAvailable flag set)" "Information"
                return $true
            }
            
            if ($updateProps.opv -and $chromeVersion -and $updateProps.opv -ne $chromeVersion) {
                Write-LogEntry "Chrome has pending update (version mismatch: current=$chromeVersion, staged=$($updateProps.opv))" "Information"
                return $true
            }
        }
        
        # Method 2: Check if Chrome.exe and new_chrome.exe both exist (indicates pending update)
        $chromePath = "C:\Program Files\Google\Chrome\Application"
        $chromePathx86 = "C:\Program Files (x86)\Google\Chrome\Application"
        
        $basePath = $null
        if (Test-Path "$chromePath\chrome.exe") {
            $basePath = $chromePath
        }
        elseif (Test-Path "$chromePathx86\chrome.exe") {
            $basePath = $chromePathx86
        }
        
        if ($basePath) {
            if (Test-Path "$basePath\new_chrome.exe") {
                Write-LogEntry "Chrome has pending update (new_chrome.exe exists)" "Information"
                return $true
            }
            
            # Check for version-specific folders with newer versions
            $installedVersionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }
            if ($installedVersionFolders.Count -gt 1) {
                Write-LogEntry "Chrome has multiple version folders - possible pending update" "Information"
                return $true
            }
        }
        
        Write-LogEntry "Chrome has no pending updates detected" "Information"
        return $false
    }
    catch {
        Write-LogEntry "Error checking Chrome pending updates: $_" "Warning"
        return $false
    }
}

function Test-FirefoxPendingUpdate {
    <#
    .SYNOPSIS
    Check if Firefox has a pending update
    #>
    try {
        # Check Firefox update status file
        $firefoxUpdatePath = "$env:ProgramData\Mozilla\updates\*\updates\*"
        $activeUpdateXml = "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*\updates.xml"
        
        # Method 1: Check for active-update.xml in Firefox program directory
        $firefoxPath = "C:\Program Files\Mozilla Firefox"
        $firefoxPathx86 = "C:\Program Files (x86)\Mozilla Firefox"
        
        $basePath = $null
        if (Test-Path "$firefoxPath\firefox.exe") {
            $basePath = $firefoxPath
        }
        elseif (Test-Path "$firefoxPathx86\firefox.exe") {
            $basePath = $firefoxPathx86
        }
        
        if ($basePath) {
            # Check for ready update directory
            $updatesReady = Get-ChildItem -Path "$basePath\updated" -ErrorAction SilentlyContinue
            if ($updatesReady) {
                Write-LogEntry "Firefox has pending update (updated folder exists)" "Information"
                return $true
            }
            
            # Check for active-update.xml
            $activeUpdateFile = "$basePath\active-update.xml"
            if (Test-Path $activeUpdateFile) {
                [xml]$updateXml = Get-Content -Path $activeUpdateFile -ErrorAction SilentlyContinue
                if ($updateXml.updates.update.type -eq "minor" -or $updateXml.updates.update.type -eq "major") {
                    Write-LogEntry "Firefox has pending update (active-update.xml contains update)" "Information"
                    return $true
                }
            }
        }
        
        Write-LogEntry "Firefox has no pending updates detected" "Information"
        return $false
    }
    catch {
        Write-LogEntry "Error checking Firefox pending updates: $_" "Warning"
        return $false
    }
}

function Test-EdgePendingUpdate {
    <#
    .SYNOPSIS
    Check if Microsoft Edge has a pending update
    #>
    try {
        # Edge uses similar update mechanism to Chrome
        $edgeVersion = $null
        $edgeRegPath = "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon"
        
        if (Test-Path $edgeRegPath) {
            $edgeVersion = (Get-ItemProperty -Path $edgeRegPath -ErrorAction SilentlyContinue).version
        }
        
        # Check Microsoft Edge Update registry
        $updateRegPath = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
        
        if (Test-Path $updateRegPath) {
            $updateProps = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue
            
            if ($updateProps.UpdateAvailable -eq 1) {
                Write-LogEntry "Edge has pending update (UpdateAvailable flag set)" "Information"
                return $true
            }
        }
        
        # Check for new_msedge.exe
        $edgePath = "C:\Program Files\Microsoft\Edge\Application"
        $edgePathx86 = "C:\Program Files (x86)\Microsoft\Edge\Application"
        
        $basePath = $null
        if (Test-Path "$edgePath\msedge.exe") {
            $basePath = $edgePath
        }
        elseif (Test-Path "$edgePathx86\msedge.exe") {
            $basePath = $edgePathx86
        }
        
        if ($basePath) {
            if (Test-Path "$basePath\new_msedge.exe") {
                Write-LogEntry "Edge has pending update (new_msedge.exe exists)" "Information"
                return $true
            }
            
            # Check for multiple version folders
            $installedVersionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }
            if ($installedVersionFolders.Count -gt 1) {
                Write-LogEntry "Edge has multiple version folders - possible pending update" "Information"
                return $true
            }
        }
        
        Write-LogEntry "Edge has no pending updates detected" "Information"
        return $false
    }
    catch {
        Write-LogEntry "Error checking Edge pending updates: $_" "Warning"
        return $false
    }
}

function Show-CountdownWarning {
    <#
    .SYNOPSIS
    Show countdown warning to user with option to cancel
    #>
    param(
        [string]$BrowserName,
        [int]$CountdownSeconds
    )
    
    # Check if running as SYSTEM
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $isSystem = $currentUser.User.Value -eq "S-1-5-18"
    
    if ($isSystem) {
        Write-LogEntry "Running as SYSTEM - showing countdown warning to user session" "Information"
        
        try {
            # Get active user sessions
            $queryResults = query user 2>$null
            if ($queryResults) {
                foreach ($line in ($queryResults | Select-Object -Skip 1)) {
                    if ($line -match 'Active' -or $line -match 'console') {
                        if ($line -match '\s+(\d+)\s+') {
                            $sessionId = $matches[1]
                            
                            # Create PowerShell script for countdown dialog
                            $markerFile = "$env:TEMP\BrowserReloadCancelled_$([guid]::NewGuid().ToString().Substring(0,8)).txt"
                            $psScriptPath = "$env:TEMP\BrowserReloadCountdown_$([guid]::NewGuid().ToString().Substring(0,8)).ps1"
                            
                            $scriptContent = @"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

`$form = New-Object System.Windows.Forms.Form
`$form.Text = 'Browser Update Required - $BrowserName'
`$form.Width = 500
`$form.Height = 250
`$form.StartPosition = 'CenterScreen'
`$form.TopMost = `$true
`$form.FormBorderStyle = 'FixedDialog'
`$form.MaximizeBox = `$false

`$label = New-Object System.Windows.Forms.Label
`$label.Location = New-Object System.Drawing.Point(20,20)
`$label.Size = New-Object System.Drawing.Size(450,60)
`$label.Text = '$BrowserName has a pending update and has not been used in over 72 hours.``n``nThe browser will automatically reload to apply updates in:'
`$label.Font = New-Object System.Drawing.Font('Segoe UI',10)
`$form.Controls.Add(`$label)

`$countdownLabel = New-Object System.Windows.Forms.Label
`$countdownLabel.Location = New-Object System.Drawing.Point(20,90)
`$countdownLabel.Size = New-Object System.Drawing.Size(450,40)
`$countdownLabel.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
`$countdownLabel.ForeColor = [System.Drawing.Color]::Red
`$countdownLabel.TextAlign = 'MiddleCenter'
`$form.Controls.Add(`$countdownLabel)

`$cancelButton = New-Object System.Windows.Forms.Button
`$cancelButton.Location = New-Object System.Drawing.Point(150,150)
`$cancelButton.Size = New-Object System.Drawing.Size(200,35)
`$cancelButton.Text = 'Cancel Reload'
`$cancelButton.Font = New-Object System.Drawing.Font('Segoe UI',10)
`$cancelButton.Add_Click({
    'cancelled' | Out-File -FilePath '$markerFile' -Force
    `$form.Close()
})
`$form.Controls.Add(`$cancelButton)

`$secondsLeft = $CountdownSeconds
`$timer = New-Object System.Windows.Forms.Timer
`$timer.Interval = 1000
`$timer.Add_Tick({
    `$script:secondsLeft--
    `$minutes = [math]::Floor(`$script:secondsLeft / 60)
    `$seconds = `$script:secondsLeft % 60
    `$countdownLabel.Text = "{0}:{1:D2}" -f `$minutes, `$seconds
    
    if (`$script:secondsLeft -le 0) {
        `$timer.Stop()
        `$form.Close()
    }
})

`$timer.Start()
`$form.Add_Shown({`$form.Activate()})
[void]`$form.ShowDialog()

if (Test-Path '$markerFile') {
    Exit 1
}
Exit 0
"@
                            $scriptContent | Out-File -FilePath $psScriptPath -Encoding UTF8 -Force
                            
                            # Create scheduled task to show dialog
                            $taskName = "BrowserReloadCountdown_$([guid]::NewGuid().ToString().Substring(0,8))"
                            
                            $createCmd = "schtasks /Create /TN `"$taskName`" /TR `"powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `'$psScriptPath`'`" /SC ONCE /ST 00:00 /RU `"*`" /RL HIGHEST /IT /F"
                            cmd.exe /c $createCmd 2>&1 | Out-Null
                            
                            cmd.exe /c "schtasks /Run /TN `"$taskName`"" 2>&1 | Out-Null
                            
                            Write-LogEntry "Countdown warning displayed to user - waiting $CountdownSeconds seconds" "Information"
                            Start-Sleep -Seconds ($CountdownSeconds + 5)
                            
                            # Check if user cancelled
                            if (Test-Path $markerFile) {
                                Write-LogEntry "User cancelled browser reload" "Information"
                                
                                # Cleanup
                                cmd.exe /c "schtasks /Delete /TN `"$taskName`" /F" 2>&1 | Out-Null
                                Remove-Item -Path $psScriptPath -Force -ErrorAction SilentlyContinue
                                Remove-Item -Path $markerFile -Force -ErrorAction SilentlyContinue
                                
                                return $false
                            }
                            
                            # Cleanup
                            cmd.exe /c "schtasks /Delete /TN `"$taskName`" /F" 2>&1 | Out-Null
                            Remove-Item -Path $psScriptPath -Force -ErrorAction SilentlyContinue
                            
                            Write-LogEntry "Countdown completed - proceeding with browser reload" "Information"
                            return $true
                        }
                    }
                }
            }
        }
        catch {
            Write-LogEntry "Error showing countdown warning: $_" "Warning"
            return $true
        }
    }
    
    return $true
}

function Invoke-BrowserReload {
    <#
    .SYNOPSIS
    Force reload a specific browser to apply pending updates
    #>
    param(
        [string]$BrowserName
    )
    
    Write-LogEntry "Forcing reload of $BrowserName to apply pending updates" "Information"
    
    try {
        switch ($BrowserName) {
            "Chrome" {
                $chromeExe = $null
                if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
                    $chromeExe = "C:\Program Files\Google\Chrome\Application\chrome.exe"
                }
                elseif (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") {
                    $chromeExe = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                }
                
                if ($chromeExe) {
                    Write-LogEntry "Closing Chrome processes" "Information"
                    Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    
                    Write-LogEntry "Restarting Chrome to apply update" "Information"
                    Start-Process -FilePath $chromeExe -WindowStyle Minimized -ErrorAction SilentlyContinue
                    
                    return $true
                }
            }
            
            "Firefox" {
                $firefoxExe = $null
                if (Test-Path "C:\Program Files\Mozilla Firefox\firefox.exe") {
                    $firefoxExe = "C:\Program Files\Mozilla Firefox\firefox.exe"
                }
                elseif (Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") {
                    $firefoxExe = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
                }
                
                if ($firefoxExe) {
                    Write-LogEntry "Closing Firefox processes" "Information"
                    Get-Process firefox -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    
                    Write-LogEntry "Restarting Firefox to apply update" "Information"
                    Start-Process -FilePath $firefoxExe -WindowStyle Minimized -ErrorAction SilentlyContinue
                    
                    return $true
                }
            }
            
            "Edge" {
                $edgeExe = $null
                if (Test-Path "C:\Program Files\Microsoft\Edge\Application\msedge.exe") {
                    $edgeExe = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
                }
                elseif (Test-Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe") {
                    $edgeExe = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                }
                
                if ($edgeExe) {
                    Write-LogEntry "Closing Edge processes" "Information"
                    Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    
                    Write-LogEntry "Restarting Edge to apply update" "Information"
                    Start-Process -FilePath $edgeExe -WindowStyle Minimized -ErrorAction SilentlyContinue
                    
                    return $true
                }
            }
        }
        
        Write-LogEntry "Browser executable not found for $BrowserName" "Warning"
        return $false
    }
    catch {
        Write-LogEntry "Error reloading $BrowserName : $_" "Error"
        return $false
    }
}

# Main script execution
Write-LogEntry "======================================" "Information"
Write-LogEntry "Browser Reload Check Started" "Information"
Write-LogEntry "======================================" "Information"

# Update session tracking for browsers
Update-BrowserSessionTracking

# Load tracking data
$tracking = Get-BrowserUsageTracking

# Check and log version status for detected browsers
$versionResults = Get-BrowserVersionStatus
Write-BrowserVersionSummary -VersionResults $versionResults

# Check each browser
$browsersToReload = @()

# Chrome
if ((Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")) {
    Write-LogEntry "Checking Chrome..." "Information"
    
    if (Test-BrowserSessionAge -BrowserName "Chrome" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) {
        if (Test-ChromePendingUpdate) {
            Write-LogEntry "Chrome meets reload criteria (session age 72+ hours + pending update)" "Information"
            $browsersToReload += "Chrome"
        }
    }
}

# Firefox
if ((Test-Path "C:\Program Files\Mozilla Firefox\firefox.exe") -or (Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe")) {
    Write-LogEntry "Checking Firefox..." "Information"
    
    if (Test-BrowserSessionAge -BrowserName "Firefox" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) {
        if (Test-FirefoxPendingUpdate) {
            Write-LogEntry "Firefox meets reload criteria (session age 72+ hours + pending update)" "Information"
            $browsersToReload += "Firefox"
        }
    }
}

# Edge
if ((Test-Path "C:\Program Files\Microsoft\Edge\Application\msedge.exe") -or (Test-Path "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")) {
    Write-LogEntry "Checking Edge..." "Information"
    
    if (Test-BrowserSessionAge -BrowserName "Edge" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) {
        if (Test-EdgePendingUpdate) {
            Write-LogEntry "Edge meets reload criteria (session age 72+ hours + pending update)" "Information"
            $browsersToReload += "Edge"
        }
    }
}

# Process browsers that need reload
if ($browsersToReload.Count -gt 0) {
    Write-LogEntry "Found $($browsersToReload.Count) browser(s) requiring reload" "Information"
    
    foreach ($browser in $browsersToReload) {
        Write-LogEntry "Processing reload for $browser" "Information"
        
        # Show countdown warning
        $proceedWithReload = Show-CountdownWarning -BrowserName $browser -CountdownSeconds $WarningTimeSeconds
        
        if ($proceedWithReload) {
            # Reload browser
            if (Invoke-BrowserReload -BrowserName $browser) {
                Write-LogEntry "$browser successfully reloaded" "Information"
                
                # Update tracking - reset session start and mark as running
                $tracking = Get-BrowserUsageTracking
                $tracking.$browser.SessionStart = (Get-Date).ToString("o")
                $tracking.$browser.IsRunning = $true
                Save-BrowserUsageTracking -TrackingData $tracking
            }
            else {
                Write-LogEntry "Failed to reload $browser" "Error"
            }
        }
        else {
            Write-LogEntry "$browser reload cancelled by user" "Information"
        }
    }
}
else {
    Write-LogEntry "No browsers require reload at this time" "Information"
}

Write-LogEntry "======================================" "Information"
Write-LogEntry "Browser Reload Check Completed" "Information"
Write-LogEntry "======================================" "Information"

Exit 0
