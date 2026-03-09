#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================
$BasePath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = Join-Path $BasePath "Detection.log"
$TrackingFile = Join-Path $BasePath "BrowserUsageTracking.json"
$QueueFile = Join-Path $BasePath "ReloadQueue.json"
$BrowserReloadThresholdHours = 24
$ApiTimeoutSeconds = 20

# =========================
# Bootstrap
# =========================
if (-not (Test-Path $BasePath)) {
    New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
}

function Write-LogEntry {
    param(
        [string]$Message,
        [ValidateSet("Information","Warning","Error")]
        [string]$Level = "Information"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    try {
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    }
    catch {}

    Write-Host $line
}

function New-BrowserUsageTracking {
    [PSCustomObject]@{
        Chrome  = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
        Firefox = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
        Edge    = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
    }
}

function Get-BrowserUsageTracking {
    $default = New-BrowserUsageTracking

    if (-not (Test-Path $TrackingFile)) {
        return $default
    }

    try {
        $raw = Get-Content -Path $TrackingFile -Raw -ErrorAction Stop
        $loaded = $raw | ConvertFrom-Json -ErrorAction Stop

        foreach ($browser in @('Chrome','Firefox','Edge')) {
            if (-not ($loaded.PSObject.Properties.Name -contains $browser)) {
                Add-Member -InputObject $loaded -MemberType NoteProperty -Name $browser -Value ([PSCustomObject]@{
                    LastStart = $null
                    LastStop  = $null
                    IsRunning = $false
                }) -Force
            }

            $browserObj = $loaded.$browser

            if (-not ($browserObj.PSObject.Properties.Name -contains 'LastStart')) {
                $migratedLastStart = $null
                if ($browserObj.PSObject.Properties.Name -contains 'SessionStart') {
                    $migratedLastStart = $browserObj.SessionStart
                }
                Add-Member -InputObject $browserObj -MemberType NoteProperty -Name LastStart -Value $migratedLastStart -Force
            }

            if (-not ($browserObj.PSObject.Properties.Name -contains 'LastStop')) {
                Add-Member -InputObject $browserObj -MemberType NoteProperty -Name LastStop -Value $null -Force
            }

            if (-not ($browserObj.PSObject.Properties.Name -contains 'IsRunning')) {
                Add-Member -InputObject $browserObj -MemberType NoteProperty -Name IsRunning -Value $false -Force
            }
        }

        return $loaded
    }
    catch {
        Write-LogEntry "Tracking file is invalid or incompatible. Recreating. Error: $($_.Exception.Message)" "Warning"
        return $default
    }
}

function Save-BrowserUsageTracking {
    param([object]$TrackingData)

    try {
        $TrackingData | ConvertTo-Json -Depth 5 | Set-Content -Path $TrackingFile -Force -Encoding UTF8
    }
    catch {
        Write-LogEntry "Failed to save tracking file: $($_.Exception.Message)" "Error"
    }
}

function Get-BrowserEarliestStartTime {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    try {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if (-not $procs) {
            return $null
        }

        $startTimes = @()
        foreach ($proc in $procs) {
            try {
                if ($proc.StartTime) {
                    $startTimes += $proc.StartTime
                }
            }
            catch {}
        }

        if ($startTimes.Count -gt 0) {
            return ($startTimes | Sort-Object | Select-Object -First 1)
        }

        return $null
    }
    catch {
        Write-LogEntry "Failed to read process start time for $ProcessName : $($_.Exception.Message)" "Warning"
        return $null
    }
}

function Update-BrowserSessionTracking {
    $tracking = Get-BrowserUsageTracking
    $now = Get-Date

    $processMap = @{
        Chrome  = 'chrome'
        Firefox = 'firefox'
        Edge    = 'msedge'
    }

    foreach ($browser in $processMap.Keys) {
        $procName = $processMap[$browser]
        $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
        $isRunningNow = [bool]$procs

        if ($isRunningNow) {
            $actualStart = Get-BrowserEarliestStartTime -ProcessName $procName

            if ($actualStart) {
                $actualStartIso = $actualStart.ToString("o")

                if (-not [bool]$tracking.$browser.IsRunning) {
                    $tracking.$browser.LastStart = $actualStartIso
                    $tracking.$browser.IsRunning = $true
                    Write-LogEntry "$browser is running. Actual process start time detected as $actualStartIso"
                }
                elseif ($tracking.$browser.LastStart -ne $actualStartIso) {
                    $tracking.$browser.LastStart = $actualStartIso
                    Write-LogEntry "$browser running session start time corrected to actual process start time $actualStartIso"
                }
            }
            else {
                if (-not [bool]$tracking.$browser.IsRunning) {
                    $tracking.$browser.LastStart = $now.ToString("o")
                    $tracking.$browser.IsRunning = $true
                    Write-LogEntry "$browser is running but actual process start time could not be read. Using observation time." "Warning"
                }
            }
        }
        else {
            if ([bool]$tracking.$browser.IsRunning) {
                $tracking.$browser.LastStop = $now.ToString("o")
                $tracking.$browser.IsRunning = $false
                Write-LogEntry "$browser is now closed. LastStop recorded as observation time $($tracking.$browser.LastStop)"
            }
        }
    }

    Save-BrowserUsageTracking -TrackingData $tracking
    return $tracking
}

function Get-BrowserStateInfo {
    param(
        [string]$BrowserName,
        [object]$TrackingData,
        [int]$ThresholdHours
    )

    $state = $TrackingData.$BrowserName
    $now = Get-Date

    if ([bool]$state.IsRunning) {
        if ([string]::IsNullOrWhiteSpace($state.LastStart)) {
            Write-LogEntry "$BrowserName is running but LastStart is unknown. Skipping on this pass." "Warning"
            return [PSCustomObject]@{
                IsRunning    = $true
                Hours        = 0
                ThresholdMet = $false
                Condition    = "Running"
            }
        }

        $started = [datetime]::Parse($state.LastStart)
        $age = $now - $started
        $hours = [math]::Round($age.TotalHours, 2)

        if ($age.TotalHours -ge $ThresholdHours) {
            Write-LogEntry "$BrowserName has been running continuously for $hours hours, which meets the threshold"
            return [PSCustomObject]@{
                IsRunning    = $true
                Hours        = $hours
                ThresholdMet = $true
                Condition    = "Running"
            }
        }

        Write-LogEntry "$BrowserName has been running for $hours hours"
        return [PSCustomObject]@{
            IsRunning    = $true
            Hours        = $hours
            ThresholdMet = $false
            Condition    = "Running"
        }
    }

    if ([string]::IsNullOrWhiteSpace($state.LastStop)) {
        Write-LogEntry "$BrowserName has not yet recorded a LastStop. Skipping on this pass."
        return [PSCustomObject]@{
            IsRunning    = $false
            Hours        = 0
            ThresholdMet = $false
            Condition    = "Closed"
        }
    }

    $stopped = [datetime]::Parse($state.LastStop)
    $inactive = $now - $stopped
    $hoursClosed = [math]::Round($inactive.TotalHours, 2)

    if ($inactive.TotalHours -ge $ThresholdHours) {
        Write-LogEntry "$BrowserName has been observed closed for $hoursClosed hours, which meets the threshold"
        return [PSCustomObject]@{
            IsRunning    = $false
            Hours        = $hoursClosed
            ThresholdMet = $true
            Condition    = "Closed"
        }
    }

    Write-LogEntry "$BrowserName has been observed closed for $hoursClosed hours"
    return [PSCustomObject]@{
        IsRunning    = $false
        Hours        = $hoursClosed
        ThresholdMet = $false
        Condition    = "Closed"
    }
}

function Get-ChromeInstallPath {
    @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-EdgeInstallPath {
    @(
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-FirefoxInstallPath {
    @(
        "C:\Program Files\Mozilla Firefox\firefox.exe",
        "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-ChromeVersion {
    try {
        $path = "HKLM:\SOFTWARE\Google\Chrome\BLBeacon"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).version
        }
    }
    catch {
        Write-LogEntry "Error reading Chrome installed version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Get-EdgeVersion {
    try {
        $path = "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue).version
        }
    }
    catch {
        Write-LogEntry "Error reading Edge installed version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Get-FirefoxVersion {
    try {
        $path = "HKLM:\SOFTWARE\Mozilla\Mozilla Firefox"
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)."CurrentVersion"
        }
    }
    catch {
        Write-LogEntry "Error reading Firefox installed version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Compare-VersionNewer {
    param(
        [string]$Installed,
        [string]$Latest
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Installed) -or [string]::IsNullOrWhiteSpace($Latest)) {
            return $null
        }

        return ([version]$Latest -gt [version]$Installed)
    }
    catch {
        Write-LogEntry "Version comparison failed. Installed='$Installed' Latest='$Latest'" "Warning"
        return $null
    }
}

function Get-LatestChromeVersion {
    try {
        $uri = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions?pageSize=1"
        $response = Invoke-RestMethod -Uri $uri -TimeoutSec $ApiTimeoutSeconds -ErrorAction Stop
        if ($response.versions -and $response.versions.Count -gt 0) {
            return $response.versions[0].version
        }
    }
    catch {
        Write-LogEntry "Unable to fetch latest Chrome version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Get-LatestFirefoxVersion {
    try {
        $uri = "https://product-details.mozilla.org/1.0/firefox_versions.json"
        $response = Invoke-RestMethod -Uri $uri -TimeoutSec $ApiTimeoutSeconds -ErrorAction Stop
        if ($response.LATEST_FIREFOX_VERSION) {
            return $response.LATEST_FIREFOX_VERSION
        }
    }
    catch {
        Write-LogEntry "Unable to fetch latest Firefox version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Get-LatestEdgeVersion {
    try {
        $uri = "https://edgeupdates.microsoft.com/api/products?view=enterprise"
        $response = Invoke-RestMethod -Uri $uri -TimeoutSec $ApiTimeoutSeconds -ErrorAction Stop

        foreach ($product in $response) {
            if ($product.Product -eq "Stable") {
                $release = $product.Releases | Sort-Object {
                    try { [version]$_.ProductVersion } catch { [version]"0.0.0.0" }
                } -Descending | Select-Object -First 1

                if ($release -and $release.ProductVersion) {
                    return $release.ProductVersion
                }
            }
        }
    }
    catch {
        Write-LogEntry "Unable to fetch latest Edge version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Test-ChromePendingUpdate {
    try {
        $chromeVersion = $null
        $chromeRegPath = "HKLM:\SOFTWARE\Google\Chrome\BLBeacon"

        if (Test-Path $chromeRegPath) {
            $chromeVersion = (Get-ItemProperty -Path $chromeRegPath -ErrorAction SilentlyContinue).version
        }

        $updateRegPath = "HKLM:\SOFTWARE\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}"
        if (Test-Path $updateRegPath) {
            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry "Chrome pending update detected via UpdateAvailable flag"
                return $true
            }

            if ($props.opv -and $chromeVersion -and $props.opv -ne $chromeVersion) {
                Write-LogEntry "Chrome pending update detected via staged/current version mismatch"
                return $true
            }
        }

        $chromeExe = Get-ChromeInstallPath
        if ($chromeExe) {
            $basePath = Split-Path $chromeExe -Parent

            if (Test-Path (Join-Path $basePath "new_chrome.exe")) {
                Write-LogEntry "Chrome pending update detected via new_chrome.exe"
                return $true
            }

            $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }

            if ($versionFolders.Count -gt 1) {
                Write-LogEntry "Chrome possible pending update detected via multiple version folders"
                return $true
            }
        }

        Write-LogEntry "No pending Chrome update detected"
        return $false
    }
    catch {
        Write-LogEntry "Chrome pending update check failed: $($_.Exception.Message)" "Warning"
        return $false
    }
}

function Test-EdgePendingUpdate {
    try {
        $edgeVersion = $null
        $edgeRegPath = "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon"
        if (Test-Path $edgeRegPath) {
            $edgeVersion = (Get-ItemProperty -Path $edgeRegPath -ErrorAction SilentlyContinue).version
        }

        $updateRegPath = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
        if (Test-Path $updateRegPath) {
            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry "Edge pending update detected via UpdateAvailable flag"
                return $true
            }

            if ($props.opv -and $edgeVersion -and $props.opv -ne $edgeVersion) {
                Write-LogEntry "Edge pending update detected via staged/current version mismatch"
                return $true
            }
        }

        $edgeExe = Get-EdgeInstallPath
        if ($edgeExe) {
            $basePath = Split-Path $edgeExe -Parent

            if (Test-Path (Join-Path $basePath "new_msedge.exe")) {
                Write-LogEntry "Edge pending update detected via new_msedge.exe"
                return $true
            }

            $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }

            if ($versionFolders.Count -gt 1) {
                Write-LogEntry "Edge possible pending update detected via multiple version folders"
                return $true
            }
        }

        Write-LogEntry "No pending Edge update detected"
        return $false
    }
    catch {
        Write-LogEntry "Edge pending update check failed: $($_.Exception.Message)" "Warning"
        return $false
    }
}

function Test-FirefoxPendingUpdate {
    try {
        $firefoxExe = Get-FirefoxInstallPath
        if (-not $firefoxExe) {
            return $false
        }

        $basePath = Split-Path $firefoxExe -Parent

        if (Test-Path (Join-Path $basePath "updated")) {
            Write-LogEntry "Firefox pending update detected via updated folder"
            return $true
        }

        $activeUpdateFile = Join-Path $basePath "active-update.xml"
        if (Test-Path $activeUpdateFile) {
            try {
                [xml]$xml = Get-Content -Path $activeUpdateFile -ErrorAction Stop
                if ($xml.updates.update) {
                    Write-LogEntry "Firefox pending update detected via active-update.xml"
                    return $true
                }
            }
            catch {
                Write-LogEntry "Firefox active-update.xml exists but could not be parsed" "Warning"
            }
        }

        Write-LogEntry "No pending Firefox update detected"
        return $false
    }
    catch {
        Write-LogEntry "Firefox pending update check failed: $($_.Exception.Message)" "Warning"
        return $false
    }
}

function Get-BrowserVersionStatus {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    switch ($BrowserName) {
        "Chrome" {
            $installed = Get-ChromeVersion
            $latest = Get-LatestChromeVersion
            $pending = Test-ChromePendingUpdate
        }
        "Firefox" {
            $installed = Get-FirefoxVersion
            $latest = Get-LatestFirefoxVersion
            $pending = Test-FirefoxPendingUpdate
        }
        "Edge" {
            $installed = Get-EdgeVersion
            $latest = Get-LatestEdgeVersion
            $pending = Test-EdgePendingUpdate
        }
    }

    $outOfDate = Compare-VersionNewer -Installed $installed -Latest $latest
    $upToDateText = if ($null -eq $outOfDate) { "Unknown" } elseif (-not $outOfDate) { "Yes" } else { "No" }
    Write-LogEntry "$BrowserName version status: Installed=$installed | Latest=$latest | UpToDate=$upToDateText | PendingUpdate=$pending"

    [PSCustomObject]@{
        Browser       = $BrowserName
        Installed     = $installed
        Latest        = $latest
        OutOfDate     = $outOfDate
        PendingUpdate = $pending
    }
}

function Get-ExistingReloadQueue {
    if (-not (Test-Path $QueueFile)) {
        return @()
    }

    try {
        $queue = Get-Content -Path $QueueFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        if (-not $queue -or -not $queue.Browsers) {
            return @()
        }

        foreach ($item in $queue.Browsers) {
            if (-not ($item.PSObject.Properties.Name -contains 'PostponeUntilUtc')) {
                $item | Add-Member -MemberType NoteProperty -Name PostponeUntilUtc -Value $null -Force
            }
            if (-not ($item.PSObject.Properties.Name -contains 'PostponeChoice')) {
                $item | Add-Member -MemberType NoteProperty -Name PostponeChoice -Value $null -Force
            }
            if (-not ($item.PSObject.Properties.Name -contains 'ScheduledTaskName')) {
                $item | Add-Member -MemberType NoteProperty -Name ScheduledTaskName -Value $null -Force
            }
        }

        return @($queue.Browsers)
    }
    catch {
        Write-LogEntry "Existing reload queue could not be read. Starting fresh. Error: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Save-ReloadQueue {
    param([array]$Browsers)

    $queue = [PSCustomObject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        Browsers   = $Browsers
    }

    $queue | ConvertTo-Json -Depth 8 | Set-Content -Path $QueueFile -Force -Encoding UTF8
}

function Add-OrUpdateQueueItem {
    param(
        [System.Collections.ArrayList]$Queue,
        [array]$ExistingQueue,
        [string]$Browser,
        [string]$Reason
    )

    $existingInNewQueue = $Queue | Where-Object { $_.Browser -eq $Browser } | Select-Object -First 1
    if ($existingInNewQueue) {
        $existingInNewQueue.Reason = $Reason
        Write-LogEntry "$Browser already queued in current run; updated reason"
        return
    }

    $existingSavedItem = $ExistingQueue | Where-Object { $_.Browser -eq $Browser } | Select-Object -First 1

    if ($existingSavedItem) {
        [void]$Queue.Add([PSCustomObject]@{
            Browser           = $Browser
            Reason            = $Reason
            PostponeUntilUtc  = $existingSavedItem.PostponeUntilUtc
            PostponeChoice    = $existingSavedItem.PostponeChoice
            ScheduledTaskName = $existingSavedItem.ScheduledTaskName
        })
        Write-LogEntry "$Browser added to reload queue with preserved postpone state"
    }
    else {
        [void]$Queue.Add([PSCustomObject]@{
            Browser           = $Browser
            Reason            = $Reason
            PostponeUntilUtc  = $null
            PostponeChoice    = $null
            ScheduledTaskName = $null
        })
        Write-LogEntry "$Browser added to reload queue"
    }
}

try {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload detection started"
    Write-LogEntry "Threshold set to $BrowserReloadThresholdHours hours"
    Write-LogEntry "======================================"

    $tracking = Update-BrowserSessionTracking
    $existingQueue = Get-ExistingReloadQueue
    $queue = [System.Collections.ArrayList]::new()

    $browserChecks = @(
        [PSCustomObject]@{ Name = "Chrome";  Present = [bool](Get-ChromeInstallPath)  }
        [PSCustomObject]@{ Name = "Firefox"; Present = [bool](Get-FirefoxInstallPath) }
        [PSCustomObject]@{ Name = "Edge";    Present = [bool](Get-EdgeInstallPath)    }
    )

    foreach ($browserCheck in $browserChecks) {
        if (-not $browserCheck.Present) {
            continue
        }

        $browser = $browserCheck.Name
        Write-LogEntry "Evaluating $browser"

        $stateInfo = Get-BrowserStateInfo -BrowserName $browser -TrackingData $tracking -ThresholdHours $BrowserReloadThresholdHours
        $versionStatus = Get-BrowserVersionStatus -BrowserName $browser

        if (-not $stateInfo.ThresholdMet) {
            Write-LogEntry "$browser does not meet the $BrowserReloadThresholdHours-hour threshold"
            continue
        }

        if ($stateInfo.IsRunning) {
            if ($versionStatus.PendingUpdate) {
                Add-OrUpdateQueueItem -Queue $queue -ExistingQueue $existingQueue -Browser $browser -Reason "Pending update and browser has been running continuously for at least 24 hours"
            }
            else {
                Write-LogEntry "$browser is running for 24+ hours but no pending update was detected"
            }
        }
        else {
            $needsNotification = $false
            $reasonParts = @()

            if ($versionStatus.PendingUpdate) {
                $needsNotification = $true
                $reasonParts += "pending update"
            }

            if ($versionStatus.OutOfDate -eq $true) {
                $needsNotification = $true
                $reasonParts += "out of date"
            }

            if ($needsNotification) {
                $reason = "Browser has not been used for at least 24 hours and is " + ($reasonParts -join " and ")
                Add-OrUpdateQueueItem -Queue $queue -ExistingQueue $existingQueue -Browser $browser -Reason $reason
            }
            else {
                Write-LogEntry "$browser has not been used for 24+ hours, but no pending update or confirmed out-of-date version was detected"
            }
        }
    }

    Save-ReloadQueue -Browsers $queue

    if ($queue.Count -gt 0) {
        Write-LogEntry "Detection complete. $($queue.Count) browser(s) queued for remediation."
        Write-LogEntry "Queue file written to $QueueFile"
        exit 10
    }
    else {
        Write-LogEntry "Detection complete. No browsers require remediation."
        exit 0
    }
}
catch {
    Write-LogEntry "Fatal detection error: $($_.Exception.Message)" "Error"
    exit 1
}
finally {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload detection finished"
    Write-LogEntry "======================================"
}