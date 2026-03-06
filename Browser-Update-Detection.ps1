# PowerShell Script to Detect Browser Update Conditions
# Script must be run in system context.

#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================
$BasePath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = Join-Path $BasePath "Detection.log"
$TrackingFile = Join-Path $BasePath "BrowserUsageTracking.json"
$QueueFile = Join-Path $BasePath "ReloadQueue.json"
$InactivityThresholdHours = 72

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
    return [PSCustomObject]@{
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
            catch {
                # Ignore any process where StartTime can't be read
            }
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

        if (-not ($tracking.$browser.PSObject.Properties.Name -contains 'LastStart')) {
            Add-Member -InputObject $tracking.$browser -MemberType NoteProperty -Name LastStart -Value $null -Force
        }
        if (-not ($tracking.$browser.PSObject.Properties.Name -contains 'LastStop')) {
            Add-Member -InputObject $tracking.$browser -MemberType NoteProperty -Name LastStop -Value $null -Force
        }
        if (-not ($tracking.$browser.PSObject.Properties.Name -contains 'IsRunning')) {
            Add-Member -InputObject $tracking.$browser -MemberType NoteProperty -Name IsRunning -Value $false -Force
        }

        if ($isRunningNow) {
            $actualStart = Get-BrowserEarliestStartTime -ProcessName $procName

            if ($actualStart) {
                $actualStartIso = $actualStart.ToString("o")

                if (-not [bool]$tracking.$browser.IsRunning) {
                    $tracking.$browser.LastStart = $actualStartIso
                    $tracking.$browser.IsRunning = $true
                    Write-LogEntry "$browser is running. Actual process start time detected as $actualStartIso"
                }
                else {
                    if ($tracking.$browser.LastStart -ne $actualStartIso) {
                        $tracking.$browser.LastStart = $actualStartIso
                        Write-LogEntry "$browser running session start time corrected to actual process start time $actualStartIso"
                    }
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

function Test-BrowserSessionAge {
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
            return $false
        }

        $started = [datetime]::Parse($state.LastStart)
        $age = $now - $started

        if ($age.TotalHours -ge $ThresholdHours) {
            Write-LogEntry "$BrowserName has been running continuously for $([math]::Round($age.TotalHours,2)) hours"
            return $true
        }

        Write-LogEntry "$BrowserName has been running for $([math]::Round($age.TotalHours,2)) hours"
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($state.LastStop)) {
        Write-LogEntry "$BrowserName has not yet recorded a LastStop. Skipping on this pass."
        return $false
    }

    $stopped = [datetime]::Parse($state.LastStop)
    $inactive = $now - $stopped

    if ($inactive.TotalHours -ge $ThresholdHours) {
        Write-LogEntry "$BrowserName has been observed closed for $([math]::Round($inactive.TotalHours,2)) hours, which meets the threshold"
        return $true
    }

    Write-LogEntry "$BrowserName has been observed closed for $([math]::Round($inactive.TotalHours,2)) hours"
    return $false
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

function Save-ReloadQueue {
    param([array]$Browsers)

    $queue = [PSCustomObject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        Browsers   = $Browsers
    }

    $queue | ConvertTo-Json -Depth 5 | Set-Content -Path $QueueFile -Force -Encoding UTF8
}

try {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload detection started"
    Write-LogEntry "======================================"

    $tracking = Update-BrowserSessionTracking
    $queue = @()

    if (Get-ChromeInstallPath) {
        Write-LogEntry "Evaluating Chrome"
        if ((Test-BrowserSessionAge -BrowserName "Chrome" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) -and
            (Test-ChromePendingUpdate)) {
            $queue += [PSCustomObject]@{
                Browser = "Chrome"
                Reason  = "Pending update and session/inactivity threshold exceeded"
            }
            Write-LogEntry "Chrome added to reload queue"
        }
    }

    if (Get-FirefoxInstallPath) {
        Write-LogEntry "Evaluating Firefox"
        if ((Test-BrowserSessionAge -BrowserName "Firefox" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) -and
            (Test-FirefoxPendingUpdate)) {
            $queue += [PSCustomObject]@{
                Browser = "Firefox"
                Reason  = "Pending update and session/inactivity threshold exceeded"
            }
            Write-LogEntry "Firefox added to reload queue"
        }
    }

    if (Get-EdgeInstallPath) {
        Write-LogEntry "Evaluating Edge"
        if ((Test-BrowserSessionAge -BrowserName "Edge" -TrackingData $tracking -ThresholdHours $InactivityThresholdHours) -and
            (Test-EdgePendingUpdate)) {
            $queue += [PSCustomObject]@{
                Browser = "Edge"
                Reason  = "Pending update and session/inactivity threshold exceeded"
            }
            Write-LogEntry "Edge added to reload queue"
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