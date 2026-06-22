#Requires -Version 5.1

<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-05
Email: Peter.James@nexusos.co.uk
#>

<#
    Author: Peter Opeyemi James
    
    Browser-Update-Detection.ps1
    Detects when browsers have pending updates that require a restart to apply.

    Script must be run as system.
#>

$ErrorActionPreference = 'Stop'

function Resolve-PositiveDoubleSetting {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [double]$DefaultValue,

        [string[]]$Aliases = @()
    )

    $namesToCheck = @($Name) + @($Aliases)

    foreach ($settingName in $namesToCheck) {
        $rawValue = [Environment]::GetEnvironmentVariable($settingName)

        if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
            $parsedValue = 0.0
            $trimmedValue = $rawValue.Trim()

            if ([double]::TryParse(
                    $trimmedValue,
                    [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$parsedValue
                ) -and $parsedValue -gt 0) {
                return [PSCustomObject]@{
                    Value   = $parsedValue
                    Source  = "environment variable '$settingName'"
                    Warning = $null
                }
            }

            return [PSCustomObject]@{
                Value   = $DefaultValue
                Source  = "script default"
                Warning = "Environment variable '$settingName' has invalid value '$trimmedValue'. Using default value $DefaultValue."
            }
        }
    }

    return [PSCustomObject]@{
        Value   = $DefaultValue
        Source  = "script default"
        Warning = $null
    }
}

function Format-HourValue {
    param([double]$Hours)

    if ($Hours -eq [math]::Floor($Hours)) {
        return ([int]$Hours).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }

    return $Hours.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

# =========================
# Configuration
# =========================
$BasePath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = Join-Path $BasePath "Detection.log"
$TrackingFile = Join-Path $BasePath "BrowserUsageTracking.json"
$QueueFile = Join-Path $BasePath "ReloadQueue.json"
$BrowserRunningWindowSetting = Resolve-PositiveDoubleSetting -Name "BrowserRunningWindowHours" -DefaultValue 24 -Aliases @("BrowserReloadThresholdHours")
$BrowserInactiveWindowSetting = Resolve-PositiveDoubleSetting -Name "BrowserInactiveWindowHours" -DefaultValue 24

$BrowserRunningWindowHours = $BrowserRunningWindowSetting.Value
$BrowserInactiveWindowHours = $BrowserInactiveWindowSetting.Value
$BrowserRunningWindowHoursText = Format-HourValue -Hours $BrowserRunningWindowHours
$BrowserInactiveWindowHoursText = Format-HourValue -Hours $BrowserInactiveWindowHours

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


function Test-IsSystemOrAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

        if ($identity.User.Value -eq 'S-1-5-18') {
            return $true
        }

        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Initialize-BrowserUpdateFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }

        if (-not (Test-IsSystemOrAdministrator)) {
            Write-LogEntry "Folder ACL check skipped because the script is not running as SYSTEM or administrator" "Warning"
            return
        }

        $grants = @(
            '*S-1-5-18:(OI)(CI)F',       # SYSTEM
            '*S-1-5-32-544:(OI)(CI)F',   # Administrators
            '*S-1-5-32-545:(OI)(CI)M'    # Builtin Users - needed for logged-in user remediation
        )

        $icaclsOutput = & icacls.exe $Path /inheritance:e /grant $grants /T /C 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-LogEntry "Folder ACL verified for $Path. Logged-in user remediation can write queue, tracking, log, and stable script files."
        }
        else {
            Write-LogEntry "Failed to apply folder ACL to $Path. icacls exit code $LASTEXITCODE. Output: $icaclsOutput" "Warning"
        }
    }
    catch {
        Write-LogEntry "Failed to initialize browser update folder '$Path': $($_.Exception.Message)" "Warning"
    }
}

Initialize-BrowserUpdateFolder -Path $BasePath

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
            elseif ([string]::IsNullOrWhiteSpace($tracking.$browser.LastStop)) {
                $tracking.$browser.LastStop = $now.ToString("o")
                $tracking.$browser.IsRunning = $false
                Write-LogEntry "$browser is closed and had no previous LastStop. LastStop initialized as observation time $($tracking.$browser.LastStop)"
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
        [double]$RunningThresholdHours,
        [double]$InactiveThresholdHours
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

        if ($age.TotalHours -ge $RunningThresholdHours) {
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

    if ($inactive.TotalHours -ge $InactiveThresholdHours) {
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


function ConvertTo-NormalizedVersionString {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $match = [regex]::Match($text, '\d+(?:\.\d+){1,3}')
    if ($match.Success) {
        return $match.Value
    }

    return $null
}

function Get-BrowserRegistryVersion {
    param(
        [string[]]$Paths,
        [string[]]$ValueNames
    )

    foreach ($path in $Paths) {
        try {
            if (-not (Test-Path $path)) {
                continue
            }

            $props = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if (-not $props) {
                continue
            }

            foreach ($valueName in $ValueNames) {
                if ($props.PSObject.Properties.Name -contains $valueName) {
                    $version = ConvertTo-NormalizedVersionString -Value $props.$valueName
                    if ($version) {
                        return [PSCustomObject]@{
                            Version = $version
                            Source  = "$path\$valueName"
                        }
                    }
                }
            }
        }
        catch {}
    }

    return $null
}

function Get-ExecutableVersion {
    param([string]$Path)

    try {
        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) {
            return $null
        }

        $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
        foreach ($candidate in @($info.ProductVersion, $info.FileVersion)) {
            $version = ConvertTo-NormalizedVersionString -Value $candidate
            if ($version) {
                return $version
            }
        }
    }
    catch {
        Write-LogEntry "Failed to read executable version from '$Path': $($_.Exception.Message)" "Warning"
    }

    return $null
}

function Get-HighestVersionFolderVersion {
    param([string]$ExecutablePath)

    try {
        if ([string]::IsNullOrWhiteSpace($ExecutablePath) -or -not (Test-Path $ExecutablePath)) {
            return $null
        }

        $basePath = Split-Path $ExecutablePath -Parent
        $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
            Sort-Object { try { [version]$_.Name } catch { [version]'0.0.0.0' } } -Descending

        $highest = $versionFolders | Select-Object -First 1
        if ($highest) {
            return $highest.Name
        }
    }
    catch {
        Write-LogEntry "Failed to inspect version folders for '$ExecutablePath': $($_.Exception.Message)" "Warning"
    }

    return $null
}

function Get-ChromeVersion {
    try {
        $registryMatch = Get-BrowserRegistryVersion -Paths @(
            "HKLM:\SOFTWARE\Google\Chrome\BLBeacon",
            "HKLM:\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
            "HKCU:\SOFTWARE\Google\Chrome\BLBeacon",
            "HKCU:\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
        ) -ValueNames @("version", "pv", "DisplayVersion")

        if ($registryMatch) {
            Write-LogEntry "Chrome installed version detected via registry $($registryMatch.Source): $($registryMatch.Version)"
            return $registryMatch.Version
        }

        $chromeExe = Get-ChromeInstallPath
        $exeVersion = Get-ExecutableVersion -Path $chromeExe
        if ($exeVersion) {
            Write-LogEntry "Chrome installed version detected via executable '$chromeExe': $exeVersion"
            return $exeVersion
        }

        $folderVersion = Get-HighestVersionFolderVersion -ExecutablePath $chromeExe
        if ($folderVersion) {
            Write-LogEntry "Chrome installed version inferred from highest version folder: $folderVersion"
            return $folderVersion
        }

        Write-LogEntry "Chrome is present but installed version could not be determined from BLBeacon, uninstall registry, executable metadata, or version folders" "Warning"
    }
    catch {
        Write-LogEntry "Error reading Chrome installed version: $($_.Exception.Message)" "Warning"
    }
    return $null
}

function Get-EdgeVersion {
    try {
        $registryMatch = Get-BrowserRegistryVersion -Paths @(
            "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon",
            "HKCU:\SOFTWARE\Microsoft\Edge\BLBeacon",
            "HKCU:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge"
        ) -ValueNames @("version", "pv", "DisplayVersion")

        if ($registryMatch) {
            Write-LogEntry "Edge installed version detected via registry $($registryMatch.Source): $($registryMatch.Version)"
            return $registryMatch.Version
        }

        $edgeExe = Get-EdgeInstallPath
        $exeVersion = Get-ExecutableVersion -Path $edgeExe
        if ($exeVersion) {
            Write-LogEntry "Edge installed version detected via executable '$edgeExe': $exeVersion"
            return $exeVersion
        }

        $folderVersion = Get-HighestVersionFolderVersion -ExecutablePath $edgeExe
        if ($folderVersion) {
            Write-LogEntry "Edge installed version inferred from highest version folder: $folderVersion"
            return $folderVersion
        }

        Write-LogEntry "Edge is present but installed version could not be determined from BLBeacon, uninstall registry, executable metadata, or version folders" "Warning"
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
        $chromeVersion = Get-ChromeVersion
        $updateRegPaths = @(
            "HKLM:\SOFTWARE\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}",
            "HKLM:\SOFTWARE\WOW6432Node\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}",
            "HKCU:\SOFTWARE\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}",
            "HKCU:\SOFTWARE\WOW6432Node\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}"
        )

        foreach ($updateRegPath in $updateRegPaths) {
            if (-not (Test-Path $updateRegPath)) {
                continue
            }

            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue
            if (-not $props) {
                continue
            }

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry "Chrome pending update detected via UpdateAvailable flag at $updateRegPath"
                return $true
            }

            $oldProductVersion = ConvertTo-NormalizedVersionString -Value $props.opv
            if ($oldProductVersion -and $chromeVersion -and $oldProductVersion -ne $chromeVersion) {
                Write-LogEntry "Chrome pending update detected via staged/current version mismatch at $updateRegPath. opv=$oldProductVersion current=$chromeVersion"
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
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
                Sort-Object { try { [version]$_.Name } catch { [version]'0.0.0.0' } } -Descending

            if (@($versionFolders).Count -gt 1) {
                $folderList = (@($versionFolders | Select-Object -ExpandProperty Name) -join ', ')
                Write-LogEntry "Chrome possible pending update detected via multiple version folders: $folderList"
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
        $edgeVersion = Get-EdgeVersion
        $updateRegPaths = @(
            "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}",
            "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}",
            "HKCU:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
        )

        foreach ($updateRegPath in $updateRegPaths) {
            if (-not (Test-Path $updateRegPath)) {
                continue
            }

            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue
            if (-not $props) {
                continue
            }

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry "Edge pending update detected via UpdateAvailable flag at $updateRegPath"
                return $true
            }

            $oldProductVersion = ConvertTo-NormalizedVersionString -Value $props.opv
            if ($oldProductVersion -and $edgeVersion -and $oldProductVersion -ne $edgeVersion) {
                Write-LogEntry "Edge pending update detected via staged/current version mismatch at $updateRegPath. opv=$oldProductVersion current=$edgeVersion"
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
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
                Sort-Object { try { [version]$_.Name } catch { [version]'0.0.0.0' } } -Descending

            if (@($versionFolders).Count -gt 1) {
                $folderList = (@($versionFolders | Select-Object -ExpandProperty Name) -join ', ')
                Write-LogEntry "Edge possible pending update detected via multiple version folders: $folderList"
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
            if (-not ($item.PSObject.Properties.Name -contains 'RemediationMode')) {
                $item | Add-Member -MemberType NoteProperty -Name RemediationMode -Value $null -Force
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

    try {
        $queue = [PSCustomObject]@{
            CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
            Browsers   = $Browsers
        }

        $queue | ConvertTo-Json -Depth 8 | Set-Content -Path $QueueFile -Force -Encoding UTF8
        return $true
    }
    catch {
        Write-LogEntry "Failed to save reload queue '$QueueFile': $($_.Exception.Message)" "Error"
        return $false
    }
}

function Add-OrUpdateQueueItem {
    param(
        [System.Collections.ArrayList]$Queue,
        [array]$ExistingQueue,
        [string]$Browser,
        [string]$Reason,
        [ValidateSet("InteractiveRestart","ClosedUpdateCycle")]
        [string]$RemediationMode = "InteractiveRestart"
    )

    $existingInNewQueue = $Queue | Where-Object { $_.Browser -eq $Browser } | Select-Object -First 1
    if ($existingInNewQueue) {
        $existingInNewQueue.Reason = $Reason

        if (-not ($existingInNewQueue.PSObject.Properties.Name -contains 'RemediationMode')) {
            $existingInNewQueue | Add-Member -MemberType NoteProperty -Name RemediationMode -Value $RemediationMode -Force
        }
        else {
            $existingInNewQueue.RemediationMode = $RemediationMode
        }

        Write-LogEntry "$Browser already queued in current run; updated reason and remediation mode to $RemediationMode"
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
            RemediationMode   = $RemediationMode
        })
        Write-LogEntry "$Browser added to reload queue with preserved postpone state and remediation mode $RemediationMode"
    }
    else {
        [void]$Queue.Add([PSCustomObject]@{
            Browser           = $Browser
            Reason            = $Reason
            PostponeUntilUtc  = $null
            PostponeChoice    = $null
            ScheduledTaskName = $null
            RemediationMode   = $RemediationMode
        })
        Write-LogEntry "$Browser added to reload queue with remediation mode $RemediationMode"
    }
}

try {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload detection started"
    Write-LogEntry "Running window set to $BrowserRunningWindowHoursText hours via $($BrowserRunningWindowSetting.Source)"
    Write-LogEntry "Inactive window set to $BrowserInactiveWindowHoursText hours via $($BrowserInactiveWindowSetting.Source)"

    if ($BrowserRunningWindowSetting.Warning) {
        Write-LogEntry $BrowserRunningWindowSetting.Warning "Warning"
    }

    if ($BrowserInactiveWindowSetting.Warning) {
        Write-LogEntry $BrowserInactiveWindowSetting.Warning "Warning"
    }

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

        $stateInfo = Get-BrowserStateInfo -BrowserName $browser -TrackingData $tracking -RunningThresholdHours $BrowserRunningWindowHours -InactiveThresholdHours $BrowserInactiveWindowHours
        $versionStatus = Get-BrowserVersionStatus -BrowserName $browser

        if (-not $stateInfo.IsRunning -and ($versionStatus.PendingUpdate -or $versionStatus.OutOfDate -eq $true)) {
            $immediateReasonParts = @()
            if ($versionStatus.PendingUpdate) {
                $immediateReasonParts += "pending update"
            }
            if ($versionStatus.OutOfDate -eq $true) {
                $immediateReasonParts += "out of date"
            }

            $immediateReason = "Browser is closed and is " + ($immediateReasonParts -join " and ") + "; queued immediately for open/update/close cycle"
            Add-OrUpdateQueueItem -Queue $queue -ExistingQueue $existingQueue -Browser $browser -Reason $immediateReason -RemediationMode "ClosedUpdateCycle"
            Write-LogEntry "$browser is closed and is $($immediateReasonParts -join ' and '). Inactive window threshold is bypassed for immediate treatment."
            continue
        }

        if (-not $stateInfo.ThresholdMet) {
            if ($stateInfo.IsRunning) {
                Write-LogEntry "$browser does not meet the $BrowserRunningWindowHoursText-hour running window"
            }
            else {
                Write-LogEntry "$browser does not meet the $BrowserInactiveWindowHoursText-hour inactive window and has no pending update or confirmed out-of-date status requiring immediate treatment"
            }
            continue
        }

        if ($stateInfo.IsRunning) {
            if ($versionStatus.PendingUpdate) {
                Add-OrUpdateQueueItem -Queue $queue -ExistingQueue $existingQueue -Browser $browser -Reason "Pending update and browser has been running continuously for at least $BrowserRunningWindowHoursText hours" -RemediationMode "InteractiveRestart"
            }
            else {
                Write-LogEntry "$browser is running for $BrowserRunningWindowHoursText+ hours but no pending update was detected"
            }
        }
        else {
            $needsClosedUpdateCycle = $false
            $reasonParts = @()

            if ($versionStatus.PendingUpdate) {
                $needsClosedUpdateCycle = $true
                $reasonParts += "pending update"
            }

            if ($versionStatus.OutOfDate -eq $true) {
                $needsClosedUpdateCycle = $true
                $reasonParts += "out of date"
            }

            if ($needsClosedUpdateCycle) {
                $reason = "Browser has not been used for at least $BrowserInactiveWindowHoursText hours and is " + ($reasonParts -join " and ") + "; queued for open/update/close cycle"
                Add-OrUpdateQueueItem -Queue $queue -ExistingQueue $existingQueue -Browser $browser -Reason $reason -RemediationMode "ClosedUpdateCycle"
            }
            else {
                Write-LogEntry "$browser has not been used for $BrowserInactiveWindowHoursText+ hours, but no pending update or confirmed out-of-date version was detected"
            }
        }
    }

    if (-not (Save-ReloadQueue -Browsers $queue)) {
        Write-LogEntry "Detection could not save the reload queue. Remediation will not be reliable until folder permissions are repaired." "Error"
        exit 1
    }

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