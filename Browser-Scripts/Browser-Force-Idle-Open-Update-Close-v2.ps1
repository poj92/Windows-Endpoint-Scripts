#Requires -Version 5.1
<#!
Browser-Force-Idle-Open-Close-VerifiedUpdatePage-v7-EnforcedHelpPage.ps1

Purpose
  Closed-browser update worker for Datto RMM.

Design
  - Does NOT download MSI/EXE installers.
  - Does NOT run vendor enterprise installers.
  - Intended to run in the logged-in user context when opening browser UI is required.
  - Processes queued closed-browser items from C:\ProgramData\Datto\BrowserUpdateCheck\ReloadQueue.json.
  - For each queued closed browser:
      1. Confirm the browser is still closed.
      2. Record installed version before action.
      3. Optionally trigger native updater/scheduled update tasks where available.
      4. Open browser to its update/help page using an isolated temporary profile.
      5. Wait for update checks/apply phase.
      6. Close only the browser instance started by this script, using the unique profile path marker.
      7. Optionally trigger native updater again.
      8. Verify installed version after action.
      9. Remove the queue item only when update success is verified.

Exit codes
  0 = completed; no queued closed-browser failures remain from this run
  1 = one or more queued closed-browser items could not be verified as updated
  2 = ReportOnly and action would have been taken
  3 = no supported browsers installed/found
  4 = fatal script error

Notes
  - This script deliberately does not touch browsers that are already running. Your separate user-context
    notification/reload remediation should handle open browsers.
  - For staged/pending updates, an open/update-page/close cycle can be sufficient.
  - For browsers that are merely out of date and have not downloaded/staged an update, the browser vendor's
    own update mechanism may still not complete during the wait window. This script reports that honestly.
  - v6 adds smarter retry bounds: fewer cycles for non-staged updates, per-item cooldown/attempt metadata, and an optional per-run browser cap.
  - v7 adds an optional local DevTools navigation fallback so Chromium browsers are forced to open the update/help page rather than just the home/start page.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$BasePath = 'C:\ProgramData\Datto\BrowserUpdateCheck',
    [string]$QueueFile = '',
    [string[]]$SupportedBrowsers = @('Chrome','Edge'),

    [bool]$ProcessQueuedClosedBrowsers = $true,
    [bool]$EnableIdleNudge = $false,
    [bool]$UseNativeUpdateEngine = $false,
    [bool]$ProcessOutOfDateWithoutPendingUpdate = $true,
    [bool]$RequireVersionChangeForQueueSuccess = $true,
    [bool]$RequireLatestVersionForQueueSuccess = $true,

    [int]$LookbackHours = 24,
    [int]$ClosedBrowserUpdateWaitSeconds = 180,
    [int]$NativeUpdatePreOpenWaitSeconds = 40,
    [int]$NativeUpdatePostCloseWaitSeconds = 60,
    [int]$PostCloseVerificationSeconds = 60,
    [int]$MaxOpenUpdateCloseCycles = 2,
    [int]$MaxOpenUpdateCloseCyclesWhenNoPendingUpdate = 1,
    [int]$InterCycleWaitSeconds = 30,
    [int]$VersionCheckIntervalSeconds = 10,
    [int]$ProcessHeartbeatSeconds = 30,
    [int]$MaxBrowsersPerRun = 0,
    [int]$QueueRetryCooldownHours = 12,
    [int]$MaxClosedWorkerAttempts = 3,
    [bool]$RemoveQueueItemAfterMaxAttempts = $false,
    [bool]$TreatRetryCooldownAsFailure = $false,
    [bool]$EnforceUpdatePageWithDevTools = $true,
    [int]$DevToolsNavigationDelaySeconds = 8,
    [int]$DevToolsNavigationTimeoutSeconds = 25,

    [switch]$ReportOnly
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($QueueFile)) {
    $QueueFile = Join-Path $BasePath 'ReloadQueue.json'
}

$LogFile = Join-Path $BasePath 'ForceIdleOpenClose.log'

function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('Information','Warning','Error')][string]$Level = 'Information'
    )

    $line = '[{0}] [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line

    try {
        if (-not (Test-Path $BasePath)) {
            New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch { }
}

function ConvertTo-BoolSetting {
    param([object]$Value, [bool]$DefaultValue)

    if ($null -eq $Value) { return $DefaultValue }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $DefaultValue }

    switch -Regex ($text.ToLowerInvariant()) {
        '^(1|true|yes|y|on)$' { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default { return $DefaultValue }
    }
}

function Resolve-StringSetting {
    param([string]$Name, [string]$DefaultValue)
    $envValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if ([string]::IsNullOrWhiteSpace($envValue)) { $envValue = [Environment]::GetEnvironmentVariable($Name, 'Machine') }
    if ([string]::IsNullOrWhiteSpace($envValue)) { return $DefaultValue }
    return $envValue
}

function Resolve-IntSetting {
    param([string]$Name, [int]$DefaultValue, [int]$MinValue = 0, [int]$MaxValue = 86400)
    $text = Resolve-StringSetting -Name $Name -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($text)) { return $DefaultValue }
    $value = 0
    if ([int]::TryParse($text.Trim(), [ref]$value)) {
        if ($value -lt $MinValue) { return $DefaultValue }
        if ($value -gt $MaxValue) { return $DefaultValue }
        return $value
    }
    return $DefaultValue
}

function Resolve-BoolSetting {
    param([string]$Name, [bool]$DefaultValue)
    $text = Resolve-StringSetting -Name $Name -DefaultValue ''
    return ConvertTo-BoolSetting -Value $text -DefaultValue $DefaultValue
}

function Resolve-StringArraySetting {
    param([string]$Name, [string[]]$DefaultValue)
    $text = Resolve-StringSetting -Name $Name -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($text)) { return $DefaultValue }
    $items = @($text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($items.Count -eq 0) { return $DefaultValue }
    return $items
}

function Initialize-SettingsFromEnvironment {
    $script:SupportedBrowsers = Resolve-StringArraySetting -Name 'SupportedBrowsers' -DefaultValue $script:SupportedBrowsers
    $script:ProcessQueuedClosedBrowsers = Resolve-BoolSetting -Name 'ProcessQueuedClosedBrowsers' -DefaultValue $script:ProcessQueuedClosedBrowsers
    $script:EnableIdleNudge = Resolve-BoolSetting -Name 'EnableIdleNudge' -DefaultValue $script:EnableIdleNudge
    $script:UseNativeUpdateEngine = Resolve-BoolSetting -Name 'UseNativeUpdateEngine' -DefaultValue $script:UseNativeUpdateEngine
    $script:ProcessOutOfDateWithoutPendingUpdate = Resolve-BoolSetting -Name 'ProcessOutOfDateWithoutPendingUpdate' -DefaultValue $script:ProcessOutOfDateWithoutPendingUpdate
    $script:RequireVersionChangeForQueueSuccess = Resolve-BoolSetting -Name 'RequireVersionChangeForQueueSuccess' -DefaultValue $script:RequireVersionChangeForQueueSuccess
    $script:RequireLatestVersionForQueueSuccess = Resolve-BoolSetting -Name 'RequireLatestVersionForQueueSuccess' -DefaultValue $script:RequireLatestVersionForQueueSuccess
    $script:LookbackHours = Resolve-IntSetting -Name 'LookbackHours' -DefaultValue $script:LookbackHours -MinValue 0 -MaxValue 8760
    $script:ClosedBrowserUpdateWaitSeconds = Resolve-IntSetting -Name 'ClosedBrowserUpdateWaitSeconds' -DefaultValue $script:ClosedBrowserUpdateWaitSeconds -MinValue 10 -MaxValue 3600
    $script:NativeUpdatePreOpenWaitSeconds = Resolve-IntSetting -Name 'NativeUpdatePreOpenWaitSeconds' -DefaultValue $script:NativeUpdatePreOpenWaitSeconds -MinValue 0 -MaxValue 1800
    $script:NativeUpdatePostCloseWaitSeconds = Resolve-IntSetting -Name 'NativeUpdatePostCloseWaitSeconds' -DefaultValue $script:NativeUpdatePostCloseWaitSeconds -MinValue 0 -MaxValue 1800
    $script:PostCloseVerificationSeconds = Resolve-IntSetting -Name 'PostCloseVerificationSeconds' -DefaultValue $script:PostCloseVerificationSeconds -MinValue 0 -MaxValue 1800
    $script:MaxOpenUpdateCloseCycles = Resolve-IntSetting -Name 'MaxOpenUpdateCloseCycles' -DefaultValue $script:MaxOpenUpdateCloseCycles -MinValue 1 -MaxValue 10
    $script:MaxOpenUpdateCloseCyclesWhenNoPendingUpdate = Resolve-IntSetting -Name 'MaxOpenUpdateCloseCyclesWhenNoPendingUpdate' -DefaultValue $script:MaxOpenUpdateCloseCyclesWhenNoPendingUpdate -MinValue 1 -MaxValue 10
    $script:InterCycleWaitSeconds = Resolve-IntSetting -Name 'InterCycleWaitSeconds' -DefaultValue $script:InterCycleWaitSeconds -MinValue 0 -MaxValue 900
    $script:VersionCheckIntervalSeconds = Resolve-IntSetting -Name 'VersionCheckIntervalSeconds' -DefaultValue $script:VersionCheckIntervalSeconds -MinValue 2 -MaxValue 300
    $script:ProcessHeartbeatSeconds = Resolve-IntSetting -Name 'ProcessHeartbeatSeconds' -DefaultValue $script:ProcessHeartbeatSeconds -MinValue 5 -MaxValue 600
    $script:MaxBrowsersPerRun = Resolve-IntSetting -Name 'MaxBrowsersPerRun' -DefaultValue $script:MaxBrowsersPerRun -MinValue 0 -MaxValue 50
    $script:QueueRetryCooldownHours = Resolve-IntSetting -Name 'QueueRetryCooldownHours' -DefaultValue $script:QueueRetryCooldownHours -MinValue 0 -MaxValue 8760
    $script:MaxClosedWorkerAttempts = Resolve-IntSetting -Name 'MaxClosedWorkerAttempts' -DefaultValue $script:MaxClosedWorkerAttempts -MinValue 1 -MaxValue 100
    $script:RemoveQueueItemAfterMaxAttempts = Resolve-BoolSetting -Name 'RemoveQueueItemAfterMaxAttempts' -DefaultValue $script:RemoveQueueItemAfterMaxAttempts
    $script:TreatRetryCooldownAsFailure = Resolve-BoolSetting -Name 'TreatRetryCooldownAsFailure' -DefaultValue $script:TreatRetryCooldownAsFailure
    $script:EnforceUpdatePageWithDevTools = Resolve-BoolSetting -Name 'EnforceUpdatePageWithDevTools' -DefaultValue $script:EnforceUpdatePageWithDevTools
    $script:DevToolsNavigationDelaySeconds = Resolve-IntSetting -Name 'DevToolsNavigationDelaySeconds' -DefaultValue $script:DevToolsNavigationDelaySeconds -MinValue 1 -MaxValue 120
    $script:DevToolsNavigationTimeoutSeconds = Resolve-IntSetting -Name 'DevToolsNavigationTimeoutSeconds' -DefaultValue $script:DevToolsNavigationTimeoutSeconds -MinValue 5 -MaxValue 300
}

function Test-IsSystem {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        return ($identity.User.Value -eq 'S-1-5-18')
    }
    catch { return $false }
}

function Get-ExecutionContextSummary {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        return 'User={0} | SID={1} | SessionId={2} | Process={3}' -f $identity.Name, $identity.User.Value, (Get-Process -Id $PID).SessionId, (Get-Process -Id $PID).ProcessName
    }
    catch {
        return 'Unable to determine execution context: {0}' -f $_.Exception.Message
    }
}

function Get-AppPathExe {
    param([Parameter(Mandatory = $true)][string]$ExeName)

    $subKeys = @(
        "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
    )

    foreach ($root in @([Microsoft.Win32.Registry]::LocalMachine, [Microsoft.Win32.Registry]::CurrentUser)) {
        foreach ($sub in $subKeys) {
            try {
                $key = $root.OpenSubKey($sub)
                if ($key) {
                    $value = $key.GetValue('')
                    $key.Close()
                    if ($value -and (Test-Path -LiteralPath $value)) { return $value }
                }
            }
            catch { }
        }
    }

    return $null
}

function Get-BrowserDefinitions {
    $programFiles = [Environment]::GetFolderPath('ProgramFiles')
    $programFilesX86 = ${env:ProgramFiles(x86)}
    if ([string]::IsNullOrWhiteSpace($programFilesX86)) { $programFilesX86 = $programFiles }

    return @(
        [pscustomobject]@{
            Name = 'Edge'
            Type = 'Chromium'
            ExeLeaf = 'msedge.exe'
            ProcessNames = @('msedge')
            CloseExeNames = @('msedge.exe')
            Candidates = @(
                (Join-Path $programFiles 'Microsoft\Edge\Application\msedge.exe'),
                (Join-Path $programFilesX86 'Microsoft\Edge\Application\msedge.exe')
            )
            UpdateUrl = 'edge://settings/help'
            PrefetchBases = @('msedge')
        },
        [pscustomobject]@{
            Name = 'Chrome'
            Type = 'Chromium'
            ExeLeaf = 'chrome.exe'
            ProcessNames = @('chrome')
            CloseExeNames = @('chrome.exe')
            Candidates = @(
                (Join-Path $programFiles 'Google\Chrome\Application\chrome.exe'),
                (Join-Path $programFilesX86 'Google\Chrome\Application\chrome.exe')
            )
            UpdateUrl = 'chrome://settings/help'
            PrefetchBases = @('chrome')
        },
        [pscustomobject]@{
            Name = 'Firefox'
            Type = 'Firefox'
            ExeLeaf = 'firefox.exe'
            ProcessNames = @('firefox')
            CloseExeNames = @('firefox.exe')
            Candidates = @(
                (Join-Path $programFiles 'Mozilla Firefox\firefox.exe'),
                (Join-Path $programFilesX86 'Mozilla Firefox\firefox.exe')
            )
            UpdateUrl = 'about:preferences'
            PrefetchBases = @('firefox')
        },
        [pscustomobject]@{
            Name = 'Brave'
            Type = 'Chromium'
            ExeLeaf = 'brave.exe'
            ProcessNames = @('brave')
            CloseExeNames = @('brave.exe')
            Candidates = @(
                (Join-Path $programFiles 'BraveSoftware\Brave-Browser\Application\brave.exe'),
                (Join-Path $programFilesX86 'BraveSoftware\Brave-Browser\Application\brave.exe')
            )
            UpdateUrl = 'brave://settings/help'
            PrefetchBases = @('brave')
        },
        [pscustomobject]@{
            Name = 'Opera'
            Type = 'Chromium'
            ExeLeaf = 'launcher.exe'
            ProcessNames = @('opera','launcher')
            CloseExeNames = @('opera.exe','launcher.exe')
            Candidates = @(
                (Join-Path $programFiles 'Opera\launcher.exe'),
                (Join-Path $programFilesX86 'Opera\launcher.exe')
            )
            UpdateUrl = 'opera://update'
            PrefetchBases = @('opera','launcher')
        }
    )
}

function Find-InstalledBrowsers {
    param([string[]]$Names)

    $definitions = Get-BrowserDefinitions
    $installed = @()

    foreach ($name in $Names) {
        $definition = $definitions | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if (-not $definition) {
            Write-LogEntry "Unsupported browser requested: $name" 'Warning'
            continue
        }

        $exePath = $null
        foreach ($candidate in $definition.Candidates) {
            if ($candidate -and (Test-Path -LiteralPath $candidate)) {
                $exePath = $candidate
                break
            }
        }

        if (-not $exePath) {
            $exePath = Get-AppPathExe -ExeName $definition.ExeLeaf
        }

        if ($exePath) {
            $installed += [pscustomobject]@{
                Name = $definition.Name
                Type = $definition.Type
                ExeLeaf = $definition.ExeLeaf
                ExePath = $exePath
                ProcessNames = $definition.ProcessNames
                CloseExeNames = $definition.CloseExeNames
                UpdateUrl = $definition.UpdateUrl
                PrefetchBases = $definition.PrefetchBases
            }
        }
    }

    return $installed
}

function Test-BrowserRunning {
    param([Parameter(Mandatory = $true)][string[]]$ProcessNames)

    foreach ($processName in $ProcessNames) {
        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) { return $true }
    }
    return $false
}

function Get-RegistryValueSafe {
    param([string]$Path, [string]$Name)
    try {
        if (Test-Path -LiteralPath $Path) {
            $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $item -and $null -ne $item.$Name) { return [string]$item.$Name }
        }
    }
    catch { }
    return $null
}

function Get-HighestVersionFolder {
    param([string]$ApplicationDirectory)

    try {
        if (-not (Test-Path -LiteralPath $ApplicationDirectory)) { return $null }
        $folders = Get-ChildItem -Path $ApplicationDirectory -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d+(\.\d+){1,4}$' }
        if (-not $folders) { return $null }
        $best = $folders | Sort-Object { try { [version]$_.Name } catch { [version]'0.0.0.0' } } -Descending | Select-Object -First 1
        if ($best) { return $best.Name }
    }
    catch { }

    return $null
}

function Get-UninstallDisplayVersion {
    param([string]$BrowserName)

    $namePatterns = switch ($BrowserName) {
        'Chrome'  { @('Google Chrome') }
        'Edge'    { @('Microsoft Edge') }
        'Firefox' { @('Mozilla Firefox') }
        'Brave'   { @('Brave') }
        'Opera'   { @('Opera') }
        default   { @($BrowserName) }
    }

    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($path in $paths) {
        try {
            $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if (-not $item.DisplayName -or -not $item.DisplayVersion) { continue }
                foreach ($pattern in $namePatterns) {
                    if ($item.DisplayName -like "*$pattern*") { return [string]$item.DisplayVersion }
                }
            }
        }
        catch { }
    }

    return $null
}

function Test-IsMachineLevelBrowserPath {
    param([string]$ExePath)

    if ([string]::IsNullOrWhiteSpace($ExePath)) { return $false }

    try {
        $expanded = [Environment]::ExpandEnvironmentVariables($ExePath).Trim()
        $pf = [Environment]::GetFolderPath('ProgramFiles')
        $pf86 = [Environment]::GetFolderPath('ProgramFilesX86')

        if (-not [string]::IsNullOrWhiteSpace($pf) -and $expanded.StartsWith($pf, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        if (-not [string]::IsNullOrWhiteSpace($pf86) -and $expanded.StartsWith($pf86, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    catch { }

    return $false
}

function Get-BrowserRegistryVersionCandidates {
    param([string]$BrowserName, [bool]$MachineLevelInstall)

    $hklm = @()
    $hkcu = @()

    switch ($BrowserName) {
        'Chrome' {
            $hklm = @(
                @{ Path = 'HKLM:\SOFTWARE\Google\Chrome\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon'; Name = 'version' }
            )
            $hkcu = @(
                @{ Path = 'HKCU:\SOFTWARE\Google\Chrome\BLBeacon'; Name = 'version' },
                @{ Path = 'HKCU:\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon'; Name = 'version' }
            )
        }
        'Edge' {
            $hklm = @(
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon'; Name = 'version' }
            )
            $hkcu = @(
                @{ Path = 'HKCU:\SOFTWARE\Microsoft\Edge\BLBeacon'; Name = 'version' },
                @{ Path = 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon'; Name = 'version' }
            )
        }
        'Brave' {
            $hklm = @(
                @{ Path = 'HKLM:\SOFTWARE\BraveSoftware\Brave-Browser\BLBeacon'; Name = 'version' }
            )
            $hkcu = @(
                @{ Path = 'HKCU:\SOFTWARE\BraveSoftware\Brave-Browser\BLBeacon'; Name = 'version' }
            )
        }
    }

    if ($MachineLevelInstall) { return @($hklm + $hkcu) }
    return @($hkcu + $hklm)
}

function Get-BrowserFileVersionInfo {
    param([Parameter(Mandatory = $true)]$Browser)

    try {
        if ($Browser.ExePath -and (Test-Path -LiteralPath $Browser.ExePath)) {
            $fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($Browser.ExePath).ProductVersion
            if ([string]::IsNullOrWhiteSpace($fileVersion)) {
                $fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($Browser.ExePath).FileVersion
            }
            if (-not [string]::IsNullOrWhiteSpace($fileVersion)) {
                $clean = ($fileVersion -replace '\s.*$', '').Trim()
                return [pscustomobject]@{ Version = $clean; Source = "file version $($Browser.ExePath)" }
            }
        }
    }
    catch { }

    return [pscustomobject]@{ Version = $null; Source = 'file version not detected' }
}

function Get-BrowserHighestFolderVersionInfo {
    param([Parameter(Mandatory = $true)]$Browser)

    try {
        if ($Browser.ExePath) {
            $applicationDirectory = Split-Path -Parent $Browser.ExePath
            $folderVersion = Get-HighestVersionFolder -ApplicationDirectory $applicationDirectory
            if (-not [string]::IsNullOrWhiteSpace($folderVersion)) {
                return [pscustomobject]@{ Version = $folderVersion; Source = "highest application folder $applicationDirectory" }
            }
        }
    }
    catch { }

    return [pscustomobject]@{ Version = $null; Source = 'highest application folder not detected' }
}

function Get-BrowserInstalledVersion {
    param([Parameter(Mandatory = $true)]$Browser)

    $browserName = $Browser.Name
    $machineLevel = Test-IsMachineLevelBrowserPath -ExePath $Browser.ExePath

    # HKCU BLBeacon can be stale or user-specific for machine-level installs. For Program Files installs,
    # prefer machine/app sources first and use HKCU only as a last fallback.
    if ($machineLevel) {
        foreach ($candidate in (Get-BrowserRegistryVersionCandidates -BrowserName $browserName -MachineLevelInstall $true | Where-Object { $_.Path -like 'HKLM:*' })) {
            $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return [pscustomobject]@{ Version = $value.Trim(); Source = ('registry {0}:{1}' -f $candidate.Path, $candidate.Name) }
            }
        }

        $uninstallVersion = Get-UninstallDisplayVersion -BrowserName $browserName
        if (-not [string]::IsNullOrWhiteSpace($uninstallVersion)) {
            return [pscustomobject]@{ Version = $uninstallVersion.Trim(); Source = 'uninstall DisplayVersion' }
        }

        $fileInfo = Get-BrowserFileVersionInfo -Browser $Browser
        if (-not [string]::IsNullOrWhiteSpace($fileInfo.Version)) { return $fileInfo }

        $folderInfo = Get-BrowserHighestFolderVersionInfo -Browser $Browser
        if (-not [string]::IsNullOrWhiteSpace($folderInfo.Version)) { return $folderInfo }

        foreach ($candidate in (Get-BrowserRegistryVersionCandidates -BrowserName $browserName -MachineLevelInstall $true | Where-Object { $_.Path -like 'HKCU:*' })) {
            $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return [pscustomobject]@{ Version = $value.Trim(); Source = ('registry fallback {0}:{1}' -f $candidate.Path, $candidate.Name) }
            }
        }
    }
    else {
        foreach ($candidate in (Get-BrowserRegistryVersionCandidates -BrowserName $browserName -MachineLevelInstall $false)) {
            $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return [pscustomobject]@{ Version = $value.Trim(); Source = ('registry {0}:{1}' -f $candidate.Path, $candidate.Name) }
            }
        }

        $fileInfo = Get-BrowserFileVersionInfo -Browser $Browser
        if (-not [string]::IsNullOrWhiteSpace($fileInfo.Version)) { return $fileInfo }

        $folderInfo = Get-BrowserHighestFolderVersionInfo -Browser $Browser
        if (-not [string]::IsNullOrWhiteSpace($folderInfo.Version)) { return $folderInfo }

        $uninstallVersion = Get-UninstallDisplayVersion -BrowserName $browserName
        if (-not [string]::IsNullOrWhiteSpace($uninstallVersion)) {
            return [pscustomobject]@{ Version = $uninstallVersion.Trim(); Source = 'uninstall DisplayVersion' }
        }
    }

    return [pscustomobject]@{ Version = $null; Source = 'not detected' }
}

function Compare-VersionText {
    param([string]$Left, [string]$Right)

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) { return $null }
    try {
        $leftVersion = [version](($Left -replace '[^0-9\.]', '').Trim('.'))
        $rightVersion = [version](($Right -replace '[^0-9\.]', '').Trim('.'))
        return $leftVersion.CompareTo($rightVersion)
    }
    catch {
        return [string]::Compare($Left, $Right, $true)
    }
}


function Get-QueueTargetVersion {
    param([object]$QueueItem)

    if ($null -eq $QueueItem) { return $null }
    foreach ($name in @('LatestVersion','TargetVersion','RequiredVersion','Latest')) {
        try {
            if ($QueueItem.PSObject.Properties.Name -contains $name) {
                $value = [string]$QueueItem.$name
                if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
            }
        }
        catch { }
    }

    return $null
}

function Test-VersionMeetsTarget {
    param([string]$Version, [string]$TargetVersion)

    if ([string]::IsNullOrWhiteSpace($Version) -or [string]::IsNullOrWhiteSpace($TargetVersion)) { return $null }
    $comparison = Compare-VersionText -Left $Version -Right $TargetVersion
    if ($null -eq $comparison) { return $null }
    return ($comparison -ge 0)
}

function New-TargetAwareUpdateResult {
    param(
        [Parameter(Mandatory = $true)]$Browser,
        [string]$BeforeVersion,
        [string]$AfterVersion,
        [string]$TargetVersion,
        [string]$Phase,
        [string]$SuccessReason
    )

    $changed = $false
    if (-not [string]::IsNullOrWhiteSpace($BeforeVersion) -and -not [string]::IsNullOrWhiteSpace($AfterVersion)) {
        $comparisonToBefore = Compare-VersionText -Left $AfterVersion -Right $BeforeVersion
        if ($null -ne $comparisonToBefore -and $comparisonToBefore -gt 0) { $changed = $true }
    }

    if ($RequireLatestVersionForQueueSuccess -and -not [string]::IsNullOrWhiteSpace($TargetVersion)) {
        $meetsTarget = Test-VersionMeetsTarget -Version $AfterVersion -TargetVersion $TargetVersion
        if ($meetsTarget -eq $true) {
            Write-LogEntry "$($Browser.Name) reached target/latest version during $Phase. Target=$TargetVersion Before=$BeforeVersion After=$AfterVersion"
            return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = $SuccessReason; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
        }

        if ($changed) {
            Write-LogEntry "$($Browser.Name) version changed during $Phase, but it is still below target/latest. Target=$TargetVersion Before=$BeforeVersion After=$AfterVersion. Leaving queued for another pass." 'Warning'
            return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'UpdatedButStillBelowTarget'; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
        }

        Write-LogEntry "$($Browser.Name) did not reach target/latest during $Phase. Target=$TargetVersion Before=$BeforeVersion After=$AfterVersion" 'Warning'
        return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'TargetNotReached'; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
    }

    if ($changed) {
        Write-LogEntry "$($Browser.Name) update verified by version change during $Phase. Before=$BeforeVersion After=$AfterVersion"
        return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = $SuccessReason; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
    }

    if (-not $RequireVersionChangeForQueueSuccess) {
        Write-LogEntry "$($Browser.Name) version change was not required for queue success. Before=$BeforeVersion After=$AfterVersion" 'Warning'
        return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = 'CompletedNoVersionRequirement'; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
    }

    return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'NoVersionChange'; Before = $BeforeVersion; After = $AfterVersion; Target = $TargetVersion }
}

function Get-PrefetchLastRun {
    param([Parameter(Mandatory = $true)][string]$ExeBaseName)

    $prefetchDir = Join-Path $env:SystemRoot 'Prefetch'
    if (-not (Test-Path -LiteralPath $prefetchDir)) { return $null }

    $pattern = ('{0}.EXE-*.pf' -f $ExeBaseName.ToUpperInvariant())
    $files = Get-ChildItem -Path $prefetchDir -Filter $pattern -ErrorAction SilentlyContinue
    if (-not $files) { return $null }

    return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
}

function Get-LastOpenedWhenNotRunning {
    param([Parameter(Mandatory = $true)]$Browser)

    $prefTimes = @()
    foreach ($base in $Browser.PrefetchBases) {
        try {
            $prefetchTime = Get-PrefetchLastRun -ExeBaseName $base
            if ($prefetchTime) { $prefTimes += $prefetchTime }
        }
        catch { }
    }

    if ($prefTimes.Count -gt 0) {
        return [pscustomobject]@{ Time = ($prefTimes | Sort-Object -Descending | Select-Object -First 1); Source = 'Prefetch' }
    }

    return [pscustomobject]@{ Time = $null; Source = 'None' }
}

function New-UniqueProfileDir {
    param([Parameter(Mandatory = $true)][string]$BrowserName)

    $root = Join-Path $env:ProgramData 'BrowserNudge'
    $browserRoot = Join-Path $root $BrowserName
    $profile = Join-Path $browserRoot ([guid]::NewGuid().ToString('N'))

    if (-not (Test-Path -LiteralPath $browserRoot)) {
        New-Item -Path $browserRoot -ItemType Directory -Force | Out-Null
    }
    New-Item -Path $profile -ItemType Directory -Force | Out-Null
    return $profile
}

function Get-LaunchArguments {
    param(
        [Parameter(Mandatory = $true)]$Browser,
        [Parameter(Mandatory = $true)][string]$ProfileDir,
        [Parameter(Mandatory = $true)][string]$Url,
        [switch]$NavigateOnly,
        [int]$DevToolsPort = 0
    )

    if ($Browser.Type -eq 'Chromium') {
        $commonArgs = @(
            "--user-data-dir=`"$ProfileDir`"",
            '--profile-directory=Default'
        )

        if ($DevToolsPort -gt 0) {
            $commonArgs += "--remote-debugging-port=$DevToolsPort"
            $commonArgs += '--remote-allow-origins=*'
        }

        $args = @($commonArgs + @(
            '--no-first-run',
            '--no-default-browser-check',
            '--disable-features=ChromeWhatsNewUI,DefaultBrowserSettingEnabled,PromotionalTabsEnabled',
            '--new-window',
            "`"$Url`""
        ))

        if ($NavigateOnly) {
            $args = @($commonArgs + @("`"$Url`""))
        }

        return $args -join ' '
    }

    if ($Browser.Type -eq 'Firefox') {
        return @(
            '-no-remote',
            '-profile', "`"$ProfileDir`"",
            '-new-window',
            "`"$Url`""
        ) -join ' '
    }

    return "`"$Url`""
}


function Get-FreeTcpPort {
    try {
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
        $listener.Start()
        $port = $listener.LocalEndpoint.Port
        $listener.Stop()
        return [int]$port
    }
    catch {
        return 0
    }
}

function Invoke-DevToolsUpdatePageNavigation {
    param(
        [Parameter(Mandatory = $true)][int]$Port,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$BrowserName,
        [int]$TimeoutSeconds = 25
    )

    if ($Port -le 0) { return $false }

    $baseUri = "http://127.0.0.1:$Port"
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $ready = $false

    while ((Get-Date) -lt $deadline) {
        try {
            Invoke-RestMethod -Uri "$baseUri/json/version" -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop | Out-Null
            $ready = $true
            break
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }

    if (-not $ready) {
        Write-LogEntry "$BrowserName local DevTools endpoint did not become ready on port $Port within $TimeoutSeconds seconds. The normal command-line URL launch was still attempted." 'Warning'
        return $false
    }

    $encodedUrl = [System.Uri]::EscapeDataString($Url)
    $candidateUris = @(
        "$baseUri/json/new?$encodedUrl",
        "$baseUri/json/new?$Url"
    )

    foreach ($candidateUri in $candidateUris) {
        foreach ($method in @('PUT','GET')) {
            try {
                Invoke-RestMethod -Uri $candidateUri -Method $method -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop | Out-Null
                Write-LogEntry "Forced $BrowserName update/help page navigation through local DevTools endpoint: $Url"
                return $true
            }
            catch {
                # Try the next method/URI form. Chrome/Edge builds differ on the accepted verb for /json/new.
            }
        }
    }

    Write-LogEntry "Could not force $BrowserName update/help page through local DevTools. The normal command-line URL launch was still attempted." 'Warning'
    return $false
}

function Stop-OnlyOurBrowserProcesses {
    param(
        [Parameter(Mandatory = $true)][string[]]$ProcessExeNames,
        [Parameter(Mandatory = $true)][string]$MarkerText,
        [int]$GraceSeconds = 10
    )

    $matched = @()
    foreach ($exeName in $ProcessExeNames) {
        try {
            $escaped = $exeName.Replace("'", "''")
            $processes = Get-CimInstance Win32_Process -Filter "Name='$escaped'" -ErrorAction SilentlyContinue
            foreach ($process in $processes) {
                if ($process.CommandLine -and ($process.CommandLine -like "*$MarkerText*")) {
                    $matched += $process
                }
            }
        }
        catch { }
    }

    if (-not $matched -or $matched.Count -eq 0) {
        Write-LogEntry "No matching browser processes found with marker '$MarkerText' during close step" 'Warning'
        return 0
    }

    Write-LogEntry "Closing $($matched.Count) process(es) started by this script using isolated profile marker"

    foreach ($process in $matched) {
        try {
            $managed = Get-Process -Id $process.ProcessId -ErrorAction SilentlyContinue
            if ($managed) {
                try { $managed.CloseMainWindow() | Out-Null } catch { }
            }
        }
        catch { }
    }

    Start-Sleep -Seconds $GraceSeconds

    $killed = 0
    foreach ($process in $matched) {
        try {
            $managed = Get-Process -Id $process.ProcessId -ErrorAction SilentlyContinue
            if ($managed) {
                Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
                $killed++
            }
        }
        catch { }
    }

    if ($killed -gt 0) {
        Write-LogEntry "Force-terminated $killed remaining process(es) started by this script" 'Warning'
    }

    return $matched.Count
}

function Wait-WithHeartbeat {
    param([int]$Seconds, [string]$Activity)

    if ($Seconds -le 0) { return }

    $remaining = $Seconds
    $elapsed = 0
    while ($remaining -gt 0) {
        $sleep = [Math]::Min($ProcessHeartbeatSeconds, $remaining)
        Start-Sleep -Seconds $sleep
        $elapsed += $sleep
        $remaining -= $sleep
        if ($remaining -gt 0) {
            Write-LogEntry "$Activity still waiting. ElapsedSeconds=$elapsed RemainingSeconds=$remaining"
        }
    }
}

function Wait-ForVersionChange {
    param(
        [Parameter(Mandatory = $true)]$Browser,
        [string]$OriginalVersion,
        [int]$TimeoutSeconds,
        [string]$Phase
    )

    $lastSeen = $OriginalVersion
    if ($TimeoutSeconds -le 0) {
        $info = Get-BrowserInstalledVersion -Browser $Browser
        return [pscustomobject]@{ Changed = $false; Version = $info.Version; Source = $info.Source }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $VersionCheckIntervalSeconds
        $info = Get-BrowserInstalledVersion -Browser $Browser
        $lastSeen = $info.Version
        Write-LogEntry "$($Browser.Name) version check during $Phase. Current=$($info.Version) Source=$($info.Source)"

        if (-not [string]::IsNullOrWhiteSpace($OriginalVersion) -and -not [string]::IsNullOrWhiteSpace($info.Version)) {
            $comparison = Compare-VersionText -Left $info.Version -Right $OriginalVersion
            if ($null -ne $comparison -and $comparison -gt 0) {
                return [pscustomobject]@{ Changed = $true; Version = $info.Version; Source = $info.Source }
            }
        }
    }

    return [pscustomobject]@{ Changed = $false; Version = $lastSeen; Source = 'last observed during wait' }
}

function Get-NativeUpdaterCandidates {
    param([Parameter(Mandatory = $true)][string]$BrowserName)

    $candidates = @()
    $programFiles = [Environment]::GetFolderPath('ProgramFiles')
    $programFilesX86 = ${env:ProgramFiles(x86)}
    if ([string]::IsNullOrWhiteSpace($programFilesX86)) { $programFilesX86 = $programFiles }

    switch ($BrowserName) {
        'Chrome' {
            foreach ($path in @(
                (Join-Path $programFiles 'Google\Update\GoogleUpdate.exe'),
                (Join-Path $programFilesX86 'Google\Update\GoogleUpdate.exe'),
                (Join-Path $env:LOCALAPPDATA 'Google\Update\GoogleUpdate.exe')
            )) {
                if ($path -and (Test-Path -LiteralPath $path)) {
                    $candidates += [pscustomobject]@{ Path = $path; Arguments = '/ua /installsource scheduler'; Description = 'Chrome native update engine' }
                }
            }
        }
        'Edge' {
            foreach ($path in @(
                (Join-Path $programFiles 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'),
                (Join-Path $programFilesX86 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'),
                (Join-Path $env:LOCALAPPDATA 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe')
            )) {
                if ($path -and (Test-Path -LiteralPath $path)) {
                    $candidates += [pscustomobject]@{ Path = $path; Arguments = '/ua /installsource scheduler'; Description = 'Edge native update engine' }
                }
            }
        }
    }

    return $candidates
}

function Start-UpdateScheduledTasks {
    param([Parameter(Mandatory = $true)][string]$BrowserName)

    $patterns = switch ($BrowserName) {
        'Chrome' { @('*GoogleUpdateTask*') }
        'Edge'   { @('*MicrosoftEdgeUpdateTask*') }
        default  { @() }
    }

    if ($patterns.Count -eq 0) { return 0 }

    $started = 0
    foreach ($pattern in $patterns) {
        try {
            $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
            foreach ($task in $tasks) {
                try {
                    Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop
                    Write-LogEntry "Started scheduled update task for ${BrowserName}: $($task.TaskPath)$($task.TaskName)"
                    $started++
                }
                catch {
                    Write-LogEntry "Could not start scheduled update task for ${BrowserName}: $($task.TaskPath)$($task.TaskName). $($_.Exception.Message)" 'Warning'
                }
            }
        }
        catch { }
    }

    if ($started -eq 0) {
        Write-LogEntry "No scheduled update tasks were started for $BrowserName" 'Warning'
    }

    return $started
}

function Invoke-NativeUpdateEngine {
    param([Parameter(Mandatory = $true)]$Browser, [string]$Phase)

    if (-not $UseNativeUpdateEngine) {
        Write-LogEntry "Native update engine trigger disabled. Skipping $($Browser.Name) $Phase trigger."
        return
    }

    Write-LogEntry "Triggering native updater for $($Browser.Name) during $Phase. No installer download will be performed."
    Start-UpdateScheduledTasks -BrowserName $Browser.Name | Out-Null

    $candidates = Get-NativeUpdaterCandidates -BrowserName $Browser.Name
    if (-not $candidates -or $candidates.Count -eq 0) {
        Write-LogEntry "No native updater executable found for $($Browser.Name)" 'Warning'
        return
    }

    foreach ($candidate in $candidates | Select-Object -First 1) {
        try {
            Write-LogEntry "Starting update trigger: $($candidate.Description) $($candidate.Arguments)"
            $process = Start-Process -FilePath $candidate.Path -ArgumentList $candidate.Arguments -WindowStyle Hidden -PassThru -ErrorAction Stop
            $waited = $process.WaitForExit(30000)
            if ($waited) {
                Write-LogEntry "Update trigger exited with code $($process.ExitCode): $($candidate.Description)"
            }
            else {
                Write-LogEntry "Update trigger still running after 30 seconds; continuing: $($candidate.Description)" 'Warning'
            }
        }
        catch {
            Write-LogEntry "Failed to start native update trigger for $($Browser.Name): $($_.Exception.Message)" 'Warning'
        }
    }
}

function Read-ReloadQueue {
    if (-not (Test-Path -LiteralPath $QueueFile)) {
        return [pscustomobject]@{ CreatedUtc = (Get-Date).ToUniversalTime().ToString('o'); Browsers = @() }
    }

    try {
        $raw = Get-Content -Path $QueueFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return [pscustomobject]@{ CreatedUtc = (Get-Date).ToUniversalTime().ToString('o'); Browsers = @() }
        }
        $queue = $raw | ConvertFrom-Json
        if ($null -eq $queue.Browsers) { $queue | Add-Member -MemberType NoteProperty -Name Browsers -Value @() -Force }
        return $queue
    }
    catch {
        Write-LogEntry "Failed to read queue file '$QueueFile': $($_.Exception.Message)" 'Error'
        return [pscustomobject]@{ CreatedUtc = (Get-Date).ToUniversalTime().ToString('o'); Browsers = @() }
    }
}

function Save-ReloadQueue {
    param([array]$Browsers)

    try {
        $queue = [pscustomobject]@{
            CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
            Browsers = @($Browsers)
        }
        $queue | ConvertTo-Json -Depth 8 | Set-Content -Path $QueueFile -Force -Encoding UTF8
        Write-LogEntry "Queue file saved successfully. RemainingItems=$(@($Browsers).Count)"
        return $true
    }
    catch {
        Write-LogEntry "Failed to save queue file '$QueueFile': $($_.Exception.Message)" 'Error'
        return $false
    }
}

function Test-IsClosedQueueItem {
    param([Parameter(Mandatory = $true)]$Item)

    $mode = [string]$Item.RemediationMode
    $reason = [string]$Item.Reason

    if ($mode -ieq 'ClosedUpdateCycle') { return $true }
    if ($reason -match 'closed' -and $reason -match 'out of date|pending update|update') { return $true }
    return $false
}


function Get-QueueStringProperty {
    param(
        [Parameter(Mandatory = $true)]$QueueItem,
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$DefaultValue = ''
    )

    try {
        $prop = $QueueItem.PSObject.Properties[$Name]
        if ($null -eq $prop -or $null -eq $prop.Value) { return $DefaultValue }
        return [string]$prop.Value
    }
    catch { return $DefaultValue }
}

function Get-QueueIntProperty {
    param(
        [Parameter(Mandatory = $true)]$QueueItem,
        [Parameter(Mandatory = $true)][string]$Name,
        [int]$DefaultValue = 0
    )

    try {
        $text = Get-QueueStringProperty -QueueItem $QueueItem -Name $Name -DefaultValue ''
        $value = 0
        if ([int]::TryParse($text, [ref]$value)) { return $value }
        return $DefaultValue
    }
    catch { return $DefaultValue }
}

function Set-QueueItemProperty {
    param(
        [Parameter(Mandatory = $true)]$QueueItem,
        [Parameter(Mandatory = $true)][string]$Name,
        $Value
    )

    try {
        $QueueItem | Add-Member -MemberType NoteProperty -Name $Name -Value $Value -Force
    }
    catch { }

    return $QueueItem
}

function Test-QueueRetryCooldownActive {
    param([Parameter(Mandatory = $true)]$QueueItem)

    if ($QueueRetryCooldownHours -le 0) {
        return [pscustomobject]@{ Active = $false; UntilUtc = $null; RemainingMinutes = 0 }
    }

    $lastAttemptText = Get-QueueStringProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastAttemptUtc' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($lastAttemptText)) {
        return [pscustomobject]@{ Active = $false; UntilUtc = $null; RemainingMinutes = 0 }
    }

    try {
        $lastAttempt = [datetime]::Parse($lastAttemptText).ToUniversalTime()
        $until = $lastAttempt.AddHours($QueueRetryCooldownHours)
        $now = (Get-Date).ToUniversalTime()
        if ($now -lt $until) {
            $remaining = [math]::Ceiling(($until - $now).TotalMinutes)
            return [pscustomobject]@{ Active = $true; UntilUtc = $until.ToString('o'); RemainingMinutes = $remaining }
        }
    }
    catch { }

    return [pscustomobject]@{ Active = $false; UntilUtc = $null; RemainingMinutes = 0 }
}

function Stamp-QueueAttemptResult {
    param(
        [Parameter(Mandatory = $true)]$QueueItem,
        [Parameter(Mandatory = $true)]$Result
    )

    $attempts = Get-QueueIntProperty -QueueItem $QueueItem -Name 'ClosedWorkerAttemptCount' -DefaultValue 0
    $attempts++

    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerAttemptCount' -Value $attempts
    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastAttemptUtc' -Value ((Get-Date).ToUniversalTime().ToString('o'))
    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastResult' -Value ([string]$Result.Reason)
    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastBeforeVersion' -Value ([string]$Result.Before)
    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastAfterVersion' -Value ([string]$Result.After)
    $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerLastTargetVersion' -Value ([string]$Result.Target)

    if ($QueueRetryCooldownHours -gt 0) {
        $next = (Get-Date).ToUniversalTime().AddHours($QueueRetryCooldownHours).ToString('o')
        $QueueItem = Set-QueueItemProperty -QueueItem $QueueItem -Name 'ClosedWorkerNextRetryUtc' -Value $next
    }

    return $QueueItem
}

function Invoke-OpenUpdateCloseCycle {
    param([Parameter(Mandatory = $true)]$Browser)

    $profileDir = $null
    try {
        $profileDir = New-UniqueProfileDir -BrowserName $Browser.Name
        $devToolsPort = 0
        if ($Browser.Type -eq 'Chromium' -and $EnforceUpdatePageWithDevTools) {
            $devToolsPort = Get-FreeTcpPort
            if ($devToolsPort -gt 0) {
                Write-LogEntry "$($Browser.Name) local DevTools navigation enforcement enabled on port $devToolsPort"
            }
            else {
                Write-LogEntry "$($Browser.Name) local DevTools navigation enforcement requested, but no free TCP port could be allocated. Falling back to normal URL launch." 'Warning'
            }
        }

        $arguments = Get-LaunchArguments -Browser $Browser -ProfileDir $profileDir -Url $Browser.UpdateUrl -DevToolsPort $devToolsPort

        Write-LogEntry "Opening $($Browser.Name) to update/help page '$($Browser.UpdateUrl)' using isolated profile. WaitSeconds=$ClosedBrowserUpdateWaitSeconds"
        Write-LogEntry "Executable: $($Browser.ExePath)"
        Write-LogEntry "Launch arguments: $arguments"

        Start-Process -FilePath $Browser.ExePath -ArgumentList $arguments -ErrorAction Stop | Out-Null

        if ($Browser.Type -eq 'Chromium') {
            Start-Sleep -Seconds $DevToolsNavigationDelaySeconds
            $secondNavigationArguments = Get-LaunchArguments -Browser $Browser -ProfileDir $profileDir -Url $Browser.UpdateUrl -NavigateOnly -DevToolsPort $devToolsPort
            Write-LogEntry "Re-issued $($Browser.Name) update/help URL after initial launch to bypass first-run/new-profile pages: $($Browser.UpdateUrl)"
            Start-Process -FilePath $Browser.ExePath -ArgumentList $secondNavigationArguments -ErrorAction SilentlyContinue | Out-Null

            if ($EnforceUpdatePageWithDevTools -and $devToolsPort -gt 0) {
                Invoke-DevToolsUpdatePageNavigation -Port $devToolsPort -Url $Browser.UpdateUrl -BrowserName $Browser.Name -TimeoutSeconds $DevToolsNavigationTimeoutSeconds | Out-Null
            }
        }

        Wait-WithHeartbeat -Seconds $ClosedBrowserUpdateWaitSeconds -Activity "$($Browser.Name) update/help page cycle"
        Stop-OnlyOurBrowserProcesses -ProcessExeNames $Browser.CloseExeNames -MarkerText $profileDir | Out-Null

        return $true
    }
    catch {
        Write-LogEntry "Open/update/close cycle failed for $($Browser.Name): $($_.Exception.Message)" 'Error'
        return $false
    }
    finally {
        if ($profileDir -and (Test-Path -LiteralPath $profileDir)) {
            try { Remove-Item -Path $profileDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
}


function Get-QueueBoolProperty {
    param(
        [Parameter(Mandatory = $true)]$QueueItem,
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$DefaultValue = $false,
        [switch]$ReturnNullWhenMissing
    )

    try {
        $prop = $QueueItem.PSObject.Properties[$Name]
        if ($null -eq $prop) {
            if ($ReturnNullWhenMissing) { return $null }
            return $DefaultValue
        }

        $value = $prop.Value
        if ($null -eq $value) {
            if ($ReturnNullWhenMissing) { return $null }
            return $DefaultValue
        }

        if ($value -is [bool]) { return [bool]$value }

        $text = ([string]$value).Trim()
        switch -Regex ($text.ToLowerInvariant()) {
            '^(1|true|yes|y|on)$' { return $true }
            '^(0|false|no|n|off)$' { return $false }
            default {
                if ($ReturnNullWhenMissing) { return $null }
                return $DefaultValue
            }
        }
    }
    catch {
        if ($ReturnNullWhenMissing) { return $null }
        return $DefaultValue
    }
}

function Invoke-QueuedClosedBrowserUpdate {
    param([Parameter(Mandatory = $true)]$Browser, [Parameter(Mandatory = $true)]$QueueItem)

    if (Test-BrowserRunning -ProcessNames $Browser.ProcessNames) {
        Write-LogEntry "$($Browser.Name) is now running. Leaving queue item for the normal user notification/reload remediation script." 'Warning'
        return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'BrowserRunning' }
    }

    $before = Get-BrowserInstalledVersion -Browser $Browser
    $targetVersion = Get-QueueTargetVersion -QueueItem $QueueItem
    Write-LogEntry "$($Browser.Name) closed-browser update-page verification started. BeforeVersion=$($before.Version) Source=$($before.Source) TargetLatestVersion=$targetVersion"

    $queuePendingUpdate = Get-QueueBoolProperty -QueueItem $QueueItem -Name 'PendingUpdate' -ReturnNullWhenMissing
    $queueOutOfDate = Get-QueueBoolProperty -QueueItem $QueueItem -Name 'OutOfDate' -ReturnNullWhenMissing
    if ($queuePendingUpdate -eq $false -and $queueOutOfDate -eq $true -and -not $ProcessOutOfDateWithoutPendingUpdate) {
        Write-LogEntry "$($Browser.Name) queue item is out of date but has PendingUpdate=False. This worker is open/update/close only, so it will not repeatedly attempt a non-staged update. Removing this closed-browser queue item. Set ProcessOutOfDateWithoutPendingUpdate=True to force an attempt anyway." 'Warning'
        return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = 'SkippedNoPendingUpdate'; Before = $before.Version; After = $before.Version; Target = $targetVersion }
    }

    if ([string]::IsNullOrWhiteSpace($before.Version)) {
        Write-LogEntry "$($Browser.Name) installed version could not be detected before action. Continuing, but success verification will be limited." 'Warning'
    }

    if ($RequireLatestVersionForQueueSuccess -and [string]::IsNullOrWhiteSpace($targetVersion)) {
        Write-LogEntry "$($Browser.Name) queue item does not include LatestVersion/TargetVersion. Falling back to version-change-only success criteria for this item." 'Warning'
    }

    if ($RequireLatestVersionForQueueSuccess -and -not [string]::IsNullOrWhiteSpace($targetVersion) -and -not [string]::IsNullOrWhiteSpace($before.Version)) {
        $alreadyCompliant = Test-VersionMeetsTarget -Version $before.Version -TargetVersion $targetVersion
        if ($alreadyCompliant -eq $true) {
            Write-LogEntry "$($Browser.Name) already meets or exceeds the queued target/latest version before any closed-browser action. Current=$($before.Version) Target=$targetVersion Source=$($before.Source). Removing stale queue item without opening the browser."
            return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = 'AlreadyCompliant'; Before = $before.Version; After = $before.Version; Target = $targetVersion }
        }
    }

    $effectiveMaxCycles = $MaxOpenUpdateCloseCycles
    if ($queuePendingUpdate -eq $false -and $queueOutOfDate -eq $true) {
        $effectiveMaxCycles = [math]::Min($MaxOpenUpdateCloseCycles, $MaxOpenUpdateCloseCyclesWhenNoPendingUpdate)
        Write-LogEntry "$($Browser.Name) has PendingUpdate=False and OutOfDate=True. Using non-staged best-effort limit: $effectiveMaxCycles cycle(s)."
    }

    if ($ReportOnly) {
        Write-LogEntry "ReportOnly: would run up to $effectiveMaxCycles open/update/close cycle(s) for $($Browser.Name) to '$($Browser.UpdateUrl)', stopping early if the target is reached."
        return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'ReportOnly' }
    }

    Invoke-NativeUpdateEngine -Browser $Browser -Phase 'pre-open'
    if ($UseNativeUpdateEngine -and $NativeUpdatePreOpenWaitSeconds -gt 0 -and -not [string]::IsNullOrWhiteSpace($before.Version)) {
        $preWait = Wait-ForVersionChange -Browser $Browser -OriginalVersion $before.Version -TimeoutSeconds $NativeUpdatePreOpenWaitSeconds -Phase 'native updater pre-open phase'
        if ($preWait.Changed) {
            $preResult = New-TargetAwareUpdateResult -Browser $Browser -BeforeVersion $before.Version -AfterVersion $preWait.Version -TargetVersion $targetVersion -Phase 'native updater pre-open phase' -SuccessReason 'UpdatedPreOpen'
            if ($preResult.Success) { return $preResult }
        }
    }

    $lastResult = $null
    $lastObservedVersion = $before.Version

    for ($cycle = 1; $cycle -le $effectiveMaxCycles; $cycle++) {
        if (Test-BrowserRunning -ProcessNames $Browser.ProcessNames) {
            Write-LogEntry "$($Browser.Name) became running before cycle $cycle. Leaving queue item for the normal user notification/reload remediation script." 'Warning'
            return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'BrowserRunning'; Before = $before.Version; After = $lastObservedVersion; Target = $targetVersion }
        }

        $cycleStart = Get-BrowserInstalledVersion -Browser $Browser
        $lastObservedVersion = $cycleStart.Version

        if ($RequireLatestVersionForQueueSuccess -and -not [string]::IsNullOrWhiteSpace($targetVersion) -and -not [string]::IsNullOrWhiteSpace($cycleStart.Version)) {
            $cycleAlreadyCompliant = Test-VersionMeetsTarget -Version $cycleStart.Version -TargetVersion $targetVersion
            if ($cycleAlreadyCompliant -eq $true) {
                Write-LogEntry "$($Browser.Name) reached the target before cycle $cycle started. Current=$($cycleStart.Version) Target=$targetVersion. Removing queue item."
                return [pscustomobject]@{ Success = $true; LeaveQueued = $false; Reason = 'AlreadyCompliantAfterWait'; Before = $before.Version; After = $cycleStart.Version; Target = $targetVersion }
            }
        }

        Write-LogEntry "$($Browser.Name) open/update/close cycle $cycle of $effectiveMaxCycles starting. CurrentVersion=$($cycleStart.Version) Target=$targetVersion"

        $opened = Invoke-OpenUpdateCloseCycle -Browser $Browser
        if (-not $opened) {
            return [pscustomobject]@{ Success = $false; LeaveQueued = $true; Reason = 'OpenCloseFailed'; Before = $before.Version; After = $lastObservedVersion; Target = $targetVersion }
        }

        Invoke-NativeUpdateEngine -Browser $Browser -Phase "post-close cycle $cycle"
        if ($UseNativeUpdateEngine -and $NativeUpdatePostCloseWaitSeconds -gt 0 -and -not [string]::IsNullOrWhiteSpace($cycleStart.Version)) {
            $postNative = Wait-ForVersionChange -Browser $Browser -OriginalVersion $cycleStart.Version -TimeoutSeconds $NativeUpdatePostCloseWaitSeconds -Phase "native updater post-close phase cycle $cycle"
            if ($postNative.Changed) {
                $postResult = New-TargetAwareUpdateResult -Browser $Browser -BeforeVersion $before.Version -AfterVersion $postNative.Version -TargetVersion $targetVersion -Phase "native updater post-close phase cycle $cycle" -SuccessReason 'UpdatedPostClose'
                $lastResult = $postResult
                $lastObservedVersion = $postNative.Version
                if ($postResult.Success) { return $postResult }
            }
        }

        if ($PostCloseVerificationSeconds -gt 0 -and -not [string]::IsNullOrWhiteSpace($cycleStart.Version)) {
            $null = Wait-ForVersionChange -Browser $Browser -OriginalVersion $cycleStart.Version -TimeoutSeconds $PostCloseVerificationSeconds -Phase "final post-close verification phase cycle $cycle"
        }

        $afterCycle = Get-BrowserInstalledVersion -Browser $Browser
        $lastObservedVersion = $afterCycle.Version
        Write-LogEntry "$($Browser.Name) cycle $cycle verification finished. InitialBefore=$($before.Version) CycleStart=$($cycleStart.Version) After=$($afterCycle.Version) Source=$($afterCycle.Source) Target=$targetVersion"

        $cycleResult = New-TargetAwareUpdateResult -Browser $Browser -BeforeVersion $before.Version -AfterVersion $afterCycle.Version -TargetVersion $targetVersion -Phase "open/update/close cycle $cycle" -SuccessReason 'UpdatedVerified'
        $lastResult = $cycleResult

        if ($cycleResult.Success) { return $cycleResult }

        if ($cycle -lt $effectiveMaxCycles) {
            Write-LogEntry "$($Browser.Name) target was not reached after cycle $cycle. Another cycle will be attempted after $InterCycleWaitSeconds seconds. Result=$($cycleResult.Reason) Current=$($afterCycle.Version) Target=$targetVersion" 'Warning'
            if ($InterCycleWaitSeconds -gt 0) {
                Wait-WithHeartbeat -Seconds $InterCycleWaitSeconds -Activity "$($Browser.Name) inter-cycle wait"
            }
        }
    }

    if ($null -eq $lastResult) {
        $after = Get-BrowserInstalledVersion -Browser $Browser
        $lastResult = New-TargetAwareUpdateResult -Browser $Browser -BeforeVersion $before.Version -AfterVersion $after.Version -TargetVersion $targetVersion -Phase 'final post-close verification phase' -SuccessReason 'UpdatedVerified'
    }

    Write-LogEntry "$($Browser.Name) did not reach the required target after $effectiveMaxCycles open/update/close cycle(s). Before=$($before.Version) LastObserved=$lastObservedVersion Target=$targetVersion Result=$($lastResult.Reason). Leaving queue item for retry." 'Warning'
    return $lastResult
}

function Invoke-LegacyIdleNudge {
    param([array]$InstalledBrowsers)

    $cutoff = (Get-Date).AddHours(-[double]$LookbackHours)
    $actions = 0

    Write-LogEntry "Legacy idle nudge is enabled. LookbackHours=$LookbackHours Cutoff=$cutoff"

    foreach ($browser in $InstalledBrowsers) {
        if (Test-BrowserRunning -ProcessNames $browser.ProcessNames) {
            Write-LogEntry "Skipping $($browser.Name): browser is currently running."
            continue
        }

        $lastOpened = Get-LastOpenedWhenNotRunning -Browser $browser
        $withinLookback = [bool]($lastOpened.Time -and $lastOpened.Time -ge $cutoff)
        $lastOpenedText = '<unknown>'
        if ($lastOpened.Time) { $lastOpenedText = $lastOpened.Time.ToString('yyyy-MM-dd HH:mm:ss') }

        Write-LogEntry "$($browser.Name): RunningNow=False LastOpened=$lastOpenedText Source=$($lastOpened.Source) OpenedWithinLookback=$withinLookback"

        if ($withinLookback) { continue }

        if ($ReportOnly) {
            Write-LogEntry "ReportOnly: would run legacy idle update-page nudge for $($browser.Name)."
            $actions++
            continue
        }

        $null = Invoke-OpenUpdateCloseCycle -Browser $browser
        $actions++
    }

    return $actions
}

# ---------------- Main ----------------
try {
    Initialize-SettingsFromEnvironment

    Write-LogEntry '======================================'
    Write-LogEntry 'Browser force idle/open-close update-page worker started'
    Write-LogEntry 'No direct MSI/EXE installer download or enterprise installer update logic is present in this script. v8 can best-effort process out-of-date closed browsers even when PendingUpdate=False, enforces the update/help page, and can process all queued browsers in one run when MaxBrowsersPerRun=0.'
    Write-LogEntry (Get-ExecutionContextSummary)
    Write-LogEntry "Settings: SupportedBrowsers=$($SupportedBrowsers -join ',') ProcessQueuedClosedBrowsers=$ProcessQueuedClosedBrowsers EnableIdleNudge=$EnableIdleNudge UseNativeUpdateEngine=$UseNativeUpdateEngine ProcessOutOfDateWithoutPendingUpdate=$ProcessOutOfDateWithoutPendingUpdate MaxBrowsersPerRun=$MaxBrowsersPerRun (0 = all queued browsers)"
    Write-LogEntry "Waits: ClosedBrowserUpdateWaitSeconds=$ClosedBrowserUpdateWaitSeconds NativeUpdatePreOpenWaitSeconds=$NativeUpdatePreOpenWaitSeconds NativeUpdatePostCloseSeconds=$NativeUpdatePostCloseWaitSeconds PostCloseVerificationSeconds=$PostCloseVerificationSeconds MaxOpenUpdateCloseCycles=$MaxOpenUpdateCloseCycles MaxOpenUpdateCloseCyclesWhenNoPendingUpdate=$MaxOpenUpdateCloseCyclesWhenNoPendingUpdate InterCycleWaitSeconds=$InterCycleWaitSeconds QueueRetryCooldownHours=$QueueRetryCooldownHours MaxClosedWorkerAttempts=$MaxClosedWorkerAttempts"
    Write-LogEntry '======================================'

    if (Test-IsSystem) {
        Write-LogEntry 'Script is running as SYSTEM. Opening browser UI as SYSTEM is not recommended. This script is intended for logged-in user context.' 'Warning'
    }

    $installed = @(Find-InstalledBrowsers -Names $SupportedBrowsers)
    if (-not $installed -or $installed.Count -eq 0) {
        Write-LogEntry 'No supported browsers found installed on this machine.' 'Warning'
        exit 3
    }

    foreach ($browser in $installed) {
        Write-LogEntry "Detected $($browser.Name): $($browser.ExePath)"
    }

    $failureCount = 0
    $actionWouldRun = 0
    $processedQueuedCount = 0

    if ($ProcessQueuedClosedBrowsers) {
        $queue = Read-ReloadQueue
        $queueItems = @($queue.Browsers)
        $closedItems = @($queueItems | Where-Object { Test-IsClosedQueueItem -Item $_ })

        Write-LogEntry "Queue items found: Total=$($queueItems.Count) ClosedBrowserItems=$($closedItems.Count)"

        $remaining = @()

        foreach ($item in $queueItems) {
            if (-not (Test-IsClosedQueueItem -Item $item)) {
                $remaining += $item
                continue
            }

            $browserName = [string]$item.Browser
            $browser = $installed | Where-Object { $_.Name -ieq $browserName } | Select-Object -First 1
            if (-not $browser) {
                Write-LogEntry "Queued browser '$browserName' is not installed or not supported by this script. Leaving queued." 'Warning'
                $remaining += $item
                $failureCount++
                continue
            }

            $attemptCount = Get-QueueIntProperty -QueueItem $item -Name 'ClosedWorkerAttemptCount' -DefaultValue 0
            if ($attemptCount -ge $MaxClosedWorkerAttempts) {
                Write-LogEntry "$($browser.Name) has already reached MaxClosedWorkerAttempts=$MaxClosedWorkerAttempts. LastResult=$(Get-QueueStringProperty -QueueItem $item -Name 'ClosedWorkerLastResult' -DefaultValue 'Unknown'). Skipping to avoid repeated long open/update/close runs." 'Warning'
                if ($RemoveQueueItemAfterMaxAttempts) {
                    Write-LogEntry "Removing $($browser.Name) from queue because RemoveQueueItemAfterMaxAttempts=True."
                }
                else {
                    $remaining += $item
                }
                continue
            }

            $cooldown = Test-QueueRetryCooldownActive -QueueItem $item
            if ($cooldown.Active) {
                Write-LogEntry "$($browser.Name) closed-browser retry cooldown is active until $($cooldown.UntilUtc) ($($cooldown.RemainingMinutes) minute(s) remaining). Leaving queued without action." 'Warning'
                $remaining += $item
                if ($TreatRetryCooldownAsFailure) { $failureCount++ }
                continue
            }

            if (($MaxBrowsersPerRun -gt 0) -and ($processedQueuedCount -ge $MaxBrowsersPerRun)) {
                Write-LogEntry "MaxBrowsersPerRun=$MaxBrowsersPerRun reached. Leaving $($browser.Name) queued for a later run." 'Information'
                $remaining += $item
                continue
            }

            $processedQueuedCount++
            if ($ReportOnly) { $actionWouldRun++ }

            Write-LogEntry "Processing queued closed-browser item for $($browser.Name). Reason=$($item.Reason)"
            $result = Invoke-QueuedClosedBrowserUpdate -Browser $browser -QueueItem $item

            if ($result.Success -and -not $result.LeaveQueued) {
                Write-LogEntry "Removing $($browser.Name) from queue. Result=$($result.Reason) Before=$($result.Before) After=$($result.After) Target=$($result.Target)"
            }
            else {
                Write-LogEntry "Leaving $($browser.Name) in queue. Result=$($result.Reason)" 'Warning'
                if ($result.Reason -ne 'ReportOnly' -and $result.Reason -ne 'BrowserRunning') {
                    $item = Stamp-QueueAttemptResult -QueueItem $item -Result $result
                }
                $remaining += $item
                if ($result.Reason -ne 'ReportOnly') { $failureCount++ }
            }
        }

        if (-not $ReportOnly) {
            if (-not (Save-ReloadQueue -Browsers $remaining)) { $failureCount++ }
        }
    }

    if ($EnableIdleNudge) {
        $legacyActions = Invoke-LegacyIdleNudge -InstalledBrowsers $installed
        if ($legacyActions -gt 0 -and $ReportOnly) { $actionWouldRun += $legacyActions }
    }

    Write-LogEntry '======================================'
    Write-LogEntry "Browser force idle/open-close update-page worker finished. ProcessedQueued=$processedQueuedCount Failures=$failureCount ReportOnly=$ReportOnly"
    Write-LogEntry '======================================'

    if ($ReportOnly -and $actionWouldRun -gt 0) { exit 2 }
    if ($failureCount -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-LogEntry "Fatal script error: $($_.Exception.Message)" 'Error'
    try { Write-LogEntry ($_.ScriptStackTrace) 'Error' } catch { }
    exit 4
}
