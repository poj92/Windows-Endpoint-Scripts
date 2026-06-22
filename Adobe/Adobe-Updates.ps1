<#
Author: Peter Opeyemi James / Nexus Open Systems Ltd
Script: DattoRMM-Adobe-Updates-v6.ps1
Purpose: Datto RMM component script to detect installed Adobe products, optionally deploy/update Adobe Remote Update Manager (RUM), run RUM, and report what was found and what changed.

Overview:
    - Discovers installed Adobe products from Windows uninstall registry hives.
    - Verifies Adobe Remote Update Manager exists.
    - Optionally deploys RUM if missing when Adobe desktop products are present.
    - Optionally updates RUM when the installed copy is below MinRumVersion or older than the supplied source copy.
    - Runs RUM --action=list first so the log shows applicable updates before changes.
    - Optionally closes running Adobe applications before update.
    - Runs RUM --action=install to download/install applicable updates.
    - Captures RUM stdout/stderr and copies RemoteUpdateManager.log into the working folder.
    - Takes a before/after installed-product snapshot and reports updated/changed products.

Datto component variables / PowerShell parameters:
    InstallUpdates             = true
        true  = list and install applicable updates
        false = list/report only

    ProductVersions            = ""
        Optional RUM filter. Example: PHSP,ILST or APRO#15.0,RDR#15.0
        Leave blank to target all installed products supported by RUM.

    ChannelIds                 = ""
        Optional legacy RUM channel filter. Leave blank in most cases.

    RumPath                    = ""
        Optional full path to RemoteUpdateManager.exe. Leave blank to auto-detect.

    BundledRumPath             = ""
        Optional path to a RemoteUpdateManager.exe bundled with the Datto component.
        If blank, the script also checks the component folder and WorkingDirectory.

    DeployRumIfMissing         = false
        true  = if Adobe desktop products are found and RUM is missing, deploy RUM from RumSourcePath, RumSourceUrl, BundledRumPath, or a bundled file/folder.
        false = do not install RUM automatically.

    UpdateRumIfOutdated        = false
        true  = if RUM exists but is older than MinRumVersion or older than the provided source copy, replace it from the source.
        false = do not update the RUM tool itself.

    ForceRumDeployment         = false
        true  = copy the provided RUM source to the target RUM directory even if installed/source versions look the same.
        false = only deploy when missing or detected as outdated.

    RumSourcePath              = ""
        Optional path to a RUM source file, folder, or zip. This can be a Datto bundled file/folder, local path, or network path accessible by the Datto execution account.

    RumSourceUrl               = ""
        Optional HTTPS/HTTP URL to a RUM source exe or zip, normally an internal repository URL. Adobe Admin Console downloads normally require authentication, so an internal URL is recommended.

    RumSourceSha256            = ""
        Optional SHA256 hash to verify the downloaded/source file before deployment.

    MinRumVersion              = ""
        Optional minimum acceptable RUM version, for example 3.1.0.0. If installed RUM is lower and UpdateRumIfOutdated=true, the script updates it from the provided source.

    RumInstallDirectory        = ""
        Optional target directory for RUM deployment. Default: C:\Program Files (x86)\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager

    CloseRunningAdobeApps      = false
        true  = attempt to close known Adobe desktop applications before RUM install
        false = only report running Adobe processes and continue

    FailIfAdobeAppsRunning     = false
        true  = fail if known Adobe desktop applications are running and CloseRunningAdobeApps is false
        false = warn and continue

    RequireAdmin               = true
        RUM should run elevated. Datto RMM normally runs components as LocalSystem.

    TimeoutMinutes             = 120
        Per-RUM-action timeout in minutes.

    FailWhenNoAdobeProducts    = false
        true  = return exit 7 when no installed Adobe products are discovered
        false = return exit 0 when nothing Adobe is found

    FailReportOnlyWhenUpdatesAvailable = false
        In report-only mode, return exit 1 if the RUM list output appears to show applicable updates.
        Detection is best-effort because Adobe's list output is not a stable machine-readable API.

    ClearRumLogBeforeRun       = true
        Backs up/deletes existing RemoteUpdateManager.log before each RUM action so copied logs are easier to read.

    ProxyUserName              = ""
    ProxyPassword              = ""
        Optional RUM proxy credentials.

    LogToFile                  = true
    LogPath                    = C:\ProgramData\DattoRMM\Logs\Adobe-RUM-Updates.log
    WorkingDirectory           = C:\ProgramData\DattoRMM\Packages\AdobeRUM

Exit codes:
    0 = success, no Adobe products found by policy, no updates needed, or updates completed
    1 = report-only mode found likely applicable updates and FailReportOnlyWhenUpdatesAvailable=true
    2 = fatal script error / not elevated when RequireAdmin=true
    3 = Adobe Remote Update Manager not found
    4 = RUM list action failed
    5 = RUM install action returned generic failure
    6 = RUM install action returned partial failure
    7 = no Adobe products found and FailWhenNoAdobeProducts=true
    8 = Adobe apps running and FailIfAdobeAppsRunning=true
    9 = RUM action timed out
    10 = RUM deployment/update requested but no valid RUM source was found
    11 = RUM deployment/update failed
#>

param(
    [object]$InstallUpdates = $null,

    [string]$ProductVersions = "",
    [string]$ChannelIds = "",
    [string]$RumPath = "",
    [string]$BundledRumPath = "",

    [object]$DeployRumIfMissing = $null,
    [object]$UpdateRumIfOutdated = $null,
    [object]$ForceRumDeployment = $null,
    [string]$RumSourcePath = "",
    [string]$RumSourceUrl = "",
    [string]$RumSourceSha256 = "",
    [string]$MinRumVersion = "",
    [string]$RumInstallDirectory = "",

    [object]$CloseRunningAdobeApps = $null,
    [object]$FailIfAdobeAppsRunning = $null,
    [object]$RequireAdmin = $null,
    [object]$FailWhenNoAdobeProducts = $null,
    [object]$FailReportOnlyWhenUpdatesAvailable = $null,
    [object]$ClearRumLogBeforeRun = $null,

    [int]$TimeoutMinutes = 120,

    [string]$ProxyUserName = "",
    [string]$ProxyPassword = "",

    [object]$LogToFile = $null,
    [string]$LogPath = "C:\ProgramData\DattoRMM\Logs\Adobe-RUM-Updates.log",
    [string]$WorkingDirectory = "C:\ProgramData\DattoRMM\Packages\AdobeRUM"
)

$script:StartTime = Get-Date
$script:TimedOut = $false
$script:ResolvedRumPath = $null
$script:RumCandidatePathsChecked = @()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line

    if ($script:ResolvedLogToFile) {
        try {
            $dir = Split-Path -Path $script:ResolvedLogPath -Parent
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $script:ResolvedLogPath -Value $line -Encoding UTF8
        }
        catch {
            Write-Host "[WARN] Failed writing log file: $($_.Exception.Message)"
        }
    }
}

function Resolve-InputValue {
    param(
        [string]$CurrentValue,
        [string]$EnvName,
        [string[]]$Aliases = @()
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return $CurrentValue.Trim()
    }

    foreach ($name in @($EnvName) + $Aliases) {
        try {
            $envValue = [Environment]::GetEnvironmentVariable($name)
            if (-not [string]::IsNullOrWhiteSpace($envValue)) {
                return $envValue.Trim()
            }
        }
        catch {}
    }

    return ""
}

function Resolve-BooleanSetting {
    param(
        [object]$CurrentValue,
        [string]$EnvName,
        [bool]$DefaultValue,
        [string[]]$Aliases = @()
    )

    $rawValue = $null
    $sourceName = $null

    if ($null -ne $CurrentValue -and -not [string]::IsNullOrWhiteSpace(([string]$CurrentValue))) {
        $rawValue = [string]$CurrentValue
        $sourceName = "parameter:$EnvName"
    }
    else {
        foreach ($name in @($EnvName) + $Aliases) {
            try {
                $envValue = [Environment]::GetEnvironmentVariable($name)
                if (-not [string]::IsNullOrWhiteSpace($envValue)) {
                    $rawValue = $envValue
                    $sourceName = "environment:$name"
                    break
                }
            }
            catch {}
        }
    }

    if ([string]::IsNullOrWhiteSpace($rawValue)) {
        return [pscustomobject]@{ Value = $DefaultValue; Source = "default"; Raw = "" }
    }

    switch -Regex ($rawValue.Trim().ToLowerInvariant()) {
        '^(\$?true|1|yes|y|on)$' {
            return [pscustomobject]@{ Value = $true; Source = $sourceName; Raw = $rawValue }
        }
        '^(\$?false|0|no|n|off)$' {
            return [pscustomobject]@{ Value = $false; Source = $sourceName; Raw = $rawValue }
        }
        default {
            return [pscustomobject]@{ Value = $DefaultValue; Source = "default-invalid-$sourceName"; Raw = $rawValue }
        }
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-Settings {
    $installSetting = Resolve-BooleanSetting -CurrentValue $InstallUpdates -EnvName "InstallUpdates" -DefaultValue $true -Aliases @("Remediate", "ApplyUpdates", "UpdateAdobe")
    $closeSetting = Resolve-BooleanSetting -CurrentValue $CloseRunningAdobeApps -EnvName "CloseRunningAdobeApps" -DefaultValue $false -Aliases @("CloseAdobeApps", "KillAdobeProcesses")
    $failRunningSetting = Resolve-BooleanSetting -CurrentValue $FailIfAdobeAppsRunning -EnvName "FailIfAdobeAppsRunning" -DefaultValue $false -Aliases @("FailOnRunningAdobeApps")
    $requireAdminSetting = Resolve-BooleanSetting -CurrentValue $RequireAdmin -EnvName "RequireAdmin" -DefaultValue $true -Aliases @("RequireElevation")
    $failNoAdobeSetting = Resolve-BooleanSetting -CurrentValue $FailWhenNoAdobeProducts -EnvName "FailWhenNoAdobeProducts" -DefaultValue $false -Aliases @("FailIfNoAdobeProducts")
    $failReportOnlySetting = Resolve-BooleanSetting -CurrentValue $FailReportOnlyWhenUpdatesAvailable -EnvName "FailReportOnlyWhenUpdatesAvailable" -DefaultValue $false -Aliases @("FailWhenUpdatesAvailable")
    $clearRumLogSetting = Resolve-BooleanSetting -CurrentValue $ClearRumLogBeforeRun -EnvName "ClearRumLogBeforeRun" -DefaultValue $true -Aliases @("ClearRemoteUpdateManagerLog")
    $deployRumSetting = Resolve-BooleanSetting -CurrentValue $DeployRumIfMissing -EnvName "DeployRumIfMissing" -DefaultValue $false -Aliases @("InstallRumIfMissing", "DeployRemoteUpdateManagerIfMissing")
    $updateRumSetting = Resolve-BooleanSetting -CurrentValue $UpdateRumIfOutdated -EnvName "UpdateRumIfOutdated" -DefaultValue $false -Aliases @("UpdateRemoteUpdateManagerIfOutdated", "UpdateRUM")
    $forceRumSetting = Resolve-BooleanSetting -CurrentValue $ForceRumDeployment -EnvName "ForceRumDeployment" -DefaultValue $false -Aliases @("ForceDeployRUM", "ForceRemoteUpdateManagerDeployment")
    $logToFileSetting = Resolve-BooleanSetting -CurrentValue $LogToFile -EnvName "LogToFile" -DefaultValue $true -Aliases @()

    $script:ShouldInstallUpdates = [bool]$installSetting.Value
    $script:ShouldCloseRunningAdobeApps = [bool]$closeSetting.Value
    $script:ShouldFailIfAdobeAppsRunning = [bool]$failRunningSetting.Value
    $script:ShouldRequireAdmin = [bool]$requireAdminSetting.Value
    $script:ShouldFailWhenNoAdobeProducts = [bool]$failNoAdobeSetting.Value
    $script:ShouldFailReportOnlyWhenUpdatesAvailable = [bool]$failReportOnlySetting.Value
    $script:ShouldClearRumLogBeforeRun = [bool]$clearRumLogSetting.Value
    $script:ShouldDeployRumIfMissing = [bool]$deployRumSetting.Value
    $script:ShouldUpdateRumIfOutdated = [bool]$updateRumSetting.Value
    $script:ShouldForceRumDeployment = [bool]$forceRumSetting.Value
    $script:ResolvedLogToFile = [bool]$logToFileSetting.Value

    $script:ResolvedProductVersions = Resolve-InputValue -CurrentValue $ProductVersions -EnvName "ProductVersions" -Aliases @("AdobeProductVersions")
    $script:ResolvedChannelIds = Resolve-InputValue -CurrentValue $ChannelIds -EnvName "ChannelIds" -Aliases @("AdobeChannelIds")
    $script:RequestedRumPath = Resolve-InputValue -CurrentValue $RumPath -EnvName "RumPath" -Aliases @("RemoteUpdateManagerPath", "AdobeRUMPath")
    $script:ResolvedBundledRumPath = Resolve-InputValue -CurrentValue $BundledRumPath -EnvName "BundledRumPath" -Aliases @("AdobeBundledRumPath", "BundledRemoteUpdateManagerPath")
    $script:ResolvedRumSourcePath = Resolve-InputValue -CurrentValue $RumSourcePath -EnvName "RumSourcePath" -Aliases @("RemoteUpdateManagerSourcePath", "AdobeRumSourcePath", "AdobeRUMSourcePath")
    $script:ResolvedRumSourceUrl = Resolve-InputValue -CurrentValue $RumSourceUrl -EnvName "RumSourceUrl" -Aliases @("RemoteUpdateManagerSourceUrl", "AdobeRumSourceUrl", "AdobeRUMSourceUrl")
    $script:ResolvedRumSourceSha256 = Resolve-InputValue -CurrentValue $RumSourceSha256 -EnvName "RumSourceSha256" -Aliases @("RemoteUpdateManagerSourceSha256", "AdobeRumSourceSha256", "AdobeRUMSourceSha256")
    $script:ResolvedMinRumVersion = Resolve-InputValue -CurrentValue $MinRumVersion -EnvName "MinRumVersion" -Aliases @("MinimumRumVersion", "MinimumRemoteUpdateManagerVersion")
    $script:ResolvedRumInstallDirectory = Resolve-InputValue -CurrentValue $RumInstallDirectory -EnvName "RumInstallDirectory" -Aliases @("RemoteUpdateManagerInstallDirectory", "AdobeRUMInstallDirectory")
    $script:ResolvedProxyUserName = Resolve-InputValue -CurrentValue $ProxyUserName -EnvName "ProxyUserName" -Aliases @("AdobeProxyUserName")
    $script:ResolvedProxyPassword = Resolve-InputValue -CurrentValue $ProxyPassword -EnvName "ProxyPassword" -Aliases @("AdobeProxyPassword")
    $script:ResolvedWorkingDirectory = Resolve-InputValue -CurrentValue $WorkingDirectory -EnvName "WorkingDirectory" -Aliases @("AdobeRUMWorkingDirectory")
    $script:ResolvedLogPath = Resolve-InputValue -CurrentValue $LogPath -EnvName "LogPath" -Aliases @("AdobeRUMLogPath")

    if ([string]::IsNullOrWhiteSpace($script:ResolvedWorkingDirectory)) {
        $script:ResolvedWorkingDirectory = "C:\ProgramData\DattoRMM\Packages\AdobeRUM"
    }
    if ([string]::IsNullOrWhiteSpace($script:ResolvedLogPath)) {
        $script:ResolvedLogPath = "C:\ProgramData\DattoRMM\Logs\Adobe-RUM-Updates.log"
    }
    if ([string]::IsNullOrWhiteSpace($script:ResolvedRumInstallDirectory)) {
        $commonX86ForRum = [Environment]::GetEnvironmentVariable("CommonProgramFiles(x86)")
        if ([string]::IsNullOrWhiteSpace($commonX86ForRum)) {
            $commonX86ForRum = "C:\Program Files (x86)\Common Files"
        }
        $script:ResolvedRumInstallDirectory = Join-Path $commonX86ForRum "Adobe\OOBE_Enterprise\RemoteUpdateManager"
    }
    try {
        $script:ResolvedTimeoutMinutes = [int]$TimeoutMinutes
    }
    catch {
        $script:ResolvedTimeoutMinutes = 120
    }

    if ($script:ResolvedTimeoutMinutes -lt 1) {
        $script:ResolvedTimeoutMinutes = 120
    }

    Ensure-Directory -Path $script:ResolvedWorkingDirectory

    Write-Log "========== Datto RMM Adobe update script =========="
    Write-Log ("InstallUpdates: {0} (source: {1})" -f $script:ShouldInstallUpdates, $installSetting.Source)
    Write-Log ("CloseRunningAdobeApps: {0} (source: {1})" -f $script:ShouldCloseRunningAdobeApps, $closeSetting.Source)
    Write-Log ("FailIfAdobeAppsRunning: {0} (source: {1})" -f $script:ShouldFailIfAdobeAppsRunning, $failRunningSetting.Source)
    Write-Log ("RequireAdmin: {0} (source: {1})" -f $script:ShouldRequireAdmin, $requireAdminSetting.Source)
    Write-Log ("FailWhenNoAdobeProducts: {0} (source: {1})" -f $script:ShouldFailWhenNoAdobeProducts, $failNoAdobeSetting.Source)
    Write-Log ("FailReportOnlyWhenUpdatesAvailable: {0} (source: {1})" -f $script:ShouldFailReportOnlyWhenUpdatesAvailable, $failReportOnlySetting.Source)
    Write-Log ("ClearRumLogBeforeRun: {0} (source: {1})" -f $script:ShouldClearRumLogBeforeRun, $clearRumLogSetting.Source)
    Write-Log ("DeployRumIfMissing: {0} (source: {1})" -f $script:ShouldDeployRumIfMissing, $deployRumSetting.Source)
    Write-Log ("UpdateRumIfOutdated: {0} (source: {1})" -f $script:ShouldUpdateRumIfOutdated, $updateRumSetting.Source)
    Write-Log ("ForceRumDeployment: {0} (source: {1})" -f $script:ShouldForceRumDeployment, $forceRumSetting.Source)
    Write-Log ("TimeoutMinutes: {0}" -f $script:ResolvedTimeoutMinutes)
    Write-Log ("ProductVersions filter: {0}" -f $(if ($script:ResolvedProductVersions) { $script:ResolvedProductVersions } else { "<none>" }))
    Write-Log ("ChannelIds filter: {0}" -f $(if ($script:ResolvedChannelIds) { $script:ResolvedChannelIds } else { "<none>" }))
    Write-Log ("RumPath override: {0}" -f $(if ($script:RequestedRumPath) { $script:RequestedRumPath } else { "<none>" }))
    Write-Log ("BundledRumPath: {0}" -f $(if ($script:ResolvedBundledRumPath) { $script:ResolvedBundledRumPath } else { "<none>" }))
    Write-Log ("RumSourcePath: {0}" -f $(if ($script:ResolvedRumSourcePath) { $script:ResolvedRumSourcePath } else { "<none>" }))
    Write-Log ("RumSourceUrl: {0}" -f $(if ($script:ResolvedRumSourceUrl) { $script:ResolvedRumSourceUrl } else { "<none>" }))
    Write-Log ("RumSourceSha256: {0}" -f $(if ($script:ResolvedRumSourceSha256) { "<provided>" } else { "<none>" }))
    Write-Log ("MinRumVersion: {0}" -f $(if ($script:ResolvedMinRumVersion) { $script:ResolvedMinRumVersion } else { "<none>" }))
    Write-Log ("RumInstallDirectory: {0}" -f $script:ResolvedRumInstallDirectory)
    Write-Log ("WorkingDirectory: {0}" -f $script:ResolvedWorkingDirectory)
    Write-Log ("Script LogPath: {0}" -f $script:ResolvedLogPath)
}

function Test-IsElevated {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Log ("Could not determine elevation status: {0}" -f $_.Exception.Message) "WARN"
        return $false
    }
}

function Normalize-VersionString {
    param([string]$Version)

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return ""
    }

    $v = $Version.Trim()
    $v = $v -replace '^[Vv]\s*', ''
    $m = [regex]::Match($v, '\d+(\.\d+){0,4}')

    if ($m.Success) {
        return $m.Value
    }

    return ""
}

function Convert-ToVersionObject {
    param([string]$Version)

    $normalized = Normalize-VersionString -Version $Version
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $parts = $normalized.Split('.')
    if ($parts.Count -gt 4) {
        $normalized = ($parts[0..3] -join '.')
    }

    try {
        return [version]$normalized
    }
    catch {
        return $null
    }
}

function Normalize-ProductName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    $n = $Name.ToLowerInvariant()
    $n = $n -replace '\s*\(64-bit\)\s*', ''
    $n = $n -replace '\s*\(32-bit\)\s*', ''
    $n = $n -replace '\s+', ' '
    return $n.Trim()
}

function Compare-VersionText {
    param(
        [string]$Before,
        [string]$After
    )

    if ([string]::IsNullOrWhiteSpace($Before) -and [string]::IsNullOrWhiteSpace($After)) {
        return "Same"
    }

    if ($Before -eq $After) {
        return "Same"
    }

    $beforeVersion = Convert-ToVersionObject -Version $Before
    $afterVersion = Convert-ToVersionObject -Version $After

    if ($beforeVersion -and $afterVersion) {
        if ($afterVersion -gt $beforeVersion) { return "Increased" }
        if ($afterVersion -lt $beforeVersion) { return "Decreased" }
        return "Changed"
    }

    return "Changed"
}

function Convert-PropertyToText {
    param([object]$Value)

    if ($null -eq $Value) { return "" }

    try {
        if ($Value -is [array]) {
            return (($Value | ForEach-Object { [string]$_ }) -join "; ")
        }
        return [string]$Value
    }
    catch {
        return ""
    }
}

function Get-InstalledApplications {
    $paths = @(
        @{ Hive = "HKLM"; Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" },
        @{ Hive = "HKLM-WOW6432Node"; Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" }
    )

    $items = @()

    foreach ($entry in $paths) {
        $path = $entry.Path
        $hive = $entry.Hive

        try {
            $apps = @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName })
        }
        catch {
            Write-Log ("Failed reading uninstall registry path {0}: {1}" -f $path, $_.Exception.Message) "WARN"
            $apps = @()
        }

        foreach ($app in $apps) {
            try {
                $items += [pscustomobject]@{
                    DisplayName          = Convert-PropertyToText $app.DisplayName
                    DisplayVersion       = Convert-PropertyToText $app.DisplayVersion
                    Publisher            = Convert-PropertyToText $app.Publisher
                    InstallDate          = Convert-PropertyToText $app.InstallDate
                    InstallLocation      = Convert-PropertyToText $app.InstallLocation
                    UninstallString      = Convert-PropertyToText $app.UninstallString
                    QuietUninstallString = Convert-PropertyToText $app.QuietUninstallString
                    PSChildName          = Convert-PropertyToText $app.PSChildName
                    RegistryPath         = Convert-PropertyToText $app.PSPath
                    RegistryHive         = $hive
                }
            }
            catch {
                Write-Log ("Skipped one uninstall registry entry under {0}: {1}" -f $path, $_.Exception.Message) "WARN"
            }
        }
    }

    return @($items)
}

function Get-InstalledAdobeProducts {
    $apps = @(Get-InstalledApplications)
    $seen = @{}
    $result = @()

    foreach ($app in $apps) {
        $name = [string]$app.DisplayName
        $publisher = [string]$app.Publisher

        $isAdobe = $false
        if ($name -match '^Adobe\b' -or $publisher -match 'Adobe') {
            $isAdobe = $true
        }

        if (-not $isAdobe) {
            continue
        }

        $key = ("{0}|{1}|{2}|{3}" -f (Normalize-ProductName -Name $name), $app.DisplayVersion, $app.Publisher, $app.RegistryHive).ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }
        $seen[$key] = $true

        $result += [pscustomobject]@{
            DisplayName       = $name
            DisplayVersion    = $app.DisplayVersion
            NormalizedVersion = (Normalize-VersionString -Version $app.DisplayVersion)
            Publisher         = $publisher
            InstallDate       = $app.InstallDate
            InstallLocation   = $app.InstallLocation
            ProductCode       = $app.PSChildName
            RegistryHive      = $app.RegistryHive
            RegistryPath      = $app.RegistryPath
            ProductKey        = (Normalize-ProductName -Name $name)
        }
    }

    return @($result | Sort-Object DisplayName, DisplayVersion, RegistryHive)
}

function Write-AdobeProductSnapshot {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Products,
        [Parameter(Mandatory)][string]$Label
    )

    Write-Log ("----- {0}: installed Adobe products -----" -f $Label)

    if (-not $Products -or $Products.Count -eq 0) {
        Write-Log ("{0}: no installed Adobe products discovered in uninstall registry hives." -f $Label) "WARN"
        return
    }

    Write-Log ("{0}: {1} Adobe product registry entr{2} found." -f $Label, $Products.Count, $(if ($Products.Count -eq 1) { "y" } else { "ies" }))

    foreach ($product in $Products) {
        $version = if ($product.DisplayVersion) { $product.DisplayVersion } else { "<blank>" }
        $publisher = if ($product.Publisher) { $product.Publisher } else { "<blank>" }
        $installLocation = if ($product.InstallLocation) { $product.InstallLocation } else { "<blank>" }
        Write-Log ("FOUND [{0}] {1} :: Version={2} :: Publisher={3} :: Location={4}" -f $product.RegistryHive, $product.DisplayName, $version, $publisher, $installLocation)
    }
}

function Get-KnownAdobeUserProcesses {
    $knownNames = @(
        "Acrobat",
        "AcroRd32",
        "AcroCEF",
        "AdobeCollabSync",
        "Photoshop",
        "Illustrator",
        "InDesign",
        "InCopy",
        "AfterFX",
        "Adobe Premiere Pro",
        "PremierePro",
        "Adobe Media Encoder",
        "Media Encoder",
        "Audition",
        "Dreamweaver",
        "Animate",
        "Bridge",
        "Lightroom",
        "LightroomClassic",
        "Character Animator",
        "Adobe Substance 3D Painter",
        "Adobe Substance 3D Designer",
        "Adobe Substance 3D Sampler",
        "Adobe Substance 3D Stager"
    )

    $matches = @()
    $processes = Get-Process -ErrorAction SilentlyContinue

    foreach ($proc in $processes) {
        foreach ($known in $knownNames) {
            if ($proc.ProcessName -ieq $known -or $proc.ProcessName -like "*$known*") {
                $matches += $proc
                break
            }
        }
    }

    return @($matches | Sort-Object ProcessName, Id -Unique)
}

function Handle-RunningAdobeProcesses {
    $running = @(Get-KnownAdobeUserProcesses)

    if ($running.Count -eq 0) {
        Write-Log "No known Adobe desktop application processes are running."
        return
    }

    foreach ($proc in $running) {
        Write-Log ("Running Adobe-related process detected: {0} PID={1}" -f $proc.ProcessName, $proc.Id) "WARN"
    }

    if ($script:ShouldCloseRunningAdobeApps) {
        Write-Log "CloseRunningAdobeApps=true. Attempting to stop detected Adobe desktop application processes." "WARN"
        foreach ($proc in $running) {
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                Write-Log ("Stopped process {0} PID={1}" -f $proc.ProcessName, $proc.Id) "SUCCESS"
            }
            catch {
                Write-Log ("Failed to stop process {0} PID={1}: {2}" -f $proc.ProcessName, $proc.Id, $_.Exception.Message) "WARN"
            }
        }
        Start-Sleep -Seconds 5
        return
    }

    if ($script:ShouldFailIfAdobeAppsRunning) {
        Write-Log "Adobe desktop applications are running and FailIfAdobeAppsRunning=true. Aborting before update." "ERROR"
        exit 8
    }

    Write-Log "Adobe desktop applications are running. Continuing because FailIfAdobeAppsRunning=false. Some updates may fail if files are locked." "WARN"
}

function Resolve-RumPath {
    $script:RumCandidatePathsChecked = @()

    if (-not [string]::IsNullOrWhiteSpace($script:RequestedRumPath)) {
        $script:RumCandidatePathsChecked += $script:RequestedRumPath
        if (Test-Path -LiteralPath $script:RequestedRumPath -PathType Leaf) {
            return (Get-Item -LiteralPath $script:RequestedRumPath).FullName
        }
        throw "Specified RumPath does not exist: $($script:RequestedRumPath)"
    }

    $candidatePaths = @(
        "C:\Program Files (x86)\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe",
        "C:\Program Files\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager.exe"
    )

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedBundledRumPath)) {
        $candidatePaths += $script:ResolvedBundledRumPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidatePaths += (Join-Path $PSScriptRoot "RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $PSScriptRoot "RemoteUpdateManager\RemoteUpdateManager.exe")
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedWorkingDirectory)) {
        $candidatePaths += (Join-Path $script:ResolvedWorkingDirectory "RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $script:ResolvedWorkingDirectory "RemoteUpdateManager\RemoteUpdateManager.exe")
    }

    $commonX86 = [Environment]::GetEnvironmentVariable("CommonProgramFiles(x86)")
    $common64 = [Environment]::GetEnvironmentVariable("CommonProgramFiles")
    $programX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")

    if (-not [string]::IsNullOrWhiteSpace($commonX86)) {
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager.exe")
    }
    if (-not [string]::IsNullOrWhiteSpace($common64)) {
        $candidatePaths += (Join-Path $common64 "Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $common64 "Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
    }
    if (-not [string]::IsNullOrWhiteSpace($programX86)) {
        $candidatePaths += (Join-Path $programX86 "Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $programX86 "Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
    }

    foreach ($path in @($candidatePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        $script:RumCandidatePathsChecked += $path
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return (Get-Item -LiteralPath $path).FullName
        }
    }

    return $null
}


function Get-StandardRumCandidatePaths {
    $candidatePaths = @(
        "C:\Program Files (x86)\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe",
        "C:\Program Files\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files\Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager\RemoteUpdateManager.exe",
        "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager.exe"
    )

    $commonX86 = [Environment]::GetEnvironmentVariable("CommonProgramFiles(x86)")
    $common64 = [Environment]::GetEnvironmentVariable("CommonProgramFiles")
    $programX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")

    if (-not [string]::IsNullOrWhiteSpace($commonX86)) {
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $commonX86 "Adobe\OOBE\PDApp\CCP\utilities\RemoteUpdateManager.exe")
    }
    if (-not [string]::IsNullOrWhiteSpace($common64)) {
        $candidatePaths += (Join-Path $common64 "Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $common64 "Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
    }
    if (-not [string]::IsNullOrWhiteSpace($programX86)) {
        $candidatePaths += (Join-Path $programX86 "Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager\RemoteUpdateManager.exe")
        $candidatePaths += (Join-Path $programX86 "Common Files\Adobe\OOBE_Enterprise\RemoteUpdateManager.exe")
    }

    return @($candidatePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Resolve-InstalledRumPath {
    if (-not [string]::IsNullOrWhiteSpace($script:RequestedRumPath)) {
        if (Test-Path -LiteralPath $script:RequestedRumPath -PathType Leaf) {
            return (Get-Item -LiteralPath $script:RequestedRumPath).FullName
        }
        return $null
    }

    foreach ($path in Get-StandardRumCandidatePaths) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return (Get-Item -LiteralPath $path).FullName
        }
    }

    return $null
}

function Get-RumFileInfo {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
        $fileVersion = Convert-PropertyToText $item.VersionInfo.FileVersion
        $productVersion = Convert-PropertyToText $item.VersionInfo.ProductVersion
        $normalized = Normalize-VersionString -Version $productVersion
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            $normalized = Normalize-VersionString -Version $fileVersion
        }

        return [pscustomobject]@{
            Path             = $item.FullName
            Directory        = $item.DirectoryName
            FileVersion      = $fileVersion
            ProductVersion   = $productVersion
            VersionText      = $normalized
            VersionObject    = (Convert-ToVersionObject -Version $normalized)
            LastWriteTimeUtc = $item.LastWriteTimeUtc
            SizeBytes        = $item.Length
        }
    }
    catch {
        Write-Log ("Could not read RUM file version info from {0}: {1}" -f $Path, $_.Exception.Message) "WARN"
        return $null
    }
}

function Write-RumFileInfo {
    param(
        [string]$Label,
        [pscustomobject]$Info
    )

    if (-not $Info) {
        Write-Log ("{0}: <not found>" -f $Label) "WARN"
        return
    }

    Write-Log ("{0}: Path={1} :: ProductVersion={2} :: FileVersion={3} :: LastWriteUtc={4}" -f $Label, $Info.Path, $(if ($Info.ProductVersion) { $Info.ProductVersion } else { "<blank>" }), $(if ($Info.FileVersion) { $Info.FileVersion } else { "<blank>" }), $Info.LastWriteTimeUtc)
}

function Get-RumRequiredAdobeProducts {
    param([object[]]$Products)

    $excludedPatterns = @(
        '^Adobe Creative Cloud$',
        '^Adobe Refresh Manager$',
        '^UXP WebView Support$',
        '^Adobe Genuine',
        '^Adobe Update',
        '^Adobe Acrobat Update Service$'
    )

    $targetPatterns = @(
        'Adobe Acrobat',
        'Adobe Reader',
        'Adobe Photoshop',
        'Adobe Illustrator',
        'Adobe InDesign',
        'Adobe InCopy',
        'Adobe After Effects',
        'Adobe Premiere Pro',
        'Adobe Media Encoder',
        'Adobe Lightroom',
        'Adobe Audition',
        'Adobe Animate',
        'Adobe Bridge',
        'Adobe Dreamweaver',
        'Adobe Character Animator',
        'Adobe Substance'
    )

    $result = @()
    foreach ($product in @($Products)) {
        $name = [string]$product.DisplayName
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $excluded = $false
        foreach ($pattern in $excludedPatterns) {
            if ($name -match $pattern) {
                $excluded = $true
                break
            }
        }
        if ($excluded) { continue }

        $isTarget = $false
        foreach ($pattern in $targetPatterns) {
            if ($name -match ([regex]::Escape($pattern))) {
                $isTarget = $true
                break
            }
        }

        if ($isTarget -or $name -match '^Adobe\b') {
            $result += $product
        }
    }

    return @($result | Sort-Object DisplayName, DisplayVersion, RegistryHive)
}

function Get-Sha256HashText {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (Get-Command Get-FileHash -ErrorAction SilentlyContinue) {
            return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
        }

        $sha = [System.Security.Cryptography.SHA256]::Create()
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $bytes = $sha.ComputeHash($stream)
            return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '').ToUpperInvariant()
        }
        finally {
            $stream.Dispose()
            $sha.Dispose()
        }
    }
    catch {
        throw "Could not calculate SHA256 for '$Path': $($_.Exception.Message)"
    }
}

function Test-RumSourceHashIfProvided {
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($script:ResolvedRumSourceSha256)) {
        return
    }

    $expected = ($script:ResolvedRumSourceSha256 -replace '\s', '').ToUpperInvariant()
    $actual = Get-Sha256HashText -Path $Path
    if ($actual -ne $expected) {
        throw "RUM source SHA256 mismatch for '$Path'. Expected $expected but got $actual."
    }

    Write-Log ("RUM source SHA256 verified for: {0}" -f $Path) "SUCCESS"
}

function Expand-RumZipSource {
    param([Parameter(Mandatory)][string]$ZipPath)

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $destination = Join-Path $script:ResolvedWorkingDirectory ("RUM-source-expanded-{0}" -f $stamp)
    Ensure-Directory -Path $destination

    try {
        if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {
            Expand-Archive -LiteralPath $ZipPath -DestinationPath $destination -Force -ErrorAction Stop
        }
        else {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $destination)
        }
    }
    catch {
        throw "Failed to expand RUM zip source '$ZipPath': $($_.Exception.Message)"
    }

    Write-Log ("Expanded RUM zip source to: {0}" -f $destination)
    return $destination
}

function Resolve-RumSourceFromPath {
    param([Parameter(Mandatory)][string]$SourcePath)

    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return $null
    }

    $item = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
    $sourceRoot = ""
    $exePath = ""
    $sourceType = ""

    if ($item.PSIsContainer) {
        $sourceType = "Folder"
        $matches = @(Get-ChildItem -LiteralPath $item.FullName -Recurse -Filter "RemoteUpdateManager.exe" -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer })
        if ($matches.Count -eq 0) {
            return $null
        }
        $exePath = ($matches | Sort-Object FullName | Select-Object -First 1).FullName
        $sourceRoot = Split-Path -Path $exePath -Parent
    }
    else {
        $extension = [System.IO.Path]::GetExtension($item.FullName).ToLowerInvariant()
        if ($item.Name -ieq "RemoteUpdateManager.exe") {
            $sourceType = "Exe"
            Test-RumSourceHashIfProvided -Path $item.FullName
            $exePath = $item.FullName
            $sourceRoot = Split-Path -Path $exePath -Parent
        }
        elseif ($extension -eq ".zip") {
            $sourceType = "Zip"
            Test-RumSourceHashIfProvided -Path $item.FullName
            $expanded = Expand-RumZipSource -ZipPath $item.FullName
            $matches = @(Get-ChildItem -LiteralPath $expanded -Recurse -Filter "RemoteUpdateManager.exe" -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer })
            if ($matches.Count -eq 0) {
                throw "RUM zip source '$($item.FullName)' did not contain RemoteUpdateManager.exe."
            }
            $exePath = ($matches | Sort-Object FullName | Select-Object -First 1).FullName
            $sourceRoot = Split-Path -Path $exePath -Parent
        }
        else {
            return $null
        }
    }

    $info = Get-RumFileInfo -Path $exePath
    return [pscustomobject]@{
        SourceType = $sourceType
        SourcePath = $item.FullName
        SourceRoot = $sourceRoot
        ExePath    = $exePath
        Info       = $info
    }
}

function Download-RumSourceFromUrl {
    if ([string]::IsNullOrWhiteSpace($script:ResolvedRumSourceUrl)) {
        return $null
    }

    try { $uri = [Uri]$script:ResolvedRumSourceUrl } catch { throw "RumSourceUrl is not a valid URL: $($script:ResolvedRumSourceUrl)" }

    $extension = [System.IO.Path]::GetExtension($uri.AbsolutePath)
    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = ".bin" }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $downloadPath = Join-Path $script:ResolvedWorkingDirectory ("RUM-source-download-{0}{1}" -f $stamp, $extension)

    Write-Log ("Downloading RUM source from URL to: {0}" -f $downloadPath)

    try {
        if (Get-Command Invoke-WebRequest -ErrorAction SilentlyContinue) {
            Invoke-WebRequest -Uri $script:ResolvedRumSourceUrl -OutFile $downloadPath -UseBasicParsing -ErrorAction Stop
        }
        else {
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($script:ResolvedRumSourceUrl, $downloadPath)
        }
    }
    catch {
        throw "Failed to download RUM source from URL: $($_.Exception.Message)"
    }

    Test-RumSourceHashIfProvided -Path $downloadPath
    return $downloadPath
}

function Resolve-RumDeploymentSource {
    $sourceCandidates = @()

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedRumSourceUrl)) {
        $downloaded = Download-RumSourceFromUrl
        if ($downloaded) { $sourceCandidates += $downloaded }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedRumSourcePath)) {
        $sourceCandidates += $script:ResolvedRumSourcePath
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedBundledRumPath)) {
        $sourceCandidates += $script:ResolvedBundledRumPath
    }

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $sourceCandidates += (Join-Path $PSScriptRoot "RemoteUpdateManager.exe")
        $sourceCandidates += (Join-Path $PSScriptRoot "RemoteUpdateManager.zip")
        $sourceCandidates += (Join-Path $PSScriptRoot "RUM.zip")
        $sourceCandidates += (Join-Path $PSScriptRoot "RemoteUpdateManager")
        $sourceCandidates += (Join-Path $PSScriptRoot "RUM")
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedWorkingDirectory)) {
        $sourceCandidates += (Join-Path $script:ResolvedWorkingDirectory "RemoteUpdateManager.exe")
        $sourceCandidates += (Join-Path $script:ResolvedWorkingDirectory "RemoteUpdateManager.zip")
        $sourceCandidates += (Join-Path $script:ResolvedWorkingDirectory "RUM.zip")
        $sourceCandidates += (Join-Path $script:ResolvedWorkingDirectory "RemoteUpdateManager")
        $sourceCandidates += (Join-Path $script:ResolvedWorkingDirectory "RUM")
    }

    foreach ($candidate in @($sourceCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        $source = Resolve-RumSourceFromPath -SourcePath $candidate
        if ($source) {
            Write-Log ("Selected RUM deployment source: Type={0} :: Exe={1}" -f $source.SourceType, $source.ExePath)
            Write-RumFileInfo -Label "RUM source version" -Info $source.Info
            return $source
        }
    }

    Write-Log "No valid RUM deployment source was found. Checked RumSourceUrl, RumSourcePath, BundledRumPath, the component folder, and the WorkingDirectory." "ERROR"
    return $null
}

function Test-RumNeedsUpdate {
    param(
        [pscustomobject]$InstalledInfo,
        [pscustomobject]$SourceInfo
    )

    if ($script:ShouldForceRumDeployment) {
        return [pscustomobject]@{ NeedsUpdate = $true; Reason = "ForceRumDeployment=true" }
    }

    if (-not $InstalledInfo) {
        return [pscustomobject]@{ NeedsUpdate = $true; Reason = "RUM is missing" }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedMinRumVersion)) {
        $minVersion = Convert-ToVersionObject -Version $script:ResolvedMinRumVersion
        if ($minVersion -and $InstalledInfo.VersionObject) {
            if ($InstalledInfo.VersionObject -lt $minVersion) {
                return [pscustomobject]@{ NeedsUpdate = $true; Reason = ("installed version {0} is below MinRumVersion {1}" -f $InstalledInfo.VersionText, $script:ResolvedMinRumVersion) }
            }
        }
        elseif ($minVersion -and -not $InstalledInfo.VersionObject) {
            return [pscustomobject]@{ NeedsUpdate = $true; Reason = ("installed RUM version could not be parsed and MinRumVersion is {0}" -f $script:ResolvedMinRumVersion) }
        }
        else {
            Write-Log ("MinRumVersion '{0}' could not be parsed as a version. Skipping MinRumVersion comparison." -f $script:ResolvedMinRumVersion) "WARN"
        }
    }

    if ($SourceInfo -and $SourceInfo.VersionObject -and $InstalledInfo.VersionObject) {
        if ($SourceInfo.VersionObject -gt $InstalledInfo.VersionObject) {
            return [pscustomobject]@{ NeedsUpdate = $true; Reason = ("source version {0} is newer than installed version {1}" -f $SourceInfo.VersionText, $InstalledInfo.VersionText) }
        }
    }

    return [pscustomobject]@{ NeedsUpdate = $false; Reason = "installed RUM is not below MinRumVersion and is not older than the supplied source" }
}

function Stop-RunningRumProcess {
    $running = @(Get-Process -Name "RemoteUpdateManager" -ErrorAction SilentlyContinue)
    foreach ($proc in $running) {
        try {
            Write-Log ("Stopping running RemoteUpdateManager process PID={0}" -f $proc.Id) "WARN"
            Stop-Process -Id $proc.Id -Force -ErrorAction Stop
        }
        catch {
            Write-Log ("Failed to stop RemoteUpdateManager process PID={0}: {1}" -f $proc.Id, $_.Exception.Message) "WARN"
        }
    }
}

function Copy-RumSourceToInstallDirectory {
    param([Parameter(Mandatory)][pscustomobject]$Source)

    $targetDirectory = $script:ResolvedRumInstallDirectory
    if (-not [string]::IsNullOrWhiteSpace($script:RequestedRumPath)) {
        $targetDirectory = Split-Path -Path $script:RequestedRumPath -Parent
    }

    Ensure-Directory -Path $targetDirectory
    Stop-RunningRumProcess

    $targetExe = Join-Path $targetDirectory "RemoteUpdateManager.exe"

    if ((Test-Path -LiteralPath $targetExe -PathType Leaf) -and ((Get-Item -LiteralPath $targetExe).FullName -ieq (Get-Item -LiteralPath $Source.ExePath).FullName)) {
        Write-Log "RUM source and target are the same file. No copy needed."
        return $targetExe
    }

    if (Test-Path -LiteralPath $targetExe -PathType Leaf) {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupDirectory = Join-Path $script:ResolvedWorkingDirectory ("RUM-backup-{0}" -f $stamp)
        try {
            Ensure-Directory -Path $backupDirectory
            Get-ChildItem -LiteralPath $targetDirectory -Force -ErrorAction SilentlyContinue | ForEach-Object {
                Copy-Item -LiteralPath $_.FullName -Destination $backupDirectory -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Log ("Backed up existing RUM directory to: {0}" -f $backupDirectory)
        }
        catch {
            Write-Log ("Could not back up existing RUM directory: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    try {
        Get-ChildItem -LiteralPath $Source.SourceRoot -Force -ErrorAction Stop | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination $targetDirectory -Recurse -Force -ErrorAction Stop
        }
    }
    catch {
        throw "Failed to copy RUM source from '$($Source.SourceRoot)' to '$targetDirectory': $($_.Exception.Message)"
    }

    if (-not (Test-Path -LiteralPath $targetExe -PathType Leaf)) {
        throw "RUM deployment copy completed but target executable was not found at '$targetExe'."
    }

    $newInfo = Get-RumFileInfo -Path $targetExe
    Write-RumFileInfo -Label "RUM target after deployment" -Info $newInfo
    Write-Log ("RUM deployment/update completed to: {0}" -f $targetExe) "SUCCESS"
    return $targetExe
}

function Ensure-RumToolState {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$RumRequiredProducts)

    if ($RumRequiredProducts.Count -eq 0) {
        Write-Log "No Adobe desktop products that normally require RUM were identified. RUM deployment/update will be skipped." "WARN"
        return
    }

    Write-Log ("Adobe desktop products requiring RUM candidate count: {0}" -f $RumRequiredProducts.Count)
    foreach ($product in ($RumRequiredProducts | Select-Object -First 20)) {
        Write-Log ("RUM candidate product: {0} :: Version={1}" -f $product.DisplayName, $(if ($product.DisplayVersion) { $product.DisplayVersion } else { "<blank>" }))
    }
    if ($RumRequiredProducts.Count -gt 20) {
        Write-Log ("RUM candidate product list truncated; total candidates: {0}" -f $RumRequiredProducts.Count) "WARN"
    }

    $installedRumPath = Resolve-InstalledRumPath
    $installedInfo = $null
    if ($installedRumPath) {
        $installedInfo = Get-RumFileInfo -Path $installedRumPath
        Write-RumFileInfo -Label "Installed RUM before deployment check" -Info $installedInfo
    }
    else {
        Write-Log "Installed RUM was not found in standard locations or RumPath override." "WARN"
    }

    if (-not $installedRumPath) {
        if (-not $script:ShouldDeployRumIfMissing) {
            return
        }

        Write-Log "DeployRumIfMissing=true and RUM is missing. Attempting RUM deployment." "WARN"
        try {
            $source = Resolve-RumDeploymentSource
        }
        catch {
            Write-Log $_.Exception.Message "ERROR"
            exit 10
        }
        if (-not $source) { exit 10 }

        try {
            $script:ResolvedRumPath = Copy-RumSourceToInstallDirectory -Source $source
        }
        catch {
            Write-Log $_.Exception.Message "ERROR"
            exit 11
        }
        return
    }

    if (-not $script:ShouldUpdateRumIfOutdated) {
        return
    }

    try {
        $sourceForComparison = Resolve-RumDeploymentSource
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
        exit 10
    }
    if (-not $sourceForComparison) { exit 10 }

    $evaluation = Test-RumNeedsUpdate -InstalledInfo $installedInfo -SourceInfo $sourceForComparison.Info
    if (-not $evaluation.NeedsUpdate) {
        Write-Log ("RUM self-update check: no update required. Reason: {0}" -f $evaluation.Reason)
        return
    }

    Write-Log ("UpdateRumIfOutdated=true. Updating RUM because {0}." -f $evaluation.Reason) "WARN"
    try {
        $script:ResolvedRumPath = Copy-RumSourceToInstallDirectory -Source $sourceForComparison
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
        exit 11
    }
}

function Protect-ArgumentValue {
    param([string]$Value)

    if ($null -eq $Value) { return "" }

    $escaped = $Value.Replace('"', '\"')
    if ($escaped -match '\s') {
        return '"' + $escaped + '"'
    }

    return $escaped
}

function New-RumArgumentString {
    param([Parameter(Mandatory)][string]$Action)

    $args = @("--action=$Action")

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedProductVersions)) {
        $args += ("--productVersions={0}" -f (Protect-ArgumentValue -Value $script:ResolvedProductVersions))
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedChannelIds)) {
        $args += ("--channelIds={0}" -f (Protect-ArgumentValue -Value $script:ResolvedChannelIds))
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedProxyUserName)) {
        $args += ("--proxyUserName={0}" -f (Protect-ArgumentValue -Value $script:ResolvedProxyUserName))
    }

    if (-not [string]::IsNullOrWhiteSpace($script:ResolvedProxyPassword)) {
        $args += ("--proxyPassword={0}" -f (Protect-ArgumentValue -Value $script:ResolvedProxyPassword))
    }

    return ($args -join " ")
}

function Get-RumLogPaths {
    $paths = @()

    foreach ($base in @($env:TEMP, $env:TMP, "C:\Windows\Temp")) {
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            $paths += (Join-Path $base "RemoteUpdateManager.log")
        }
    }

    try {
        if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
            $paths += (Join-Path $env:LOCALAPPDATA "Temp\RemoteUpdateManager.log")
        }
    }
    catch {}

    return @($paths | Select-Object -Unique)
}

function Clear-RumLogsForAction {
    param([Parameter(Mandatory)][string]$Action)

    if (-not $script:ShouldClearRumLogBeforeRun) {
        return
    }

    foreach ($path in Get-RumLogPaths) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            try {
                $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
                $backupName = "RemoteUpdateManager-$Action-existing-$stamp.log"
                $backupPath = Join-Path $script:ResolvedWorkingDirectory $backupName
                Copy-Item -LiteralPath $path -Destination $backupPath -Force -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                Write-Log ("Backed up and cleared existing RUM log: {0}" -f $path)
            }
            catch {
                Write-Log ("Could not clear existing RUM log {0}: {1}" -f $path, $_.Exception.Message) "WARN"
            }
        }
    }
}

function Copy-RumLogAfterAction {
    param([Parameter(Mandatory)][string]$Action)

    $existing = @(Get-RumLogPaths | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($existing.Count -eq 0) {
        Write-Log ("No RemoteUpdateManager.log was found after RUM action '{0}'." -f $Action) "WARN"
        return [pscustomobject]@{ SourcePath = ""; CopiedPath = ""; Text = "" }
    }

    $newest = $existing | Sort-Object { (Get-Item -LiteralPath $_).LastWriteTimeUtc } -Descending | Select-Object -First 1
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $destination = Join-Path $script:ResolvedWorkingDirectory ("RemoteUpdateManager-{0}-{1}.log" -f $Action, $stamp)

    try {
        Copy-Item -LiteralPath $newest -Destination $destination -Force -ErrorAction Stop
        Write-Log ("Copied RUM log for action '{0}' to: {1}" -f $Action, $destination)
    }
    catch {
        Write-Log ("Failed to copy RUM log from {0}: {1}" -f $newest, $_.Exception.Message) "WARN"
        $destination = ""
    }

    $text = ""
    try {
        $text = Get-Content -LiteralPath $newest -Raw -ErrorAction SilentlyContinue
    }
    catch {}

    return [pscustomobject]@{ SourcePath = $newest; CopiedPath = $destination; Text = $text }
}

function Write-CapturedTextToLog {
    param(
        [Parameter(Mandatory)][string]$Label,
        [string]$Text,
        [int]$MaxLines = 120
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Write-Log ("{0}: <no output>" -f $Label)
        return
    }

    $lines = @($Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) {
        Write-Log ("{0}: <no output>" -f $Label)
        return
    }

    Write-Log ("{0}: showing {1} of {2} non-blank line(s)." -f $Label, [Math]::Min($MaxLines, $lines.Count), $lines.Count)

    $take = $lines | Select-Object -First $MaxLines
    foreach ($line in $take) {
        Write-Log ("{0}> {1}" -f $Label, $line.Trim())
    }

    if ($lines.Count -gt $MaxLines) {
        Write-Log ("{0}: output truncated in Datto console log. Full output/log files are in {1}." -f $Label, $script:ResolvedWorkingDirectory) "WARN"
    }
}

function Invoke-ExternalProcess {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$Action
    )

    $stdoutPath = Join-Path $script:ResolvedWorkingDirectory ("RUM-{0}-stdout-{1}.txt" -f $Action, (Get-Date -Format "yyyyMMdd-HHmmss"))
    $stderrPath = Join-Path $script:ResolvedWorkingDirectory ("RUM-{0}-stderr-{1}.txt" -f $Action, (Get-Date -Format "yyyyMMdd-HHmmss"))

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    $psi.WorkingDirectory = Split-Path -Path $FilePath -Parent
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    Write-Log ('Executing RUM action ''{0}'': "{1}" {2}' -f $Action, $FilePath, $Arguments)

    $null = $process.Start()
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()

    $timeoutMs = [int]($script:ResolvedTimeoutMinutes * 60 * 1000)
    $finished = $process.WaitForExit($timeoutMs)

    if (-not $finished) {
        $script:TimedOut = $true
        try { $process.Kill() } catch {}
        try { $process.WaitForExit(10000) | Out-Null } catch {}

        $stdout = ""
        $stderr = ""
        try { $stdout = $stdoutTask.Result } catch {}
        try { $stderr = $stderrTask.Result } catch {}

        if ($stdout) { Set-Content -LiteralPath $stdoutPath -Value $stdout -Encoding UTF8 -Force }
        if ($stderr) { Set-Content -LiteralPath $stderrPath -Value $stderr -Encoding UTF8 -Force }

        throw "RUM action '$Action' timed out after $($script:ResolvedTimeoutMinutes) minute(s)."
    }

    $process.WaitForExit()

    $stdout = ""
    $stderr = ""
    try { $stdout = $stdoutTask.Result } catch {}
    try { $stderr = $stderrTask.Result } catch {}

    Set-Content -LiteralPath $stdoutPath -Value $stdout -Encoding UTF8 -Force
    Set-Content -LiteralPath $stderrPath -Value $stderr -Encoding UTF8 -Force

    Write-Log ("RUM action '{0}' exit code: {1}" -f $Action, $process.ExitCode)
    Write-Log ("RUM action '{0}' stdout captured to: {1}" -f $Action, $stdoutPath)
    Write-Log ("RUM action '{0}' stderr captured to: {1}" -f $Action, $stderrPath)

    return [pscustomobject]@{
        ExitCode   = $process.ExitCode
        StdOut     = $stdout
        StdErr     = $stderr
        StdOutPath = $stdoutPath
        StdErrPath = $stderrPath
    }
}

function Invoke-RumAction {
    param([Parameter(Mandatory)][ValidateSet("list","install","download")][string]$Action)

    Clear-RumLogsForAction -Action $Action

    $argumentString = New-RumArgumentString -Action $Action
    $processResult = Invoke-ExternalProcess -FilePath $script:ResolvedRumPath -Arguments $argumentString -Action $Action
    $rumLog = Copy-RumLogAfterAction -Action $Action

    Write-CapturedTextToLog -Label ("RUM {0} stdout" -f $Action) -Text $processResult.StdOut -MaxLines 160
    Write-CapturedTextToLog -Label ("RUM {0} stderr" -f $Action) -Text $processResult.StdErr -MaxLines 80

    if (-not [string]::IsNullOrWhiteSpace($rumLog.Text)) {
        $errorLines = @($rumLog.Text -split "`r?`n" | Where-Object { $_ -match '\[ERROR\]' })
        if ($errorLines.Count -gt 0) {
            Write-Log ("RUM {0} log contains {1} [ERROR] line(s)." -f $Action, $errorLines.Count) "WARN"
            foreach ($line in ($errorLines | Select-Object -First 30)) {
                Write-Log ("RUM {0} error> {1}" -f $Action, $line.Trim()) "WARN"
            }
        }
    }

    return [pscustomobject]@{
        Action        = $Action
        ExitCode      = $processResult.ExitCode
        StdOut        = $processResult.StdOut
        StdErr        = $processResult.StdErr
        StdOutPath    = $processResult.StdOutPath
        StdErrPath    = $processResult.StdErrPath
        RumLogPath    = $rumLog.CopiedPath
        RumLogText    = $rumLog.Text
        SourceLogPath = $rumLog.SourcePath
    }
}

function Get-RumResultCombinedText {
    param([pscustomobject]$RumListResult)

    $text = ""
    if ($RumListResult.StdOut) { $text += "`n" + [string]$RumListResult.StdOut }
    if ($RumListResult.StdErr) { $text += "`n" + [string]$RumListResult.StdErr }
    if ($RumListResult.RumLogText) { $text += "`n" + [string]$RumListResult.RumLogText }
    return $text
}

function Get-RumApplicableUpdateLines {
    param([pscustomobject]$RumListResult)

    $text = Get-RumResultCombinedText -RumListResult $RumListResult
    $updates = @()

    if ([string]::IsNullOrWhiteSpace($text)) {
        return @()
    }

    $inApplicableSection = $false
    $lines = @($text -split "`r?`n")

    foreach ($line in $lines) {
        $clean = [string]$line
        $clean = $clean -replace '^\s*\[[^\]]+\]\s+\[[^\]]+\]\s*RUM\s+list\s+stdout>\s*', ''
        $clean = $clean -replace '^\s*RUM\s+list\s+stdout>\s*', ''
        $clean = $clean.Trim()

        if ([string]::IsNullOrWhiteSpace($clean)) {
            continue
        }

        if ($clean -match '(?i)Following\s+Updates\s+are\s+applicable') {
            $inApplicableSection = $true
            continue
        }

        if ($clean -match '(?i)RemoteUpdateManager\s+exiting|Return\s+Code') {
            $inApplicableSection = $false
            continue
        }

        if ($clean -match '^\*{5,}$') {
            continue
        }

        if ($inApplicableSection) {
            if ($clean -match '(?i)^No\s+(new\s+)?updates') {
                continue
            }

            if ($clean -match '^\(([^\)]+)\)$') {
                $updates += $Matches[1]
                continue
            }

            if ($clean -match '^([A-Z0-9]{2,15}[/_][^\s]+)$') {
                $updates += $Matches[1]
                continue
            }

            if ($clean -match '\d+\.\d+') {
                $updates += $clean
                continue
            }
        }
        else {
            $matches = [regex]::Matches($clean, '\(([A-Z0-9]{2,15}/[^\)]+)\)')
            foreach ($match in $matches) {
                if ($match.Groups.Count -gt 1) {
                    $updates += $match.Groups[1].Value
                }
            }
        }
    }

    $deduped = @()
    $seen = @{}
    foreach ($update in @($updates)) {
        $value = ([string]$update).Trim()
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }
        if (-not $seen.ContainsKey($value)) {
            $seen[$value] = $true
            $deduped += $value
        }
    }

    return @($deduped)
}

function Test-RumListAppearsToHaveUpdates {
    param([pscustomobject]$RumListResult)

    $text = Get-RumResultCombinedText -RumListResult $RumListResult

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    $applicableUpdates = @(Get-RumApplicableUpdateLines -RumListResult $RumListResult)
    if ($applicableUpdates.Count -gt 0) {
        return $true
    }

    if ($text -match '(?i)Following\s+Updates\s+are\s+applicable') {
        return $true
    }

    if ($text -match '(?i)Adobe[A-Za-z]+DC-\d+\.\d+|[A-Z]{2,15}_\d+(\.\d+)+_|[A-Z]{2,15}/\d+(\.\d+)+') {
        return $true
    }

    return $false
}

function Write-RumListInterpretation {
    param([pscustomobject]$RumListResult)

    $applicableUpdates = @(Get-RumApplicableUpdateLines -RumListResult $RumListResult)
    $appearsToHaveUpdates = Test-RumListAppearsToHaveUpdates -RumListResult $RumListResult

    if ($applicableUpdates.Count -gt 0) {
        Write-Log ("RUM list interpretation: {0} applicable update item(s) detected from RUM output." -f $applicableUpdates.Count) "WARN"
        foreach ($update in ($applicableUpdates | Select-Object -First 50)) {
            Write-Log ("RUM applicable update: {0}" -f $update) "WARN"
        }
    }
    elseif ($appearsToHaveUpdates) {
        Write-Log "RUM list interpretation: RUM output indicates applicable updates, but exact update identifiers could not be parsed. See captured stdout and copied RemoteUpdateManager list log for exact Adobe output." "WARN"
    }
    else {
        Write-Log "RUM list interpretation: no applicable updates detected from the captured RUM output/log. See copied RUM log for the authoritative Adobe detail."
    }

    return $appearsToHaveUpdates
}

function New-ProductLookup {
    param([object[]]$Products)

    $lookup = @{}
    foreach ($product in @($Products)) {
        $key = $product.ProductKey
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }
        if (-not $lookup.ContainsKey($key)) {
            $lookup[$key] = @()
        }
        $lookup[$key] += $product
    }
    return $lookup
}

function Select-BestProductMatch {
    param(
        [pscustomobject]$BeforeProduct,
        [object[]]$Candidates
    )

    if (-not $Candidates -or $Candidates.Count -eq 0) {
        return $null
    }

    $sameHive = @($Candidates | Where-Object { $_.RegistryHive -eq $BeforeProduct.RegistryHive })
    if ($sameHive.Count -gt 0) {
        return ($sameHive | Select-Object -First 1)
    }

    return ($Candidates | Select-Object -First 1)
}

function Write-UpdateComparison {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$BeforeProducts,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AfterProducts
    )

    Write-Log "----- Update result comparison -----"

    $beforeLookup = New-ProductLookup -Products $BeforeProducts
    $afterLookup = New-ProductLookup -Products $AfterProducts
    $reportedAfterKeys = @{}

    $changedCount = 0
    $unchangedCount = 0
    $missingCount = 0

    foreach ($before in @($BeforeProducts | Sort-Object DisplayName, DisplayVersion, RegistryHive)) {
        $key = $before.ProductKey
        $after = $null
        if ($afterLookup.ContainsKey($key)) {
            $after = Select-BestProductMatch -BeforeProduct $before -Candidates @($afterLookup[$key])
        }

        if ($after) {
            $reportedAfterKeys[$key] = $true
            $comparison = Compare-VersionText -Before $before.DisplayVersion -After $after.DisplayVersion
            if ($comparison -eq "Same") {
                $unchangedCount++
                Write-Log ("UNCHANGED :: {0} :: Version={1}" -f $before.DisplayName, $(if ($before.DisplayVersion) { $before.DisplayVersion } else { "<blank>" }))
            }
            else {
                $changedCount++
                $level = "SUCCESS"
                if ($comparison -eq "Decreased") { $level = "WARN" }
                Write-Log ("UPDATED/CHANGED :: {0} :: {1} -> {2} :: Result={3}" -f $before.DisplayName, $(if ($before.DisplayVersion) { $before.DisplayVersion } else { "<blank>" }), $(if ($after.DisplayVersion) { $after.DisplayVersion } else { "<blank>" }), $comparison) $level
            }
        }
        else {
            $missingCount++
            Write-Log ("REMOVED/NOT-DETECTED-AFTER :: {0} :: PreviousVersion={1}" -f $before.DisplayName, $(if ($before.DisplayVersion) { $before.DisplayVersion } else { "<blank>" })) "WARN"
        }
    }

    $newCount = 0
    foreach ($after in @($AfterProducts | Sort-Object DisplayName, DisplayVersion, RegistryHive)) {
        $key = $after.ProductKey
        if (-not $beforeLookup.ContainsKey($key)) {
            $newCount++
            Write-Log ("NEWLY-DETECTED :: {0} :: Version={1}" -f $after.DisplayName, $(if ($after.DisplayVersion) { $after.DisplayVersion } else { "<blank>" })) "SUCCESS"
        }
    }

    Write-Log ("Comparison summary: Changed={0}; Unchanged={1}; RemovedOrNotDetected={2}; NewlyDetected={3}" -f $changedCount, $unchangedCount, $missingCount, $newCount)

    return [pscustomobject]@{
        ChangedCount   = $changedCount
        UnchangedCount = $unchangedCount
        MissingCount   = $missingCount
        NewCount       = $newCount
    }
}

function Write-FinalArtifactsSummary {
    param(
        [pscustomobject]$ListResult,
        [pscustomobject]$InstallResult
    )

    Write-Log "----- Artifact summary -----"
    if ($ListResult) {
        Write-Log ("RUM list stdout: {0}" -f $ListResult.StdOutPath)
        Write-Log ("RUM list stderr: {0}" -f $ListResult.StdErrPath)
        if ($ListResult.RumLogPath) { Write-Log ("RUM list log copy: {0}" -f $ListResult.RumLogPath) }
    }
    if ($InstallResult) {
        Write-Log ("RUM install stdout: {0}" -f $InstallResult.StdOutPath)
        Write-Log ("RUM install stderr: {0}" -f $InstallResult.StdErrPath)
        if ($InstallResult.RumLogPath) { Write-Log ("RUM install log copy: {0}" -f $InstallResult.RumLogPath) }
    }
    Write-Log ("Main script log: {0}" -f $script:ResolvedLogPath)
}

Initialize-Settings

try {
    if ($script:ShouldRequireAdmin -and -not (Test-IsElevated)) {
        Write-Log "Script is not running elevated. Adobe Remote Update Manager should be run with elevated privileges." "ERROR"
        exit 2
    }

    Write-Log "Starting installed Adobe product discovery from uninstall registry hives."
    $beforeProducts = @(Get-InstalledAdobeProducts)
    Write-AdobeProductSnapshot -Products $beforeProducts -Label "Before"

    if ($beforeProducts.Count -eq 0) {
        if ($script:ShouldFailWhenNoAdobeProducts) {
            Write-Log "No Adobe products were found and FailWhenNoAdobeProducts=true." "ERROR"
            exit 7
        }
        else {
            Write-Log "No Adobe products were found. Nothing to update." "SUCCESS"
            exit 0
        }
    }

    $rumRequiredProducts = @(Get-RumRequiredAdobeProducts -Products $beforeProducts)
    Ensure-RumToolState -RumRequiredProducts $rumRequiredProducts

    try {
        $script:ResolvedRumPath = Resolve-RumPath
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
        exit 3
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedRumPath)) {
        Write-Log "Adobe Remote Update Manager was not found. Adobe products were discovered, but no RUM executable exists in the standard or bundled locations." "ERROR"
        Write-Log "Searched RUM paths:" "INFO"
        foreach ($checkedPath in @($script:RumCandidatePathsChecked | Select-Object -Unique)) {
            Write-Log ("  - {0}" -f $checkedPath) "INFO"
        }
        Write-Log "Fix: deploy an Adobe Admin Console package with Remote Update Manager enabled, or set DeployRumIfMissing=true and provide/bundle a valid RUM source using RumSourcePath, RumSourceUrl, BundledRumPath, RemoteUpdateManager.exe, or RemoteUpdateManager.zip." "ERROR"
        exit 3
    }

    Write-Log ("Adobe Remote Update Manager path: {0}" -f $script:ResolvedRumPath)

    $listResult = Invoke-RumAction -Action "list"

    if ($listResult.ExitCode -ne 0) {
        Write-Log ("RUM list failed with exit code {0}. Install will not be attempted." -f $listResult.ExitCode) "ERROR"
        Write-FinalArtifactsSummary -ListResult $listResult -InstallResult $null
        exit 4
    }

    $listAppearsToHaveUpdates = Write-RumListInterpretation -RumListResult $listResult

    if (-not $script:ShouldInstallUpdates) {
        $afterReportOnly = @(Get-InstalledAdobeProducts)
        Write-AdobeProductSnapshot -Products $afterReportOnly -Label "After report-only"
        $null = Write-UpdateComparison -BeforeProducts $beforeProducts -AfterProducts $afterReportOnly
        Write-FinalArtifactsSummary -ListResult $listResult -InstallResult $null

        if ($script:ShouldFailReportOnlyWhenUpdatesAvailable -and $listAppearsToHaveUpdates) {
            Write-Log "Report-only result: applicable updates appear available." "WARN"
            exit 1
        }

        Write-Log "Report-only result: completed. No install action was performed." "SUCCESS"
        exit 0
    }

    Handle-RunningAdobeProcesses

    $installResult = Invoke-RumAction -Action "install"

    $afterProducts = @(Get-InstalledAdobeProducts)
    Write-AdobeProductSnapshot -Products $afterProducts -Label "After"
    $comparison = Write-UpdateComparison -BeforeProducts $beforeProducts -AfterProducts $afterProducts
    Write-FinalArtifactsSummary -ListResult $listResult -InstallResult $installResult

    if ($installResult.ExitCode -eq 0) {
        if ($comparison.ChangedCount -gt 0 -or $comparison.NewCount -gt 0) {
            Write-Log "Adobe update result: RUM completed successfully and installed product changes were detected." "SUCCESS"
        }
        else {
            Write-Log "Adobe update result: RUM completed successfully. No installed product version changes were detected; machine may already have been current or updates may not change DisplayVersion." "SUCCESS"
        }
        exit 0
    }

    if ($installResult.ExitCode -eq 2) {
        Write-Log "Adobe update result: RUM reported partial failure. Some updates may have installed; review comparison and copied RUM logs." "ERROR"
        exit 6
    }

    Write-Log ("Adobe update result: RUM install failed with exit code {0}." -f $installResult.ExitCode) "ERROR"
    exit 5
}
catch {
    if ($script:TimedOut) {
        Write-Log ("Fatal timeout: {0}" -f $_.Exception.Message) "ERROR"
        exit 9
    }

    Write-Log ("Fatal error: {0}" -f $_.Exception.Message) "ERROR"
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
        Write-Log ("Fatal location: {0}" -f (($_.InvocationInfo.PositionMessage -replace "`r?`n", " | "))) "ERROR"
    }
    if ($_.ScriptStackTrace) {
        Write-Log ("Fatal stack: {0}" -f (($_.ScriptStackTrace -replace "`r?`n", " | "))) "ERROR"
    }
    exit 2
}
finally {
    $elapsed = (Get-Date) - $script:StartTime
    Write-Log ("Completed in {0:n1} seconds" -f $elapsed.TotalSeconds)
}
