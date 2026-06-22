param(
    [string]$Mode = "",

    [string]$SkipNonDell = "",
    [string]$LogToFile = "",
    [string]$LogPath = "",
    [string]$WorkingDirectory = "",
    [string]$ForceTls12 = "",

    [string]$UpdateDCU = "",
    [string]$InstallIfMissingDCU = "",
    [string]$DCUUrl = "",
    [string]$DCUVersion = "",
    [string]$DCUSha256 = "",
    [string]$DCUFileName = "",
    [string]$DCUSilentArgs = "",

    [string]$UpdateDellOptimizer = "",
    [string]$InstallIfMissingDellOptimizer = "",
    [string]$DellOptimizerUrl = "",
    [string]$DellOptimizerVersion = "",
    [string]$DellOptimizerSha256 = "",
    [string]$DellOptimizerFileName = "",
    [string]$DellOptimizerSilentArgs = "",

    [string]$UpdateSupportAssist = "",
    [string]$InstallIfMissingSupportAssist = "",
    [string]$SupportAssistUrl = "",
    [string]$SupportAssistVersion = "",
    [string]$SupportAssistSha256 = "",
    [string]$SupportAssistFileName = "",
    [string]$SupportAssistSilentArgs = ""
)

$script:StartTime = Get-Date
$script:Failures = New-Object System.Collections.Generic.List[string]
$script:NeedsAction = $false

function Resolve-InputValue {
    param(
        [string]$CurrentValue,
        [string]$EnvName,
        [string]$DefaultValue = ""
    )

    if ($null -ne $CurrentValue) {
        $trimmedCurrent = [string]$CurrentValue
        if (-not [string]::IsNullOrWhiteSpace($trimmedCurrent)) {
            return $trimmedCurrent.Trim()
        }
    }

    try {
        $envValue = [Environment]::GetEnvironmentVariable($EnvName)
        if (-not [string]::IsNullOrWhiteSpace($envValue)) {
            return ([string]$envValue).Trim()
        }
    }
    catch {
    }

    return $DefaultValue
}

function Convert-ToBoolean {
    param(
        [string]$Value,
        [bool]$DefaultValue = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultValue
    }

    switch ($Value.Trim().ToLowerInvariant()) {
        "true"  { return $true }
        "false" { return $false }
        "1"     { return $true }
        "0"     { return $false }
        "yes"   { return $true }
        "no"    { return $false }
        "y"     { return $true }
        "n"     { return $false }
        "on"    { return $true }
        "off"   { return $false }
        default { return $DefaultValue }
    }
}

function Normalize-VersionString {
    param([string]$Version)

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return $null
    }

    $v = $Version.Trim()
    $v = $v -replace '^[Vv]\s*', ''
    $m = [regex]::Match($v, '\d+(\.\d+){0,3}')

    if ($m.Success) {
        return $m.Value
    }

    return $null
}

function Convert-ToVersionObject {
    param([string]$Version)

    $normalized = Normalize-VersionString -Version $Version
    if (-not $normalized) {
        return $null
    }

    try {
        return [version]$normalized
    }
    catch {
        return $null
    }
}

function Test-VersionLessThan {
    param(
        [string]$VersionA,
        [string]$VersionB
    )

    $va = Convert-ToVersionObject -Version $VersionA
    $vb = Convert-ToVersionObject -Version $VersionB

    if (-not $va -or -not $vb) {
        return $false
    }

    return ($va -lt $vb)
}

function Test-VersionGreaterOrEqual {
    param(
        [string]$VersionA,
        [string]$VersionB
    )

    $va = Convert-ToVersionObject -Version $VersionA
    $vb = Convert-ToVersionObject -Version $VersionB

    if (-not $va -or -not $vb) {
        return $false
    }

    return ($va -ge $vb)
}

function Resolve-Configuration {
    $script:ResolvedMode = Resolve-InputValue -CurrentValue $Mode -EnvName "Mode" -DefaultValue "Remediate"
    if ($script:ResolvedMode -notin @("Detect","Remediate")) {
        $script:ResolvedMode = "Remediate"
    }

    $script:ResolvedSkipNonDell = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $SkipNonDell -EnvName "SkipNonDell" -DefaultValue "true") -DefaultValue $true
    $script:ResolvedLogToFile = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $LogToFile -EnvName "LogToFile" -DefaultValue "true") -DefaultValue $true
    $script:ResolvedForceTls12 = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $ForceTls12 -EnvName "ForceTls12" -DefaultValue "false") -DefaultValue $false

    $script:ResolvedLogPath = Resolve-InputValue -CurrentValue $LogPath -EnvName "LogPath" -DefaultValue "C:\ProgramData\DattoRMM\Logs\Dell-App-Update-Component.log"
    $script:ResolvedWorkingDirectory = Resolve-InputValue -CurrentValue $WorkingDirectory -EnvName "WorkingDirectory" -DefaultValue "C:\ProgramData\DattoRMM\Packages\DellApps"

    $script:ResolvedUpdateDCU = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $UpdateDCU -EnvName "UpdateDCU" -DefaultValue "true") -DefaultValue $true
    $script:ResolvedInstallIfMissingDCU = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $InstallIfMissingDCU -EnvName "InstallIfMissingDCU" -DefaultValue "true") -DefaultValue $true
    $script:ResolvedDCUUrl = Resolve-InputValue -CurrentValue $DCUUrl -EnvName "DCUUrl" -DefaultValue ""
    $script:ResolvedDCUVersion = Resolve-InputValue -CurrentValue $DCUVersion -EnvName "DCUVersion" -DefaultValue ""
    $script:ResolvedDCUSha256 = Resolve-InputValue -CurrentValue $DCUSha256 -EnvName "DCUSha256" -DefaultValue ""
    $script:ResolvedDCUFileName = Resolve-InputValue -CurrentValue $DCUFileName -EnvName "DCUFileName" -DefaultValue ""
    $script:ResolvedDCUSilentArgs = Resolve-InputValue -CurrentValue $DCUSilentArgs -EnvName "DCUSilentArgs" -DefaultValue "/s"

    $script:ResolvedUpdateDellOptimizer = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $UpdateDellOptimizer -EnvName "UpdateDellOptimizer" -DefaultValue "true") -DefaultValue $true
    $script:ResolvedInstallIfMissingDellOptimizer = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $InstallIfMissingDellOptimizer -EnvName "InstallIfMissingDellOptimizer" -DefaultValue "false") -DefaultValue $false
    $script:ResolvedDellOptimizerUrl = Resolve-InputValue -CurrentValue $DellOptimizerUrl -EnvName "DellOptimizerUrl" -DefaultValue ""
    $script:ResolvedDellOptimizerVersion = Resolve-InputValue -CurrentValue $DellOptimizerVersion -EnvName "DellOptimizerVersion" -DefaultValue ""
    $script:ResolvedDellOptimizerSha256 = Resolve-InputValue -CurrentValue $DellOptimizerSha256 -EnvName "DellOptimizerSha256" -DefaultValue ""
    $script:ResolvedDellOptimizerFileName = Resolve-InputValue -CurrentValue $DellOptimizerFileName -EnvName "DellOptimizerFileName" -DefaultValue ""
    $script:ResolvedDellOptimizerSilentArgs = Resolve-InputValue -CurrentValue $DellOptimizerSilentArgs -EnvName "DellOptimizerSilentArgs" -DefaultValue "/s"

    $script:ResolvedUpdateSupportAssist = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $UpdateSupportAssist -EnvName "UpdateSupportAssist" -DefaultValue "false") -DefaultValue $false
    $script:ResolvedInstallIfMissingSupportAssist = Convert-ToBoolean -Value (Resolve-InputValue -CurrentValue $InstallIfMissingSupportAssist -EnvName "InstallIfMissingSupportAssist" -DefaultValue "false") -DefaultValue $false
    $script:ResolvedSupportAssistUrl = Resolve-InputValue -CurrentValue $SupportAssistUrl -EnvName "SupportAssistUrl" -DefaultValue ""
    $script:ResolvedSupportAssistVersion = Resolve-InputValue -CurrentValue $SupportAssistVersion -EnvName "SupportAssistVersion" -DefaultValue ""
    $script:ResolvedSupportAssistSha256 = Resolve-InputValue -CurrentValue $SupportAssistSha256 -EnvName "SupportAssistSha256" -DefaultValue ""
    $script:ResolvedSupportAssistFileName = Resolve-InputValue -CurrentValue $SupportAssistFileName -EnvName "SupportAssistFileName" -DefaultValue ""
    $script:ResolvedSupportAssistSilentArgs = Resolve-InputValue -CurrentValue $SupportAssistSilentArgs -EnvName "SupportAssistSilentArgs" -DefaultValue "/s"
}

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
            Add-Content -Path $script:ResolvedLogPath -Value $line
        }
        catch {
            Write-Host "[WARN] Failed writing log file: $($_.Exception.Message)"
        }
    }
}

function Add-Failure {
    param([string]$Message)
    $script:Failures.Add($Message) | Out-Null
    Write-Log -Message $Message -Level "ERROR"
}

function Ensure-WorkingDirectory {
    if (-not (Test-Path -LiteralPath $script:ResolvedWorkingDirectory)) {
        New-Item -Path $script:ResolvedWorkingDirectory -ItemType Directory -Force | Out-Null
    }
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -Algorithm SHA256 -Path $Path).Hash.ToLowerInvariant()
}

function Test-IsDellDevice {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

        $manufacturer = [string]$cs.Manufacturer
        $model = [string]$cs.Model
        $serial = if ($bios) { [string]$bios.SerialNumber } else { "" }

        Write-Log ("Manufacturer: {0}" -f $manufacturer)
        Write-Log ("Model: {0}" -f $model)
        if ($serial) { Write-Log ("Serial: {0}" -f $serial) }

        return ($manufacturer -match "Dell")
    }
    catch {
        throw "Unable to determine system manufacturer: $($_.Exception.Message)"
    }
}

function Get-InstalledApplications {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $items = foreach ($path in $paths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Select-Object `
                @{ Name = "DisplayName"; Expression = { $_.DisplayName } },
                @{ Name = "DisplayVersion"; Expression = { $_.DisplayVersion } },
                @{ Name = "Publisher"; Expression = { $_.Publisher } },
                @{ Name = "InstallLocation"; Expression = { $_.InstallLocation } },
                @{ Name = "UninstallString"; Expression = { $_.UninstallString } },
                @{ Name = "QuietUninstallString"; Expression = { $_.QuietUninstallString } },
                @{ Name = "RegistryPath"; Expression = { $_.PSPath } }
    }

    return @($items)
}

function Resolve-PackageConfig {
    param(
        [string]$Name,
        [bool]$Enabled,
        [bool]$InstallIfMissing,
        [string]$Url,
        [string]$Version,
        [string]$Sha256,
        [string]$FileName,
        [string]$SilentArgs,
        [string[]]$DetectNames,
        [string]$PrimaryDisplayName
    )

    $resolvedUrl = $Url
    $resolvedVersion = $Version
    $resolvedSha256 = $Sha256
    $resolvedFileName = $FileName
    $resolvedSilentArgs = $SilentArgs

    if ([string]::IsNullOrWhiteSpace($resolvedFileName) -and -not [string]::IsNullOrWhiteSpace($resolvedUrl)) {
        try {
            $resolvedFileName = [System.IO.Path]::GetFileName(([System.Uri]$resolvedUrl).AbsolutePath)
        }
        catch {
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolvedSilentArgs)) {
        $resolvedSilentArgs = "/s"
    }

    [pscustomobject]@{
        Name               = $Name
        Enabled            = $Enabled
        InstallIfMissing   = $InstallIfMissing
        Url                = $resolvedUrl
        TargetVersion      = $resolvedVersion
        Sha256             = $(if ($resolvedSha256) { $resolvedSha256.ToLowerInvariant() } else { "" })
        FileName           = $resolvedFileName
        SilentArgs         = $resolvedSilentArgs
        DetectNames        = $DetectNames
        PrimaryDisplayName = $PrimaryDisplayName
    }
}

function Find-InstalledProduct {
    param(
        [pscustomobject]$Package,
        [array]$InstalledApps
    )

    $matches = @(
        $InstalledApps | Where-Object {
            $name = [string]$_.DisplayName
            $Package.DetectNames -contains $name
        }
    )

    if (-not $matches -or $matches.Count -eq 0) {
        return $null
    }

    $sorted = $matches | Sort-Object {
        $v = Convert-ToVersionObject -Version $_.DisplayVersion
        if ($v) { $v } else { [version]"0.0" }
    } -Descending

    return $sorted[0]
}

function Test-PackageMetadataPresent {
    param([pscustomobject]$Package)

    if ([string]::IsNullOrWhiteSpace($Package.Url)) { return $false }
    if ([string]::IsNullOrWhiteSpace($Package.FileName)) { return $false }
    return $true
}

function Download-Package {
    param([pscustomobject]$Package)

    if (-not (Test-PackageMetadataPresent -Package $Package)) {
        throw ("{0}: URL and file name are required for download." -f $Package.Name)
    }

    Ensure-WorkingDirectory

    if ($script:ResolvedForceTls12) {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {
            Write-Log ("Could not force TLS 1.2 for {0}: {1}" -f $Package.Name, $_.Exception.Message) "WARN"
        }
    }

    $destination = Join-Path $script:ResolvedWorkingDirectory $Package.FileName

    if (Test-Path -LiteralPath $destination) {
        if (-not [string]::IsNullOrWhiteSpace($Package.Sha256)) {
            try {
                $existingHash = Get-FileSha256 -Path $destination
                if ($existingHash -eq $Package.Sha256) {
                    Write-Log ("{0}: installer already present and checksum matches." -f $Package.Name)
                    return $destination
                }

                Write-Log ("{0}: existing installer checksum mismatch. Re-downloading." -f $Package.Name) "WARN"
                Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log ("{0}: failed to validate existing installer: {1}" -f $Package.Name, $_.Exception.Message) "WARN"
                Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-Log ("{0}: installer already present; no checksum specified, reusing file." -f $Package.Name)
            return $destination
        }
    }

    Write-Log ("{0}: downloading package." -f $Package.Name)
    Write-Log ("{0}: source: {1}" -f $Package.Name, $Package.Url)
    Write-Log ("{0}: destination: {1}" -f $Package.Name, $destination)

    $ProgressPreference = 'SilentlyContinue'

    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "DattoRMM-Dell-App-Update")
        $wc.DownloadFile($Package.Url, $destination)
    }
    catch {
        throw ("{0}: download failed: {1}" -f $Package.Name, $_.Exception.Message)
    }

    if (-not (Test-Path -LiteralPath $destination)) {
        throw ("{0}: download did not produce file at {1}" -f $Package.Name, $destination)
    }

    if (-not [string]::IsNullOrWhiteSpace($Package.Sha256)) {
        $hash = Get-FileSha256 -Path $destination
        Write-Log ("{0}: downloaded SHA256: {1}" -f $Package.Name, $hash)

        if ($hash -ne $Package.Sha256) {
            Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
            throw ("{0}: checksum verification failed." -f $Package.Name)
        }

        Write-Log ("{0}: checksum verified successfully." -f $Package.Name) "SUCCESS"
    }
    else {
        Write-Log ("{0}: no checksum specified; skipping hash validation." -f $Package.Name) "WARN"
    }

    return $destination
}

function Install-Package {
    param(
        [pscustomobject]$Package,
        [string]$InstallerPath
    )

    Write-Log ("{0}: installing from {1}" -f $Package.Name, $InstallerPath)
    Write-Log ("{0}: silent arguments: {1}" -f $Package.Name, $Package.SilentArgs)

    $p = Start-Process -FilePath $InstallerPath -ArgumentList $Package.SilentArgs -Wait -PassThru -WindowStyle Hidden
    Write-Log ("{0}: installer exit code: {1}" -f $Package.Name, $p.ExitCode)

    if ($p.ExitCode -notin 0,1641,3010) {
        throw ("{0}: installer returned unexpected exit code {1}" -f $Package.Name, $p.ExitCode)
    }

    Start-Sleep -Seconds 20
}

function Get-DesiredState {
    param(
        [pscustomobject]$Package,
        $InstalledApp
    )

    $result = [pscustomobject]@{
        Installed        = $false
        NeedsAction      = $false
        Reason           = ""
        InstalledName    = ""
        InstalledVersion = ""
    }

    if ($InstalledApp) {
        $result.Installed = $true
        $result.InstalledName = $InstalledApp.DisplayName
        $result.InstalledVersion = $InstalledApp.DisplayVersion

        if (-not $Package.Enabled) {
            $result.NeedsAction = $false
            $result.Reason = "Disabled by policy"
            return $result
        }

        if (-not [string]::IsNullOrWhiteSpace($Package.TargetVersion)) {
            if (Test-VersionLessThan -VersionA $InstalledApp.DisplayVersion -VersionB $Package.TargetVersion) {
                $result.NeedsAction = $true
                $result.Reason = ("Installed version {0} is below target {1}" -f $InstalledApp.DisplayVersion, $Package.TargetVersion)
            }
            else {
                $result.NeedsAction = $false
                $result.Reason = ("Installed version {0} meets or exceeds target {1}" -f $InstalledApp.DisplayVersion, $Package.TargetVersion)
            }
        }
        else {
            if ($script:ResolvedMode -eq "Detect") {
                $result.NeedsAction = $false
                $result.Reason = "Installed and enabled; no target version supplied, so no action is required in Detect mode"
            }
            else {
                $result.NeedsAction = $true
                $result.Reason = "Installed and enabled; no target version supplied, so package is treated as update candidate in Remediate mode"
            }
        }

        return $result
    }

    if (-not $Package.Enabled) {
        $result.NeedsAction = $false
        $result.Reason = "Disabled by policy"
        return $result
    }

    if ($Package.InstallIfMissing) {
        $result.NeedsAction = $true
        $result.Reason = "Not installed and policy allows install if missing"
    }
    else {
        $result.NeedsAction = $false
        $result.Reason = "Not installed and policy does not allow install if missing"
    }

    return $result
}

function Verify-PackagePostInstall {
    param(
        [pscustomobject]$Package,
        [array]$InstalledApps
    )

    $post = Find-InstalledProduct -Package $Package -InstalledApps $InstalledApps
    if (-not $post) {
        throw ("{0}: package was not detected after installation." -f $Package.Name)
    }

    Write-Log ("{0}: detected after install as '{1}' version '{2}'." -f $Package.Name, $post.DisplayName, $post.DisplayVersion)

    if (-not [string]::IsNullOrWhiteSpace($Package.TargetVersion)) {
        if (-not (Test-VersionGreaterOrEqual -VersionA $post.DisplayVersion -VersionB $Package.TargetVersion)) {
            throw ("{0}: detected version '{1}' is below target '{2}' after install." -f $Package.Name, $post.DisplayVersion, $Package.TargetVersion)
        }
    }
}

function Process-Package {
    param([pscustomobject]$Package)

    Write-Log ("========== Processing {0} ==========" -f $Package.Name)

    if (-not $Package.Enabled) {
        Write-Log ("{0}: disabled by policy. Skipping." -f $Package.Name)
        return
    }

    $installedApps = Get-InstalledApplications
    $installed = Find-InstalledProduct -Package $Package -InstalledApps $installedApps
    $state = Get-DesiredState -Package $Package -InstalledApp $installed

    if ($state.Installed) {
        Write-Log ("{0}: detected '{1}' version '{2}'." -f $Package.Name, $state.InstalledName, $state.InstalledVersion)
    }
    else {
        Write-Log ("{0}: not detected." -f $Package.Name)
    }

    Write-Log ("{0}: decision: {1}" -f $Package.Name, $state.Reason)

    if ($state.NeedsAction) {
        $script:NeedsAction = $true
    }

    if ($script:ResolvedMode -eq "Detect") {
        return
    }

    if (-not $state.NeedsAction) {
        return
    }

    try {
        if (-not (Test-PackageMetadataPresent -Package $Package)) {
            throw ("{0}: package metadata missing. URL/file name required for remediation." -f $Package.Name)
        }

        $installer = Download-Package -Package $Package
        Install-Package -Package $Package -InstallerPath $installer

        $postApps = Get-InstalledApplications
        Verify-PackagePostInstall -Package $Package -InstalledApps $postApps

        Write-Log ("{0}: remediation completed successfully." -f $Package.Name) "SUCCESS"
    }
    catch {
        Add-Failure ("{0}: {1}" -f $Package.Name, $_.Exception.Message)
    }
}

Resolve-Configuration

Write-Log "========== Dell application update component =========="
Write-Log ("Resolved Mode: {0}" -f $script:ResolvedMode)
Write-Log ("Resolved SkipNonDell: {0}" -f $script:ResolvedSkipNonDell)
Write-Log ("Resolved LogToFile: {0}" -f $script:ResolvedLogToFile)
Write-Log ("Resolved LogPath: {0}" -f $script:ResolvedLogPath)
Write-Log ("Resolved WorkingDirectory: {0}" -f $script:ResolvedWorkingDirectory)
Write-Log ("Resolved ForceTls12: {0}" -f $script:ResolvedForceTls12)
Write-Log ("Resolved UpdateDCU: {0}" -f $script:ResolvedUpdateDCU)
Write-Log ("Resolved InstallIfMissingDCU: {0}" -f $script:ResolvedInstallIfMissingDCU)
Write-Log ("Resolved UpdateDellOptimizer: {0}" -f $script:ResolvedUpdateDellOptimizer)
Write-Log ("Resolved InstallIfMissingDellOptimizer: {0}" -f $script:ResolvedInstallIfMissingDellOptimizer)
Write-Log ("Resolved UpdateSupportAssist: {0}" -f $script:ResolvedUpdateSupportAssist)
Write-Log ("Resolved InstallIfMissingSupportAssist: {0}" -f $script:ResolvedInstallIfMissingSupportAssist)

try {
    if ($script:ResolvedSkipNonDell -and -not (Test-IsDellDevice)) {
        Write-Log "Non-Dell device detected. Skipping by policy." "SUCCESS"
        exit 3
    }

    $packages = @(
        (Resolve-PackageConfig `
            -Name "Dell Command Update" `
            -Enabled $script:ResolvedUpdateDCU `
            -InstallIfMissing $script:ResolvedInstallIfMissingDCU `
            -Url $script:ResolvedDCUUrl `
            -Version $script:ResolvedDCUVersion `
            -Sha256 $script:ResolvedDCUSha256 `
            -FileName $script:ResolvedDCUFileName `
            -SilentArgs $script:ResolvedDCUSilentArgs `
            -DetectNames @(
                "Dell Command | Update",
                "Dell Command Update",
                "Dell Command | Update for Windows Universal"
            ) `
            -PrimaryDisplayName "Dell Command | Update"),

        (Resolve-PackageConfig `
            -Name "Dell Optimizer" `
            -Enabled $script:ResolvedUpdateDellOptimizer `
            -InstallIfMissing $script:ResolvedInstallIfMissingDellOptimizer `
            -Url $script:ResolvedDellOptimizerUrl `
            -Version $script:ResolvedDellOptimizerVersion `
            -Sha256 $script:ResolvedDellOptimizerSha256 `
            -FileName $script:ResolvedDellOptimizerFileName `
            -SilentArgs $script:ResolvedDellOptimizerSilentArgs `
            -DetectNames @("Dell Optimizer") `
            -PrimaryDisplayName "Dell Optimizer"),

        (Resolve-PackageConfig `
            -Name "Dell SupportAssist" `
            -Enabled $script:ResolvedUpdateSupportAssist `
            -InstallIfMissing $script:ResolvedInstallIfMissingSupportAssist `
            -Url $script:ResolvedSupportAssistUrl `
            -Version $script:ResolvedSupportAssistVersion `
            -Sha256 $script:ResolvedSupportAssistSha256 `
            -FileName $script:ResolvedSupportAssistFileName `
            -SilentArgs $script:ResolvedSupportAssistSilentArgs `
            -DetectNames @(
                "SupportAssist",
                "Dell SupportAssist",
                "SupportAssist for Business PCs",
                "Dell SupportAssist Remediation",
                "Dell SupportAssist OS Recovery Plugin for Dell Update"
            ) `
            -PrimaryDisplayName "Dell SupportAssist")
    )

    foreach ($pkg in $packages) {
        Process-Package -Package $pkg
    }

    if ($script:ResolvedMode -eq "Detect") {
        if ($script:NeedsAction) {
            Write-Log "Detection result: one or more enabled Dell applications require action." "WARN"
            exit 1
        }
        else {
            Write-Log "Detection result: all enabled Dell applications are compliant with policy." "SUCCESS"
            exit 0
        }
    }

    if ($script:Failures.Count -gt 0) {
        Write-Log ("Remediation completed with failures: {0}" -f ($script:Failures -join " | ")) "ERROR"
        exit 4
    }

    Write-Log "Remediation completed successfully." "SUCCESS"
    exit 0
}
catch {
    Write-Log ("Fatal error: {0}" -f $_.Exception.Message) "ERROR"
    exit 2
}
finally {
    $elapsed = (Get-Date) - $script:StartTime
    Write-Log ("Completed in {0:n1} seconds" -f $elapsed.TotalSeconds)
}