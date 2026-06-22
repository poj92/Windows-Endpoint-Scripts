<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-05
Email: Peter.James@nexusos.co.uk
#>

<#
Detects vulnerable Dell Peripheral Manager / Dell Display and Peripheral Manager installs on
Dell systems and optionally remediates them. The component skips non-Dell devices by default.
When Remediate is set to true, it removes any DPM/DDPM version below the configured minimum,
verifies removal, and only then downloads and installs the configured DDPM package. When
Remediate is false, it reports compliance only and makes no changes. Installer URL, version,
and SHA-256 are supplied as Datto variables for long-term maintainability.

Datto Variables:
    Remediate = false
    MinimumSafeVersion = 1.7.6
    SkipNonDell = true
    LatestDDPMUrl = https://dl.dell.com/FOLDERxxxxx/1/DDPM-Setup_x.x.x.xx.exe
    LatestDDPMVersion = 2.2.1.16
    LatestDDPMSha256 = ae6e965495ad54b78fab64a580f143cfa7c73e82ee5a473a1a925a6ac5c203a7
    LatestDDPMFileName = DDPM-Setup_2.2.1.16.exe
    LatestDDPMSilentArgs = /S
    LogToFile = true
    LogPath = C:\ProgramData\DattoRMM\Logs\DellPeripheral-DDPM-Compliance.log
    WorkingDirectory = C:\ProgramData\DattoRMM\Packages\DellDDPM
#>

param(
    # Primary Datto control. Accepts true/false, yes/no, 1/0, on/off.
    # false = report only; true = remediate when non-compliant.
    [object]$Remediate = $null,

    [string]$MinimumSafeVersion = "1.7.6",

    [bool]$SkipNonDell = $true,

    [string]$LatestDDPMUrl = "",
    [string]$LatestDDPMVersion = "",
    [string]$LatestDDPMSha256 = "",
    [string]$LatestDDPMFileName = "",
    [string]$LatestDDPMSilentArgs = "/S",

    [bool]$LogToFile = $true,
    [string]$LogPath = "C:\ProgramData\DattoRMM\Logs\DellPeripheral-DDPM-Compliance.log",
    [string]$WorkingDirectory = "C:\ProgramData\DattoRMM\Packages\DellDDPM",

    [switch]$ForceTls12
)

# Exit codes
# 0 = compliant / remediation succeeded
# 1 = non-compliant detected in report-only mode
# 2 = fatal script error
# 3 = skipped non-Dell device
# 4 = below-minimum software found, but uninstall verification failed; install not attempted
# 5 = older version removed successfully, but DDPM package metadata missing
# 6 = DDPM downloaded/installed, but not detected afterwards

$script:StartTime = Get-Date
$script:LatestDDPM = $null
$script:LatestDDPMFileNameResolved = $null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line

    if ($LogToFile) {
        try {
            $dir = Split-Path -Path $LogPath -Parent
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
            }
            Add-Content -Path $LogPath -Value $line
        }
        catch {
            Write-Host "[WARN] Failed writing log file: $($_.Exception.Message)"
        }
    }
}

function Resolve-InputValue {
    param(
        [string]$CurrentValue,
        [string]$EnvName
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
        return $CurrentValue.Trim()
    }

    try {
        $envValue = [Environment]::GetEnvironmentVariable($EnvName)
        if (-not [string]::IsNullOrWhiteSpace($envValue)) {
            return $envValue.Trim()
        }
    }
    catch {}

    return ""
}

function Resolve-BooleanSetting {
    param(
        [object]$CurrentValue,
        [string]$EnvName,
        [bool]$DefaultValue,
        [string[]]$Aliases = @()
    )

    $candidateNames = @($EnvName) + $Aliases
    $rawValue = $null
    $sourceName = $null

    if ($null -ne $CurrentValue -and -not [string]::IsNullOrWhiteSpace(([string]$CurrentValue))) {
        $rawValue = [string]$CurrentValue
        $sourceName = "parameter:$EnvName"
    }
    else {
        foreach ($name in $candidateNames) {
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
        return [pscustomobject]@{
            Value  = $DefaultValue
            Source = "default"
            Raw    = ""
        }
    }

    switch -Regex ($rawValue.Trim().ToLowerInvariant()) {
        '^(\$?true|1|yes|y|on)$' {
            return [pscustomobject]@{ Value = $true; Source = $sourceName; Raw = $rawValue }
        }
        '^(\$?false|0|no|n|off)$' {
            return [pscustomobject]@{ Value = $false; Source = $sourceName; Raw = $rawValue }
        }
        default {
            Write-Log ("Invalid boolean value '{0}' for {1}. Using default: {2}" -f $rawValue, $EnvName, $DefaultValue) "WARN"
            return [pscustomobject]@{
                Value  = $DefaultValue
                Source = "default-invalid-$sourceName"
                Raw    = $rawValue
            }
        }
    }
}

function Try-DeriveVersionFromText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $m = [regex]::Match($Text, '\d+(\.\d+){1,3}')
    if ($m.Success) {
        return $m.Value
    }

    return ""
}

function Resolve-RemediationInputs {
    $script:ResolvedLatestDDPMUrl        = Resolve-InputValue -CurrentValue $LatestDDPMUrl        -EnvName "LatestDDPMUrl"
    $script:ResolvedLatestDDPMVersion    = Resolve-InputValue -CurrentValue $LatestDDPMVersion    -EnvName "LatestDDPMVersion"
    $script:ResolvedLatestDDPMSha256     = Resolve-InputValue -CurrentValue $LatestDDPMSha256     -EnvName "LatestDDPMSha256"
    $script:ResolvedLatestDDPMFileName   = Resolve-InputValue -CurrentValue $LatestDDPMFileName   -EnvName "LatestDDPMFileName"
    $script:ResolvedLatestDDPMSilentArgs = Resolve-InputValue -CurrentValue $LatestDDPMSilentArgs -EnvName "LatestDDPMSilentArgs"

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMSilentArgs)) {
        $script:ResolvedLatestDDPMSilentArgs = "/S"
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMFileName) -and
        -not [string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMUrl)) {
        try {
            $script:ResolvedLatestDDPMFileName = [System.IO.Path]::GetFileName(([System.Uri]$script:ResolvedLatestDDPMUrl).AbsolutePath)
        }
        catch {}
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMVersion)) {
        $script:ResolvedLatestDDPMVersion = Try-DeriveVersionFromText -Text $script:ResolvedLatestDDPMFileName
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMVersion)) {
        $script:ResolvedLatestDDPMVersion = Try-DeriveVersionFromText -Text $script:ResolvedLatestDDPMUrl
    }

    Write-Log ("Resolved LatestDDPMUrl supplied: {0}" -f (-not [string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMUrl)))
    Write-Log ("Resolved LatestDDPMVersion supplied/derived: {0}" -f (-not [string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMVersion)))
    Write-Log ("Resolved LatestDDPMSha256 supplied: {0}" -f (-not [string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMSha256)))
    Write-Log ("Resolved LatestDDPMFileName supplied/derived: {0}" -f (-not [string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMFileName)))
    Write-Log ("Resolved LatestDDPMSilentArgs: {0}" -f $script:ResolvedLatestDDPMSilentArgs)
}

function Require-RemediationPackageMetadata {
    Resolve-RemediationInputs

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMUrl)) {
        throw "LatestDDPMUrl is required when installation is needed."
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMSha256)) {
        throw "LatestDDPMSha256 is required when installation is needed."
    }

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLatestDDPMFileName)) {
        throw "LatestDDPMFileName was blank and could not be derived from LatestDDPMUrl."
    }

    $script:LatestDDPM = [pscustomobject]@{
        Version            = $script:ResolvedLatestDDPMVersion
        FileName           = $script:ResolvedLatestDDPMFileName
        Url                = $script:ResolvedLatestDDPMUrl
        Sha256             = $script:ResolvedLatestDDPMSha256.ToLowerInvariant()
        SilentArgs         = $script:ResolvedLatestDDPMSilentArgs
        ProductDisplayName = "Dell Display and Peripheral Manager"
    }

    Write-Log ("Target DDPM package version: {0}" -f $(if ($script:LatestDDPM.Version) { $script:LatestDDPM.Version } else { "Not specified" }))
    Write-Log ("Target DDPM URL: {0}" -f $script:LatestDDPM.Url)
    Write-Log ("Target DDPM file name: {0}" -f $script:LatestDDPM.FileName)
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
        Write-Log ("Version compare failed: '{0}' vs '{1}'. Treating as below minimum." -f $VersionA, $VersionB) "WARN"
        return $true
    }

    return ($va -lt $vb)
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
                @{ Name = "PSChildName"; Expression = { $_.PSChildName } },
                @{ Name = "RegistryPath"; Expression = { $_.PSPath } }
    }

    return $items
}

function Get-DellRelevantApps {
    $apps = Get-InstalledApplications
    $result = New-Object System.Collections.Generic.List[object]

    foreach ($app in $apps) {
        $name = [string]$app.DisplayName

        if ($name -eq "Dell Peripheral Manager" -or
            $name -eq "Dell Display and Peripheral Manager") {

            $family = if ($name -eq "Dell Peripheral Manager") { "DPM" } else { "DDPM" }

            $result.Add([pscustomobject]@{
                Family               = $family
                DisplayName          = $app.DisplayName
                DisplayVersion       = $app.DisplayVersion
                NormalizedVersion    = (Normalize-VersionString -Version $app.DisplayVersion)
                UninstallString      = $app.UninstallString
                QuietUninstallString = $app.QuietUninstallString
                InstallLocation      = $app.InstallLocation
                RegistryPath         = $app.RegistryPath
                PSChildName          = $app.PSChildName
                Publisher            = $app.Publisher
            })
        }
    }

    return $result
}

function Get-ComplianceSnapshot {
    $apps = Get-DellRelevantApps

    $snapshot = [pscustomobject]@{
        Apps             = @($apps)
        DPMApps          = @()
        DDPMApps         = @()
        BelowMinimumApps = @()
        IsCompliant      = $true
        FoundAnything    = ($apps.Count -gt 0)
    }

    foreach ($app in $apps) {
        if ($app.Family -eq "DPM")  { $snapshot.DPMApps += $app }
        if ($app.Family -eq "DDPM") { $snapshot.DDPMApps += $app }

        if (Test-VersionLessThan -VersionA $app.DisplayVersion -VersionB $MinimumSafeVersion) {
            $snapshot.BelowMinimumApps += $app
            $snapshot.IsCompliant = $false
        }
    }

    return $snapshot
}

function Stop-DellProcesses {
    $patterns = @(
        "DellPeripheralManager",
        "Dell Display And Peripheral Manager",
        "DDPM",
        "DPM"
    )

    $all = Get-Process -ErrorAction SilentlyContinue
    foreach ($proc in $all) {
        foreach ($p in $patterns) {
            if ($proc.ProcessName -like "*$p*") {
                try {
                    Write-Log ("Stopping process {0} PID {1}" -f $proc.ProcessName, $proc.Id)
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                }
                catch {
                    Write-Log ("Unable to stop process {0}: {1}" -f $proc.ProcessName, $_.Exception.Message) "WARN"
                }
                break
            }
        }
    }
}

function Remove-RegistryEntryForApp {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$App
    )

    if (-not [string]::IsNullOrWhiteSpace($App.RegistryPath)) {
        try {
            Write-Log ("Removing stale uninstall registry entry: {0}" -f $App.RegistryPath) "WARN"
            Remove-Item -Path $App.RegistryPath -Recurse -Force -ErrorAction Stop
            return
        }
        catch {
            Write-Log ("Failed to remove registry path {0}: {1}" -f $App.RegistryPath, $_.Exception.Message) "WARN"
        }
    }

    throw "Unable to remove stale uninstall registry entry for $($App.DisplayName)"
}

function Split-CommandToExecutableAndArgs {
    param(
        [Parameter(Mandatory)]
        [string]$Command
    )

    $trimmed = $Command.Trim()

    if ($trimmed -match '^(?i)"?msiexec(\.exe)?"?\b') {
        return [pscustomobject]@{
            Executable = "msiexec.exe"
            Arguments  = ($trimmed -replace '^(?i)"?msiexec(\.exe)?"?\s*', '')
            IsMSI      = $true
        }
    }

    if ($trimmed.StartsWith('"')) {
        $parts = $trimmed -split '"', 3
        return [pscustomobject]@{
            Executable = $parts[1]
            Arguments  = $(if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" })
            IsMSI      = $false
        }
    }

    if (Test-Path -LiteralPath $trimmed) {
        return [pscustomobject]@{
            Executable = $trimmed
            Arguments  = ""
            IsMSI      = $false
        }
    }

    $matches = [regex]::Matches($trimmed, '\.exe', 'IgnoreCase')
    for ($m = $matches.Count - 1; $m -ge 0; $m--) {
        $exeEnd = $matches[$m].Index + $matches[$m].Length
        $possibleExe = $trimmed.Substring(0, $exeEnd).Trim('"',' ')
        $possibleArgs = $trimmed.Substring($exeEnd).Trim()

        if (Test-Path -LiteralPath $possibleExe) {
            return [pscustomobject]@{
                Executable = $possibleExe
                Arguments  = $possibleArgs
                IsMSI      = $false
            }
        }
    }

    return [pscustomobject]@{
        Executable = $null
        Arguments  = $null
        IsMSI      = $false
    }
}

function Invoke-UninstallCommand {
    param(
        [Parameter(Mandatory)]
        [string]$Command
    )

    $trimmed = $Command.Trim()
    Write-Log ("Original uninstall command: {0}" -f $trimmed)

    $parsed = Split-CommandToExecutableAndArgs -Command $trimmed

    if ($parsed.IsMSI) {
        $guidMatch = [regex]::Match($trimmed, '\{[A-Z0-9\-]+\}', 'IgnoreCase')
        if ($guidMatch.Success) {
            $guid = $guidMatch.Value
            $args = "/x $guid /qn /norestart"
        }
        else {
            $args = $parsed.Arguments
            if ($args -notmatch '(/quiet|/qn|/passive)') {
                $args += " /qn /norestart"
            }
        }

        Write-Log ("Executing MSI uninstall: msiexec.exe {0}" -f $args)
        $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
        return $p.ExitCode
    }

    if (-not $parsed.Executable) {
        throw "Could not parse uninstall command into a valid executable path: $Command"
    }

    $args = $parsed.Arguments
    if ($args -notmatch '(/quiet|/qn|/passive|/s|/silent|/verysilent)') {
        $args = "$args /S".Trim()
    }

    Write-Log ("Executing EXE uninstall: {0} {1}" -f $parsed.Executable, $args)
    $p = Start-Process -FilePath $parsed.Executable -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    return $p.ExitCode
}

function Uninstall-App {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$App
    )

    Write-Log ("Uninstalling {0} version '{1}'" -f $App.DisplayName, $App.DisplayVersion)

    $cmd = if ($App.QuietUninstallString) { $App.QuietUninstallString } else { $App.UninstallString }
    if (-not $cmd) {
        Write-Log ("No uninstall command found for {0}. Treating as stale entry and removing registry entry." -f $App.DisplayName) "WARN"
        Remove-RegistryEntryForApp -App $App
        return
    }

    $parsed = Split-CommandToExecutableAndArgs -Command $cmd

    if (-not $parsed.IsMSI) {
        if (-not $parsed.Executable -or -not (Test-Path -LiteralPath $parsed.Executable)) {
            Write-Log ("Uninstall target not found on disk: {0}" -f $(if ($parsed.Executable) { $parsed.Executable } else { "Unresolved path" })) "WARN"
            Write-Log ("Treating {0} as stale uninstall entry and removing registry entry." -f $App.DisplayName) "WARN"
            Remove-RegistryEntryForApp -App $App
            return
        }
    }

    Stop-DellProcesses
    $exitCode = Invoke-UninstallCommand -Command $cmd
    Write-Log ("Uninstall exit code: {0}" -f $exitCode)

    if ($exitCode -notin 0,1605,1614,1641,3010) {
        throw "Unexpected uninstall exit code $exitCode for $($App.DisplayName)"
    }

    Start-Sleep -Seconds 5
}

function Clear-PathAttributes {
    param([string]$TargetPath)

    try {
        Get-ChildItem -LiteralPath $TargetPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $_.Attributes = 'Normal'
            }
            catch {}
        }

        if (Test-Path -LiteralPath $TargetPath) {
            $item = Get-Item -LiteralPath $TargetPath -Force -ErrorAction SilentlyContinue
            if ($item) {
                try { $item.Attributes = 'Directory' } catch {}
            }
        }
    }
    catch {}
}

function Remove-StaleAppDirectories {
    $paths = @(
        "C:\Program Files\Dell\Dell Peripheral Manager",
        "C:\Program Files (x86)\Dell\Dell Peripheral Manager",
        "C:\Program Files\Dell\Dell Display and Peripheral Manager",
        "C:\Program Files (x86)\Dell\Dell Display and Peripheral Manager",
        "C:\ProgramData\Dell\Dell Peripheral Manager",
        "C:\ProgramData\Dell\Dell Display and Peripheral Manager"
    )

    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            try {
                Write-Log ("Removing stale directory: {0}" -f $path)
                Clear-PathAttributes -TargetPath $path
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-Log ("Could not remove stale directory {0}: {1}" -f $path, $_.Exception.Message) "WARN"
            }
        }
    }
}

function Ensure-WorkingDirectory {
    if (-not (Test-Path -LiteralPath $WorkingDirectory)) {
        New-Item -Path $WorkingDirectory -ItemType Directory -Force | Out-Null
    }
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -Algorithm SHA256 -Path $Path).Hash.ToLowerInvariant()
}

function Download-LatestDDPM {
    Ensure-WorkingDirectory

    if ($ForceTls12) {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {
            Write-Log ("Could not force TLS 1.2: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    $destination = Join-Path $WorkingDirectory $script:LatestDDPM.FileName

    if (Test-Path -LiteralPath $destination) {
        try {
            $existingHash = Get-FileSha256 -Path $destination
            if ($existingHash -eq $script:LatestDDPM.Sha256) {
                Write-Log ("Installer already present and checksum matches: {0}" -f $destination)
                return $destination
            }

            Write-Log "Existing installer checksum mismatch. Re-downloading." "WARN"
            Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log ("Failed to validate existing installer: {0}" -f $_.Exception.Message) "WARN"
            Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Log "Downloading DDPM package"
    Write-Log ("Source: {0}" -f $script:LatestDDPM.Url)
    Write-Log ("Destination: {0}" -f $destination)

    $ProgressPreference = 'SilentlyContinue'

    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", "DattoRMM-DDPM-Compliance")
        $wc.DownloadFile($script:LatestDDPM.Url, $destination)
    }
    catch {
        throw "Download failed: $($_.Exception.Message)"
    }

    if (-not (Test-Path -LiteralPath $destination)) {
        throw "Download did not produce file at $destination"
    }

    $hash = Get-FileSha256 -Path $destination
    Write-Log ("Downloaded file SHA256: {0}" -f $hash)

    if ($hash -ne $script:LatestDDPM.Sha256) {
        Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
        throw "Checksum verification failed for downloaded DDPM package."
    }

    Write-Log "Checksum verified successfully." "SUCCESS"
    return $destination
}

function Install-DDPM {
    param(
        [Parameter(Mandatory)]
        [string]$InstallerPath
    )

    Write-Log ("Installing DDPM from {0}" -f $InstallerPath)
    Write-Log ("Silent arguments: {0}" -f $script:LatestDDPM.SilentArgs)

    $p = Start-Process -FilePath $InstallerPath -ArgumentList $script:LatestDDPM.SilentArgs -Wait -PassThru -WindowStyle Hidden
    Write-Log ("Installer exit code: {0}" -f $p.ExitCode)

    if ($p.ExitCode -notin 0,1641,3010) {
        throw "DDPM installer returned unexpected exit code $($p.ExitCode)"
    }

    Start-Sleep -Seconds 20
}

function Get-DDPMAfterInstall {
    $apps = Get-DellRelevantApps | Where-Object { $_.Family -eq "DDPM" }

    if (-not $apps -or $apps.Count -eq 0) {
        return $null
    }

    $sorted = $apps | Sort-Object {
        $v = Convert-ToVersionObject -Version $_.DisplayVersion
        if ($v) { $v } else { [version]"0.0" }
    } -Descending

    return $sorted[0]
}

function Assert-NoBelowMinimumRemains {
    $post = Get-ComplianceSnapshot

    if ($post.BelowMinimumApps.Count -gt 0) {
        $left = $post.BelowMinimumApps | ForEach-Object {
            "$($_.DisplayName) [$($_.DisplayVersion)]"
        } | Sort-Object -Unique

        throw "Below-minimum software still remains: $($left -join '; ')"
    }

    Write-Log "Verified: no below-minimum DPM/DDPM entries remain." "SUCCESS"
}

function Assert-TargetDDPMDetected {
    $installed = Get-DDPMAfterInstall
    if (-not $installed) {
        throw "DDPM install completed but DDPM was not detected afterwards."
    }

    Write-Log ("Installed DDPM detected: {0}" -f $installed.DisplayVersion)

    if (Test-VersionLessThan -VersionA $installed.DisplayVersion -VersionB $MinimumSafeVersion) {
        throw "Installed DDPM version '$($installed.DisplayVersion)' is still below minimum '$MinimumSafeVersion'."
    }

    if (-not [string]::IsNullOrWhiteSpace($script:LatestDDPM.Version)) {
        Write-Log ("Expected DDPM version (advisory only): {0}" -f $script:LatestDDPM.Version)
    }
}

$remediationSetting = Resolve-BooleanSetting -CurrentValue $Remediate -EnvName "Remediate" -DefaultValue $false -Aliases @("PerformRemediation", "EnableRemediation", "AllowRemediation")
$script:ShouldRemediate = [bool]$remediationSetting.Value

Write-Log "========== Dell DPM/DDPM compliance script =========="
Write-Log ("Remediate: {0} (source: {1})" -f $script:ShouldRemediate, $remediationSetting.Source)
Write-Log ("MinimumSafeVersion: {0}" -f $MinimumSafeVersion)
Write-Log ("SkipNonDell: {0}" -f $SkipNonDell)
Write-Log ("WorkingDirectory: {0}" -f $WorkingDirectory)

try {
    if ($SkipNonDell -and -not (Test-IsDellDevice)) {
        Write-Log "Non-Dell device detected. Skipping by policy." "SUCCESS"
        exit 3
    }

    $initial = Get-ComplianceSnapshot

    if ($initial.Apps.Count -eq 0) {
        Write-Log "No DPM/DDPM detected."
    }
    else {
        foreach ($app in $initial.Apps) {
            $status = if ($initial.BelowMinimumApps -contains $app) { "BELOW-MINIMUM" } else { "OK" }
            Write-Log ("{0} :: {1} :: Version={2} :: Normalized={3}" -f $status, $app.DisplayName, $app.DisplayVersion, $app.NormalizedVersion)
        }
    }

    if (-not $script:ShouldRemediate) {
        if ($initial.BelowMinimumApps.Count -gt 0) {
            Write-Log "Report-only result: NON-COMPLIANT. Remediate=false, so no uninstall, cleanup, download, or install actions were performed." "WARN"
            exit 1
        }
        else {
            Write-Log "Report-only result: COMPLIANT. Remediate=false, so no changes were performed." "SUCCESS"
            exit 0
        }
    }

    Write-Log "Remediate=true. Non-compliant installs will be removed and the configured DDPM package will be installed if required."

    if ($initial.BelowMinimumApps.Count -eq 0) {
        Write-Log "No below-minimum DPM/DDPM found. No remediation required, and DDPM install will not be attempted." "SUCCESS"
        exit 0
    }

    Write-Log "Below-minimum software detected. Starting remediation."
    $removedAnyOlderVersion = $false

    foreach ($app in $initial.BelowMinimumApps) {
        Uninstall-App -App $app
        $removedAnyOlderVersion = $true
    }

    Remove-StaleAppDirectories

    try {
        Assert-NoBelowMinimumRemains
    }
    catch {
        Write-Log "Uninstall verification failed. DDPM install will not be attempted." "ERROR"
        Write-Log $_.Exception.Message "ERROR"
        exit 4
    }

    if (-not $removedAnyOlderVersion) {
        Write-Log "No older versions were removed. DDPM install will not be attempted." "SUCCESS"
        exit 0
    }

    try {
        Require-RemediationPackageMetadata
    }
    catch {
        Write-Log "Older version was removed successfully, but DDPM package metadata is missing or invalid." "ERROR"
        Write-Log $_.Exception.Message "ERROR"
        exit 5
    }

    $installer = Download-LatestDDPM
    Install-DDPM -InstallerPath $installer

    try {
        Assert-TargetDDPMDetected
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
        exit 6
    }

    Assert-NoBelowMinimumRemains

    Write-Log "Remediation result: COMPLIANT" "SUCCESS"
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