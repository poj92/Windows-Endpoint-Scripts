<#
Detects and remediates vulnerable Dell Peripheral Manager / Dell Display and Peripheral Manager 
installs on Dell systems. The component skips non-Dell devices by default, removes any DPM/DDPM 
version below the configured minimum, verifies removal, and only then downloads and installs 
the configured DDPM package. Installer URL, version, and SHA-256 are supplied as Datto 
variables for long-term maintainability.

Datto Variables:
    Mode = Remediate
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
    [ValidateSet("Detect","Remediate")]
    [string]$Mode = "Remediate",

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
# 1 = non-compliant detected (Detect mode)
# 2 = fatal script error
# 3 = skipped non-Dell device
# 4 = below-minimum software found, but uninstall verification failed; install not attempted

$script:StartTime = Get-Date

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

function Require-RemediationPackageMetadata {
    if ([string]::IsNullOrWhiteSpace($LatestDDPMUrl)) {
        throw "LatestDDPMUrl is required in Remediate mode."
    }
    if ([string]::IsNullOrWhiteSpace($LatestDDPMVersion)) {
        throw "LatestDDPMVersion is required in Remediate mode."
    }
    if ([string]::IsNullOrWhiteSpace($LatestDDPMSha256)) {
        throw "LatestDDPMSha256 is required in Remediate mode."
    }

    if ([string]::IsNullOrWhiteSpace($LatestDDPMFileName)) {
        try {
            $script:LatestDDPMFileNameResolved = [System.IO.Path]::GetFileName(([System.Uri]$LatestDDPMUrl).AbsolutePath)
        }
        catch {
            throw "LatestDDPMFileName was blank and could not be derived from LatestDDPMUrl."
        }
    }
    else {
        $script:LatestDDPMFileNameResolved = $LatestDDPMFileName
    }

    $script:LatestDDPM = [pscustomobject]@{
        Version            = $LatestDDPMVersion
        FileName           = $script:LatestDDPMFileNameResolved
        Url                = $LatestDDPMUrl
        Sha256             = $LatestDDPMSha256.ToLowerInvariant()
        SilentArgs         = $LatestDDPMSilentArgs
        ProductDisplayName = "Dell Display and Peripheral Manager"
    }

    Write-Log "Target DDPM package version: $($script:LatestDDPM.Version)"
    Write-Log "Target DDPM URL: $($script:LatestDDPM.Url)"
    Write-Log "Target DDPM file name: $($script:LatestDDPM.FileName)"
}

function Test-IsDellDevice {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

        $manufacturer = [string]$cs.Manufacturer
        $model = [string]$cs.Model
        $serial = if ($bios) { [string]$bios.SerialNumber } else { "" }

        Write-Log "Manufacturer: $manufacturer"
        Write-Log "Model: $model"
        if ($serial) { Write-Log "Serial: $serial" }

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
        Write-Log "Version compare failed: '$VersionA' vs '$VersionB'. Treating as below minimum." "WARN"
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
                    Write-Log "Stopping process $($proc.ProcessName) PID $($proc.Id)"
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                }
                catch {
                    Write-Log "Unable to stop process $($proc.ProcessName): $($_.Exception.Message)" "WARN"
                }
                break
            }
        }
    }
}

function Invoke-UninstallCommand {
    param(
        [Parameter(Mandatory)]
        [string]$Command
    )

    $trimmed = $Command.Trim()
    Write-Log "Original uninstall command: $trimmed"

    if ($trimmed -match '(?i)msiexec(\.exe)?\s') {
        $guidMatch = [regex]::Match($trimmed, '\{[A-Z0-9\-]+\}', 'IgnoreCase')
        if ($guidMatch.Success) {
            $guid = $guidMatch.Value
            $args = "/x $guid /qn /norestart"
            Write-Log "Executing MSI uninstall: msiexec.exe $args"
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
            return $p.ExitCode
        }
        else {
            $args = $trimmed -replace '^(?i)"?msiexec(\.exe)?"?\s*', ''
            if ($args -notmatch '(/quiet|/qn|/passive)') {
                $args += " /qn /norestart"
            }
            Write-Log "Executing MSI uninstall with args: msiexec.exe $args"
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
            return $p.ExitCode
        }
    }

    $exe = $null
    $args = $null

    if ($trimmed.StartsWith('"')) {
        $parts = $trimmed -split '"', 3
        $exe = $parts[1]
        $args = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }
    }
    else {
        $space = $trimmed.IndexOf(" ")
        if ($space -gt 0) {
            $exe = $trimmed.Substring(0, $space)
            $args = $trimmed.Substring($space + 1).Trim()
        }
        else {
            $exe = $trimmed
            $args = ""
        }
    }

    if (-not $exe) {
        throw "Could not parse uninstall command: $Command"
    }

    if ($args -notmatch '(/quiet|/qn|/passive|/s|/silent|/verysilent)') {
        $args = "$args /S".Trim()
    }

    Write-Log "Executing EXE uninstall: $exe $args"
    $p = Start-Process -FilePath $exe -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
    return $p.ExitCode
}

function Uninstall-App {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$App
    )

    Write-Log "Uninstalling $($App.DisplayName) version '$($App.DisplayVersion)'"

    $cmd = if ($App.QuietUninstallString) { $App.QuietUninstallString } else { $App.UninstallString }
    if (-not $cmd) {
        throw "No uninstall command found for $($App.DisplayName)"
    }

    Stop-DellProcesses
    $exitCode = Invoke-UninstallCommand -Command $cmd
    Write-Log "Uninstall exit code: $exitCode"

    if ($exitCode -notin 0,1605,1614,1641,3010) {
        throw "Unexpected uninstall exit code $exitCode for $($App.DisplayName)"
    }

    Start-Sleep -Seconds 5
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
                Write-Log "Removing stale directory: $path"
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
            }
            catch {
                Write-Log "Could not remove stale directory $path: $($_.Exception.Message)" "WARN"
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
            Write-Log "Could not force TLS 1.2: $($_.Exception.Message)" "WARN"
        }
    }

    $destination = Join-Path $WorkingDirectory $script:LatestDDPM.FileName

    if (Test-Path -LiteralPath $destination) {
        try {
            $existingHash = Get-FileSha256 -Path $destination
            if ($existingHash -eq $script:LatestDDPM.Sha256) {
                Write-Log "Installer already present and checksum matches: $destination"
                return $destination
            }

            Write-Log "Existing installer checksum mismatch. Re-downloading." "WARN"
            Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log "Failed to validate existing installer: $($_.Exception.Message)" "WARN"
            Remove-Item -LiteralPath $destination -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Log "Downloading DDPM package"
    Write-Log "Source: $($script:LatestDDPM.Url)"
    Write-Log "Destination: $destination"

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
    Write-Log "Downloaded file SHA256: $hash"

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

    Write-Log "Installing DDPM from $InstallerPath"
    Write-Log "Silent arguments: $($script:LatestDDPM.SilentArgs)"

    $p = Start-Process -FilePath $InstallerPath -ArgumentList $script:LatestDDPM.SilentArgs -Wait -PassThru -WindowStyle Hidden
    Write-Log "Installer exit code: $($p.ExitCode)"

    if ($p.ExitCode -notin 0,1641,3010) {
        throw "DDPM installer returned unexpected exit code $($p.ExitCode)"
    }

    Start-Sleep -Seconds 15
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

    Write-Log "Installed DDPM detected: $($installed.DisplayVersion)"

    if (Test-VersionLessThan -VersionA $installed.DisplayVersion -VersionB $MinimumSafeVersion) {
        throw "Installed DDPM version '$($installed.DisplayVersion)' is still below minimum '$MinimumSafeVersion'."
    }
}

Write-Log "========== Dell DPM/DDPM compliance script =========="
Write-Log "Mode: $Mode"
Write-Log "MinimumSafeVersion: $MinimumSafeVersion"
Write-Log "SkipNonDell: $SkipNonDell"
Write-Log "WorkingDirectory: $WorkingDirectory"

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
            Write-Log "$status :: $($app.DisplayName) :: Version=$($app.DisplayVersion) :: Normalized=$($app.NormalizedVersion)"
        }
    }

    if ($Mode -eq "Detect") {
        if ($initial.BelowMinimumApps.Count -gt 0) {
            Write-Log "Detection result: NON-COMPLIANT" "WARN"
            exit 1
        }
        else {
            Write-Log "Detection result: COMPLIANT" "SUCCESS"
            exit 0
        }
    }

    Require-RemediationPackageMetadata

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

    $installer = Download-LatestDDPM
    Install-DDPM -InstallerPath $installer
    Assert-TargetDDPMDetected
    Assert-NoBelowMinimumRemains

    Write-Log "Remediation result: COMPLIANT" "SUCCESS"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
    exit 2
}
finally {
    $elapsed = (Get-Date) - $script:StartTime
    Write-Log ("Completed in {0:n1} seconds" -f $elapsed.TotalSeconds)
}