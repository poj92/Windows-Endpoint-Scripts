<#
.SYNOPSIS
    Datto RMM - Install Required .NET Runtimes

.DESCRIPTION
    Installs/verifies the .NET runtimes commonly required by apps:
    This script will install .Net even if not found on the machine before running the script.

Author: Peter Opeyemi James

.NOTES
    Intended for Datto RMM running as SYSTEM/Admin.

    Variables:
    TargetChannel
        The .NET channel to target. Example: 8.0
    TargetVersion
        The minimum version to install. Example: 8.0.27
    Architecture
        The architecture to install. Valid values: x64, x86, arm64
    InstallAspNetCoreRuntime
        Whether to install the ASP.NET Core Runtime. Valid values: true, false
    InstallWindowsDesktopRuntime
        Whether to install the .NET Desktop Runtime. Valid values: true, false
    ReinstallIfPresent
        Whether to reinstall the runtime even if a compliant version is already installed. Valid values: true, false
    FailIfMissingAfterInstall
        Whether to fail the script if the required runtimes are still missing after installation. Valid values: true, false
#>

$ErrorActionPreference = 'Stop'

# ============================================================
# Environment helpers
# ============================================================

function Get-EnvString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [string]$Default = ''
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')

    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($Name, 'Machine')
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    return $value
}

function Get-EnvBool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [bool]$Default
    )

    $raw = Get-EnvString -Name $Name -Default $null

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $Default
    }

    switch -Regex ($raw.Trim()) {
        '^(1|true|yes|y|on)$'  { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default                { return $Default }
    }
}

# ============================================================
# Datto RMM variables
# ============================================================

$script:TargetChannel = Get-EnvString `
    -Name 'TargetChannel' `
    -Default '8.0'

$script:MinimumVersion = Get-EnvString `
    -Name 'MinimumVersion' `
    -Default '8.0.27'

$script:Architecture = Get-EnvString `
    -Name 'Architecture' `
    -Default 'x64'

$script:InstallAspNetCoreRuntime = Get-EnvBool `
    -Name 'InstallAspNetCoreRuntime' `
    -Default $true

$script:InstallWindowsDesktopRuntime = Get-EnvBool `
    -Name 'InstallWindowsDesktopRuntime' `
    -Default $true

$script:ReinstallIfPresent = Get-EnvBool `
    -Name 'ReinstallIfPresent' `
    -Default $false

$script:FailIfMissingAfterInstall = Get-EnvBool `
    -Name 'FailIfMissingAfterInstall' `
    -Default $true

$script:ForceTls12 = Get-EnvBool `
    -Name 'ForceTls12' `
    -Default $true

$script:WorkingDirectory = Get-EnvString `
    -Name 'WorkingDirectory' `
    -Default 'C:\ProgramData\DattoRMM\Packages\DotNet-Runtimes'

$script:LogPath = Get-EnvString `
    -Name 'LogPath' `
    -Default 'C:\ProgramData\DattoRMM\Logs\DotNet-Runtimes.log'

$script:OfflineInstallerFolder = Get-EnvString `
    -Name 'OfflineInstallerFolder' `
    -Default ''

# ============================================================
# Validation and setup
# ============================================================

if ($script:Architecture -notin @('x64', 'x86', 'arm64')) {
    throw ("Invalid Architecture value [{0}]. Valid values: x64, x86, arm64." -f $script:Architecture)
}

if ($script:TargetChannel -notmatch '^\d+\.\d+$') {
    throw ("Invalid TargetChannel [{0}]. Expected format like 8.0." -f $script:TargetChannel)
}

try {
    [version]$script:MinimumVersionObject = $script:MinimumVersion
}
catch {
    throw ("Invalid MinimumVersion [{0}]. Expected format like 8.0.27." -f $script:MinimumVersion)
}

New-Item -Path $script:WorkingDirectory -ItemType Directory -Force | Out-Null
New-Item -Path (Split-Path -Path $script:LogPath -Parent) -ItemType Directory -Force | Out-Null

# ============================================================
# Logging
# ============================================================

function Write-Log {
    param(
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line

    try {
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    }
    catch {
        Write-Host ("Could not write to log file. {0}" -f $_.Exception.Message)
    }
}

# ============================================================
# Globals
# ============================================================

if ($script:ForceTls12) {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    catch {
        Write-Log ("Could not force TLS 1.2. {0}" -f $_.Exception.Message) 'WARN'
    }
}

$script:ProgramFiles64 = $env:ProgramW6432

if ([string]::IsNullOrWhiteSpace($script:ProgramFiles64)) {
    $script:ProgramFiles64 = $env:ProgramFiles
}

$script:ProgramFilesX86 = ${env:ProgramFiles(x86)}
$script:Failures = New-Object System.Collections.Generic.List[string]

# ============================================================
# Helper functions
# ============================================================

function Convert-VersionSafe {
    param(
        [string]$Version
    )

    try {
        return [version]$Version
    }
    catch {
        return [version]'0.0.0'
    }
}

function Get-FrameworkComponentInfo {
    param(
        [string]$Framework
    )

    switch ($Framework) {
        'Microsoft.AspNetCore.App' {
            return [pscustomobject]@{
                AkaName = 'aspnetcore-runtime'
                Prefix  = 'aspnetcore-runtime'
                Label   = 'ASP.NET Core Runtime'
            }
        }
        'Microsoft.WindowsDesktop.App' {
            return [pscustomobject]@{
                AkaName = 'windowsdesktop-runtime'
                Prefix  = 'windowsdesktop-runtime'
                Label   = '.NET Desktop Runtime'
            }
        }
        'Microsoft.NETCore.App' {
            return [pscustomobject]@{
                AkaName = 'dotnet-runtime'
                Prefix  = 'dotnet-runtime'
                Label   = '.NET Runtime'
            }
        }
        default {
            throw ("Unsupported framework: {0}" -f $Framework)
        }
    }
}

function Get-DotNetSharedBasePath {
    param(
        [string]$Arch
    )

    if ($Arch -eq 'x86') {
        if (-not [string]::IsNullOrWhiteSpace($script:ProgramFilesX86)) {
            return (Join-Path -Path $script:ProgramFilesX86 -ChildPath 'dotnet\shared')
        }

        return $null
    }

    return (Join-Path -Path $script:ProgramFiles64 -ChildPath 'dotnet\shared')
}

function Get-InstalledRuntimeVersions {
    param(
        [string]$Framework,
        [string]$Arch
    )

    $base = Get-DotNetSharedBasePath -Arch $Arch

    if ([string]::IsNullOrWhiteSpace($base)) {
        return @()
    }

    $frameworkPath = Join-Path -Path $base -ChildPath $Framework

    if (-not (Test-Path -LiteralPath $frameworkPath)) {
        return @()
    }

    return @(
        Get-ChildItem -Path $frameworkPath -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+\.\d+\.\d+' } |
            Sort-Object { Convert-VersionSafe -Version $_.Name } -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    Framework = $Framework
                    Version   = $_.Name
                    VersionObject = Convert-VersionSafe -Version $_.Name
                    Arch      = $Arch
                    Path      = $_.FullName
                }
            }
    )
}

function Get-HighestRuntimeForChannel {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    return @(
        Get-InstalledRuntimeVersions -Framework $Framework -Arch $Arch |
            Where-Object { $_.Version -like "$Channel.*" } |
            Sort-Object VersionObject -Descending |
            Select-Object -First 1
    )
}

function Test-RuntimeMinimumInstalled {
    param(
        [string]$Framework,
        [string]$Channel,
        [version]$MinimumVersion,
        [string]$Arch
    )

    $runtime = Get-HighestRuntimeForChannel -Framework $Framework -Channel $Channel -Arch $Arch

    if (-not $runtime) {
        return $false
    }

    if ($runtime.VersionObject -ge $MinimumVersion) {
        return $true
    }

    return $false
}

function Find-OfflineInstaller {
    param(
        [string]$Framework,
        [string]$Arch
    )

    if ([string]::IsNullOrWhiteSpace($script:OfflineInstallerFolder)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $script:OfflineInstallerFolder)) {
        Write-Log ("OfflineInstallerFolder does not exist: {0}" -f $script:OfflineInstallerFolder) 'WARN'
        return $null
    }

    $component = Get-FrameworkComponentInfo -Framework $Framework
    $rid = "win-$Arch"

    $files = Get-ChildItem -Path $script:OfflineInstallerFolder -Filter '*.exe' -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "$($component.Prefix)*$rid.exe"
        } |
        Sort-Object LastWriteTime -Descending

    $file = $files | Select-Object -First 1

    if ($file) {
        return $file.FullName
    }

    return $null
}

function Get-AkaInstallerUrl {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    $component = Get-FrameworkComponentInfo -Framework $Framework

    return ("https://aka.ms/dotnet/{0}/{1}-win-{2}.exe" -f $Channel, $component.AkaName, $Arch)
}

function Download-Installer {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    $component = Get-FrameworkComponentInfo -Framework $Framework

    $offline = Find-OfflineInstaller -Framework $Framework -Arch $Arch

    if ($offline) {
        Write-Log ("Using offline installer: {0}" -f $offline)
        return $offline
    }

    $installerDir = Join-Path -Path $script:WorkingDirectory -ChildPath 'Installers'
    New-Item -Path $installerDir -ItemType Directory -Force | Out-Null

    $installerName = "{0}-{1}-latest-win-{2}.exe" -f $component.Prefix, $Channel, $Arch
    $installerPath = Join-Path -Path $installerDir -ChildPath $installerName

    $url = Get-AkaInstallerUrl -Framework $Framework -Channel $Channel -Arch $Arch

    Write-Log ("Downloading {0} for .NET {1} {2} from {3}" -f $component.Label, $Channel, $Arch, $url)

    if (Test-Path -LiteralPath $installerPath) {
        Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
    }

    Invoke-WebRequest `
        -Uri $url `
        -OutFile $installerPath `
        -UseBasicParsing `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $installerPath)) {
        throw ("Download did not create installer: {0}" -f $installerPath)
    }

    $length = (Get-Item -LiteralPath $installerPath).Length

    if ($length -lt 1000000) {
        throw ("Downloaded file looks too small to be a runtime installer. Path: {0}; Size: {1} bytes" -f $installerPath, $length)
    }

    Write-Log ("Downloaded installer: {0} ({1} bytes)" -f $installerPath, $length)

    return $installerPath
}

function Install-DotNetRuntime {
    param(
        [string]$Framework,
        [string]$Channel,
        [version]$MinimumVersion,
        [string]$Arch
    )

    $component = Get-FrameworkComponentInfo -Framework $Framework

    Write-Log ("Checking {0} {1} {2} or later..." -f $Framework, $MinimumVersion.ToString(), $Arch)

    $existing = Get-HighestRuntimeForChannel -Framework $Framework -Channel $Channel -Arch $Arch

    if ($existing) {
        Write-Log ("Highest installed {0} {1}: {2} [{3}]" -f $Framework, $Arch, $existing.Version, $existing.Path)
    }
    else {
        Write-Log ("No installed {0} {1} runtime found for channel {2}." -f $Framework, $Arch, $Channel) 'WARN'
    }

    if ((Test-RuntimeMinimumInstalled -Framework $Framework -Channel $Channel -MinimumVersion $MinimumVersion -Arch $Arch) -and -not $script:ReinstallIfPresent) {
        Write-Log ("Already compliant: {0} {1} {2} or later is installed." -f $Framework, $MinimumVersion.ToString(), $Arch) 'SUCCESS'
        return $true
    }

    try {
        $installerPath = Download-Installer -Framework $Framework -Channel $Channel -Arch $Arch

        $installLogDir = Join-Path -Path $script:WorkingDirectory -ChildPath 'InstallerLogs'
        New-Item -Path $installLogDir -ItemType Directory -Force | Out-Null

        $safeName = "{0}-{1}-{2}-{3}" -f $Framework, $Channel, $Arch, ((Get-Date).ToString('yyyyMMddHHmmss'))
        $safeName = $safeName -replace '[^\w\.-]', '_'

        $installerLog = Join-Path -Path $installLogDir -ChildPath "$safeName.log"

        Write-Log ("Installing {0} using [{1}]" -f $component.Label, $installerPath)

        $arguments = @(
            '/install',
            '/quiet',
            '/norestart',
            '/log',
            $installerLog
        )

        $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru

        Write-Log ("Installer exit code for {0} {1} {2}: {3}" -f $Framework, $Channel, $Arch, $process.ExitCode)

        if ($process.ExitCode -notin @(0, 3010, 1641)) {
            throw ("Installer failed with exit code {0}. Installer log: {1}" -f $process.ExitCode, $installerLog)
        }

        $post = Get-HighestRuntimeForChannel -Framework $Framework -Channel $Channel -Arch $Arch

        if ($post) {
            Write-Log ("Post-install highest {0} {1}: {2} [{3}]" -f $Framework, $Arch, $post.Version, $post.Path)
        }

        if (Test-RuntimeMinimumInstalled -Framework $Framework -Channel $Channel -MinimumVersion $MinimumVersion -Arch $Arch) {
            Write-Log ("Verified installed/compliant: {0} {1} {2} or later." -f $Framework, $MinimumVersion.ToString(), $Arch) 'SUCCESS'
            return $true
        }

        throw ("Post-install verification failed for {0} {1} {2} or later." -f $Framework, $MinimumVersion.ToString(), $Arch)
    }
    catch {
        $message = "Failed to install {0} {1} {2} or later. {3}" -f $Framework, $MinimumVersion.ToString(), $Arch, $_.Exception.Message
        Write-Log $message 'ERROR'
        [void]$script:Failures.Add($message)
        return $false
    }
}

# ============================================================
# Main
# ============================================================

Write-Log '========== .NET Runtime Installer =========='
Write-Log ("TargetChannel: {0}" -f $script:TargetChannel)
Write-Log ("MinimumVersion: {0}" -f $script:MinimumVersion)
Write-Log ("Architecture: {0}" -f $script:Architecture)
Write-Log ("InstallAspNetCoreRuntime: {0}" -f $script:InstallAspNetCoreRuntime)
Write-Log ("InstallWindowsDesktopRuntime: {0}" -f $script:InstallWindowsDesktopRuntime)
Write-Log ("ReinstallIfPresent: {0}" -f $script:ReinstallIfPresent)
Write-Log ("FailIfMissingAfterInstall: {0}" -f $script:FailIfMissingAfterInstall)
Write-Log ("WorkingDirectory: {0}" -f $script:WorkingDirectory)
Write-Log ("LogPath: {0}" -f $script:LogPath)
Write-Log ("OfflineInstallerFolder: {0}" -f $script:OfflineInstallerFolder)

try {
    if ($script:InstallAspNetCoreRuntime) {
        [void](Install-DotNetRuntime `
            -Framework 'Microsoft.AspNetCore.App' `
            -Channel $script:TargetChannel `
            -MinimumVersion $script:MinimumVersionObject `
            -Arch $script:Architecture)
    }
    else {
        Write-Log "InstallAspNetCoreRuntime=False. ASP.NET Core Runtime skipped." 'WARN'
    }

    if ($script:InstallWindowsDesktopRuntime) {
        [void](Install-DotNetRuntime `
            -Framework 'Microsoft.WindowsDesktop.App' `
            -Channel $script:TargetChannel `
            -MinimumVersion $script:MinimumVersionObject `
            -Arch $script:Architecture)
    }
    else {
        Write-Log "InstallWindowsDesktopRuntime=False. .NET Desktop Runtime skipped." 'WARN'
    }

    Write-Log '========== Installed .NET shared runtimes after install =========='

    foreach ($framework in @(
        'Microsoft.NETCore.App',
        'Microsoft.AspNetCore.App',
        'Microsoft.WindowsDesktop.App'
    )) {
        $installed = Get-InstalledRuntimeVersions -Framework $framework -Arch $script:Architecture

        foreach ($runtime in $installed) {
            Write-Log ("{0} {1} {2} [{3}]" -f $runtime.Framework, $runtime.Version, $runtime.Arch, $runtime.Path)
        }
    }

    $missing = New-Object System.Collections.Generic.List[string]

    if ($script:InstallAspNetCoreRuntime) {
        if (-not (Test-RuntimeMinimumInstalled -Framework 'Microsoft.AspNetCore.App' -Channel $script:TargetChannel -MinimumVersion $script:MinimumVersionObject -Arch $script:Architecture)) {
            [void]$missing.Add(("Microsoft.AspNetCore.App {0} {1} or later" -f $script:MinimumVersion, $script:Architecture))
        }
    }

    if ($script:InstallWindowsDesktopRuntime) {
        if (-not (Test-RuntimeMinimumInstalled -Framework 'Microsoft.WindowsDesktop.App' -Channel $script:TargetChannel -MinimumVersion $script:MinimumVersionObject -Arch $script:Architecture)) {
            [void]$missing.Add(("Microsoft.WindowsDesktop.App {0} {1} or later" -f $script:MinimumVersion, $script:Architecture))
        }
    }

    if ($script:Failures.Count -gt 0) {
        Write-Log ("Install failures recorded: {0}" -f $script:Failures.Count) 'ERROR'

        foreach ($failure in $script:Failures) {
            Write-Log ("Failure: {0}" -f $failure) 'ERROR'
        }

        exit 1
    }

    if ($missing.Count -gt 0) {
        Write-Log ("Missing required runtimes after install: {0}" -f ($missing -join '; ')) 'ERROR'

        if ($script:FailIfMissingAfterInstall) {
            exit 1
        }

        exit 0
    }

    Write-Log "Required .NET runtimes are installed." 'SUCCESS'
    exit 0
}
catch {
    Write-Log ("Unhandled error: {0}" -f $_.Exception.Message) 'ERROR'
    exit 1
}