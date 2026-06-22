<#
Author: Peter Opeyemi James
Company: Nexus Open Ssystems Ltd

.SYNOPSIS    
Datto RMM - Generic .NET App Roll-Forward Repair

.DESCRIPTION
    Repairs selected framework-dependent .NET applications by patching the app's
    *.runtimeconfig.json file and optionally installing the required newer .NET
    runtime family.

    This is intended for apps that keep asking for an older .NET runtime, such as
    .NET 6, where you want the app to use a later installed/supported runtime.

    Key behaviour:
      - Does NOT set DOTNET_ROLL_FORWARD globally
      - Does NOT reinstall old/EOL .NET runtimes
      - Patches only selected app runtimeconfig.json files
      - Backs up runtimeconfig.json before changing it
      - Can install target .NET Runtime / ASP.NET Core Runtime / Windows Desktop Runtime
      - Supports x86, x64, auto, or both architecture handling
      - Logs all actions clearly

.EXAMPLE - IPSMonitor / Brother iPrint&Scan
    AppFriendlyName=IPSMonitor
    TargetExecutableNames=IPSMonitor.exe
    PathIncludeRegex=IPSMonitor
    RuntimeFamilies=Microsoft.AspNetCore.App
    Architecture=x86
    TargetRuntimeChannel=10.0
    RollForward=LatestMajor
    InstallRuntime=true
    Remediate=true
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

function Get-EnvInt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [int]$Default
    )

    $raw = Get-EnvString -Name $Name -Default $null

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $Default
    }

    $parsed = 0

    if ([int]::TryParse($raw.Trim(), [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function Split-List {
    param(
        [AllowNull()]
        [string]$Raw
    )

    if ([string]::IsNullOrWhiteSpace($Raw)) {
        return @()
    }

    return @(
        $Raw -split '[;\r\n,]+' |
            ForEach-Object { $_.Trim().Trim('"') } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

# ============================================================
# Datto RMM variables
# ============================================================

$script:Remediate = Get-EnvBool `
    -Name 'Remediate' `
    -Default $false

$script:AppFriendlyName = Get-EnvString `
    -Name 'AppFriendlyName' `
    -Default 'Generic .NET App'

$script:TargetExecutableNamesRaw = Get-EnvString `
    -Name 'TargetExecutableNames' `
    -Default ''

$script:RuntimeConfigPathsRaw = Get-EnvString `
    -Name 'RuntimeConfigPaths' `
    -Default ''

$script:SearchRootsRaw = Get-EnvString `
    -Name 'SearchRoots' `
    -Default ''

$script:PathIncludeRegex = Get-EnvString `
    -Name 'PathIncludeRegex' `
    -Default ''

$script:RuntimeFamiliesRaw = Get-EnvString `
    -Name 'RuntimeFamilies' `
    -Default 'auto'

$script:Architecture = Get-EnvString `
    -Name 'Architecture' `
    -Default 'auto'

$script:RollForward = Get-EnvString `
    -Name 'RollForward' `
    -Default 'LatestMajor'

$script:TargetRuntimeChannel = Get-EnvString `
    -Name 'TargetRuntimeChannel' `
    -Default '8.0'

$script:InstallRuntime = Get-EnvBool `
    -Name 'InstallRuntime' `
    -Default $false

$script:AddBaseRuntimeDependency = Get-EnvBool `
    -Name 'AddBaseRuntimeDependency' `
    -Default $true

$script:CreateBackup = Get-EnvBool `
    -Name 'CreateBackup' `
    -Default $true

$script:StopProcesses = Get-EnvBool `
    -Name 'StopProcesses' `
    -Default $true

$script:FailIfNoTarget = Get-EnvBool `
    -Name 'FailIfNoTarget' `
    -Default $true

$script:FailIfInstallFails = Get-EnvBool `
    -Name 'FailIfInstallFails' `
    -Default $true

$script:ForceTls12 = Get-EnvBool `
    -Name 'ForceTls12' `
    -Default $true

$script:MaxRuntimeConfigFiles = Get-EnvInt `
    -Name 'MaxRuntimeConfigFiles' `
    -Default 5000

$script:WorkingDirectory = Get-EnvString `
    -Name 'WorkingDirectory' `
    -Default 'C:\ProgramData\DattoRMM\Packages\DotNet-App-RollForward'

$script:LogPath = Get-EnvString `
    -Name 'LogPath' `
    -Default 'C:\ProgramData\DattoRMM\Logs\DotNet-App-RollForward.log'

$script:OfflineInstallerFolder = Get-EnvString `
    -Name 'OfflineInstallerFolder' `
    -Default ''

$script:ReleaseIndexUrl = Get-EnvString `
    -Name 'ReleaseIndexUrl' `
    -Default 'https://builds.dotnet.microsoft.com/dotnet/release-metadata/releases-index.json'

# ============================================================
# Validation
# ============================================================

$script:TargetExecutableNames = Split-List -Raw $script:TargetExecutableNamesRaw
$script:RuntimeConfigPaths = Split-List -Raw $script:RuntimeConfigPathsRaw

$validRollForwardValues = @(
    'Minor',
    'Major',
    'LatestPatch',
    'LatestMinor',
    'LatestMajor',
    'Disable'
)

if ($script:RollForward -notin $validRollForwardValues) {
    throw ("Invalid RollForward value [{0}]. Valid values: {1}" -f $script:RollForward, ($validRollForwardValues -join ', '))
}

if ($script:Architecture -notin @('auto', 'x86', 'x64', 'both')) {
    throw ("Invalid Architecture value [{0}]. Valid values: auto, x86, x64, both" -f $script:Architecture)
}

if ($script:TargetRuntimeChannel -notmatch '^\d+\.\d+$') {
    throw ("Invalid TargetRuntimeChannel [{0}]. Expected format like 8.0 or 10.0" -f $script:TargetRuntimeChannel)
}

if (
    $script:TargetExecutableNames.Count -eq 0 -and
    $script:RuntimeConfigPaths.Count -eq 0 -and
    [string]::IsNullOrWhiteSpace($script:PathIncludeRegex)
) {
    throw "You must set at least one of: TargetExecutableNames, RuntimeConfigPaths, or PathIncludeRegex. This prevents accidental broad patching."
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
$script:ReleaseIndexCache = $null
$script:ReleaseJsonCache = @{}
$script:Failures = New-Object System.Collections.Generic.List[string]

# ============================================================
# Utility functions
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

function Normalize-FrameworkName {
    param(
        [AllowNull()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    switch -Regex ($Name.Trim()) {
        '^(Microsoft\.)?NETCore(\.App)?$|^runtime$|^netcore$|^dotnet$' {
            return 'Microsoft.NETCore.App'
        }
        '^(Microsoft\.)?AspNetCore(\.App)?$|^aspnet$|^aspnetcore$|^asp\.net$' {
            return 'Microsoft.AspNetCore.App'
        }
        '^(Microsoft\.)?WindowsDesktop(\.App)?$|^desktop$|^windowsdesktop$' {
            return 'Microsoft.WindowsDesktop.App'
        }
        default {
            return $null
        }
    }
}

function Get-FrameworkComponentInfo {
    param(
        [string]$Framework
    )

    switch ($Framework) {
        'Microsoft.NETCore.App' {
            return [pscustomobject]@{
                Node   = 'runtime'
                Prefix = 'dotnet-runtime'
                Label  = '.NET Runtime'
            }
        }
        'Microsoft.AspNetCore.App' {
            return [pscustomobject]@{
                Node   = 'aspnetcore-runtime'
                Prefix = 'aspnetcore-runtime'
                Label  = 'ASP.NET Core Runtime'
            }
        }
        'Microsoft.WindowsDesktop.App' {
            return [pscustomobject]@{
                Node   = 'windowsdesktop'
                Prefix = 'windowsdesktop-runtime'
                Label  = '.NET Windows Desktop Runtime'
            }
        }
        default {
            return $null
        }
    }
}

function Get-TargetProcessNames {
    $names = New-Object System.Collections.Generic.List[string]

    foreach ($exe in $script:TargetExecutableNames) {
        $clean = [System.IO.Path]::GetFileNameWithoutExtension($exe)

        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            [void]$names.Add($clean)
        }
    }

    foreach ($path in $script:RuntimeConfigPaths) {
        $fileName = [System.IO.Path]::GetFileName($path)

        if ($fileName -like '*.runtimeconfig.json') {
            $stem = $fileName -replace '\.runtimeconfig\.json$', ''

            if (-not [string]::IsNullOrWhiteSpace($stem)) {
                [void]$names.Add($stem)
            }
        }
    }

    return @($names | Select-Object -Unique)
}

function Get-PeArchitecture {
    param(
        [string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        $fs = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )

        try {
            $br = New-Object System.IO.BinaryReader($fs)

            if ($br.ReadUInt16() -ne 0x5A4D) {
                return $null
            }

            $fs.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
            $peOffset = $br.ReadInt32()

            if ($peOffset -lt 0 -or $peOffset -gt ($fs.Length - 6)) {
                return $null
            }

            $fs.Seek($peOffset, [System.IO.SeekOrigin]::Begin) | Out-Null

            if ($br.ReadUInt32() -ne 0x00004550) {
                return $null
            }

            $machine = $br.ReadUInt16()

            switch ($machine) {
                0x8664 { return 'x64' }
                0x014c { return 'x86' }
                default { return $null }
            }
        }
        finally {
            $fs.Dispose()
        }
    }
    catch {
        return $null
    }
}

function Get-ExePathForRuntimeConfig {
    param(
        [System.IO.FileInfo]$RuntimeConfig
    )

    $stem = $RuntimeConfig.Name -replace '\.runtimeconfig\.json$', ''
    $sameNameExe = Join-Path -Path $RuntimeConfig.DirectoryName -ChildPath "$stem.exe"

    if (Test-Path -LiteralPath $sameNameExe) {
        return $sameNameExe
    }

    foreach ($exeName in $script:TargetExecutableNames) {
        $candidate = Join-Path -Path $RuntimeConfig.DirectoryName -ChildPath $exeName

        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-ArchitecturesForRuntimeConfig {
    param(
        [System.IO.FileInfo]$RuntimeConfig
    )

    if ($script:Architecture -eq 'x86') {
        return @('x86')
    }

    if ($script:Architecture -eq 'x64') {
        return @('x64')
    }

    if ($script:Architecture -eq 'both') {
        return @('x64', 'x86')
    }

    $exePath = Get-ExePathForRuntimeConfig -RuntimeConfig $RuntimeConfig

    if ($exePath) {
        $peArch = Get-PeArchitecture -Path $exePath

        if ($peArch -in @('x86', 'x64')) {
            return @($peArch)
        }
    }

    if (
        -not [string]::IsNullOrWhiteSpace($script:ProgramFilesX86) -and
        $RuntimeConfig.FullName.StartsWith($script:ProgramFilesX86, [StringComparison]::OrdinalIgnoreCase)
    ) {
        return @('x86')
    }

    return @('x64')
}

function Get-SearchRoots {
    if (-not [string]::IsNullOrWhiteSpace($script:SearchRootsRaw)) {
        return @(
            Split-List -Raw $script:SearchRootsRaw |
                Where-Object { Test-Path -LiteralPath $_ }
        )
    }

    $roots = New-Object System.Collections.Generic.List[string]

    foreach ($root in @($script:ProgramFiles64, $script:ProgramFilesX86, 'C:\ProgramData')) {
        if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-Path -LiteralPath $root)) {
            [void]$roots.Add($root)
        }
    }

    return @($roots | Select-Object -Unique)
}

function Get-FrameworksFromRuntimeConfigJson {
    param(
        $Json
    )

    $frameworks = @()

    try {
        if ($Json.runtimeOptions.PSObject.Properties.Name -contains 'framework') {
            if ($null -ne $Json.runtimeOptions.framework) {
                $frameworks += $Json.runtimeOptions.framework
            }
        }

        if ($Json.runtimeOptions.PSObject.Properties.Name -contains 'frameworks') {
            if ($null -ne $Json.runtimeOptions.frameworks) {
                foreach ($framework in @($Json.runtimeOptions.frameworks)) {
                    $frameworks += $framework
                }
            }
        }
    }
    catch {
        return @()
    }

    return @($frameworks)
}

function Test-RuntimeConfigPathMatchesTarget {
    param(
        [System.IO.FileInfo]$File
    )

    $stem = $File.Name -replace '\.runtimeconfig\.json$', ''

    foreach ($exeName in $script:TargetExecutableNames) {
        $targetStem = [System.IO.Path]::GetFileNameWithoutExtension($exeName)

        if ($stem -ieq $targetStem) {
            return $true
        }

        $candidateExe = Join-Path -Path $File.DirectoryName -ChildPath $exeName

        if (Test-Path -LiteralPath $candidateExe) {
            return $true
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:PathIncludeRegex)) {
        if ($File.FullName -match $script:PathIncludeRegex) {
            return $true
        }
    }

    return $false
}

# ============================================================
# Runtimeconfig discovery
# ============================================================

function Find-TargetRuntimeConfigs {
    $targets = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($path in $script:RuntimeConfigPaths) {
        try {
            if (Test-Path -LiteralPath $path) {
                $item = Get-Item -LiteralPath $path -ErrorAction Stop

                if (-not $item.PSIsContainer -and $item.Name -like '*.runtimeconfig.json') {
                    [void]$targets.Add($item)
                    Write-Log ("Added explicit runtimeconfig: {0}" -f $item.FullName)
                }
            }
            else {
                Write-Log ("Explicit runtimeconfig path not found: {0}" -f $path) 'WARN'
            }
        }
        catch {
            Write-Log ("Could not access explicit runtimeconfig path [{0}]. {1}" -f $path, $_.Exception.Message) 'WARN'
        }
    }

    foreach ($root in (Get-SearchRoots)) {
        Write-Log ("Searching for target runtimeconfig files under: {0}" -f $root)

        try {
            $files = Get-ChildItem -Path $root -Filter '*.runtimeconfig.json' -File -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First $script:MaxRuntimeConfigFiles

            foreach ($file in $files) {
                if (Test-RuntimeConfigPathMatchesTarget -File $file) {
                    [void]$targets.Add($file)
                }
            }
        }
        catch {
            Write-Log ("Could not scan root [{0}]. {1}" -f $root, $_.Exception.Message) 'WARN'
        }
    }

    return @($targets | Sort-Object FullName -Unique)
}

# ============================================================
# Installed runtime detection
# ============================================================

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
                    Arch      = $Arch
                    Path      = $_.FullName
                }
            }
    )
}

function Test-RuntimeChannelInstalled {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    $installed = Get-InstalledRuntimeVersions -Framework $Framework -Arch $Arch

    foreach ($runtime in $installed) {
        if ($runtime.Version -like "$Channel.*") {
            return $true
        }
    }

    return $false
}

# ============================================================
# Microsoft release metadata / installer download
# ============================================================

function Get-ReleaseJsonForChannel {
    param(
        [string]$Channel
    )

    if ($script:ReleaseJsonCache.ContainsKey($Channel)) {
        return $script:ReleaseJsonCache[$Channel]
    }

    if (-not $script:ReleaseIndexCache) {
        Write-Log "Downloading .NET release index metadata..."
        $script:ReleaseIndexCache = Invoke-RestMethod -Uri $script:ReleaseIndexUrl -UseBasicParsing -ErrorAction Stop
    }

    $entry = @($script:ReleaseIndexCache.'releases-index') |
        Where-Object { $_.'channel-version' -eq $Channel } |
        Select-Object -First 1

    if (-not $entry) {
        throw ("Could not find .NET release metadata for channel {0}" -f $Channel)
    }

    Write-Log ("Downloading release metadata for .NET channel {0}" -f $Channel)

    $releaseJson = Invoke-RestMethod -Uri $entry.'releases.json' -UseBasicParsing -ErrorAction Stop

    $script:ReleaseJsonCache[$Channel] = $releaseJson

    return $releaseJson
}

function Find-OfflineInstaller {
    param(
        [string]$Framework,
        [string]$Channel,
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

    if (-not $component) {
        return $null
    }

    $rid = "win-$Arch"

    $files = Get-ChildItem -Path $script:OfflineInstallerFolder -Filter '*.exe' -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "$($component.Prefix)*-$rid.exe" -and
            $_.Name -match [regex]::Escape("$Channel.")
        } |
        Sort-Object LastWriteTime -Descending

    $file = $files | Select-Object -First 1

    if ($file) {
        return [pscustomobject]@{
            Name      = $file.Name
            Url       = $null
            Version   = $Channel
            LocalPath = $file.FullName
        }
    }

    return $null
}

function Get-InstallerAsset {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    $offline = Find-OfflineInstaller -Framework $Framework -Channel $Channel -Arch $Arch

    if ($offline) {
        Write-Log ("Using offline installer: {0}" -f $offline.LocalPath)
        return $offline
    }

    $component = Get-FrameworkComponentInfo -Framework $Framework

    if (-not $component) {
        throw ("Unsupported framework for installer lookup: {0}" -f $Framework)
    }

    $rid = "win-$Arch"
    $releaseJson = Get-ReleaseJsonForChannel -Channel $Channel

    $releases = @($releaseJson.releases) |
        Sort-Object { Convert-VersionSafe -Version $_.'release-version' } -Descending

    foreach ($release in $releases) {
        $prop = $release.PSObject.Properties[$component.Node]

        if (-not $prop) {
            continue
        }

        $node = $prop.Value

        if (-not $node -or -not $node.files) {
            continue
        }

        $file = @($node.files) |
            Where-Object {
                $_.rid -eq $rid -and
                $_.url -and
                $_.name -like "$($component.Prefix)*-$rid.exe"
            } |
            Select-Object -First 1

        if ($file) {
            return [pscustomobject]@{
                Name      = $file.name
                Url       = $file.url
                Version   = $node.version
                LocalPath = $null
            }
        }
    }

    throw ("Could not find installer asset for {0} {1} {2}" -f $Framework, $Channel, $Arch)
}

function Install-DotNetRuntime {
    param(
        [string]$Framework,
        [string]$Channel,
        [string]$Arch
    )

    $component = Get-FrameworkComponentInfo -Framework $Framework

    if (-not $component) {
        Write-Log ("Skipping unsupported framework: {0}" -f $Framework) 'WARN'
        return $true
    }

    if (Test-RuntimeChannelInstalled -Framework $Framework -Channel $Channel -Arch $Arch) {
        Write-Log ("{0} {1} {2} is already installed." -f $Framework, $Channel, $Arch) 'SUCCESS'
        return $true
    }

    if (-not $script:Remediate) {
        Write-Log ("Would install {0} {1} {2}, but Remediate=False." -f $Framework, $Channel, $Arch) 'WARN'
        return $true
    }

    try {
        $asset = Get-InstallerAsset -Framework $Framework -Channel $Channel -Arch $Arch

        $installerDir = Join-Path -Path $script:WorkingDirectory -ChildPath 'Installers'
        New-Item -Path $installerDir -ItemType Directory -Force | Out-Null

        if ($asset.LocalPath) {
            $installerPath = $asset.LocalPath
        }
        else {
            $installerPath = Join-Path -Path $installerDir -ChildPath $asset.Name

            if (-not (Test-Path -LiteralPath $installerPath)) {
                Write-Log ("Downloading {0} {1} {2}..." -f $component.Label, $asset.Version, $Arch)
                Invoke-WebRequest -Uri $asset.Url -OutFile $installerPath -UseBasicParsing -ErrorAction Stop
            }
            else {
                Write-Log ("Using cached installer: {0}" -f $installerPath)
            }
        }

        $installLogDir = Join-Path -Path $script:WorkingDirectory -ChildPath 'InstallerLogs'
        New-Item -Path $installLogDir -ItemType Directory -Force | Out-Null

        $safeName = "{0}-{1}-{2}-{3}" -f $Framework, $Channel, $Arch, ((Get-Date).ToString('yyyyMMddHHmmss'))
        $safeName = $safeName -replace '[^\w\.-]', '_'

        $installerLog = Join-Path -Path $installLogDir -ChildPath "$safeName.log"

        Write-Log ("Installing {0} {1} {2} using [{3}]" -f $Framework, $Channel, $Arch, $installerPath)

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

        if (Test-RuntimeChannelInstalled -Framework $Framework -Channel $Channel -Arch $Arch) {
            Write-Log ("Verified installed: {0} {1} {2}" -f $Framework, $Channel, $Arch) 'SUCCESS'
            return $true
        }

        throw ("Post-install verification failed for {0} {1} {2}" -f $Framework, $Channel, $Arch)
    }
    catch {
        $message = "Failed to install {0} {1} {2}. {3}" -f $Framework, $Channel, $Arch, $_.Exception.Message
        Write-Log $message 'ERROR'
        [void]$script:Failures.Add($message)

        if ($script:FailIfInstallFails) {
            return $false
        }

        return $true
    }
}

# ============================================================
# Runtime requirement collection
# ============================================================

function Get-ForcedRuntimeFamilies {
    if ($script:RuntimeFamiliesRaw.Trim().ToLower() -eq 'auto') {
        return @()
    }

    $families = New-Object System.Collections.Generic.List[string]

    foreach ($item in (Split-List -Raw $script:RuntimeFamiliesRaw)) {
        $normalized = Normalize-FrameworkName -Name $item

        if ($normalized) {
            [void]$families.Add($normalized)
        }
        else {
            Write-Log ("Ignoring unsupported RuntimeFamilies item: {0}" -f $item) 'WARN'
        }
    }

    return @($families | Select-Object -Unique)
}

function Get-RuntimeRequirements {
    param(
        [System.IO.FileInfo[]]$RuntimeConfigs
    )

    $requirements = [ordered]@{}
    $forcedFamilies = @(Get-ForcedRuntimeFamilies)

    foreach ($runtimeConfig in $RuntimeConfigs) {
        try {
            $raw = Get-Content -LiteralPath $runtimeConfig.FullName -Raw -ErrorAction Stop
            $json = $raw | ConvertFrom-Json -ErrorAction Stop
            $arches = @(Get-ArchitecturesForRuntimeConfig -RuntimeConfig $runtimeConfig)

            $frameworkNames = New-Object System.Collections.Generic.List[string]

            if ($forcedFamilies.Count -gt 0) {
                foreach ($family in $forcedFamilies) {
                    [void]$frameworkNames.Add($family)
                    Write-Log ("Forced runtime family for {0}: {1}" -f $runtimeConfig.Name, $family)
                }
            }
            else {
                $frameworks = @(Get-FrameworksFromRuntimeConfigJson -Json $json)

                foreach ($framework in $frameworks) {
                    $frameworkNameRaw = ''
                    $frameworkVersionRaw = ''

                    try { $frameworkNameRaw = [string]$framework.name } catch {}
                    try { $frameworkVersionRaw = [string]$framework.version } catch {}

                    $normalized = Normalize-FrameworkName -Name $frameworkNameRaw

                    if ($normalized) {
                        [void]$frameworkNames.Add($normalized)
                        Write-Log ("Detected framework in {0}: {1} {2}" -f $runtimeConfig.Name, $normalized, $frameworkVersionRaw)
                    }
                }
            }

            if ($frameworkNames.Count -eq 0) {
                Write-Log ("No supported framework family found in {0}. Set RuntimeFamilies explicitly if needed." -f $runtimeConfig.FullName) 'WARN'
                continue
            }

            foreach ($arch in $arches) {
                foreach ($frameworkName in ($frameworkNames | Select-Object -Unique)) {
                    $key = "{0}|{1}|{2}" -f $frameworkName, $script:TargetRuntimeChannel, $arch

                    if (-not $requirements.Contains($key)) {
                        $requirements[$key] = [pscustomobject]@{
                            Framework = $frameworkName
                            Channel   = $script:TargetRuntimeChannel
                            Arch      = $arch
                            Source    = $runtimeConfig.FullName
                        }
                    }

                    if ($script:AddBaseRuntimeDependency -and $frameworkName -ne 'Microsoft.NETCore.App') {
                        $baseKey = "Microsoft.NETCore.App|{0}|{1}" -f $script:TargetRuntimeChannel, $arch

                        if (-not $requirements.Contains($baseKey)) {
                            $requirements[$baseKey] = [pscustomobject]@{
                                Framework = 'Microsoft.NETCore.App'
                                Channel   = $script:TargetRuntimeChannel
                                Arch      = $arch
                                Source    = ("Base runtime dependency for {0} from {1}" -f $frameworkName, $runtimeConfig.FullName)
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Log ("Could not collect runtime requirement from {0}. {1}" -f $runtimeConfig.FullName, $_.Exception.Message) 'WARN'
        }
    }

    return @(
        $requirements.Values |
            Sort-Object `
                @{ Expression = {
                    switch ($_.Framework) {
                        'Microsoft.NETCore.App' { 1 }
                        'Microsoft.AspNetCore.App' { 2 }
                        'Microsoft.WindowsDesktop.App' { 3 }
                        default { 9 }
                    }
                }},
                Framework,
                Arch
    )
}

# ============================================================
# Process stop / runtimeconfig patching
# ============================================================

function Stop-TargetProcesses {
    if (-not $script:StopProcesses) {
        return
    }

    foreach ($processName in (Get-TargetProcessNames)) {
        try {
            $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue

            foreach ($process in $processes) {
                Write-Log ("Stopping process {0} PID={1}" -f $process.ProcessName, $process.Id)
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log ("Could not stop process [{0}]. {1}" -f $processName, $_.Exception.Message) 'WARN'
        }
    }
}

function Patch-RuntimeConfig {
    param(
        [System.IO.FileInfo]$RuntimeConfig
    )

    Write-Log ("Processing runtimeconfig: {0}" -f $RuntimeConfig.FullName)

    $raw = Get-Content -LiteralPath $RuntimeConfig.FullName -Raw -ErrorAction Stop

    try {
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Log ("Runtimeconfig is not valid JSON and will not be modified: {0}. {1}" -f $RuntimeConfig.FullName, $_.Exception.Message) 'ERROR'
        return $false
    }

    if (-not $json.runtimeOptions) {
        Write-Log ("No runtimeOptions section found. Skipping: {0}" -f $RuntimeConfig.FullName) 'WARN'
        return $false
    }

    try {
        $frameworks = @(Get-FrameworksFromRuntimeConfigJson -Json $json)

        foreach ($framework in $frameworks) {
            $frameworkName = ''
            $frameworkVersion = ''

            try { $frameworkName = [string]$framework.name } catch {}
            try { $frameworkVersion = [string]$framework.version } catch {}

            if (-not [string]::IsNullOrWhiteSpace($frameworkName)) {
                Write-Log ("Existing framework requirement: {0} {1}" -f $frameworkName, $frameworkVersion)
            }
        }
    }
    catch {
        Write-Log ("Could not enumerate framework requirements from {0}. {1}" -f $RuntimeConfig.FullName, $_.Exception.Message) 'WARN'
    }

    $oldRollForward = $null

    try {
        if ($json.runtimeOptions.PSObject.Properties.Name -contains 'rollForward') {
            $oldRollForward = [string]$json.runtimeOptions.rollForward
        }
    }
    catch {
        $oldRollForward = $null
    }

    if ($oldRollForward -eq $script:RollForward) {
        Write-Log ("Runtimeconfig already has rollForward={0}" -f $script:RollForward) 'SUCCESS'
        return $true
    }

    if (-not $script:Remediate) {
        Write-Log ("Would change rollForward from [{0}] to [{1}], but Remediate=False." -f $oldRollForward, $script:RollForward) 'WARN'
        return $true
    }

    Stop-TargetProcesses

    if ($script:CreateBackup) {
        $backupPath = "{0}.bak-{1}" -f $RuntimeConfig.FullName, ((Get-Date).ToString('yyyyMMdd-HHmmss'))
        Copy-Item -LiteralPath $RuntimeConfig.FullName -Destination $backupPath -Force
        Write-Log ("Backup created: {0}" -f $backupPath)
    }

    try {
        $item = Get-Item -LiteralPath $RuntimeConfig.FullName -ErrorAction Stop

        if ($item.IsReadOnly) {
            $item.IsReadOnly = $false
            Write-Log ("Removed read-only attribute from {0}" -f $RuntimeConfig.FullName)
        }
    }
    catch {
        Write-Log ("Could not check/remove read-only attribute. {0}" -f $_.Exception.Message) 'WARN'
    }

    $newRaw = $raw

    if ($newRaw -match '"rollForward"\s*:') {
        $newRaw = [regex]::Replace(
            $newRaw,
            '"rollForward"\s*:\s*"[^"]*"',
            ('"rollForward": "{0}"' -f $script:RollForward),
            1
        )
    }
    else {
        $runtimeOptionsMatch = [regex]::Match($newRaw, '"runtimeOptions"\s*:\s*\{')

        if (-not $runtimeOptionsMatch.Success) {
            Write-Log ("Could not locate runtimeOptions object text in {0}" -f $RuntimeConfig.FullName) 'ERROR'
            return $false
        }

        $insertAt = $runtimeOptionsMatch.Index + $runtimeOptionsMatch.Length
        $insertText = "`r`n    `"rollForward`": `"$($script:RollForward)`","
        $newRaw = $newRaw.Insert($insertAt, $insertText)
    }

    try {
        $verifyJson = $newRaw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-Log ("Updated JSON failed validation. Original file was not modified. {0}" -f $_.Exception.Message) 'ERROR'
        return $false
    }

    $verifiedRollForward = $null

    try {
        $verifiedRollForward = [string]$verifyJson.runtimeOptions.rollForward
    }
    catch {
        $verifiedRollForward = $null
    }

    if ($verifiedRollForward -ne $script:RollForward) {
        Write-Log ("Post-update validation failed. rollForward was not set to {0}" -f $script:RollForward) 'ERROR'
        return $false
    }

    Set-Content -LiteralPath $RuntimeConfig.FullName -Value $newRaw -Encoding UTF8 -Force

    Write-Log ("Updated rollForward from [{0}] to [{1}]" -f $oldRollForward, $script:RollForward) 'SUCCESS'

    return $true
}

# ============================================================
# Main
# ============================================================

Write-Log '========== Generic .NET App Roll-Forward Repair =========='
Write-Log ("AppFriendlyName: {0}" -f $script:AppFriendlyName)
Write-Log ("Remediate: {0}" -f $script:Remediate)
Write-Log ("TargetExecutableNames: {0}" -f $script:TargetExecutableNamesRaw)
Write-Log ("RuntimeConfigPaths: {0}" -f $script:RuntimeConfigPathsRaw)
Write-Log ("SearchRoots: {0}" -f $script:SearchRootsRaw)
Write-Log ("PathIncludeRegex: {0}" -f $script:PathIncludeRegex)
Write-Log ("RuntimeFamilies: {0}" -f $script:RuntimeFamiliesRaw)
Write-Log ("Architecture: {0}" -f $script:Architecture)
Write-Log ("RollForward: {0}" -f $script:RollForward)
Write-Log ("TargetRuntimeChannel: {0}" -f $script:TargetRuntimeChannel)
Write-Log ("InstallRuntime: {0}" -f $script:InstallRuntime)
Write-Log ("AddBaseRuntimeDependency: {0}" -f $script:AddBaseRuntimeDependency)
Write-Log ("CreateBackup: {0}" -f $script:CreateBackup)
Write-Log ("StopProcesses: {0}" -f $script:StopProcesses)
Write-Log ("FailIfNoTarget: {0}" -f $script:FailIfNoTarget)
Write-Log ("FailIfInstallFails: {0}" -f $script:FailIfInstallFails)
Write-Log ("WorkingDirectory: {0}" -f $script:WorkingDirectory)
Write-Log ("LogPath: {0}" -f $script:LogPath)
Write-Log ("OfflineInstallerFolder: {0}" -f $script:OfflineInstallerFolder)

try {
    $targets = @(Find-TargetRuntimeConfigs)

    if ($targets.Count -eq 0) {
        $message = "No target runtimeconfig files found for {0}." -f $script:AppFriendlyName

        if ($script:FailIfNoTarget) {
            Write-Log $message 'ERROR'
            exit 1
        }

        Write-Log $message 'WARN'
        exit 0
    }

    Write-Log ("Target runtimeconfig files found: {0}" -f $targets.Count)

    foreach ($target in $targets) {
        Write-Log ("Target: {0}" -f $target.FullName)
    }

    if ($script:InstallRuntime) {
        $requirements = @(Get-RuntimeRequirements -RuntimeConfigs $targets)

        if ($requirements.Count -eq 0) {
            Write-Log "No runtime requirements discovered. Runtime installation skipped." 'WARN'
        }
        else {
            Write-Log ("Runtime requirements to validate/install: {0}" -f $requirements.Count)

            foreach ($requirement in $requirements) {
                Write-Log ("Requirement: {0} {1} {2} | Source: {3}" -f $requirement.Framework, $requirement.Channel, $requirement.Arch, $requirement.Source)

                $ok = Install-DotNetRuntime `
                    -Framework $requirement.Framework `
                    -Channel $requirement.Channel `
                    -Arch $requirement.Arch

                if (-not $ok -and $script:FailIfInstallFails) {
                    Write-Log "Stopping because runtime installation failed and FailIfInstallFails=True." 'ERROR'
                    exit 1
                }
            }
        }
    }
    else {
        Write-Log "InstallRuntime=False. Runtime installation skipped." 'WARN'
    }

    $patchedCount = 0

    foreach ($target in $targets) {
        if (Patch-RuntimeConfig -RuntimeConfig $target) {
            $patchedCount++
        }
    }

    Write-Log ("Runtimeconfig files processed successfully: {0} / {1}" -f $patchedCount, $targets.Count) 'SUCCESS'

    Write-Log '========== Installed .NET shared runtimes after repair =========='

    foreach ($arch in @('x64', 'x86')) {
        foreach ($framework in @(
            'Microsoft.NETCore.App',
            'Microsoft.AspNetCore.App',
            'Microsoft.WindowsDesktop.App'
        )) {
            $installed = Get-InstalledRuntimeVersions -Framework $framework -Arch $arch

            foreach ($runtime in $installed) {
                Write-Log ("{0} {1} {2} [{3}]" -f $runtime.Framework, $runtime.Version, $runtime.Arch, $runtime.Path)
            }
        }
    }

    if ($script:Failures.Count -gt 0) {
        Write-Log ("Completed with failures: {0}" -f $script:Failures.Count) 'ERROR'

        foreach ($failure in $script:Failures) {
            Write-Log ("Failure: {0}" -f $failure) 'ERROR'
        }

        exit 1
    }

    Write-Log ("{0} .NET roll-forward repair completed." -f $script:AppFriendlyName) 'SUCCESS'
    exit 0
}
catch {
    Write-Log ("Unhandled error: {0}" -f $_.Exception.Message) 'ERROR'
    exit 1
}