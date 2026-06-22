#requires -version 5.1
<#
.SYNOPSIS
    Datto RMM component to detect and remediate Windows unquoted service paths.

.DESCRIPTION
    Finds Windows service ImagePath registry values where the executable path is unquoted
    and contains spaces. The script quotes only the executable portion and preserves any
    service arguments.

    Example:
        C:\Program Files\Vendor\App\service.exe -k run
    becomes:
        "C:\Program Files\Vendor\App\service.exe" -k run

    It preserves REG_SZ / REG_EXPAND_SZ where possible and backs up changed service keys.

.NOTES
    Recommended first run:
        ReportOnly = true

    Remediation run:
        ReportOnly = false
        RestartChangedRunningServices = false

    Exit codes:
        0 = Success / no unresolved vulnerable paths
        1 = Vulnerable paths remain or report-only found vulnerable paths
        2 = Fatal script error

    VAriables
    | Variable                        |    Type | Recommended default | Purpose                                                                               |
| ------------------------------- | ------: | ------------------: | ------------------------------------------------------------------------------------- |
| `ReportOnly`                    | Boolean |             `false` | `true` detects only. `false` fixes.                                                   |
| `RestartChangedRunningServices` | Boolean |             `false` | Restart services after fixing their registry path. Usually leave false.               |
| `BackupBeforeChange`            | Boolean |              `true` | Exports each changed service registry key before editing.                             |
| `IncludeDriverServices`         | Boolean |             `false` | Usually false. This targets normal Windows services.                                  |
| `ForceFixAmbiguousPaths`        | Boolean |             `false` | Only use true if a scanner still flags paths where the executable cannot be verified. |
| `FailIfVulnerableFound`         | Boolean |              `true` | Returns exit code `1` if unresolved vulnerable paths remain.                          |
| `ServiceNameIncludeRegex`       |  String |               blank | Optional include filter.                                                              |
| `ServiceNameExcludeRegex`       |  String |               blank | Optional exclude filter.                                                              |
| `WorkingDirectory`              |  String |               blank | Optional. Defaults to `C:\ProgramData\DattoRMM\UnquotedServicePathFix`.               |
    
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Get-EnvString {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$Default = ''
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')

    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($Name, 'Machine')
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    return $value.Trim()
}

function ConvertTo-Bool {
    param(
        [string]$Value,
        [bool]$Default = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    switch -Regex ($Value.Trim()) {
        '^(1|true|yes|y|on)$'  { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default                { return $Default }
    }
}

function Get-EnvBool {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [bool]$Default = $false
    )

    return ConvertTo-Bool -Value (Get-EnvString -Name $Name -Default '') -Default $Default
}

$ReportOnly                    = Get-EnvBool   -Name 'ReportOnly' -Default $false
$RestartChangedRunningServices = Get-EnvBool   -Name 'RestartChangedRunningServices' -Default $false
$BackupBeforeChange            = Get-EnvBool   -Name 'BackupBeforeChange' -Default $true
$IncludeDriverServices         = Get-EnvBool   -Name 'IncludeDriverServices' -Default $false
$ForceFixAmbiguousPaths        = Get-EnvBool   -Name 'ForceFixAmbiguousPaths' -Default $false
$FailIfVulnerableFound         = Get-EnvBool   -Name 'FailIfVulnerableFound' -Default $true
$ServiceNameIncludeRegex       = Get-EnvString -Name 'ServiceNameIncludeRegex' -Default ''
$ServiceNameExcludeRegex       = Get-EnvString -Name 'ServiceNameExcludeRegex' -Default ''
$WorkingDirectory              = Get-EnvString -Name 'WorkingDirectory' -Default 'C:\ProgramData\DattoRMM\UnquotedServicePathFix'

$LogPath = Join-Path $WorkingDirectory 'Unquoted-Service-Path-Fix.log'
$BackupDirectory = Join-Path $WorkingDirectory 'Backups'

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

Ensure-Directory -Path $WorkingDirectory
Ensure-Directory -Path $BackupDirectory

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-TestablePath {
    param([Parameter(Mandatory)][string]$Path)

    $expanded = [Environment]::ExpandEnvironmentVariables($Path)

    if ($expanded -match '^\\SystemRoot\\(.+)$') {
        return Join-Path $env:SystemRoot $Matches[1]
    }

    if ($expanded -match '^\\\?\?\\([A-Za-z]:\\.*)$') {
        return $Matches[1]
    }

    if ($expanded -match '^\\\\\?\\([A-Za-z]:\\.*)$') {
        return $Matches[1]
    }

    return $expanded
}

function Get-ServiceImagePathAnalysis {
    param(
        [Parameter(Mandatory)]
        [string]$ImagePath
    )

    $original = $ImagePath
    $trimmed = $ImagePath.Trim()

    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return [pscustomobject]@{
            IsVulnerable        = $false
            IsAlreadyQuoted     = $false
            CanSafelyFix        = $false
            BinaryOriginal      = $null
            Arguments           = $null
            ProposedImagePath   = $null
            Reason              = 'Empty ImagePath'
        }
    }

    if ($trimmed -match '^\s*"([^"]+)"(.*)$') {
        return [pscustomobject]@{
            IsVulnerable        = $false
            IsAlreadyQuoted     = $true
            CanSafelyFix        = $false
            BinaryOriginal      = $Matches[1]
            Arguments           = $Matches[2]
            ProposedImagePath   = $null
            Reason              = 'Already quoted'
        }
    }

    $exeMatches = [regex]::Matches($trimmed, '(?i)\.exe')

    if ($exeMatches.Count -eq 0) {
        return [pscustomobject]@{
            IsVulnerable        = $false
            IsAlreadyQuoted     = $false
            CanSafelyFix        = $false
            BinaryOriginal      = $null
            Arguments           = $null
            ProposedImagePath   = $null
            Reason              = 'No .exe found in ImagePath'
        }
    }

    $candidates = New-Object System.Collections.Generic.List[object]

    foreach ($match in $exeMatches) {
        $binaryEnd = $match.Index + $match.Length
        $binaryOriginal = $trimmed.Substring(0, $binaryEnd)
        $arguments = $trimmed.Substring($binaryEnd)

        $testablePath = ConvertTo-TestablePath -Path $binaryOriginal

        $exists = $false
        try {
            $exists = Test-Path -LiteralPath $testablePath -PathType Leaf
        }
        catch {
            $exists = $false
        }

        $hasSpaceInOriginal = $binaryOriginal -match '\s'
        $hasSpaceWhenExpanded = $testablePath -match '\s'

        $candidates.Add([pscustomobject]@{
            BinaryOriginal       = $binaryOriginal
            Arguments            = $arguments
            TestablePath         = $testablePath
            Exists               = $exists
            HasSpaceInOriginal   = $hasSpaceInOriginal
            HasSpaceWhenExpanded = $hasSpaceWhenExpanded
            HasSpace             = ($hasSpaceInOriginal -or $hasSpaceWhenExpanded)
        })
    }

    $selected = $candidates | Where-Object { $_.Exists } | Select-Object -First 1

    if (-not $selected) {
        $spaceCandidates = @($candidates | Where-Object { $_.HasSpace })

        if ($spaceCandidates.Count -eq 1) {
            $selected = $spaceCandidates[0]
        }
        elseif ($ForceFixAmbiguousPaths -and $spaceCandidates.Count -gt 1) {
            $selected = $spaceCandidates[0]
        }
    }

    if (-not $selected) {
        return [pscustomobject]@{
            IsVulnerable        = $false
            IsAlreadyQuoted     = $false
            CanSafelyFix        = $false
            BinaryOriginal      = $null
            Arguments           = $null
            ProposedImagePath   = $null
            Reason              = 'Could not safely identify executable portion'
        }
    }

    if (-not $selected.HasSpace) {
        return [pscustomobject]@{
            IsVulnerable        = $false
            IsAlreadyQuoted     = $false
            CanSafelyFix        = $false
            BinaryOriginal      = $selected.BinaryOriginal
            Arguments           = $selected.Arguments
            ProposedImagePath   = $null
            Reason              = 'Executable portion does not contain spaces'
        }
    }

    $proposed = '"' + $selected.BinaryOriginal + '"' + $selected.Arguments

    return [pscustomobject]@{
        IsVulnerable        = $true
        IsAlreadyQuoted     = $false
        CanSafelyFix        = $true
        BinaryOriginal      = $selected.BinaryOriginal
        Arguments           = $selected.Arguments
        ProposedImagePath   = $proposed
        Reason              = 'Unquoted executable path contains spaces'
    }
}

function Export-ServiceRegistryBackup {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName
    )

    $safeName = $ServiceName -replace '[^\w\.\-]', '_'
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupFile = Join-Path $BackupDirectory "$safeName-$timestamp.reg"
    $regPath = "HKLM\SYSTEM\CurrentControlSet\Services\$ServiceName"

    $process = Start-Process -FilePath "$env:SystemRoot\System32\reg.exe" `
        -ArgumentList @('export', $regPath, $backupFile, '/y') `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath $backupFile)) {
        Write-Log "Registry backup created for service '$ServiceName': $backupFile" 'INFO'
        return $true
    }

    Write-Log "Registry backup failed for service '$ServiceName'. reg.exe exit code: $($process.ExitCode)" 'WARN'
    return $false
}

function Restart-ServiceIfRequested {
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName
    )

    if (-not $RestartChangedRunningServices) {
        return
    }

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop

        if ($service.Status -ne 'Running') {
            Write-Log "Service '$ServiceName' changed but not restarted because it is not currently running. Current status: $($service.Status)" 'INFO'
            return
        }

        Write-Log "RestartChangedRunningServices=true. Restarting service '$ServiceName'." 'WARN'

        Restart-Service -Name $ServiceName -Force -ErrorAction Stop

        $service.WaitForStatus('Running', [TimeSpan]::FromSeconds(60))
        Write-Log "Service '$ServiceName' restarted successfully." 'SUCCESS'
    }
    catch {
        Write-Log "Failed to restart service '$ServiceName': $($_.Exception.Message)" 'WARN'
    }
}

try {
    Write-Log '========== Datto RMM Unquoted Service Path remediation started =========='
    Write-Log "Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log "Computer: $env:COMPUTERNAME"
    Write-Log "ReportOnly: $ReportOnly"
    Write-Log "RestartChangedRunningServices: $RestartChangedRunningServices"
    Write-Log "BackupBeforeChange: $BackupBeforeChange"
    Write-Log "IncludeDriverServices: $IncludeDriverServices"
    Write-Log "ForceFixAmbiguousPaths: $ForceFixAmbiguousPaths"
    Write-Log "FailIfVulnerableFound: $FailIfVulnerableFound"
    Write-Log "ServiceNameIncludeRegex: $ServiceNameIncludeRegex"
    Write-Log "ServiceNameExcludeRegex: $ServiceNameExcludeRegex"
    Write-Log "WorkingDirectory: $WorkingDirectory"

    if (-not (Test-IsAdministrator)) {
        throw 'This component must run elevated as Administrator or SYSTEM.'
    }

    $registryRootPath = 'SYSTEM\CurrentControlSet\Services'
    $servicesRoot = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($registryRootPath, $false)

    if (-not $servicesRoot) {
        throw "Unable to open HKLM:\$registryRootPath"
    }

    $totalServicesChecked = 0
    $vulnerableFound = 0
    $fixedCount = 0
    $failedFixCount = 0
    $skippedCount = 0
    $remainingVulnerable = 0

    foreach ($serviceName in $servicesRoot.GetSubKeyNames()) {
        if (-not [string]::IsNullOrWhiteSpace($ServiceNameIncludeRegex)) {
            if ($serviceName -notmatch $ServiceNameIncludeRegex) {
                continue
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($ServiceNameExcludeRegex)) {
            if ($serviceName -match $ServiceNameExcludeRegex) {
                Write-Log "Skipping service '$serviceName' because it matches ServiceNameExcludeRegex." 'INFO'
                continue
            }
        }

        $serviceKeyPath = "$registryRootPath\$serviceName"

        $readKey = $null
        $writeKey = $null

        try {
            $readKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($serviceKeyPath, $false)

            if (-not $readKey) {
                continue
            }

            $imagePath = $readKey.GetValue(
                'ImagePath',
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
            )

            if ([string]::IsNullOrWhiteSpace($imagePath)) {
                continue
            }

            $typeValue = $readKey.GetValue('Type', 0)
            $displayName = $readKey.GetValue('DisplayName', $serviceName)

            $isWin32Service = (($typeValue -band 0x10) -ne 0) -or (($typeValue -band 0x20) -ne 0)

            if (-not $IncludeDriverServices -and -not $isWin32Service) {
                continue
            }

            $totalServicesChecked++

            $analysis = Get-ServiceImagePathAnalysis -ImagePath ([string]$imagePath)

            if (-not $analysis.IsVulnerable) {
                continue
            }

            $vulnerableFound++

            Write-Log "VULNERABLE: Service='$serviceName' DisplayName='$displayName'" 'WARN'
            Write-Log "  Current ImagePath : $imagePath" 'WARN'
            Write-Log "  Proposed ImagePath: $($analysis.ProposedImagePath)" 'WARN'

            if ($ReportOnly) {
                $remainingVulnerable++
                continue
            }

            if (-not $analysis.CanSafelyFix) {
                $skippedCount++
                $remainingVulnerable++
                Write-Log "Skipping service '$serviceName': $($analysis.Reason)" 'WARN'
                continue
            }

            if ($BackupBeforeChange) {
                [void](Export-ServiceRegistryBackup -ServiceName $serviceName)
            }

            $writeKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($serviceKeyPath, $true)

            if (-not $writeKey) {
                $failedFixCount++
                $remainingVulnerable++
                Write-Log "Unable to open service '$serviceName' registry key for writing." 'ERROR'
                continue
            }

            try {
                $valueKind = $writeKey.GetValueKind('ImagePath')
            }
            catch {
                $valueKind = [Microsoft.Win32.RegistryValueKind]::String
            }

            $writeKey.SetValue('ImagePath', $analysis.ProposedImagePath, $valueKind)

            $verifyValue = $writeKey.GetValue(
                'ImagePath',
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
            )

            if ([string]$verifyValue -eq [string]$analysis.ProposedImagePath) {
                $fixedCount++
                Write-Log "FIXED: Service '$serviceName' ImagePath updated successfully." 'SUCCESS'
                Restart-ServiceIfRequested -ServiceName $serviceName
            }
            else {
                $failedFixCount++
                $remainingVulnerable++
                Write-Log "FAILED: Service '$serviceName' ImagePath verification failed after write." 'ERROR'
                Write-Log "  Registry now contains: $verifyValue" 'ERROR'
            }
        }
        catch {
            $failedFixCount++
            Write-Log "Error processing service '$serviceName': $($_.Exception.Message)" 'ERROR'
        }
        finally {
            if ($readKey) {
                $readKey.Close()
            }

            if ($writeKey) {
                $writeKey.Close()
            }
        }
    }

    $servicesRoot.Close()

    Write-Log '========== Summary =========='
    Write-Log "Services checked: $totalServicesChecked"
    Write-Log "Vulnerable services found: $vulnerableFound"
    Write-Log "Fixed: $fixedCount"
    Write-Log "Skipped: $skippedCount"
    Write-Log "Failed fixes: $failedFixCount"
    Write-Log "Remaining vulnerable/report-only findings: $remainingVulnerable"
    Write-Log "Log path: $LogPath"
    Write-Log '========== Datto RMM Unquoted Service Path remediation finished =========='

    if ($ReportOnly -and $vulnerableFound -gt 0 -and $FailIfVulnerableFound) {
        exit 1
    }

    if ($remainingVulnerable -gt 0 -and $FailIfVulnerableFound) {
        exit 1
    }

    if ($failedFixCount -gt 0 -and $FailIfVulnerableFound) {
        exit 1
    }

    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    exit 2
}