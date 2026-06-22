#requires -version 5.1
<#
Datto RMM Component: PJ-LEAP-Adobe-Repair-Simplified

Purpose:
1. Close Adobe and LEAP processes.
2. Uninstall Adobe Acrobat/Reader.
3. Run Adobe Acrobat Cleaner as best-effort.
4. Reinstall Adobe Reader from packaged installer.
5. Repair existing LEAP installation using MSI repair.
6. Do not install LEAP if LEAP is not already installed.

Required packaged files:
- AcroRdrDCx64*.exe
- AdobeAcroCleaner*.exe

Optional variables:
- ReportOnly=false
- LogPath=C:\ProgramData\DattoRMM\Logs\LEAP-Adobe-Repair.log
- AdobeInstallerPath=
- AcrobatCleanerPath=
- RebootAfterCompletion=false
#>

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

$script:ComponentName = 'PJ-LEAP-Adobe-Repair-Simplified'
$script:RebootRequired = $false

function Get-EnvValue {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,

        [AllowNull()]
        [string]$Default = $null
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    return $value.Trim()
}

function ConvertTo-Bool {
    param(
        [AllowNull()]
        [string]$Value,

        [bool]$Default = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    switch -Regex ($Value.Trim().ToLowerInvariant()) {
        '^(1|true|yes|y|on)$'  { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default                { return $Default }
    }
}

function ConvertTo-Int {
    param(
        [AllowNull()]
        [string]$Value,

        [int]$Default
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    $parsed = 0

    if ([int]::TryParse($Value.Trim(), [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

$script:ReportOnly = ConvertTo-Bool -Value (Get-EnvValue -Name 'ReportOnly' -Default 'false') -Default $false
$script:RebootAfterCompletion = ConvertTo-Bool -Value (Get-EnvValue -Name 'RebootAfterCompletion' -Default 'false') -Default $false

$script:LogPath = Get-EnvValue -Name 'LogPath' -Default 'C:\ProgramData\DattoRMM\Logs\LEAP-Adobe-Repair.log'
$script:WorkingDirectory = Get-EnvValue -Name 'WorkingDirectory' -Default 'C:\ProgramData\DattoRMM\Packages\LEAPAdobeRepair'

$script:AdobeInstallerPath = Get-EnvValue -Name 'AdobeInstallerPath' -Default ''
$script:AcrobatCleanerPath = Get-EnvValue -Name 'AcrobatCleanerPath' -Default ''

$script:UninstallTimeoutMinutes = ConvertTo-Int -Value (Get-EnvValue -Name 'UninstallTimeoutMinutes' -Default '20') -Default 20
$script:AdobeInstallTimeoutMinutes = ConvertTo-Int -Value (Get-EnvValue -Name 'AdobeInstallTimeoutMinutes' -Default '30') -Default 30
$script:CleanerTimeoutMinutes = ConvertTo-Int -Value (Get-EnvValue -Name 'CleanerTimeoutMinutes' -Default '10') -Default 10
$script:LeapRepairTimeoutMinutes = ConvertTo-Int -Value (Get-EnvValue -Name 'LeapRepairTimeoutMinutes' -Default '30') -Default 30

$script:AdobeInstallerArguments = '/sAll /rs /rps /msi EULA_ACCEPT=YES'
$script:LeapRepairFlags = '/faums'

$script:AdobeProductNameRegex = '(?i)\b(Adobe\s+)?Acrobat(\s+\(64-bit\)|\s+(Reader|DC|Pro|Professional|Standard))?\b|\bAdobe\s+Reader\b'
$script:LeapProductNameRegex = '(?i)\bLEAP\b'
$script:ProcessNameRegex = '(?i)^(Acrobat|AcroRd32|AcroCEF|RdrCEF|AdobeCollabSync|AdobeARM|AdobeNotificationClient|AdobeIPCBroker|LEAP|LEAPDesktop|LEAPCloud|LEAPOffice|LEAPUpdater|LEAPService|LEAPUpdate|leapsystray|TrayIconUtil|LeaPDF|leap-calc|BASupSrvc|BASupSrvcUpdater|BaSupSrvcCnfg)$'

function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level,

        [string]$Message
    )

    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    [Console]::Out.WriteLine($line)

    try {
        $logDir = Split-Path -Parent $script:LogPath

        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    } catch {
        [Console]::Out.WriteLine("[WARN] Failed to write to log file: $($_.Exception.Message)")
    }
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-SafeFileNamePart {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 'Unknown'
    }

    return ($Text -replace '[^\w\.-]+', '_')
}

function Get-SearchRoots {
    $roots = New-Object System.Collections.Generic.List[string]

    $candidates = @(
        $env:CustomComponentDirectory,
        $env:ComponentDirectory,
        $env:CS_PROFILE_PATH,
        $PSScriptRoot,
        (Get-Location).Path,
        $script:WorkingDirectory
    )

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $expanded = [Environment]::ExpandEnvironmentVariables($candidate)

            if ((Test-Path -LiteralPath $expanded) -and -not $roots.Contains($expanded)) {
                $roots.Add($expanded)
            }
        }
    }

    return $roots.ToArray()
}

function Resolve-RequiredFile {
    param(
        [AllowNull()]
        [string]$ExplicitPath,

        [Parameter(Mandatory=$true)]
        [string]$Purpose,

        [Parameter(Mandatory=$true)]
        [string[]]$Patterns
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $expanded = [Environment]::ExpandEnvironmentVariables($ExplicitPath)

        if (Test-Path -LiteralPath $expanded) {
            $resolved = (Resolve-Path -LiteralPath $expanded).Path
            Write-Log INFO "$Purpose resolved from explicit path: $resolved"
            return $resolved
        }

        throw "$Purpose path was supplied but does not exist: $ExplicitPath"
    }

    $matches = @()

    foreach ($root in Get-SearchRoots) {
        foreach ($pattern in $Patterns) {
            $found = Get-ChildItem -LiteralPath $root -Filter $pattern -File -Recurse -ErrorAction SilentlyContinue

            if ($found) {
                $matches += $found
            }
        }
    }

    $matches = $matches |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -Unique

    if ($matches.Count -eq 0) {
        $searched = (Get-SearchRoots) -join '; '
        throw "$Purpose was not found. Patterns: $($Patterns -join ', '). Searched: $searched"
    }

    $selected = $matches[0].FullName
    Write-Log INFO "$Purpose auto-detected: $selected"

    return $selected
}

function Resolve-OptionalFile {
    param(
        [AllowNull()]
        [string]$ExplicitPath,

        [Parameter(Mandatory=$true)]
        [string]$Purpose,

        [Parameter(Mandatory=$true)]
        [string[]]$Patterns
    )

    try {
        return Resolve-RequiredFile -ExplicitPath $ExplicitPath -Purpose $Purpose -Patterns $Patterns
    } catch {
        Write-Log WARN "$Purpose was not found: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-Process {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [AllowNull()]
        [string]$Arguments = '',

        [int[]]$SuccessCodes = @(0, 3010, 1605, 1614),

        [int]$TimeoutMinutes = 30,

        [switch]$IgnoreFailure
    )

    if ($script:ReportOnly) {
        Write-Log INFO "ReportOnly: would run with timeout ${TimeoutMinutes} minute(s): `"$FilePath`" $Arguments"
        return 0
    }

    Write-Log INFO "Running with timeout ${TimeoutMinutes} minute(s): `"$FilePath`" $Arguments"

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru
    } catch {
        if ($IgnoreFailure) {
            Write-Log WARN "Failed to start process. Continuing. Command: `"$FilePath`" $Arguments. Error: $($_.Exception.Message)"
            return -999999
        }

        throw
    }

    $timeoutMs = [Math]::Max(1, $TimeoutMinutes) * 60 * 1000
    $completed = $process.WaitForExit($timeoutMs)

    if (-not $completed) {
        Write-Log ERROR "Process timed out after $TimeoutMinutes minute(s). Killing process tree. PID=$($process.Id). Command: `"$FilePath`" $Arguments"

        try {
            Start-Process -FilePath 'taskkill.exe' -ArgumentList "/PID $($process.Id) /T /F" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null
        } catch {
            Write-Log WARN "Failed to taskkill PID $($process.Id): $($_.Exception.Message)"
        }

        if ($IgnoreFailure) {
            return -888888
        }

        throw "Process timed out after $TimeoutMinutes minute(s): `"$FilePath`" $Arguments"
    }

    try {
        $exitCode = [int]$process.ExitCode
    } catch {
        $exitCode = -777777
    }

    if ($SuccessCodes -contains $exitCode) {
        if ($exitCode -eq 3010) {
            $script:RebootRequired = $true
            Write-Log WARN 'Process completed with exit code 3010: reboot required.'
        } else {
            Write-Log INFO "Process completed with exit code $exitCode."
        }

        return $exitCode
    }

    $message = "Process failed with exit code $exitCode. Command: `"$FilePath`" $Arguments"

    if ($IgnoreFailure) {
        Write-Log WARN $message
        return $exitCode
    }

    throw $message
}

function Get-InstalledPrograms {
    $registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $guidRegex = '\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}'

    foreach ($path in $registryPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.DisplayName)
            } |
            ForEach-Object {
                $productCode = $null

                if ($_.PSChildName -match "^$guidRegex$") {
                    $productCode = $_.PSChildName
                } elseif ($_.UninstallString -match $guidRegex) {
                    $productCode = $Matches[0]
                } elseif ($_.QuietUninstallString -match $guidRegex) {
                    $productCode = $Matches[0]
                }

                [pscustomobject]@{
                    DisplayName          = $_.DisplayName
                    DisplayVersion       = $_.DisplayVersion
                    Publisher            = $_.Publisher
                    InstallLocation      = $_.InstallLocation
                    UninstallString      = $_.UninstallString
                    QuietUninstallString = $_.QuietUninstallString
                    WindowsInstaller     = $_.WindowsInstaller
                    ProductCode          = $productCode
                    RegistryPath         = $_.PSPath
                }
            }
    }
}

function Get-AcrobatProducts {
    $excludeRegex = '(?i)(Update Service|Genuine|Creative Cloud|AIR|Flash|Shockwave|Digital Editions|Refresh Manager|Synchronizer)'

    Get-InstalledPrograms |
        Where-Object {
            $_.DisplayName -match $script:AdobeProductNameRegex -and
            $_.DisplayName -notmatch $excludeRegex
        } |
        Sort-Object DisplayName, DisplayVersion, ProductCode -Unique
}

function Get-LeapProducts {
    Get-InstalledPrograms |
        Where-Object {
            $_.DisplayName -match $script:LeapProductNameRegex
        } |
        Sort-Object DisplayName, DisplayVersion, ProductCode -Unique
}

function Write-ProgramList {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,

        [Parameter(Mandatory=$true)]
        $Programs
    )

    Write-Log INFO $Title

    $array = @($Programs)

    if ($array.Count -eq 0) {
        Write-Log INFO '  None found.'
        return
    }

    foreach ($program in $array) {
        Write-Log INFO ("  {0} | Version={1} | Publisher={2} | ProductCode={3}" -f `
            $program.DisplayName,
            $program.DisplayVersion,
            $program.Publisher,
            $program.ProductCode)
    }
}

function Stop-TargetProcesses {
    $processes = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProcessName -match $script:ProcessNameRegex
        }

    if (-not $processes) {
        Write-Log INFO 'No matching Adobe/LEAP processes currently running.'
        return
    }

    foreach ($process in $processes) {
        try {
            Write-Log WARN "Stopping process: $($process.ProcessName) PID=$($process.Id)"

            if (-not $script:ReportOnly) {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
            }
        } catch {
            Write-Log WARN "Failed to stop process $($process.ProcessName) PID=$($process.Id): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 3
}

function Uninstall-Program {
    param(
        [Parameter(Mandatory=$true)]
        $Program,

        [int]$TimeoutMinutes = 20
    )

    Write-Log INFO "Preparing uninstall: $($Program.DisplayName) $($Program.DisplayVersion)"

    if (-not [string]::IsNullOrWhiteSpace($Program.ProductCode)) {
        $null = Invoke-Process `
            -FilePath 'msiexec.exe' `
            -Arguments "/x $($Program.ProductCode) /qn /norestart REBOOT=ReallySuppress" `
            -SuccessCodes @(0, 3010, 1605, 1614) `
            -TimeoutMinutes $TimeoutMinutes

        return
    }

    if (-not [string]::IsNullOrWhiteSpace($Program.QuietUninstallString)) {
        $null = Invoke-Process `
            -FilePath $env:ComSpec `
            -Arguments "/c `"$($Program.QuietUninstallString)`"" `
            -SuccessCodes @(0, 3010, 1605, 1614) `
            -TimeoutMinutes $TimeoutMinutes

        return
    }

    if (-not [string]::IsNullOrWhiteSpace($Program.UninstallString) -and $Program.UninstallString -match '\{[0-9A-Fa-f-]{36}\}') {
        $guid = $Matches[0]

        $null = Invoke-Process `
            -FilePath 'msiexec.exe' `
            -Arguments "/x $guid /qn /norestart REBOOT=ReallySuppress" `
            -SuccessCodes @(0, 3010, 1605, 1614) `
            -TimeoutMinutes $TimeoutMinutes

        return
    }

    Write-Log WARN "No MSI product code or quiet uninstall string found for: $($Program.DisplayName). Skipping."
}

function Reinstall-AdobeReader {
    $adobeBefore = @(Get-AcrobatProducts)
    Write-ProgramList -Title 'Adobe Acrobat/Reader entries before Adobe reinstall:' -Programs $adobeBefore

    foreach ($product in $adobeBefore) {
        Uninstall-Program -Program $product -TimeoutMinutes $script:UninstallTimeoutMinutes
    }

    Start-Sleep -Seconds 5

    $adobeAfterUninstall = @(Get-AcrobatProducts)
    Write-ProgramList -Title 'Adobe Acrobat/Reader entries after normal uninstall:' -Programs $adobeAfterUninstall

    $cleaner = Resolve-OptionalFile `
        -ExplicitPath $script:AcrobatCleanerPath `
        -Purpose 'Adobe Acrobat Cleaner' `
        -Patterns @(
            'AdobeAcroCleaner*.exe',
            'AcroCleaner*.exe',
            '*Cleaner*Acro*.exe'
        )

    if (-not [string]::IsNullOrWhiteSpace($cleaner)) {
        foreach ($productId in @(1, 0)) {
            $productName = if ($productId -eq 1) { 'Reader' } else { 'Acrobat' }

            Write-Log INFO "Running Adobe Acrobat Cleaner for $productName. ProductId=$productId"

            $exitCode = Invoke-Process `
                -FilePath $cleaner `
                -Arguments "/silent /product=$productId /installpath=Default /cleanlevel=0 /scanforothers=1" `
                -SuccessCodes @(0, 3010) `
                -TimeoutMinutes $script:CleanerTimeoutMinutes `
                -IgnoreFailure

            if (@(0, 3010) -notcontains [int]$exitCode) {
                Write-Log WARN "Adobe Acrobat Cleaner for $productName returned exit code $exitCode. Continuing."
            }
        }
    } else {
        Write-Log WARN 'Adobe Acrobat Cleaner was not found. Continuing with Adobe reinstall.'
    }

    Start-Sleep -Seconds 5

    $installer = Resolve-RequiredFile `
        -ExplicitPath $script:AdobeInstallerPath `
        -Purpose 'Adobe Reader installer' `
        -Patterns @(
            'AcroRdrDCx64*.exe',
            'AcroRdrDC*.exe',
            'AdobeReader*.exe',
            'AdobeAcrobatReader*.exe',
            'AcroRead.msi'
        )

    if ($installer -match '\.msi$') {
        $null = Invoke-Process `
            -FilePath 'msiexec.exe' `
            -Arguments "/i `"$installer`" /qn /norestart EULA_ACCEPT=YES" `
            -SuccessCodes @(0, 3010) `
            -TimeoutMinutes $script:AdobeInstallTimeoutMinutes
    } elseif ($installer -match '\.exe$') {
        $null = Invoke-Process `
            -FilePath $installer `
            -Arguments $script:AdobeInstallerArguments `
            -SuccessCodes @(0, 3010) `
            -TimeoutMinutes $script:AdobeInstallTimeoutMinutes
    } else {
        throw "Unsupported Adobe installer type: $installer"
    }

    Start-Sleep -Seconds 5

    $adobeFinal = @(Get-AcrobatProducts)
    Write-ProgramList -Title 'Adobe Acrobat/Reader entries after Adobe reinstall:' -Programs $adobeFinal

    if ($adobeFinal.Count -eq 0) {
        throw 'Adobe reinstall completed, but no Adobe Acrobat/Reader entry was detected afterwards.'
    }

    Write-Log SUCCESS 'Adobe reinstall completed successfully.'
}

function Repair-Leap {
    $leapProducts = @(Get-LeapProducts)
    Write-ProgramList -Title 'LEAP entries before LEAP repair:' -Programs $leapProducts

    if ($leapProducts.Count -eq 0) {
        throw 'No LEAP installation was detected. This component is repair-only for LEAP and will not install LEAP when absent.'
    }

    Stop-TargetProcesses

    $successCount = 0
    $failureCount = 0
    $index = 0

    foreach ($product in $leapProducts) {
        $index++

        if ([string]::IsNullOrWhiteSpace($product.ProductCode)) {
            Write-Log WARN "Cannot repair $($product.DisplayName) because no MSI ProductCode was detected."
            $failureCount++
            continue
        }

        $safeCode = Get-SafeFileNamePart -Text $product.ProductCode
        $repairLog = Join-Path 'C:\ProgramData\DattoRMM\Logs' ("LEAP-Repair-{0}-{1}.log" -f $index, $safeCode)

        Ensure-Directory -Path (Split-Path -Parent $repairLog)

        Write-Log INFO "Repairing LEAP product: $($product.DisplayName) $($product.DisplayVersion)"
        Write-Log INFO "LEAP product code: $($product.ProductCode)"
        Write-Log INFO "LEAP repair flags: $script:LeapRepairFlags"
        Write-Log INFO "LEAP repair log path: $repairLog"

        $exitCode = Invoke-Process `
            -FilePath 'msiexec.exe' `
            -Arguments "$script:LeapRepairFlags $($product.ProductCode) /qn /norestart REBOOT=ReallySuppress /L*v `"$repairLog`"" `
            -SuccessCodes @(0, 3010) `
            -TimeoutMinutes $script:LeapRepairTimeoutMinutes `
            -IgnoreFailure

        if (@(0, 3010) -contains [int]$exitCode) {
            $successCount++
            Write-Log SUCCESS "LEAP repair succeeded for product code $($product.ProductCode) with exit code $exitCode."
        } else {
            $failureCount++
            Write-Log WARN "LEAP repair failed for product code $($product.ProductCode) with exit code $exitCode. Continuing to next LEAP entry."
        }

        Start-Sleep -Seconds 5
    }

    $leapAfter = @(Get-LeapProducts)
    Write-ProgramList -Title 'LEAP entries after LEAP repair:' -Programs $leapAfter

    if ($successCount -eq 0) {
        throw "All LEAP repair attempts failed. Failed count: $failureCount. Check C:\ProgramData\DattoRMM\Logs\LEAP-Repair-*.log"
    }

    if ($leapAfter.Count -eq 0) {
        throw 'LEAP repair completed, but no LEAP entry was detected afterwards.'
    }

    Write-Log SUCCESS "LEAP repair completed. Successful repairs: $successCount. Failed repairs: $failureCount."
}

function Main {
    Ensure-Directory -Path $script:WorkingDirectory
    Ensure-Directory -Path (Split-Path -Parent $script:LogPath)
    Ensure-Directory -Path 'C:\ProgramData\DattoRMM\Logs'

    Write-Log INFO "========== $script:ComponentName started =========="
    Write-Log INFO "ReportOnly: $script:ReportOnly"
    Write-Log INFO "LogPath: $script:LogPath"
    Write-Log INFO "WorkingDirectory: $script:WorkingDirectory"
    Write-Log INFO "AdobeInstallerPath: $script:AdobeInstallerPath"
    Write-Log INFO "AcrobatCleanerPath: $script:AcrobatCleanerPath"
    Write-Log INFO "AdobeInstallerArguments: $script:AdobeInstallerArguments"
    Write-Log INFO "LeapRepairFlags: $script:LeapRepairFlags"
    Write-Log INFO "UninstallTimeoutMinutes: $script:UninstallTimeoutMinutes"
    Write-Log INFO "AdobeInstallTimeoutMinutes: $script:AdobeInstallTimeoutMinutes"
    Write-Log INFO "CleanerTimeoutMinutes: $script:CleanerTimeoutMinutes"
    Write-Log INFO "LeapRepairTimeoutMinutes: $script:LeapRepairTimeoutMinutes"

    $adobeBefore = @(Get-AcrobatProducts)
    $leapBefore = @(Get-LeapProducts)

    Write-ProgramList -Title 'Adobe Acrobat/Reader entries at start:' -Programs $adobeBefore
    Write-ProgramList -Title 'LEAP entries at start:' -Programs $leapBefore

    if ($script:ReportOnly) {
        Write-Log SUCCESS 'ReportOnly completed. No changes made.'
        return
    }

    Write-Log INFO 'Step 1: Closing Adobe and LEAP processes.'
    Stop-TargetProcesses

    Write-Log INFO 'Step 2: Reinstalling and cleaning Adobe Acrobat/Reader.'
    Reinstall-AdobeReader

    Write-Log INFO 'Step 3: Repairing existing LEAP installation.'
    Repair-Leap

    $adobeFinal = @(Get-AcrobatProducts)
    $leapFinal = @(Get-LeapProducts)

    Write-ProgramList -Title 'Adobe Acrobat/Reader entries at end:' -Programs $adobeFinal
    Write-ProgramList -Title 'LEAP entries at end:' -Programs $leapFinal

    if ($script:RebootRequired) {
        Write-Log WARN 'A reboot is required or recommended after this remediation.'
    }

    if ($script:RebootAfterCompletion) {
        Write-Log WARN 'RebootAfterCompletion=true. Restarting computer now.'
        Restart-Computer -Force
    }

    Write-Log SUCCESS "========== $script:ComponentName completed successfully =========="
}

try {
    Main
    exit 0
} catch {
    Write-Log ERROR "Component failed: $($_.Exception.Message)"
    exit 1
}