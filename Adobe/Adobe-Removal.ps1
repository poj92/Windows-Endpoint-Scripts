<#

Author: Peter Opeyemi James
Company: Nexus Open Systems


Datto RMM Component: Force Remove Selected Adobe Products
Platform: Windows
Run as: LocalSystem / Administrator
PowerShell: 5.1+

Datto RMM Variables:

  ReportOnly
    true/false
    Default: true
    When true, no uninstall actions are performed.
    Lists detected Adobe products and the planned action.

  TargetProducts
    Semicolon-separated product name fragments to remove.
    Example:
      Adobe After Effects 2024; Adobe Illustrator 2024; Adobe Photoshop 2024

  ExcludeProducts
    Semicolon-separated product name fragments to preserve.
    Example:
      Adobe Acrobat; Adobe Acrobat Reader; Adobe Acrobat Standard; Adobe Acrobat Pro

    Behaviour:
      - TargetProducts populated:
          Remove matching TargetProducts unless also matched by ExcludeProducts.
      - TargetProducts empty and ExcludeProducts populated:
          Remove all detected Adobe products except ExcludeProducts.
      - TargetProducts empty and ExcludeProducts empty and ReportOnly=false:
          No broad registry removal is attempted unless RemoveCreativeCloudAll is safely allowed.

  KillAdobeProcesses
    true/false
    Default: true
    Ignored when ReportOnly=true.

  ForceExeFallback
    true/false
    Default: false
    Allows fallback use of non-MSI uninstallers.
    Keep false unless you specifically want to test registry EXE uninstall strings.

  UseAdobeUninstaller
    true/false
    Default: true
    Uses AdobeUninstaller.exe if found.

  AdobeUninstallerPath
    Optional explicit path to AdobeUninstaller.exe.
    Example:
      C:\ProgramData\DattoRMM\Packages\AdobeUninstaller.exe

  AdobeSapCodes
    Optional Adobe SAP codes.
    Semicolon or comma separated.
    Example:
      PHSP#25.9.1; ILST#28.6; IDSN#19.5

  AutoAdobeSapCodes
    true/false
    Default: true
    If UseAdobeUninstaller=true and AdobeSapCodes is empty, the script derives SAP codes
    from detected Creative Cloud registry product codes like PHSP_25_9_1.

  RemoveCreativeCloudAll
    true/false
    Default: false
    Requests AdobeUninstaller.exe --all.

    SAFETY CHANGE:
      --all will NOT run if TargetProducts is populated unless AllowRemoveAllWhenTargetsSpecified=true.
      --all will NOT run if ExcludeProducts is populated unless AllowRemoveAllWithExclusions=true.

    This is to avoid removing untargeted products such as Adobe Acrobat/Reader.

  AllowRemoveAllWhenTargetsSpecified
    true/false
    Default: false
    Explicitly allows --all even when TargetProducts is populated.

  AllowRemoveAllWithExclusions
    true/false
    Default: false
    Explicitly allows --all even when ExcludeProducts is populated.
    Not recommended, because --all cannot honour exclusions.

  UseHDBoxSetupFallback
    true/false
    Default: true
    After AdobeUninstaller targeted removal, tries Adobe HDBox Setup.exe fallback for remaining
    Creative Cloud-managed apps using:
      Setup.exe --uninstall=1 --sapCode=<SAP> --baseVersion=<major>.0 --platform=win64

  DeleteUserPreferences
    true/false
    Default: false
    Passed to HDBox Setup.exe fallback as --deleteUserPreferences=true/false.

  RebootIfRequired
    true/false
    Default: false

  FailIfRemaining
    true/false
    Default: true
    If true, exits 1 when targeted products remain after all removal attempts.

  VerificationDelaySeconds
    Number
    Default: 30
#>

$ErrorActionPreference = "Continue"

# -----------------------------
# Helpers
# -----------------------------

function Get-EnvFirst {
    param(
        [string[]]$Names,
        [string]$Default = ""
    )

    foreach ($name in $Names) {
        foreach ($scope in @("Process", "Machine", "User")) {
            try {
                $value = [Environment]::GetEnvironmentVariable($name, $scope)
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    return $value
                }
            } catch {
                # Ignore scope lookup errors
            }
        }
    }

    return $Default
}

function Get-Bool {
    param(
        [string]$Value,
        [bool]$Default = $false
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    switch -Regex ($Value.Trim().ToLower()) {
        '^(true|1|yes|y)$'  { return $true }
        '^(false|0|no|n)$'  { return $false }
        default             { return $Default }
    }
}

function Get-Int {
    param(
        [string]$Value,
        [int]$Default = 30
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    $parsed = 0
    if ([int]::TryParse($Value, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function Split-SemicolonList {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split ';' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
    )
}

function Split-SapCodeList {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @(
        $Value -split '[;,]' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
    )
}

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $Message"

    # Do not use Write-Output here.
    # Write-Output pollutes function return values and can corrupt exit-code handling.
    [Console]::Out.WriteLine($line)

    try {
        $logDir = "C:\ProgramData\DattoRMM\Logs"
        if (!(Test-Path -LiteralPath $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        $logPath = Join-Path -Path $logDir -ChildPath "ForceAdobeRemoval.log"
        Add-Content -Path $logPath -Value $line -Encoding UTF8
    } catch {
        [Console]::Out.WriteLine("[$timestamp] Unable to write log file: $($_.Exception.Message)")
    }
}

function Test-StringContainsAny {
    param(
        [string]$Value,
        [string[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    if (!$Patterns -or $Patterns.Count -eq 0) {
        return $false
    }

    foreach ($pattern in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) {
            continue
        }

        if ($Value.IndexOf($pattern, [System.StringComparison]::InvariantCultureIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

function Get-ScriptRootSafe {
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }

    if ($MyInvocation.MyCommand.Path) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }

    return (Get-Location).Path
}

function New-CommandSpec {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string]$Arguments = "",

        [string]$DisplayCommand = ""
    )

    if ([string]::IsNullOrWhiteSpace($DisplayCommand)) {
        if ([string]::IsNullOrWhiteSpace($Arguments)) {
            $DisplayCommand = $FilePath
        } else {
            $DisplayCommand = "$FilePath $Arguments"
        }
    }

    return [PSCustomObject]@{
        FilePath = $FilePath
        Arguments = $Arguments
        DisplayCommand = $DisplayCommand
    }
}

function Convert-CommandLineToSpec {
    param([string]$CommandLine)

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $null
    }

    $cmd = [Environment]::ExpandEnvironmentVariables($CommandLine.Trim())

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        return $null
    }

    if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') {
        $file = $Matches[1]
        $args = $Matches[2]
        return New-CommandSpec -FilePath $file -Arguments $args -DisplayCommand $cmd
    }

    $exeMatches = [regex]::Matches($cmd, '(?i)\.exe')
    $selectedPath = $null
    $selectedArgs = $null

    foreach ($match in $exeMatches) {
        $candidate = $cmd.Substring(0, $match.Index + 4).Trim()
        if (Test-Path -LiteralPath $candidate) {
            $selectedPath = $candidate
            $selectedArgs = $cmd.Substring($match.Index + 4).Trim()
        }
    }

    if ($selectedPath) {
        return New-CommandSpec -FilePath $selectedPath -Arguments $selectedArgs -DisplayCommand $cmd
    }

    if ($cmd -match '^\s*(\S+)\s*(.*)$') {
        $file = $Matches[1]
        $args = $Matches[2]
        return New-CommandSpec -FilePath $file -Arguments $args -DisplayCommand $cmd
    }

    return $null
}

function Invoke-CommandSpec {
    param(
        [Parameter(Mandatory = $true)]
        $Spec
    )

    Write-Log "Running: $($Spec.DisplayCommand)"

    if ($ReportOnly) {
        Write-Log "ReportOnly=true. Command not executed."
        return [int]0
    }

    try {
        if ([string]::IsNullOrWhiteSpace($Spec.Arguments)) {
            $process = Start-Process `
                -FilePath $Spec.FilePath `
                -Wait `
                -PassThru `
                -WindowStyle Hidden
        } else {
            $process = Start-Process `
                -FilePath $Spec.FilePath `
                -ArgumentList $Spec.Arguments `
                -Wait `
                -PassThru `
                -WindowStyle Hidden
        }

        $code = [int]$process.ExitCode
        Write-Log "Exit code: $code"
        return $code
    } catch {
        Write-Log "FAILED to run command: $($_.Exception.Message)"
        return [int]9999
    }
}

function Invoke-CommandLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandLine
    )

    $spec = Convert-CommandLineToSpec -CommandLine $CommandLine

    if (!$spec) {
        Write-Log "Unable to parse command line: $CommandLine"
        return [int]9997
    }

    return [int](Invoke-CommandSpec -Spec $spec)
}

function Add-ExitCode {
    param(
        [System.Collections.Generic.List[int]]$List,
        [int]$Code
    )

    [void]$List.Add([int]$Code)
}

# -----------------------------
# Installed app discovery
# -----------------------------

function Get-InstalledApps {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $apps = foreach ($path in $paths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -and
                (
                    $_.UninstallString -or
                    $_.QuietUninstallString -or
                    $_.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$'
                )
            } |
            Select-Object `
                DisplayName,
                DisplayVersion,
                Publisher,
                InstallDate,
                UninstallString,
                QuietUninstallString,
                WindowsInstaller,
                PSChildName,
                @{Name = "RegistryPath"; Expression = { $_.PSPath }}
    }

    $apps |
        Sort-Object DisplayName, DisplayVersion, PSChildName -Unique
}

function Get-AdobeInstalledApps {
    Get-InstalledApps |
        Where-Object {
            ($_.DisplayName -match '(?i)\bAdobe\b') -or
            ($_.Publisher -match '(?i)Adobe')
        } |
        Sort-Object DisplayName, DisplayVersion
}

function Get-RemovalDecision {
    param($App)

    $isExcluded = Test-StringContainsAny -Value $App.DisplayName -Patterns $ExcludeProducts
    $isTargeted = Test-StringContainsAny -Value $App.DisplayName -Patterns $TargetProducts

    if ($isExcluded) {
        return [PSCustomObject]@{
            Action = "PRESERVE"
            Reason = "Matched ExcludeProducts"
        }
    }

    if ($TargetProducts.Count -gt 0) {
        if ($isTargeted) {
            return [PSCustomObject]@{
                Action = "REMOVE"
                Reason = "Matched TargetProducts"
            }
        }

        return [PSCustomObject]@{
            Action = "IGNORE"
            Reason = "Did not match TargetProducts"
        }
    }

    if ($TargetProducts.Count -eq 0 -and $ExcludeProducts.Count -gt 0) {
        return [PSCustomObject]@{
            Action = "REMOVE"
            Reason = "TargetProducts empty; removing all Adobe products except ExcludeProducts"
        }
    }

    return [PSCustomObject]@{
        Action = "IGNORE"
        Reason = "No TargetProducts or ExcludeProducts configured"
    }
}

function Get-MsiGuidFromString {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    if ($Value -match '\{[0-9A-Fa-f-]{36}\}') {
        return $Matches[0]
    }

    return $null
}

function Test-AdobeDesktopManagedApp {
    param($App)

    $productCode = [string]$App.PSChildName
    $uninstallString = [string]$App.UninstallString
    $quietUninstallString = [string]$App.QuietUninstallString
    $displayName = [string]$App.DisplayName

    if ($productCode -match '^[A-Z0-9]+(?:BETA)?_\d+(?:_\d+){0,3}$') {
        return $true
    }

    if ($uninstallString -match '(?i)Adobe Desktop Common\\HDBox\\Uninstaller\.exe') {
        return $true
    }

    if ($quietUninstallString -match '(?i)Adobe Desktop Common\\HDBox\\Uninstaller\.exe') {
        return $true
    }

    if ($displayName -match '(?i)^Adobe Creative Cloud$') {
        return $true
    }

    if ($uninstallString -match '(?i)Creative Cloud Uninstaller\.exe') {
        return $true
    }

    return $false
}

function Test-IsProblematicAdobeExeFallback {
    param($Spec)

    if (!$Spec) {
        return $false
    }

    $combined = "$($Spec.FilePath) $($Spec.Arguments)"

    if ($combined -match '(?i)Adobe Desktop Common\\HDBox\\Uninstaller\.exe') {
        return $true
    }

    if ($combined -match '(?i)Creative Cloud Uninstaller\.exe') {
        return $true
    }

    return $false
}

function Get-UninstallCommand {
    param($App)

    $guid = $null

    if ($App.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
        $guid = $App.PSChildName
    }

    if (!$guid -and $App.UninstallString) {
        $guid = Get-MsiGuidFromString -Value $App.UninstallString
    }

    if ($guid) {
        return New-CommandSpec `
            -FilePath "msiexec.exe" `
            -Arguments "/x $guid /qn /norestart REBOOT=ReallySuppress" `
            -DisplayCommand "msiexec.exe /x $guid /qn /norestart REBOOT=ReallySuppress"
    }

    $cmd = $null
    $usedQuietString = $false

    if ($App.QuietUninstallString) {
        $cmd = $App.QuietUninstallString.Trim()
        $usedQuietString = $true
    } elseif ($App.UninstallString) {
        $cmd = $App.UninstallString.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        return $null
    }

    if ($cmd -match '(?i)msiexec') {
        $cmd = $cmd -replace '(?i)/I', '/X'

        if ($cmd -notmatch '(?i)\s/q') {
            $cmd += " /qn"
        }

        if ($cmd -notmatch '(?i)norestart|reallysuppress') {
            $cmd += " /norestart REBOOT=ReallySuppress"
        }

        return Convert-CommandLineToSpec -CommandLine $cmd
    }

    $spec = Convert-CommandLineToSpec -CommandLine $cmd

    if (!$spec) {
        return $null
    }

    if ((Test-IsProblematicAdobeExeFallback -Spec $spec) -and -not $ForceExeFallback) {
        Write-Log "Skipping local Adobe EXE fallback for $($App.DisplayName). Use AdobeUninstaller.exe, AdobeSapCodes, HDBox Setup fallback, or set ForceExeFallback=true."
        return $null
    }

    if (!$usedQuietString -and $ForceExeFallback) {
        $combined = "$($spec.FilePath) $($spec.Arguments)"

        if ($combined -notmatch '(?i)(/quiet|/qn|/s|/silent|--silent|VERYSILENT|--mode=2)') {
            $spec.Arguments = ($spec.Arguments + " /quiet /norestart").Trim()
            $spec.DisplayCommand = "$($spec.FilePath) $($spec.Arguments)"
        }
    }

    return $spec
}

# -----------------------------
# Adobe process/service handling
# -----------------------------

function Stop-AdobeProcesses {
    if ($ReportOnly) {
        Write-Log "ReportOnly=true. Adobe processes will not be stopped."
        return
    }

    $processNames = @(
        "Acrobat",
        "AcroRd32",
        "Adobe Desktop Service",
        "AdobeIPCBroker",
        "AdobeNotificationClient",
        "AdobeCollabSync",
        "CCXProcess",
        "CoreSync",
        "Creative Cloud",
        "Creative Cloud Helper",
        "Creative Cloud UI Helper",
        "armsvc"
    )

    foreach ($name in $processNames) {
        try {
            Get-Process -Name $name -ErrorAction SilentlyContinue |
                ForEach-Object {
                    Write-Log "Stopping process: $($_.ProcessName) PID $($_.Id)"
                    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                }
        } catch {
            Write-Log "Could not stop process $name`: $($_.Exception.Message)"
        }
    }

    $serviceNames = @(
        "AdobeARMservice",
        "AGMService",
        "AGSService",
        "AdobeUpdateService"
    )

    foreach ($svc in $serviceNames) {
        try {
            $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
            if ($service -and $service.Status -eq "Running") {
                Write-Log "Stopping service: $svc"
                Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Log "Could not stop service $svc`: $($_.Exception.Message)"
        }
    }
}

# -----------------------------
# AdobeUninstaller.exe support
# -----------------------------

function Get-AdobeUninstallerTool {
    $scriptRoot = Get-ScriptRootSafe
    $currentDir = (Get-Location).Path

    $possiblePaths = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($AdobeUninstallerPath)) {
        [void]$possiblePaths.Add($AdobeUninstallerPath)
    }

    [void]$possiblePaths.Add((Join-Path -Path $scriptRoot -ChildPath "AdobeUninstaller.exe"))
    [void]$possiblePaths.Add((Join-Path -Path $currentDir -ChildPath "AdobeUninstaller.exe"))
    [void]$possiblePaths.Add("C:\ProgramData\DattoRMM\Packages\AdobeUninstaller.exe")
    [void]$possiblePaths.Add("C:\Program Files\Common Files\Adobe\Adobe Desktop Common\HDBox\AdobeUninstaller.exe")
    [void]$possiblePaths.Add("C:\Program Files (x86)\Common Files\Adobe\Adobe Desktop Common\HDBox\AdobeUninstaller.exe")

    foreach ($path in ($possiblePaths | Select-Object -Unique)) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    return $null
}

function Get-AutoSapCodeFromApp {
    param($App)

    $productCode = [string]$App.PSChildName

    # Examples:
    #   PHSP_25_9_1     -> PHSP#25.9.1
    #   ILST_28_6       -> ILST#28.6
    #   ILSTBETA_28_8   -> ILSTBETA#28.8
    #   AEFT_24_5       -> AEFT#24.5
    if ($productCode -match '^(?<sap>[A-Z0-9]+(?:BETA)?)_(?<ver>\d+(?:_\d+){0,3})$') {
        $sap = $Matches["sap"]
        $ver = $Matches["ver"] -replace '_', '.'
        return "$sap#$ver"
    }

    return $null
}

function Get-AutoSapCodesFromPlan {
    param(
        [object[]]$PlannedItems
    )

    $codes = foreach ($item in $PlannedItems) {
        $code = Get-AutoSapCodeFromApp -App $item.App
        if ($code) {
            $code
        }
    }

    return @($codes | Sort-Object -Unique)
}

function Test-CanRunAdobeUninstallerAll {
    if (-not $RemoveCreativeCloudAll) {
        return $false
    }

    if ($TargetProducts.Count -gt 0 -and -not $AllowRemoveAllWhenTargetsSpecified) {
        Write-Log "SAFETY: RemoveCreativeCloudAll=true, but TargetProducts is populated. AdobeUninstaller.exe --all will NOT run because it can remove untargeted products such as Acrobat/Reader."
        Write-Log "SAFETY: The script will use targeted AdobeSapCodes instead. Set AllowRemoveAllWhenTargetsSpecified=true only if you intentionally want broad Adobe removal."
        return $false
    }

    if ($ExcludeProducts.Count -gt 0 -and -not $AllowRemoveAllWithExclusions) {
        Write-Log "SAFETY: RemoveCreativeCloudAll=true, but ExcludeProducts is populated. AdobeUninstaller.exe --all will NOT run because it cannot honour exclusions."
        Write-Log "SAFETY: The script will use targeted AdobeSapCodes instead. Set AllowRemoveAllWithExclusions=true only if you intentionally want broad Adobe removal."
        return $false
    }

    return $true
}

function Invoke-AdobeUninstallerPass {
    param(
        [object[]]$PlannedItems,
        [bool]$AllowAllThisPass
    )

    $exitCodes = New-Object System.Collections.Generic.List[int]

    $result = [PSCustomObject]@{
        ToolFound = $false
        RanAll = $false
        RanProducts = $false
        ExitCodes = @()
    }

    if (!$UseAdobeUninstaller) {
        return $result
    }

    $tool = Get-AdobeUninstallerTool

    if (!$tool) {
        Write-Log "UseAdobeUninstaller=true, but AdobeUninstaller.exe was not found. Registry uninstall fallback will be used where safe."
        $result.ExitCodes = $exitCodes.ToArray()
        return $result
    }

    $result.ToolFound = $true
    Write-Log "Using AdobeUninstaller.exe: $tool"

    if ($ReportOnly) {
        Write-Log "ReportOnly=true. AdobeUninstaller.exe will not be executed."
        $result.ExitCodes = $exitCodes.ToArray()
        return $result
    }

    if ($AllowAllThisPass -and (Test-CanRunAdobeUninstallerAll)) {
        $spec = New-CommandSpec `
            -FilePath $tool `
            -Arguments "--all" `
            -DisplayCommand "`"$tool`" --all"

        $code = [int](Invoke-CommandSpec -Spec $spec)
        Add-ExitCode -List $exitCodes -Code $code

        $result.RanAll = $true
        $result.ExitCodes = $exitCodes.ToArray()
        return $result
    }

    $sapCodes = @(Split-SapCodeList -Value $AdobeSapCodes)

    if ($sapCodes.Count -eq 0 -and $AutoAdobeSapCodes) {
        $sapCodes = @(Get-AutoSapCodesFromPlan -PlannedItems $PlannedItems)

        if ($sapCodes.Count -gt 0) {
            Write-Log "Auto-generated AdobeSapCodes from selected products: $($sapCodes -join '; ')"
        }
    }

    if ($sapCodes.Count -gt 0) {
        $sapArgument = ($sapCodes | Sort-Object -Unique) -join ','

        $spec = New-CommandSpec `
            -FilePath $tool `
            -Arguments "--products=$sapArgument --skipNotInstalled" `
            -DisplayCommand "`"$tool`" --products=$sapArgument --skipNotInstalled"

        $code = [int](Invoke-CommandSpec -Spec $spec)
        Add-ExitCode -List $exitCodes -Code $code

        $result.RanProducts = $true
        $result.ExitCodes = $exitCodes.ToArray()
        return $result
    }

    Write-Log "UseAdobeUninstaller=true, but no AdobeSapCodes were supplied or auto-generated."

    $result.ExitCodes = $exitCodes.ToArray()
    return $result
}

# -----------------------------
# HDBox Setup.exe fallback
# -----------------------------

function Get-AdobeBaseVersionFromApp {
    param($App)

    $productCode = [string]$App.PSChildName

    if ($productCode -match '^(?<sap>[A-Z0-9]+(?:BETA)?)_(?<major>\d+)(?:_\d+){0,3}$') {
        return "$($Matches["major"]).0"
    }

    return $null
}

function Get-AdobeSapOnlyFromApp {
    param($App)

    $productCode = [string]$App.PSChildName

    if ($productCode -match '^(?<sap>[A-Z0-9]+(?:BETA)?)_\d+(?:_\d+){0,3}$') {
        return $Matches["sap"]
    }

    return $null
}

function Get-HDBoxSetupPath {
    $paths = @(
        "C:\Program Files (x86)\Common Files\Adobe\Adobe Desktop Common\HDBox\Setup.exe",
        "C:\Program Files\Common Files\Adobe\Adobe Desktop Common\HDBox\Setup.exe"
    )

    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    return $null
}

function Invoke-HDBoxSetupFallback {
    param(
        [object[]]$RemainingItems
    )

    $exitCodes = New-Object System.Collections.Generic.List[int]
    $setupPath = Get-HDBoxSetupPath

    if (!$setupPath) {
        Write-Log "HDBox Setup.exe fallback requested, but Setup.exe was not found."
        return $exitCodes.ToArray()
    }

    foreach ($item in $RemainingItems) {
        $sap = Get-AdobeSapOnlyFromApp -App $item.App
        $baseVersion = Get-AdobeBaseVersionFromApp -App $item.App

        if (!$sap -or !$baseVersion) {
            Write-Log "Cannot derive SAP/baseVersion for HDBox Setup fallback: $($item.DisplayName) ProductCode=$($item.ProductCode)"
            continue
        }

        $deletePrefsValue = "false"
        if ($DeleteUserPreferences) {
            $deletePrefsValue = "true"
        }

        $arguments = "--uninstall=1 --sapCode=$sap --baseVersion=$baseVersion --platform=win64 --deleteUserPreferences=$deletePrefsValue"

        $spec = New-CommandSpec `
            -FilePath $setupPath `
            -Arguments $arguments `
            -DisplayCommand "`"$setupPath`" $arguments"

        $code = [int](Invoke-CommandSpec -Spec $spec)
        Add-ExitCode -List $exitCodes -Code $code
    }

    return $exitCodes.ToArray()
}

# -----------------------------
# Planning helpers
# -----------------------------

function New-PlanFromCurrentAdobeApps {
    $currentAdobeApps = @(Get-AdobeInstalledApps)

    return @(
        foreach ($app in $currentAdobeApps) {
            $decision = Get-RemovalDecision -App $app

            [PSCustomObject]@{
                DisplayName = $app.DisplayName
                DisplayVersion = $app.DisplayVersion
                Publisher = $app.Publisher
                ProductCode = $app.PSChildName
                Action = $decision.Action
                Reason = $decision.Reason
                App = $app
            }
        }
    )
}

function Get-RemainingTargetedItems {
    param(
        [object[]]$CurrentPlan
    )

    return @($CurrentPlan | Where-Object { $_.Action -eq "REMOVE" })
}

# -----------------------------
# Read Datto variables
# -----------------------------

$TargetProductsRaw = Get-EnvFirst -Names @("TargetProducts", "TARGET_PRODUCTS", "target_products")
$ExcludeProductsRaw = Get-EnvFirst -Names @("ExcludeProducts", "EXCLUDE_PRODUCTS", "exclude_products")

$TargetProducts = @(Split-SemicolonList -Value $TargetProductsRaw)
$ExcludeProducts = @(Split-SemicolonList -Value $ExcludeProductsRaw)

$ReportOnly = Get-Bool (Get-EnvFirst -Names @("ReportOnly", "REPORT_ONLY", "REPORTONLY") -Default "true") $true
$KillAdobeProcesses = Get-Bool (Get-EnvFirst -Names @("KillAdobeProcesses", "KILL_ADOBE_PROCESSES") -Default "true") $true
$ForceExeFallback = Get-Bool (Get-EnvFirst -Names @("ForceExeFallback", "FORCE_EXE_FALLBACK") -Default "false") $false
$UseAdobeUninstaller = Get-Bool (Get-EnvFirst -Names @("UseAdobeUninstaller", "USE_ADOBE_UNINSTALLER") -Default "true") $true
$AdobeUninstallerPath = Get-EnvFirst -Names @("AdobeUninstallerPath", "ADOBE_UNINSTALLER_PATH")
$AdobeSapCodes = Get-EnvFirst -Names @("AdobeSapCodes", "ADOBE_SAP_CODES")
$AutoAdobeSapCodes = Get-Bool (Get-EnvFirst -Names @("AutoAdobeSapCodes", "AUTO_ADOBE_SAP_CODES") -Default "true") $true
$RemoveCreativeCloudAll = Get-Bool (Get-EnvFirst -Names @("RemoveCreativeCloudAll", "REMOVE_CREATIVE_CLOUD_ALL") -Default "false") $false
$AllowRemoveAllWhenTargetsSpecified = Get-Bool (Get-EnvFirst -Names @("AllowRemoveAllWhenTargetsSpecified", "ALLOW_REMOVE_ALL_WHEN_TARGETS_SPECIFIED") -Default "false") $false
$AllowRemoveAllWithExclusions = Get-Bool (Get-EnvFirst -Names @("AllowRemoveAllWithExclusions", "ALLOW_REMOVE_ALL_WITH_EXCLUSIONS") -Default "false") $false
$UseHDBoxSetupFallback = Get-Bool (Get-EnvFirst -Names @("UseHDBoxSetupFallback", "USE_HDBOX_SETUP_FALLBACK") -Default "true") $true
$DeleteUserPreferences = Get-Bool (Get-EnvFirst -Names @("DeleteUserPreferences", "DELETE_USER_PREFERENCES") -Default "false") $false
$RebootIfRequired = Get-Bool (Get-EnvFirst -Names @("RebootIfRequired", "REBOOT_IF_REQUIRED") -Default "false") $false
$FailIfRemaining = Get-Bool (Get-EnvFirst -Names @("FailIfRemaining", "FAIL_IF_REMAINING") -Default "true") $true
$VerificationDelaySeconds = Get-Int (Get-EnvFirst -Names @("VerificationDelaySeconds", "VERIFICATION_DELAY_SECONDS") -Default "30") 30

# -----------------------------
# Start
# -----------------------------

Write-Log "========== Adobe forced removal component started =========="
Write-Log "ReportOnly: $ReportOnly"
Write-Log "TargetProducts raw: $TargetProductsRaw"
Write-Log "ExcludeProducts raw: $ExcludeProductsRaw"
Write-Log "TargetProducts parsed: $($TargetProducts -join '; ')"
Write-Log "ExcludeProducts parsed: $($ExcludeProducts -join '; ')"
Write-Log "KillAdobeProcesses: $KillAdobeProcesses"
Write-Log "ForceExeFallback: $ForceExeFallback"
Write-Log "UseAdobeUninstaller: $UseAdobeUninstaller"
Write-Log "AdobeUninstallerPath: $AdobeUninstallerPath"
Write-Log "AdobeSapCodes: $AdobeSapCodes"
Write-Log "AutoAdobeSapCodes: $AutoAdobeSapCodes"
Write-Log "RemoveCreativeCloudAll requested: $RemoveCreativeCloudAll"
Write-Log "AllowRemoveAllWhenTargetsSpecified: $AllowRemoveAllWhenTargetsSpecified"
Write-Log "AllowRemoveAllWithExclusions: $AllowRemoveAllWithExclusions"
Write-Log "UseHDBoxSetupFallback: $UseHDBoxSetupFallback"
Write-Log "DeleteUserPreferences: $DeleteUserPreferences"
Write-Log "RebootIfRequired: $RebootIfRequired"
Write-Log "FailIfRemaining: $FailIfRemaining"
Write-Log "VerificationDelaySeconds: $VerificationDelaySeconds"

$adobeApps = @(Get-AdobeInstalledApps)

Write-Log "Detected Adobe products: $($adobeApps.Count)"

if ($adobeApps.Count -eq 0) {
    Write-Log "No Adobe products detected in uninstall registry."
    exit 0
}

Write-Log "Detected Adobe product names for TargetProducts/ExcludeProducts:"
$uniqueNames = @($adobeApps | Select-Object -ExpandProperty DisplayName -Unique)
foreach ($name in $uniqueNames) {
    Write-Log "  $name"
}

Write-Log "Detailed detected Adobe products:"
foreach ($app in $adobeApps) {
    Write-Log "DETECTED: $($app.DisplayName) | Version=$($app.DisplayVersion) | Publisher=$($app.Publisher) | ProductCode=$($app.PSChildName)"
}

$plan = @(New-PlanFromCurrentAdobeApps)

Write-Log "Action plan:"
foreach ($item in $plan) {
    Write-Log "$($item.Action): $($item.DisplayName) $($item.DisplayVersion) | ProductCode=$($item.ProductCode) | $($item.Reason)"
}

$removeItems = @(Get-RemainingTargetedItems -CurrentPlan $plan)

if ($ReportOnly) {
    if ($removeItems.Count -gt 0) {
        $suggestedSapCodes = @(Get-AutoSapCodesFromPlan -PlannedItems $removeItems)

        if ($suggestedSapCodes.Count -gt 0) {
            Write-Log "Suggested AdobeSapCodes for selected Creative Cloud products:"
            Write-Log "  $($suggestedSapCodes -join '; ')"
        }
    }

    if ($RemoveCreativeCloudAll) {
        if ($TargetProducts.Count -gt 0) {
            Write-Log "ReportOnly safety note: RemoveCreativeCloudAll=true, but TargetProducts is populated. --all would be blocked unless AllowRemoveAllWhenTargetsSpecified=true."
        } elseif ($ExcludeProducts.Count -gt 0) {
            Write-Log "ReportOnly safety note: RemoveCreativeCloudAll=true, but ExcludeProducts is populated. --all would be blocked unless AllowRemoveAllWithExclusions=true."
        } else {
            Write-Log "ReportOnly note: RemoveCreativeCloudAll=true would run AdobeUninstaller.exe --all if ReportOnly=false."
        }
    }

    Write-Log "ReportOnly=true. No uninstall actions were performed."
    Write-Log "Use the detected product names above in TargetProducts or ExcludeProducts, separated with semicolons."
    exit 0
}

if ($TargetProducts.Count -eq 0 -and $ExcludeProducts.Count -eq 0 -and -not $RemoveCreativeCloudAll) {
    Write-Log "ReportOnly=false, but TargetProducts and ExcludeProducts are both empty, and RemoveCreativeCloudAll=false. Refusing to remove anything."
    exit 2
}

if ($removeItems.Count -eq 0 -and -not $RemoveCreativeCloudAll) {
    Write-Log "No Adobe products selected for removal."
    exit 0
}

# -----------------------------
# Execute
# -----------------------------

$globalExitCodes = New-Object System.Collections.Generic.List[int]
$rebootRequired = $false

if ($KillAdobeProcesses) {
    Stop-AdobeProcesses
} else {
    Write-Log "KillAdobeProcesses=false. Adobe processes will not be stopped."
}

# First pass with AdobeUninstaller.exe.
# Important: --all is automatically blocked when TargetProducts or ExcludeProducts are populated,
# unless the explicit override variables are enabled.
$adobeToolResult = Invoke-AdobeUninstallerPass -PlannedItems $removeItems -AllowAllThisPass $true

foreach ($code in @($adobeToolResult.ExitCodes)) {
    Add-ExitCode -List $globalExitCodes -Code ([int]$code)

    if ([int]$code -in @(3010, 1641)) {
        $rebootRequired = $true
    }
}

$skipAdobeDesktopManagedRegistryFallback = $false

if (($adobeToolResult.RanAll -or $adobeToolResult.RanProducts) -and -not $ForceExeFallback) {
    $skipAdobeDesktopManagedRegistryFallback = $true
}

# Registry fallback pass for MSI/non-Creative Cloud-managed apps.
foreach ($item in $removeItems) {
    $app = $item.App

    if ($skipAdobeDesktopManagedRegistryFallback -and (Test-AdobeDesktopManagedApp -App $app)) {
        Write-Log "Skipping registry fallback for $($app.DisplayName) because AdobeUninstaller.exe was already attempted for Creative Cloud-managed apps."
        continue
    }

    Write-Log "Preparing registry uninstall for: $($app.DisplayName) $($app.DisplayVersion)"

    $spec = Get-UninstallCommand -App $app

    if (!$spec) {
        Write-Log "No safe registry uninstall command selected for: $($app.DisplayName)"
        continue
    }

    $exitCode = [int](Invoke-CommandSpec -Spec $spec)
    Add-ExitCode -List $globalExitCodes -Code $exitCode

    switch ($exitCode) {
        0 {
            Write-Log "Successfully removed or uninstall command completed: $($app.DisplayName)"
        }
        3010 {
            Write-Log "Removed with reboot required: $($app.DisplayName)"
            $rebootRequired = $true
        }
        1641 {
            Write-Log "Removed and installer initiated reboot: $($app.DisplayName)"
            $rebootRequired = $true
        }
        1605 {
            Write-Log "Product already absent: $($app.DisplayName)"
        }
        1614 {
            Write-Log "Product already uninstalled: $($app.DisplayName)"
        }
        default {
            Write-Log "Non-success exit code $exitCode for: $($app.DisplayName)"
        }
    }
}

# -----------------------------
# Verification and second-pass cleanup
# -----------------------------

Write-Log "Waiting $VerificationDelaySeconds seconds before verification."
Start-Sleep -Seconds $VerificationDelaySeconds

$remainingPlan = @(New-PlanFromCurrentAdobeApps)
$remainingTargeted = @(Get-RemainingTargetedItems -CurrentPlan $remainingPlan)

if ($remainingTargeted.Count -gt 0) {
    Write-Log "Second pass required. Remaining targeted Adobe products:"
    foreach ($item in $remainingTargeted) {
        Write-Log "REMAINING BEFORE SECOND PASS: $($item.DisplayName) $($item.DisplayVersion) | ProductCode=$($item.ProductCode)"
    }

    if ($UseAdobeUninstaller) {
        $secondPassSapCodes = @(Get-AutoSapCodesFromPlan -PlannedItems $remainingTargeted)

        if ($secondPassSapCodes.Count -gt 0) {
            $savedAdobeSapCodes = $AdobeSapCodes
            $AdobeSapCodes = ($secondPassSapCodes -join '; ')

            Write-Log "Running AdobeUninstaller.exe second pass with SAP codes: $AdobeSapCodes"

            $secondPassResult = Invoke-AdobeUninstallerPass -PlannedItems $remainingTargeted -AllowAllThisPass $false

            foreach ($code in @($secondPassResult.ExitCodes)) {
                Add-ExitCode -List $globalExitCodes -Code ([int]$code)

                if ([int]$code -in @(3010, 1641)) {
                    $rebootRequired = $true
                }
            }

            $AdobeSapCodes = $savedAdobeSapCodes
        } else {
            Write-Log "No SAP codes could be generated for AdobeUninstaller.exe second pass."
        }
    }

    Write-Log "Waiting $VerificationDelaySeconds seconds before second verification."
    Start-Sleep -Seconds $VerificationDelaySeconds

    $remainingPlan = @(New-PlanFromCurrentAdobeApps)
    $remainingTargeted = @(Get-RemainingTargetedItems -CurrentPlan $remainingPlan)

    if ($remainingTargeted.Count -gt 0 -and $UseHDBoxSetupFallback) {
        $hdBoxCandidates = @(
            $remainingTargeted |
                Where-Object {
                    (Get-AdobeSapOnlyFromApp -App $_.App) -and
                    (Get-AdobeBaseVersionFromApp -App $_.App)
                }
        )

        if ($hdBoxCandidates.Count -gt 0) {
            Write-Log "AdobeUninstaller.exe second pass did not remove all targets. Trying HDBox Setup.exe fallback for eligible Creative Cloud-managed apps."

            foreach ($code in @(Invoke-HDBoxSetupFallback -RemainingItems $hdBoxCandidates)) {
                Add-ExitCode -List $globalExitCodes -Code ([int]$code)

                if ([int]$code -in @(3010, 1641)) {
                    $rebootRequired = $true
                }
            }

            Write-Log "Waiting $VerificationDelaySeconds seconds before final verification."
            Start-Sleep -Seconds $VerificationDelaySeconds

            $remainingPlan = @(New-PlanFromCurrentAdobeApps)
            $remainingTargeted = @(Get-RemainingTargetedItems -CurrentPlan $remainingPlan)
        } else {
            Write-Log "No remaining targeted products are eligible for HDBox Setup.exe fallback."
        }
    }
}

if ($remainingTargeted.Count -gt 0) {
    Write-Log "Remaining targeted Adobe products after all uninstall attempts:"
    foreach ($item in $remainingTargeted) {
        Write-Log "REMAINING: $($item.DisplayName) $($item.DisplayVersion) | ProductCode=$($item.ProductCode) | $($item.Reason)"
    }

    $ccDesktopStillRemaining = @(
        $remainingTargeted |
            Where-Object { $_.DisplayName -match '(?i)^Adobe Creative Cloud$' }
    )

    if ($ccDesktopStillRemaining.Count -gt 0 -and $TargetProducts.Count -gt 0 -and -not $AllowRemoveAllWhenTargetsSpecified) {
        Write-Log "NOTE: Adobe Creative Cloud Desktop remains targeted, but --all was not run because TargetProducts is populated."
        Write-Log "NOTE: This is intentional to avoid removing untargeted products such as Acrobat/Reader."
        Write-Log "NOTE: If you intentionally want broad removal, set AllowRemoveAllWhenTargetsSpecified=true and keep RemoveCreativeCloudAll=true."
    }

    if ($FailIfRemaining) {
        Write-Log "Adobe forced removal completed with remaining targeted products. FailIfRemaining=true, exiting 1."
        exit 1
    } else {
        Write-Log "Adobe forced removal completed with remaining targeted products. FailIfRemaining=false, continuing."
    }
}

$badExitCodes = @(
    $globalExitCodes |
        Where-Object { [int]$_ -notin @(0, 1605, 1614, 1641, 3010) }
)

if ($badExitCodes.Count -gt 0) {
    Write-Log "Adobe forced removal completed with one or more non-success exit codes: $($badExitCodes -join ', ')"
    exit 1
}

if ($rebootRequired -and $RebootIfRequired) {
    Write-Log "Reboot required and RebootIfRequired=true. Restarting endpoint in 60 seconds."
    shutdown.exe /r /t 60 /c "Adobe product removal completed. Reboot required."
}

Write-Log "Adobe forced removal completed successfully."
exit 0