#Requires -RunAsAdministrator
[CmdletBinding()]
param()

# =========================
# Datto Component Variables
# =========================
# Create Datto input variables with these EXACT names:
# MinimumVersionToKeep      (String)  e.g. 6.3.0.0
# ReportOnly                (Boolean) true/false
# StopZoomProcesses         (Boolean) true/false
# CleanupResiduals          (Boolean) true/false
# OperationTimeoutSec       (String/Number) e.g. 600

$MinimumVersionToKeep = if ($env:MinimumVersionToKeep) { $env:MinimumVersionToKeep } else { '6.3.0.0' }
$ReportOnly           = if ($env:ReportOnly)           { $env:ReportOnly }           else { 'true' }
$StopZoomProcesses    = if ($env:StopZoomProcesses)    { $env:StopZoomProcesses }    else { 'true' }
$CleanupResiduals     = if ($env:CleanupResiduals)     { $env:CleanupResiduals }     else { 'true' }
$OperationTimeoutSec  = if ($env:OperationTimeoutSec)  { $env:OperationTimeoutSec }  else { '600' }

$LogPath = "$env:ProgramData\Datto\Logs\Zoom-Cleanup.log"
$script:Actions = New-Object System.Collections.Generic.List[string]
$script:Errors  = New-Object System.Collections.Generic.List[string]

function To-Bool {
    param($Value, [bool]$Default = $false)
    if ($null -eq $Value) { return $Default }
    switch ($Value.ToString().Trim().ToLowerInvariant()) {
        '1'     { return $true }
        'true'  { return $true }
        'yes'   { return $true }
        'y'     { return $true }
        'on'    { return $true }
        '0'     { return $false }
        'false' { return $false }
        'no'    { return $false }
        'n'     { return $false }
        'off'   { return $false }
        default { return $Default }
    }
}

function To-Version {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $parts = (($Text -replace '[^\d\.]', '') -split '\.') | Where-Object { $_ -match '^\d+$' }
    if (-not $parts) { return $null }
    while ($parts.Count -lt 2) { $parts += '0' }
    if ($parts.Count -gt 4) { $parts = $parts[0..3] }
    try { return [version]($parts -join '.') } catch { return $null }
}

function Log([string]$Message) {
    try {
        $dir = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        Add-Content -LiteralPath $LogPath -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
    } catch {}
}

function Add-Action([string]$Text) {
    $script:Actions.Add($Text) | Out-Null
    Log $Text
}

function Add-ErrorText([string]$Text) {
    $script:Errors.Add($Text) | Out-Null
    Log "ERROR: $Text"
}

function Out-Result {
    param(
        [string]$Status,
        [string]$Summary,
        [int]$ExitCode
    )
    Write-Host "STATUS=$Status"
    Write-Host "SUMMARY=$Summary"
    Write-Host "LOG=$LogPath"
    foreach ($line in $script:Actions) { Write-Host "DETAIL=$line" }
    foreach ($line in $script:Errors)  { Write-Host "ERROR=$line" }
    exit $ExitCode
}

function Test-IsZoomDesktopName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -notmatch '(?i)^Zoom') { return $false }
    if ($Name -match '(?i)(VDI|Rooms|Outlook|Plugin|Scheduler|Add-?in|Lync|Skype|Notes|Docs|Mail|AOMhost|Auto-Update|Universal Installer)') { return $false }
    return $true
}

function Get-CommandExecutablePath {
    param([string]$CommandText)

    if ([string]::IsNullOrWhiteSpace($CommandText)) { return $null }

    $text = ($CommandText -split ',')[0].Trim()

    if ($text -match '^\s*"([^"]+)"') {
        return $matches[1]
    }

    if ($text -match '^\s*([^\s]+(?:\.exe|\.msi))') {
        return $matches[1]
    }

    return $null
}

function Test-EntryHasFiles {
    param($Entry)

    $paths = @()

    if ($Entry.InstallLocation) {
        $paths += $Entry.InstallLocation
    }

    $iconPath = Get-CommandExecutablePath $Entry.DisplayIcon
    if ($iconPath) {
        $paths += $iconPath
    }

    $quietPath = Get-CommandExecutablePath $Entry.QuietUninstallString
    if ($quietPath) {
        $paths += $quietPath
    }

    foreach ($p in ($paths | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (Test-Path -LiteralPath $p) { return $true }
        try {
            $parent = Split-Path -Path $p -Parent
            if ($parent -and (Test-Path -LiteralPath $parent)) { return $true }
        } catch {}
    }

    return $false
}

function Get-ZoomEntries {
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $entries = @()

    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        foreach ($sub in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            try {
                $p = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction Stop
            } catch {
                continue
            }

            if (-not (Test-IsZoomDesktopName $p.DisplayName)) { continue }

            $obj = [pscustomobject]@{
                KeyPath              = $sub.PSPath
                KeyName              = $sub.PSChildName
                DisplayName          = $p.DisplayName
                DisplayVersion       = $p.DisplayVersion
                Version              = To-Version $p.DisplayVersion
                InstallLocation      = $p.InstallLocation
                DisplayIcon          = $p.DisplayIcon
                UninstallString      = $p.UninstallString
                QuietUninstallString = $p.QuietUninstallString
                WindowsInstaller     = [int]($p.WindowsInstaller)
            }

            $obj | Add-Member -NotePropertyName HasFiles -NotePropertyValue (Test-EntryHasFiles $obj) -Force
            $entries += $obj
        }
    }

    return $entries
}

function Build-UninstallCommand {
    param($Entry)

    if ($Entry.QuietUninstallString) {
        return $Entry.QuietUninstallString
    }

    $guid = $null
    if ($Entry.KeyName -match '^\{[A-Fa-f0-9\-]+\}$') {
        $guid = $Entry.KeyName
    } elseif ($Entry.UninstallString -match '(\{[A-Fa-f0-9\-]+\})') {
        $guid = $matches[1]
    }

    if (($Entry.WindowsInstaller -eq 1) -and $guid) {
        return "msiexec.exe /x $guid /qn /norestart"
    }

    if ($Entry.UninstallString -and $Entry.UninstallString -match '(?i)(/quiet|/silent|/qn\b|/s\b)') {
        return $Entry.UninstallString
    }

    return $null
}

function Run-Command {
    param(
        [string]$Command,
        [int]$TimeoutSec = 600
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\cmd.exe"
    $psi.Arguments = "/c $Command"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()

    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        try { & taskkill.exe /PID $proc.Id /T /F *> $null } catch {}
        throw "Timed out after $TimeoutSec seconds"
    }

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    if (-not [string]::IsNullOrWhiteSpace($stdout)) { Log "STDOUT: $($stdout.Trim())" }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) { Log "STDERR: $($stderr.Trim())" }

    return $proc.ExitCode
}

function Remove-ItemSafe {
    param(
        [string]$Path,
        [string]$Why
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        Add-Action "REMOVED path [$Path] ($Why)"
    } catch {
        Add-ErrorText "Failed removing path [$Path]: $($_.Exception.Message)"
    }
}

function Remove-RegKeySafe {
    param(
        [string]$KeyPath,
        [string]$Why
    )

    if (-not (Test-Path -LiteralPath $KeyPath)) { return }

    try {
        Remove-Item -LiteralPath $KeyPath -Recurse -Force -ErrorAction Stop
        Add-Action "REMOVED regkey [$KeyPath] ($Why)"
    } catch {
        Add-ErrorText "Failed removing regkey [$KeyPath]: $($_.Exception.Message)"
    }
}

function Get-UserProfiles {
    $systemSids     = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')
    $profileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $profiles       = @()

    foreach ($key in Get-ChildItem $profileListKey -ErrorAction SilentlyContinue) {
        $sid = $key.PSChildName
        if ($sid -in $systemSids) { continue }
        if ($sid -notmatch '^S-1-5-21-') { continue }

        try {
            $props = Get-ItemProperty $key.PSPath -ErrorAction Stop
            $path  = $props.ProfileImagePath
            if ($path -and (Test-Path -LiteralPath $path)) {
                $profiles += [pscustomobject]@{
                    SID         = $sid
                    ProfilePath = $path
                }
            }
        } catch {}
    }

    return $profiles
}

function Get-PerUserZoomInstalls {
    $installs = @()

    foreach ($profile in @(Get-UserProfiles)) {
        $profilePath = $profile.ProfilePath

        foreach ($relBin in @('AppData\Roaming\Zoom\bin\Zoom.exe', 'AppData\Local\Zoom\bin\Zoom.exe')) {
            $zoomExe = Join-Path $profilePath $relBin
            if (-not (Test-Path -LiteralPath $zoomExe)) { continue }

            $fileVer = try {
                (Get-Item -LiteralPath $zoomExe -ErrorAction Stop).VersionInfo.FileVersion
            } catch { $null }

            $version  = To-Version $fileVer
            $zoomRoot = Split-Path (Split-Path $zoomExe -Parent) -Parent

            $uninstallExe = Join-Path $zoomRoot 'uninstall\Installer.exe'
            if (-not (Test-Path -LiteralPath $uninstallExe)) { $uninstallExe = $null }

            $installs += [pscustomobject]@{
                SID            = $profile.SID
                ProfilePath    = $profilePath
                ZoomRoot       = $zoomRoot
                ZoomExe        = $zoomExe
                DisplayVersion = $fileVer
                Version        = $version
                InstallerPath  = $uninstallExe
                DisplayName    = "Zoom (per-user: $(Split-Path $profilePath -Leaf))"
            }
            break
        }
    }

    return $installs
}

function Remove-PerUserZoomRegistry {
    param(
        [string]$SID,
        [string]$ProfilePath
    )

    # If the user's hive is already loaded (user currently logged in), use it directly
    $loadedHivePath = "Registry::HKEY_USERS\$SID"
    if (Test-Path $loadedHivePath) {
        $uninstallPath = "$loadedHivePath\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        if (Test-Path $uninstallPath) {
            foreach ($sub in Get-ChildItem $uninstallPath -ErrorAction SilentlyContinue) {
                try {
                    $props = Get-ItemProperty $sub.PSPath -ErrorAction Stop
                    if (Test-IsZoomDesktopName $props.DisplayName) {
                        Remove-RegKeySafe -KeyPath $sub.PSPath -Why 'per-user Zoom HKCU cleanup'
                    }
                } catch {}
            }
        }
        return
    }

    # User is not logged in; temporarily load the offline hive
    $hivePath = Join-Path $ProfilePath 'NTUSER.DAT'
    if (-not (Test-Path -LiteralPath $hivePath)) { return }

    $mountKey = "ZoomClean_$($SID -replace '[^A-Za-z0-9]', '_')"
    try {
        $null = & reg.exe load "HKU\$mountKey" $hivePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            Add-Action "Skipped HKCU registry cleanup for $ProfilePath (hive could not be loaded)"
            return
        }

        $uninstallPath = "Registry::HKEY_USERS\$mountKey\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        if (Test-Path $uninstallPath) {
            foreach ($sub in Get-ChildItem $uninstallPath -ErrorAction SilentlyContinue) {
                try {
                    $props = Get-ItemProperty $sub.PSPath -ErrorAction Stop
                    if (Test-IsZoomDesktopName $props.DisplayName) {
                        Remove-RegKeySafe -KeyPath $sub.PSPath -Why 'per-user Zoom HKCU cleanup'
                    }
                } catch {}
            }
        }
    } finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $null = & reg.exe unload "HKU\$mountKey" 2>&1
    }
}

try {
    $reportOnly       = To-Bool $ReportOnly $true
    $stopZoom         = To-Bool $StopZoomProcesses $true
    $cleanupResiduals = To-Bool $CleanupResiduals $true
    $timeout          = [int]$OperationTimeoutSec
    $minVersion       = To-Version $MinimumVersionToKeep

    if (-not $minVersion) {
        Out-Result -Status 'Failed' -Summary "Invalid MinimumVersionToKeep [$MinimumVersionToKeep]" -ExitCode 1
    }

    Add-Action "Starting Zoom version check"
    Add-Action "Settings: MinimumVersionToKeep=$minVersion ReportOnly=$reportOnly CleanupResiduals=$cleanupResiduals TimeoutSec=$timeout"

    $entries      = @(Get-ZoomEntries)
    $perUserItems = @(Get-PerUserZoomInstalls)

    if (-not $entries.Count -and -not $perUserItems.Count) {
        Add-Action "No Zoom desktop uninstall entries or per-user installs found"
        Out-Result -Status 'OK' -Summary 'No Zoom found. Nothing to do.' -ExitCode 0
    }

    Add-Action "Logging discovered Zoom entries"
    foreach ($e in $entries) {
        $versionText = if ($e.Version) { $e.Version.ToString() } else { 'Unknown' }
        $stateText   = if ($e.HasFiles) { 'Active' } else { 'RegistryOnly' }
        Add-Action "FOUND Name=[$($e.DisplayName)] Version=[$versionText] RawVersion=[$($e.DisplayVersion)] State=[$stateText]"
    }
    foreach ($pu in $perUserItems) {
        $versionText = if ($pu.Version) { $pu.Version.ToString() } else { 'Unknown' }
        Add-Action "FOUND Name=[$($pu.DisplayName)] Version=[$versionText] RawVersion=[$($pu.DisplayVersion)] Profile=[$($pu.ProfilePath)] State=[Active]"
    }

    $activeEntries = @($entries | Where-Object { $_.HasFiles })
    if (-not $activeEntries.Count -and -not $perUserItems.Count) {
        Add-Action "No active Zoom install detected. Registry-only entries found; exiting without cleanup."
        Out-Result -Status 'OK' -Summary 'No active Zoom install detected. Logged findings only.' -ExitCode 0
    }

    $candidates        = @($activeEntries | Where-Object { $_.Version -and $_.Version -lt $minVersion })
    $perUserCandidates = @($perUserItems  | Where-Object { $_.Version -and $_.Version -lt $minVersion })

    if (-not $candidates.Count -and -not $perUserCandidates.Count) {
        Add-Action "Active Zoom detected, but no versions below minimum [$minVersion]"
        Out-Result -Status 'OK' -Summary "No Zoom versions below minimum [$minVersion]. Nothing to do." -ExitCode 0
    }

    Add-Action "Candidates below minimum version:"
    foreach ($e in $candidates) {
        Add-Action "CANDIDATE Name=[$($e.DisplayName)] Version=[$($e.Version)]"
    }
    foreach ($pu in $perUserCandidates) {
        Add-Action "CANDIDATE Name=[$($pu.DisplayName)] Version=[$($pu.Version)] Profile=[$($pu.ProfilePath)]"
    }

    if ($reportOnly) {
        foreach ($e in $candidates) {
            $cmd = Build-UninstallCommand $e
            if ($cmd) {
                Add-Action "REPORT uninstall Name=[$($e.DisplayName)] Version=[$($e.Version)] Command=[$cmd]"
            } else {
                Add-Action "REPORT skip Name=[$($e.DisplayName)] Version=[$($e.Version)] Reason=[No silent uninstall command]"
            }
        }
        foreach ($pu in $perUserCandidates) {
            Add-Action "REPORT uninstall Name=[$($pu.DisplayName)] Version=[$($pu.Version)] Command=[Remove $($pu.ZoomRoot)]"
        }

        $totalCandidates = $candidates.Count + $perUserCandidates.Count
        Out-Result -Status 'ReportOnly' -Summary "Found $totalCandidates Zoom entries below minimum [$minVersion]. No changes made." -ExitCode 0
    }

    if ($stopZoom) {
        Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.ProcessName -match '^(?i)(Zoom|CptService)$' } |
            ForEach-Object {
                try {
                    Stop-Process -Id $_.Id -Force -ErrorAction Stop
                    Add-Action "STOPPED process $($_.ProcessName) PID=$($_.Id)"
                } catch {
                    Add-ErrorText "Failed stopping process $($_.ProcessName) PID=$($_.Id)"
                }
            }
        Start-Sleep -Seconds 3  # Allow OS to release file handles before uninstaller runs
    }

    foreach ($e in $candidates) {
        $cmd   = Build-UninstallCommand $e
        $label = "$($e.DisplayName) [$($e.Version)]"

        if (-not $cmd) {
            Add-ErrorText "Skipped $label because no silent uninstall command was available"
            continue
        }

        try {
            Add-Action "START uninstall $label"
            $exitCode = Run-Command -Command $cmd -TimeoutSec $timeout
            if ($exitCode -in @(0,1605,1614,3010)) {
                Add-Action "END uninstall $label exitcode=$exitCode"
            } else {
                Add-ErrorText "Uninstall failed for $label exitcode=$exitCode"
            }
        } catch {
            Add-ErrorText "Uninstall exception for $label: $($_.Exception.Message)"
        }
    }

    foreach ($pu in $perUserCandidates) {
        $label = "$($pu.DisplayName) [$($pu.Version)]"

        try {
            Add-Action "START per-user cleanup $label"
            Remove-ItemSafe -Path $pu.ZoomRoot -Why "per-user Zoom below minimum [$minVersion]"
            Remove-PerUserZoomRegistry -SID $pu.SID -ProfilePath $pu.ProfilePath
            Add-Action "END per-user cleanup $label"
        } catch {
            Add-ErrorText "Per-user cleanup exception for $label: $($_.Exception.Message)"
        }
    }

    if ($cleanupResiduals) {
        $remaining        = @(Get-ZoomEntries | Where-Object { $_.HasFiles })
        $remainingPerUser = @(Get-PerUserZoomInstalls)

        if (-not $remaining.Count -and -not $remainingPerUser.Count) {
            Add-Action "No active Zoom install remains. Performing residual cleanup."

            foreach ($p in @(
                "$env:ProgramFiles\Zoom",
                "${env:ProgramFiles(x86)}\Zoom",
                "$env:ProgramData\Zoom"
            )) {
                if ($p) { Remove-ItemSafe -Path $p -Why 'no active Zoom remains' }
            }

            foreach ($e in $entries) {
                if (-not $e.HasFiles) {
                    Remove-RegKeySafe -KeyPath $e.KeyPath -Why 'registry-only Zoom entry after removal'
                }
            }

            foreach ($profile in @(Get-UserProfiles)) {
                foreach ($relPath in @('AppData\Roaming\Zoom', 'AppData\Local\Zoom')) {
                    $zoomPath = Join-Path $profile.ProfilePath $relPath
                    if (Test-Path -LiteralPath $zoomPath) {
                        Remove-ItemSafe -Path $zoomPath -Why 'per-user Zoom residual cleanup'
                    }
                }
                Remove-PerUserZoomRegistry -SID $profile.SID -ProfilePath $profile.ProfilePath
            }
        } else {
            Add-Action "An active Zoom install still remains. Residual shared-folder cleanup skipped."
        }
    }

    $totalProcessed = $candidates.Count + $perUserCandidates.Count

    if ($script:Errors.Count -gt 0) {
        Out-Result -Status 'CompletedWithErrors' -Summary "Processed $totalProcessed candidate(s) below minimum [$minVersion] with errors." -ExitCode 1
    }

    Out-Result -Status 'Success' -Summary "Processed $totalProcessed candidate(s) below minimum [$minVersion]." -ExitCode 0
}
catch {
    Add-ErrorText $_.Exception.Message
    Out-Result -Status 'Failed' -Summary 'Script failed unexpectedly.' -ExitCode 1
}