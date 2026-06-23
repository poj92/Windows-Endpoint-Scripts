<#
.SYNOPSIS
    Force-remove WPS Office (Kingsoft WPS) from a Windows machine.

.DESCRIPTION
    - Searches registry uninstall entries for "WPS" / "Kingsoft" / "WPS Office".
    - Tries to run the product's uninstall string silently.
    - Kills related processes, removes leftover files, registry keys, scheduled tasks.
    - Dry-run by default. Use -ForceRemove to actually remove.

.PARAMETER ForceRemove
    When supplied, the script performs destructive actions. Without it, script only reports what it would do.

.PARAMETER LogPath
    Path to log file. Default: .\WPS-removal-log.txt

.EXAMPLE
    # Dry run (default)
    .\Remove-WPS.ps1

    # Real removal
    .\Remove-WPS.ps1 -ForceRemove -LogPath C:\temp\wps-uninstall.log
#>

param (
    [switch]$ForceRemove = $false,
    [string]$LogPath = ".\WPS-removal-log.txt"
)

function Log {
    param ($Message)
    $time = (Get-Date).ToString("u")
    $line = "$time`t$Message"
    Write-Output $line
    Add-Content -Path $LogPath -Value $line
}

# Ensure admin
if (-not ([bool]([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

# initial
if (Test-Path $LogPath) { Remove-Item $LogPath -ErrorAction SilentlyContinue }
Log "Starting WPS removal script. ForceRemove=$ForceRemove"

# helper: find uninstall entries in registry
function Get-UninstallEntries {
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($k in $keys) {
        try {
            Get-ItemProperty -Path $k -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName   = $_.DisplayName
                    DisplayVersion= $_.DisplayVersion
                    Publisher     = $_.Publisher
                    UninstallString = $_.UninstallString
                    InstallLocation = $_.InstallLocation
                    RegistryPath  = $k -replace '\*\z','' + "\$($_.PSChildName)"
                    PSChildName   = $_.PSChildName
                }
            }
        } catch { }
    }
}

# find candidate entries (common name patterns)
$pattern = "WPS","Kingsoft","wps office","wps-office"
$entries = Get-UninstallEntries | Where-Object {
    $_.DisplayName -and ($pattern | ForEach-Object { $_ } | ForEach-Object { $_ } ) # placeholder
}

# better filter
$entries = Get-UninstallEntries | Where-Object {
    if (-not $_.DisplayName) { return $false }
    $dn = $_.DisplayName.ToLower()
    $patternList = @("wps","kingsoft","wps office","kingsoft office","wps-office","wps office 2019")
    foreach ($p in $patternList) {
        if ($dn -like "*$p*") { return $true }
    }
    return $false
}

if (-not $entries -or $entries.Count -eq 0) {
    Log "No registry uninstall entries matched common WPS names."
} else {
    Log "Found uninstall registry entries:"
    $entries | ForEach-Object { Log " - $($_.DisplayName) | UninstallString: $($_.UninstallString) | InstallLocation: $($_.InstallLocation)" }
}

# Common install paths to check
$possiblePaths = @(
    "$env:ProgramFiles\Kingsoft\WPS Office",
    "$env:ProgramFiles(x86)\Kingsoft\WPS Office",
    "$env:ProgramFiles\WPS Office",
    "$env:ProgramFiles(x86)\WPS Office",
    "$env:LOCALAPPDATA\Kingsoft\WPS Office",
    "$env:ProgramFiles\Kingsoft\office6",
    "$env:ProgramFiles(x86)\Kingsoft\office6"
) | Where-Object { $_ }

# Detect processes to kill
$procNames = @("wps","wpscloudsvr","wpp","et","wppshell","ksafe") # wpp (presentation), et (spreadsheets)
$running = Get-Process -ErrorAction SilentlyContinue | Where-Object { $procNames -contains $_.ProcessName.ToLower() }

if ($running) {
    Log "Found running WPS-related processes: $($running | Select-Object -ExpandProperty ProcessName -Unique -ErrorAction SilentlyContinue -Verbose:$false | Sort-Object -Unique -Join ', ')"
    foreach ($p in $running) {
        if ($ForceRemove) {
            try {
                Log "Stopping process $($p.ProcessName) (Id $($p.Id))"
                Stop-Process -Id $p.Id -Force -ErrorAction Stop
            } catch {
                Log "Failed to stop $($p.ProcessName): $($_.Exception.Message)"
            }
        } else {
            Log "DRY RUN: Would stop process $($p.ProcessName) (Id $($p.Id))"
        }
    }
} else {
    Log "No known WPS processes running."
}

# Attempt uninstall via uninstallstring from registry entries
foreach ($entry in $entries) {
    $u = $entry.UninstallString
    if ([string]::IsNullOrWhiteSpace($u)) {
        Log "No uninstall string for $($entry.DisplayName)"
        continue
    }

    Log "Attempting uninstall for '$($entry.DisplayName)' using uninstallstring: $u"
    # Some uninstall strings include a path with quotes or are msiexec GUIDs. Normalize.
    # If it's msiexec, ensure silent args added.
    $cmd = $null
    if ($u -match "msiexec" -or $u -match "/I{") {
        # convert to msiexec /x GUID /qn /norestart
        if ($u -match "({[0-9A-Fa-f\-]+})") {
            $guid = $matches[1]
            $cmd = "msiexec.exe /x $guid /qn /norestart"
        } else {
            # fallback: run uninstall string but add /qn if msiexec present
            if ($u -notmatch "/qn") { $cmd = "$u /qn /norestart" } else { $cmd = $u }
        }
    } else {
        # many vendors include an uninstaller exe path (uninst.exe or uninstall.exe). Add silent args if typical.
        $cmd = $u
        # add common silent switches if not present
        if ($cmd -notmatch "/S" -and $cmd -notmatch "/qn" -and $cmd -notmatch "/silent" -and $cmd -notmatch "/VERYSILENT") {
            # try /S and /qn as possible
            $cmd = "$cmd /S"
        }
    }

    if ($ForceRemove) {
        try {
            Log "Executing: $cmd"
            $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
            Log "Uninstall process exited with code $($proc.ExitCode)"
        } catch {
            Log "Error executing uninstall command: $($_.Exception.Message)"
        }
    } else {
        Log "DRY RUN: Would run uninstall command: $cmd"
    }
}

# If registry uninstall not found or uninstall didn't remove, remove folders
foreach ($p in $possiblePaths) {
    if (Test-Path $p) {
        if ($ForceRemove) {
            try {
                Log "Removing folder $p"
                Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop
                Log "Removed $p"
            } catch {
                Log "Failed to remove $p: $($_.Exception.Message)"
            }
        } else {
            Log "DRY RUN: Would remove folder $p"
        }
    } else {
        Log "Folder not found: $p"
    }
}

# Remove leftover registry keys commonly used by WPS
$regPaths = @(
    "HKCU:\Software\Kingsoft",
    "HKLM:\SOFTWARE\Kingsoft",
    "HKLM:\SOFTWARE\Wow6432Node\Kingsoft",
    "HKCU:\Software\WPS Office",
    "HKLM:\SOFTWARE\WPS Office",
    "HKLM:\SOFTWARE\Wow6432Node\WPS Office"
)

foreach ($rk in $regPaths) {
    try {
        if (Test-Path $rk) {
            if ($ForceRemove) {
                Log "Removing registry key $rk"
                Remove-Item -Path $rk -Recurse -Force -ErrorAction Stop
                Log "Removed $rk"
            } else {
                Log "DRY RUN: Would remove registry key $rk"
            }
        } else {
            Log "Registry key not present: $rk"
        }
    } catch {
        Log "Failed removing registry key $rk: $($_.Exception.Message)"
    }
}

# Remove scheduled tasks named like Kingsoft/WPS
try {
    $tasks = schtasks /query /fo LIST | Out-String
    # simple string check; will log if any tasks contain 'wps' or 'kingsoft'
    if ($tasks -match "(?i)wps|kingsoft") {
        Log "Scheduled tasks output contains WPS/Kingsoft markers. Searching tasks..."
        $allTasks = schtasks /query /fo CSV /v 2>&1 | ConvertFrom-Csv -ErrorAction SilentlyContinue
        foreach ($t in $allTasks) {
            $taskName = $t."TaskName"
            if ($taskName -and ($taskName -match "(?i)wps|kingsoft")) {
                if ($ForceRemove) {
                    Log "Deleting scheduled task $taskName"
                    schtasks /Delete /TN $taskName /F | Out-Null
                } else {
                    Log "DRY RUN: Would delete scheduled task $taskName"
                }
            }
        }
    } else {
        Log "No scheduled tasks matched WPS/Kingsoft."
    }
} catch {
    Log "Could not enumerate scheduled tasks: $($_.Exception.Message)"
}

# Remove leftover start menu / shortcuts in common locations
$shortcutLocations = @(
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\WPS Office*",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Kingsoft*",
    "$env:Public\Desktop\WPS*",
    "$env:UserProfile\Desktop\WPS*"
)

foreach ($s in $shortcutLocations) {
    $matches = Get-ChildItem -Path $s -ErrorAction SilentlyContinue
    if ($matches) {
        foreach ($m in $matches) {
            if ($ForceRemove) {
                try {
                    Log "Removing shortcut $($m.FullName)"
                    Remove-Item -LiteralPath $m.FullName -Force -ErrorAction Stop
                } catch {
                    Log "Failed to remove $($m.FullName): $($_.Exception.Message)"
                }
            } else {
                Log "DRY RUN: Would remove shortcut $($m.FullName)"
            }
        }
    } else {
        Log "No start menu/shortcut matches for pattern $s"
    }
}

# Attempt to remove any MSI product codes that match WPS by name using Win32_Product is NOT recommended (can trigger repair).
# Instead, check presence via Get-Package (if available) and remove.
try {
    $pkgs = Get-Package -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "(?i)wps|kingsoft" }
    if ($pkgs) {
        foreach ($p in $pkgs) {
            Log "Found package via Get-Package: $($p.Name) | Provider: $($p.ProviderName)"
            if ($ForceRemove) {
                try {
                    Log "Uninstalling package $($p.Name) via Uninstall-Package"
                    Uninstall-Package -InputObject $p -Force -ErrorAction Stop
                } catch {
                    Log "Failed to uninstall package $($p.Name): $($_.Exception.Message)"
                }
            } else {
                Log "DRY RUN: Would uninstall package $($p.Name)"
            }
        }
    } else {
        Log "No package matches in Get-Package for WPS/Kingsoft."
    }
} catch {
    Log "Get-Package not available or failed: $($_.Exception.Message)"
}

# Final check: see if any files / processes still present
$remainingProcs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $procNames -contains $_.ProcessName.ToLower() }
$remainingFolders = $possiblePaths | Where-Object { Test-Path $_ }

if ($remainingProcs.Count -gt 0 -or $remainingFolders.Count -gt 0) {
    Log "Post-clean check: remaining processes: $($remainingProcs | Select-Object -ExpandProperty ProcessName -Unique -ErrorAction SilentlyContinue -Join ', ')"
    Log "Remaining folders: $($remainingFolders -join '; ')"
    Log "You may need a reboot to complete removal or manual inspection."
} else {
    Log "Post-clean check: no known WPS processes or folders detected."
}

Log "WPS removal script finished."
if (-not $ForceRemove) {
    Log "NOTE: This was a DRY RUN. Re-run with -ForceRemove to perform actual removals."
}
