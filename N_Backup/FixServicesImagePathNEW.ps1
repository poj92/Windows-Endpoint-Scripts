<#
.SYNOPSIS
  Quote ImagePath entries that contain spaces, handle NT prefixes (\??\, \\?\), skip core Windows services.

.DESCRIPTION
  Scans HKLM:\SYSTEM\CurrentControlSet\Services
  - Expands environment variables like %SystemRoot%/%ProgramFiles%,
  - Detects the executable/driver path by matching known file extensions (.exe, .sys, .dll, .com, .drv),
  - Correctly handles NT-style prefixes (\??\ and \\?\),
  - Ensures paths containing spaces are quoted (only the path, not args),
  - Skips core Windows binaries and anything under %SystemRoot% to avoid boot breakage,
  - Creates a registry backup, logs actions, and prints a summary.
  
.NOTES
  Run as Administrator.
  Dry-run by default. Set $applyChanges = $true to actually apply changes.
#>

$timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath        = "C:\ServicePathFix_$timestamp.log"
$backupPath     = "C:\Services_Backup_$timestamp.reg"
$applyChanges   = $false    # Set to $true to apply changes

# core binary file names to skip (lowercase)
$coreBinaryNames = @(
    "svchost.exe","wininit.exe","winlogon.exe","lsass.exe","services.exe",
    "smss.exe","csrss.exe","spoolsv.exe","win32k.sys","ntoskrnl.exe"
)

# core service names to skip (service key names)
$coreServices = @(
    "AudioSrv","BFE","CryptSvc","Dhcp","Dnscache","EventLog","LanmanWorkstation",
    "LanmanServer","PlugPlay","Power","RpcSs","SamSs","Schedule","SENS",
    "ShellHWDetection","Spooler","Themes","W32Time","WinDefend","Winmgmt",
    "WlanSvc","wuauserv","Wininit","Winlogon","lsass","services","TrustedInstaller",
    "Appinfo","UserManager","ProfSvc","nsi","PolicyAgent","MpsSvc","IKEEXT","DcomLaunch"
)

[int]$fixedCount = 0
[int]$skippedCount = 0
[int]$totalCount = 0
[int]$coreSkippedCount = 0

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Write-Output $entry
    Add-Content -Path $logPath -Value $entry
}

Write-Log "=== Starting Universal Service Path Quote Check (robust parsing, skip core binaries) ==="
Write-Log "Log file: $logPath"
Write-Log "Backup file: $backupPath"
Write-Log "Dry run mode: $(!$applyChanges)"
Write-Log ""

# create registry backup
try {
    Write-Log "Creating backup of HKLM\SYSTEM\CurrentControlSet\Services to $backupPath ..."
    reg export "HKLM\SYSTEM\CurrentControlSet\Services" $backupPath /y | Out-Null
    Write-Log "Backup completed successfully.`n"
} catch {
    Write-Log "ERROR: Failed to create registry backup. $_"
    exit 1
}

$systemRootExpanded = [Environment]::ExpandEnvironmentVariables("%SystemRoot%").TrimEnd('\').ToLower()

$servicesPath = "HKLM:\SYSTEM\CurrentControlSet\Services"
$services = Get-ChildItem $servicesPath

foreach ($svc in $services) {
    $totalCount++
    $svcName = $svc.PSChildName

    # Skip if core Windows service key
    if ($coreServices -contains $svcName) {
        Write-Log "[$svcName] Skipped (core Windows service key)."
        $coreSkippedCount++
        continue
    }

    # read ImagePath if present
    try {
        $props = Get-ItemProperty -Path $svc.PSPath -ErrorAction Stop
        $imagePath = $props.ImagePath
    } catch {
        $skippedCount++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($imagePath)) {
        $skippedCount++
        continue
    }

    $rawPath = $imagePath.Trim()
    $expandedPath = [Environment]::ExpandEnvironmentVariables($rawPath)

    # --- Parse exe/driver path robustly by matching known extensions ---
    # Attempt to capture extended NT prefixes (\??\ or \\?\) and everything up to the file extension.
    # This avoids splitting inside "Program Files" etc.
    $exePart = $expandedPath
    $argsPart = ""

    # If already quoted, extract inside quotes
    if ($expandedPath.StartsWith('"')) {
        if ($expandedPath -match '^"([^"]+)"(.*)$') {
            $exePart = $matches[1].Trim()
            $argsPart = $matches[2].Trim()
        }
    } else {
        # Try to match path up to a known extension (.exe .sys .dll .com .drv), capturing a possible NT prefix
        # Use case-insensitive match
        if ($expandedPath -match '(?i)^((?:\\\?\?\\|\\\\\?\\)?[^\r\n]*?\.(?:exe|sys|dll|com|drv))\b(.*)$') {
            $exePart = $matches[1].Trim()
            $argsPart = $matches[2].Trim()
        } else {
            # Fallback: split on first whitespace (older logic)
            if ($expandedPath -match '^([^\s]+)\s+(.*)$') {
                $exePart = $matches[1].Trim()
                $argsPart = $matches[2].Trim()
            } else {
                # no args, the whole expandedPath is exePart
                $exePart = $expandedPath.Trim()
                $argsPart = ""
            }
        }
    }

    # Normalize exePart lower for core checks
    $exePartLower = $exePart.ToLower()

    # Auto-skip core binaries by path: anything under %SystemRoot% (e.g. C:\Windows\ or C:\Windows\System32\)
    # Also skip well-known core binary filenames regardless of path.
    $isUnderSystem = $false
    try {
        if ($exePartLower.StartsWith($systemRootExpanded.ToLower())) {
            $isUnderSystem = $true
        } elseif ($exePartLower -match '\\windows\\' -or $exePartLower -match '\\system32\\' -or $exePartLower -match '\\syswow64\\') {
            $isUnderSystem = $true
        }
    } catch {
        $isUnderSystem = $false
    }

    $exeFileName = [IO.Path]::GetFileName($exePartLower)
    if ($coreBinaryNames -contains $exeFileName -or $isUnderSystem) {
        Write-Log "[$svcName] Skipped (core binary or under %SystemRoot%: $exePart)"
        $coreSkippedCount++
        continue
    }

    # --- Determine if quoting is needed: check raw or expanded exe part contains spaces ---
    $hasSpaces = ($rawPath -match '\s' -or $exePart -match '\s')
    if (-not $hasSpaces) {
        Write-Log "[$svcName] Skipped (no spaces detected: $exePart)"
        $skippedCount++
        continue
    }

    # --- Build corrected path: quote only the exePart (preserve args) ---
    $fixedPath = "`"$exePart`""
    if ($argsPart) { $fixedPath += " $argsPart" }

    # Only write/update if different from original raw registry string
    if ($fixedPath -ne $rawPath) {
        Write-Log "[$svcName] OLD: $rawPath"
        Write-Log "[$svcName] NEW: $fixedPath"

        if ($applyChanges) {
            try {
                Set-ItemProperty -Path $svc.PSPath -Name ImagePath -Value $fixedPath
                Write-Log "[$svcName] FIXED successfully."
                $fixedCount++
            } catch {
                Write-Log "[$svcName] ERROR: Failed to update - $($_.Exception.Message)"
                $skippedCount++
            }
        } else {
            Write-Log "[$svcName] Would fix (dry run)."
            $fixedCount++
        }
    } else {
        Write-Log "[$svcName] No change required (already quoted correctly)."
        $skippedCount++
    }
}

# --- Summary ---
Write-Log "`n=== SUMMARY REPORT ==="
Write-Log ("Total Entries Enumerated : {0}" -f $totalCount)
Write-Log ("Entries Fixed (or would fix) : {0}" -f $fixedCount)
Write-Log ("Entries Skipped (core Windows binaries or under %SystemRoot%) : {0}" -f $coreSkippedCount)
Write-Log ("Entries Skipped (no spaces or already correct) : {0}" -f $skippedCount)
Write-Log "Backup File: $backupPath"
Write-Log "Log File: $logPath"
Write-Log "Dry Run Mode: $(!$applyChanges)"
Write-Log "=== Completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
