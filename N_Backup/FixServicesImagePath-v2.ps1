<#
.SYNOPSIS
  Quote ImagePath entries containing spaces (handles NT prefixes, systemroot paths, skips only core binaries).

  No change by default — set $applyChanges = $true to apply changes.
#>

$timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath        = "C:\ServicePathFix_$timestamp.log"
$backupPath     = "C:\Services_Backup_$timestamp.reg"
$applyChanges   = $false    # Set to $true to apply changes

# --- Core binary file names to skip (lowercase) ---
$coreBinaryNames = @(
    "svchost.exe","wininit.exe","winlogon.exe","lsass.exe","services.exe",
    "smss.exe","csrss.exe","spoolsv.exe","win32k.sys","ntoskrnl.exe"
)

# --- Core Windows service names to skip ---
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

Write-Log "=== Starting Universal Service Path Quote Fix (handles %SystemRoot%, skips only core binaries) ==="
Write-Log "Log file: $logPath"
Write-Log "Backup file: $backupPath"
Write-Log "Dry run mode: $(!$applyChanges)"
Write-Log ""

# --- Create registry backup ---
try {
    Write-Log "Creating backup of HKLM\SYSTEM\CurrentControlSet\Services ..."
    reg export "HKLM\SYSTEM\CurrentControlSet\Services" $backupPath /y | Out-Null
    Write-Log "Backup completed successfully.`n"
} catch {
    Write-Log "ERROR: Failed to create registry backup. $_"
    exit 1
}

$servicesPath = "HKLM:\SYSTEM\CurrentControlSet\Services"
$services = Get-ChildItem $servicesPath

foreach ($svc in $services) {
    $totalCount++
    $svcName = $svc.PSChildName

    # Skip known critical Windows service keys
    if ($coreServices -contains $svcName) {
        Write-Log "[$svcName] Skipped (core Windows service name)"
        $coreSkippedCount++
        continue
    }

    # Retrieve ImagePath
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

    # --- Parse executable/driver path robustly (handles \??\ and \\?\ prefixes) ---
    $exePart = $expandedPath
    $argsPart = ""

    if ($expandedPath.StartsWith('"')) {
        if ($expandedPath -match '^"([^"]+)"(.*)$') {
            $exePart = $matches[1].Trim()
            $argsPart = $matches[2].Trim()
        }
    } else {
        if ($expandedPath -match '(?i)^((?:\\\?\?\\|\\\\\?\\)?[^\r\n]*?\.(?:exe|sys|dll|com|drv))\b(.*)$') {
            $exePart = $matches[1].Trim()
            $argsPart = $matches[2].Trim()
        } elseif ($expandedPath -match '^([^\s]+)\s+(.*)$') {
            $exePart = $matches[1].Trim()
            $argsPart = $matches[2].Trim()
        }
    }

    $exeFileName = [IO.Path]::GetFileName($exePart.ToLower())

    # --- Skip only true core binaries ---
    if ($coreBinaryNames -contains $exeFileName) {
        Write-Log "[$svcName] Skipped (core binary: $exeFileName)"
        $coreSkippedCount++
        continue
    }

    # --- Determine if path has spaces ---
    $hasSpaces = ($rawPath -match '\s' -or $exePart -match '\s')
    if (-not $hasSpaces) {
        Write-Log "[$svcName] Skipped (no spaces detected: $exePart)"
        $skippedCount++
        continue
    }

    # --- Build properly quoted path ---
    $fixedPath = "`"$exePart`""
    if ($argsPart) { $fixedPath += " $argsPart" }

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
        Write-Log "[$svcName] No change required (already correct)."
        $skippedCount++
    }
}

# --- Summary ---
Write-Log "`n=== SUMMARY REPORT ==="
Write-Log ("Total Entries Enumerated : {0}" -f $totalCount)
Write-Log ("Entries Fixed : {0}" -f $fixedCount)
Write-Log ("Entries Skipped (core binaries only) : {0}" -f $coreSkippedCount)
Write-Log ("Entries Skipped (no spaces or already correct) : {0}" -f $skippedCount)
Write-Log "Backup File: $backupPath"
Write-Log "Log File: $logPath"
Write-Log "No change mode: $(!$applyChanges)"
Write-Log "=== Completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
