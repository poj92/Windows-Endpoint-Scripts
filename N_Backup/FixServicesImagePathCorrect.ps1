# Define base registry path
$basePath = "HKLM:\SYSTEM\CurrentControlSet\Services"

# Define log file path
$logDir = "C:\Logs"
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory | Out-Null
}

$logFile = Join-Path $logDir ("ServiceImagePathChanges_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

# Function to write log entries
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "Gray"
    )
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $logFile -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}

Write-Log "=== Script started at $(Get-Date) ===" "Cyan"

# Common core Windows service executables and paths to skip
$coreExecutables = @(
    "svchost.exe",
    "lsass.exe",
    "services.exe",
    "wininit.exe",
    "smss.exe",
    "winlogon.exe",
    "spoolsv.exe",
    "csrss.exe",
    "dwm.exe",
    "taskhostw.exe"
)

$corePaths = @(
    "$env:SystemRoot\System32",
    "$env:SystemRoot\SysWOW64"
)

# Get all subkeys under the Services key
Get-ChildItem -Path $basePath | ForEach-Object {
    $serviceKey = $_.PSPath
    $serviceName = $_.PSChildName

    try {
        $imagePath = (Get-ItemProperty -Path $serviceKey -Name ImagePath -ErrorAction Stop).ImagePath

        if ($null -ne $imagePath -and $imagePath -match '\s') {
            # Skip already quoted
            if ($imagePath -notmatch '^".*"$') {

                # Split into exe path and parameters
                $exePath, $params = $imagePath -split '\s+', 2
                $exeName = Split-Path $exePath -Leaf

                # Check if it's a core Windows service
                $isCore = $false
                foreach ($core in $coreExecutables) {
                    if ($exeName -ieq $core) { $isCore = $true; break }
                }
                foreach ($coreDir in $corePaths) {
                    if ($exePath -like "$coreDir*") { $isCore = $true; break }
                }

                if ($isCore) {
                    Write-Log "Skipping core Windows service '$serviceName' ($exePath)" "DarkGray"
                    return
                }

                # If looks like a valid path, quote only the path portion
                if ($exePath -match '^[A-Za-z]:\\' -or $exePath -match '^%[A-Za-z_]+%\\') {
                    $newPath = '"' + $exePath + '"'
                    if ($params) { $newPath += " $params" }

                    Write-Log "Updating service '$serviceName':" "Yellow"
                    Write-Log "  Old: $imagePath"
                    Write-Log "  New: $newPath" "Green"

                    # Update the registry
                    Set-ItemProperty -Path $serviceKey -Name ImagePath -Value $newPath
                } else {
                    Write-Log "Skipping '$serviceName' (does not look like a valid path): $imagePath" "DarkGray"
                }
            } else {
                Write-Log "Service '$serviceName' already quoted, skipping." "DarkGray"
            }
        } else {
            Write-Log "Service '$serviceName' has no space in ImagePath or no ImagePath." "DarkGray"
        }
    } catch {
        Write-Log "Skipping '$serviceName': $($_.Exception.Message)" "Red"
    }
}

Write-Log "=== Script completed at $(Get-Date) ===" "Cyan"
Write-Host "`nLog saved to: $logFile" -ForegroundColor Cyan
