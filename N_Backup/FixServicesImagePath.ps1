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

# Get all subkeys under the Services key
Get-ChildItem -Path $basePath | ForEach-Object {
    $serviceKey = $_.PSPath
    $serviceName = $_.PSChildName

    try {
        # Attempt to get the ImagePath value
        $imagePath = (Get-ItemProperty -Path $serviceKey -Name ImagePath -ErrorAction Stop).ImagePath

        if ($null -ne $imagePath -and $imagePath -match '\s') {
            # Check if it's already wrapped in quotes
            if ($imagePath -notmatch '^".*"$') {
                $newPath = '"' + $imagePath + '"'
                Write-Log "Updating service '$serviceName':" "Yellow"
                Write-Log "  Old: $imagePath"
                Write-Log "  New: $newPath" "Green"

                # Update the registry
                Set-ItemProperty -Path $serviceKey -Name ImagePath -Value $newPath
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
