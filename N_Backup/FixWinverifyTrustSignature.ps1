# Ensure running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host " Please run this script as Administrator." -ForegroundColor Red
    exit
}

# Define registry paths
$paths = @(
    "HKLM:\Software\Microsoft\Cryptography\Wintrust\Config",
    "HKLM:\Software\Wow6432Node\Microsoft\Cryptography\Wintrust\Config"
)

# Desired value name and data
$valueName = "EnableCertPaddingCheck"
$valueData = 1

foreach ($path in $paths) {
    try {
        # Create the key if it doesn't exist
        if (-not (Test-Path $path)) {
            Write-Host "Creating missing key: $path" -ForegroundColor Yellow
            New-Item -Path $path -Force | Out-Null
        }

        # Set the DWORD value
        Write-Host "Setting $valueName=1 at $path" -ForegroundColor Cyan
        Set-ItemProperty -Path $path -Name $valueName -Value $valueData -Type DWord

        # Verify result
        $currentValue = (Get-ItemProperty -Path $path -Name $valueName).$valueName
        Write-Host "✔ $path\$valueName = $currentValue" -ForegroundColor Green
    }
    catch {
        Write-Host " Failed to update $path: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n Completed. Reboot may be required for the setting to take effect." -ForegroundColor Cyan
