# Check if the service exists
$service = Get-Service CagService -ErrorAction SilentlyContinue

if ($service) {
    Write-Output "Datto RMM Agent already installed on this device."
    exit 0  # Success exit code
} else {
    Write-Output "Datto RMM Agent not found on this device."
    exit 1  # Failure exit code
}