<#
    This PowerShell script checks whether the ConnectSecure agent is installed and if not,
    it will run the installer from the same directory as this script.
#>

# Check if ConnectSecure agent is installed
$agentPath = "C:\Program Files (x86)\CyberCNSAgent\cybercnsagent.exe"
if (Test-Path $agentPath) {
    Write-Host "ConnectSecure agent is already installed."
} else {
    Write-Host "ConnectSecure agent is not installed. Starting installation..."
    $installerPath = Join-Path $PSScriptRoot "ConnectSecure-Install.ps1"
    & $installerPath
    if (Test-Path $agentPath) {
        Write-Host "ConnectSecure agent installation completed successfully."
    } else {
        Write-Host "ConnectSecure agent installation failed. Please check the installation script for errors."
    }
}
