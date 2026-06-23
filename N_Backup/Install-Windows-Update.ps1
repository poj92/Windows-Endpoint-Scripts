<#
.SYNOPSIS
    Installs Windows Update using PSWindowsUpdate.
.DESCRIPTION
    This script:
    1. Checks for the PSWindowsUpdate module.
    2. Installs it if missing.
    3. Imports the module.
    4. Installs the specified KB update automatically.
#>

<# Ensure script is running as Administrator
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    Exit
}
#>

# Define the KB number
$KBID = "KB5070773"

# Check if PSWindowsUpdate module is installed
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "PSWindowsUpdate module not found. Installing..." -ForegroundColor Yellow
    
    # Make sure PowerShell Gallery is trusted
    $repo = Get-PSRepository -Name "PSGallery" -ErrorAction SilentlyContinue
    if ($null -eq $repo) {
        Write-Host "Registering PowerShell Gallery repository..." -ForegroundColor Cyan
        Register-PSRepository -Default
    }

    if ($repo.InstallationPolicy -ne 'Trusted') {
        Write-Host "Setting PSGallery as trusted source..." -ForegroundColor Cyan
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    }

    # Install the module
    Install-Module -Name PSWindowsUpdate -Force
    Write-Host "PSWindowsUpdate module installed successfully." -ForegroundColor Green
} else {
    Write-Host "PSWindowsUpdate module is already installed." -ForegroundColor Green
}

# Import the module
Import-Module PSWindowsUpdate -Force

# Install the specified KB update
Write-Host "Searching and installing update $KBID..." -ForegroundColor Cyan
Get-WUInstall -KBArticleID $KBID -AcceptAll -Verbose -IgnoreReboot

Write-Host "Update $KBID installation process completed. Check Windows Update History for confirmation." -ForegroundColor Green
