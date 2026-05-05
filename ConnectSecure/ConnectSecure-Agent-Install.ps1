<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-05
Email: Peter.James@nexusos.co.uk
#>

<#
This script is built towork with Datto RMM variables.

In the component editor, go to Variables and add:

CompanyID → String
TenantID → String
Secret → String
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$agentPath   = "C:\Program Files (x86)\CyberCNSAgent\cybercnsagent.exe"
$CompanyID   = "$env:CompanyID".Trim()
$TenantID    = "$env:TenantID".Trim()
$Secret      = "$env:Secret".Trim()

if (Test-Path $agentPath) {
    Write-Host "ConnectSecure agent is already installed."
    exit 0
}

if ([string]::IsNullOrWhiteSpace($CompanyID)) { throw "CompanyID is empty." }
if ([string]::IsNullOrWhiteSpace($TenantID))  { throw "TenantID is empty." }
if ([string]::IsNullOrWhiteSpace($Secret))    { throw "Secret is empty." }

Write-Host "ConnectSecure agent is not installed. Starting installation..."

$source = Invoke-RestMethod -Method Get -Uri 'https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows'
$destination = Join-Path $env:TEMP 'cybercnsagent.exe'

Invoke-WebRequest -Uri $source -OutFile $destination

if (-not (Test-Path $destination)) {
    throw "Failed to download cybercnsagent.exe"
}

$proc = Start-Process -FilePath $destination -ArgumentList @(
    '-c', $CompanyID,
    '-e', $TenantID,
    '-j', $Secret,
    '-i'
) -Wait -PassThru -NoNewWindow

Write-Host "Installer exit code: $($proc.ExitCode)"

Start-Sleep -Seconds 5

if (Test-Path $agentPath) {
    Write-Host "ConnectSecure agent installation completed successfully."
    exit 0
} else {
    throw "ConnectSecure agent installation failed. Please check the installation logs/output."
}