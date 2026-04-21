<#
This script is built towork with Datto RMM variables.

In the component editor, go to Variables and add:

CompanyID → String
TenantID → String
Secret → String
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$source = Invoke-RestMethod -Method Get -Uri 'https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows'
$destination = 'cybercnsagent.exe'

Invoke-WebRequest -Uri $source -OutFile $destination

if (-not $env:CompanyID -or -not $env:TenantID -or -not $env:Secret) {
    throw "Missing Datto RMM variable: CompanyID, TenantID, or Secret."
}

.\cybercnsagent.exe -c "$env:CompanyID" -e "$env:TenantID" -j "$env:Secret" -i