<#
.SYNOPSIS
    Process leaver in hybrid environment (on-premises AD + Entra ID) by disabling account, converting mailbox to shared, removing licenses, clearing extension attributes, and moving to leavers OU.
.DESCRIPTION
    This script processes leaver users in hybrid environments by performing the following actions:
    - Imports user details from a CSV file
    - Disables user account in on-premises AD and Entra ID
    - Converts the user's mailbox to a shared mailbox
    - Removes all assigned licenses from Entra ID
    - Clears extension attributes (1-15)
    - Removes user from all Entra ID security groups
    - Moves disabled account to _Leavers OU
    - Logs all actions to a timestamped log file
.PARAMETER CsvPath
    Path to the CSV file containing user details. If not provided, defaults to "UserImport.csv" in the script directory.
.EXAMPLE
    .\Leaver_Script.ps1
    .\Leaver_Script.ps1 -CsvPath "C:\Path\To\UserImport.csv"
.NOTES
    Author: Peter James
    Date: December 2025
    Version: 2.0
    Requires: Active Directory PowerShell module, ExchangeOnlineManagement, Microsoft.Graph PowerShell SDK
    If you have any questions or need further assistance, please contact Peter James at peter.james@nexusos.co.uk
#>

param(
    [string]$CsvPath
)

try {
    if (-not $CsvPath) {
        $executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
        $CsvPath = Join-Path $executingScriptDirectory "UserImport.csv"
    }

    if (-not (Test-Path -Path $CsvPath)) {
        Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
        return
    }

    # Prepare log file
    $executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $logFile = Join-Path $executingScriptDirectory "Leaver_Script_$timestamp.log"

    function Write-Log {
        param([string]$Message, [ConsoleColor]$Color = 'White')
        $time = (Get-Date).ToString('s')
        $line = "[$time] $Message"
        Write-Host $Message -ForegroundColor $Color
        $line | Out-File -FilePath $logFile -Append -Encoding utf8
    }

    # Ensure required modules are available
    Write-Log "Checking required modules..." Cyan
    
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Log "ERROR: Active Directory module not found. Please ensure RSAT tools are installed." Red
        return
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Installing ExchangeOnlineManagement module..." Yellow
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log "Installing Microsoft.Graph module..." Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue

    # Connect to Exchange Online
    Write-Log "Connecting to Exchange Online..." Cyan
    try {
        Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        Write-Log "Connected to Exchange Online successfully" Green
    }
    catch {
        Write-Log "Warning: Could not connect to Exchange Online. Mailbox conversion will be skipped. $($_.Exception.Message)" Yellow
    }

    # Connect to Microsoft Graph
    Write-Log "Connecting to Microsoft Graph..." Cyan
    try {
        Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All','Mail.ReadWrite','Group.ReadWrite.All','MailboxSettings.ReadWrite','Organization.ReadWrite.All' -ErrorAction Stop | Out-Null
        Write-Log "Connected to Microsoft Graph successfully" Green
    }
    catch {
        Write-Log "Warning: Could not connect to Microsoft Graph. License and group operations will be skipped. $($_.Exception.Message)" Yellow
    }

    # Load the csv file
    Write-Log "Loading CSV file: $CsvPath" Cyan
    try {
        $users = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Log "CSV file loaded successfully - $($users.Count) user(s) found" Green
    }
    catch {
        Write-Log "Failed to import CSV file: $_" Red
        return
    }

    # Process each user
    foreach ($user in $users) {
        try {
            $userPrincipalName = $user.'User principal name'
            if (-not $userPrincipalName) {
                Write-Log "Skipping row: missing User principal name" Yellow
                continue
            }

            Write-Log "Processing user: $userPrincipalName" Cyan
            
            # Check if account status is specified and set to Disabled
            if ([string]::IsNullOrWhiteSpace($user.'Account status')) {
                Write-Log "  - Account status is empty for $userPrincipalName - skipping" Yellow
                continue
            }

            if ($user.'Account status' -ne 'Disabled') {
                Write-Log "  - Account status is not 'Disabled' for $userPrincipalName - skipping" Yellow
                continue
            }

            # Get AD User
            try {
                $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -ErrorAction Stop
                $samAccountName = $adUser.SamAccountName
                $distinguishedName = $adUser.DistinguishedName
            }
            catch {
                Write-Log "  - User not found in Active Directory: $userPrincipalName - $_" Red
                continue
            }

            # Disable in on-premises AD
            try {
                Set-ADUser $samAccountName -Enabled $false -ErrorAction Stop
                Write-Log "  - Disabled account in on-premises AD" Green
            }
            catch {
                Write-Log "  - Failed to disable AD account: $_" Red
            }

            # Convert mailbox to shared
            try {
                Set-Mailbox -Identity $userPrincipalName -Type Shared -ErrorAction Stop
                Write-Log "  - Converted mailbox to shared" Green
            }
            catch {
                Write-Log "  - Failed to convert mailbox to shared: $_" Yellow
            }

            # Clear extension attributes (1-15)
            try {
                Get-ADUser -Filter "UserPrincipalName -eq '$userPrincipalName'" -Properties extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15 | ForEach-Object {
                    for ($i = 1; $i -le 15; $i++) {
                        Set-ADUser $_ -Clear "extensionAttribute$i" -ErrorAction SilentlyContinue
                    }
                }
                Write-Log "  - Cleared extension attributes (1-15)" Cyan
            }
            catch {
                Write-Log "  - Failed to clear extension attributes: $_" Yellow
            }

            # Remove licenses via Graph API (if connected)
            try {
                $mgUser = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
                $assignedLicenses = Get-MgUserLicenseDetail -UserId $mgUser.Id -ErrorAction Stop

                if ($assignedLicenses) {
                    $removeLicenses = @()
                    foreach ($license in $assignedLicenses.SkuId) {
                        $removeLicenses += $license
                    }
                    
                    $assignBody = @{ addLicenses = @(); removeLicenses = $removeLicenses } | ConvertTo-Json -Depth 5
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($mgUser.Id)/assignLicense" -Body $assignBody -ErrorAction Stop | Out-Null
                    Write-Log "  - Removed $($removeLicenses.Count) license(s)" Cyan
                }
                else {
                    Write-Log "  - No licenses to remove" Cyan
                }
            }
            catch {
                Write-Log "  - Failed to remove licenses: $_" Yellow
            }

            # Remove from all groups (Graph API)
            try {
                $mgUser = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
                $groups = Get-MgUserMemberOf -UserId $mgUser.Id -All -ErrorAction Stop | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }

                if ($groups) {
                    $groupsRemoved = 0
                    foreach ($group in $groups) {
                        try {
                            $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members/$($mgUser.Id)/`$ref"
                            Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop | Out-Null
                            Write-Log "    - Removed from group: $($group.AdditionalProperties['displayName'])" Cyan
                            $groupsRemoved++
                        }
                        catch {
                            Write-Log "    - Failed to remove from group: $_" Yellow
                        }
                    }
                    Write-Log "  - Removed from $groupsRemoved group(s)" Cyan
                }
                else {
                    Write-Log "  - No groups to remove" Cyan
                }
            }
            catch {
                Write-Log "  - Failed to remove groups: $_" Yellow
            }

            # Move to _Leavers OU
            try {
                Move-ADObject -Identity $distinguishedName -TargetPath 'OU=_Leavers,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk' -ErrorAction Stop
                Write-Log "  - Moved to _Leavers OU" Green
            }
            catch {
                Write-Log "  - Failed to move to _Leavers OU: $_" Yellow
            }

            Write-Log "Successfully processed leaver: $userPrincipalName" Green
        }
        catch {
            Write-Log "Error processing user $($user.'User principal name'): $_" Red
        }
    }

    Write-Log "Leaver processing completed. Full log: $logFile" Green
}
catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
}

Pause
