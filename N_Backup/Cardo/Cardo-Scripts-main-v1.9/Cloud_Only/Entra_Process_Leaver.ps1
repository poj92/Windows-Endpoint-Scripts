<#
.SYNOPSIS
    Process leaver in Entra ID by disabling user account, converting mailbox to shared, removing licenses, and logging actions.
.DESCRIPTION
    This script processes leaver users in Entra ID by performing the following actions:
    - Imports user details from a CSV file, automatically uses "UserImport.csv" in the script directory if not specified
    - Disables the user account
    - Converts the user's mailbox to a shared mailbox
    - Removes all assigned licenses from the user
    - Logs all actions taken to a timestamped log file
.PARAMETER CsvPath
    Path to the CSV file containing user details. If not provided, defaults to "UserImport.csv" in the script directory.
.EXAMPLE
    .\Entra_Process_Leaver.ps1
    .\Entra_Process_Leaver.ps1 -CsvPath "C:\Path\To\UserImport.csv"
.NOTES
    Author: Peter James
    Date: December 2025
    Version: 1.0
    Requires: Microsoft.Graph PowerShell SDK
    If you have any questions or need further assistance, please contact Peter James at peter.james@nexusos.co.uk
#>

param(
    [string]$CsvPath
)

try {
    if (-not $CsvPath) {
        $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
        $CsvPath = Join-Path $scriptDir "UserImport.csv"
    }

    if (-not (Test-Path -Path $CsvPath)) {
        Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
        return
    }

    # Prepare log file
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $logFile = Join-Path $scriptDir "Entra_Process_Leaver_$timestamp.log"

    function Write-Log {
        param([string]$Message, [ConsoleColor]$Color = 'White')
        $time = (Get-Date).ToString('s')
        $line = "[$time] $Message"
        Write-Host $Message -ForegroundColor $Color
        $line | Out-File -FilePath $logFile -Append -Encoding utf8
    }

    # Ensure Microsoft Graph module is available
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host "Microsoft.Graph module not found. Installing for current user..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }

    # Ensure ExchangeOnlineManagement module is available for mailbox conversion
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found. Installing for current user..." -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }

    # Import ExchangeOnlineManagement module
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

    # Connect to Microsoft Graph
    try {
        Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All','Mail.ReadWrite','Group.ReadWrite.All','MailboxSettings.ReadWrite','Organization.ReadWrite.All' | Out-Null
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    # Connect to Exchange Online for mailbox conversion
    $exchangeConnected = $false
    try {
        Write-Log "Connecting to Exchange Online..." Cyan
        Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        $exchangeConnected = $true
        Write-Log "Connected to Exchange Online successfully" Green
    }
    catch {
        Write-Log "Warning: Could not connect to Exchange Online. Mailbox conversion will be skipped. $($_.Exception.Message)" Yellow
    }

    # Import user data from CSV file
    try {
        $users = Import-Csv -Path $CsvPath -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to import CSV file at $CsvPath. Please ensure the file exists and is accessible. $_" Red
        return
    }
    # Process each user in the CSV
    foreach ($user in $users) {
        try {
            $userPrincipalName = $user.'User principal name'
            if (-not $userPrincipalName) {
                Write-Log "Skipping row: missing User principal name." Yellow
                continue
            }

            # Get the UserID of the user
            try {
                $userId = (Get-MgUser -UserId $userPrincipalName -ErrorAction Stop).Id
            }
            catch {
                Write-Log "Failed to find user $userPrincipalName - $($_.Exception.Message)" Red
                continue
            }

            # Disable the user account
            try {
                Update-MgUser -UserId $userId -AccountEnabled:$false -ErrorAction Stop
                Write-Log "Disabled account for user: $userPrincipalName" Green
            }
            catch {
                Write-Log "Failed to disable account for $userPrincipalName - $($_.Exception.Message)" Red
                continue
            }

            # Convert mailbox to shared mailbox
            if ($exchangeConnected) {
                try {
                    Set-Mailbox -Identity $userPrincipalName -Type Shared -ErrorAction Stop
                    Write-Log "Converted mailbox to shared for user: $userPrincipalName" Green
                }
                catch {
                    Write-Log "Failed to convert mailbox to shared for $userPrincipalName - $($_.Exception.Message)" Yellow
                }
            }
            else {
                Write-Log "Skipped mailbox conversion for $userPrincipalName - Exchange Online not connected" Yellow
            }

            # Delete entries in extension attributes 1-15 for the user
            try {
                for ($i = 1; $i -le 15; $i++) {
                    $extensionAttribute = "extensionAttribute$i"
                    Update-MgUser -UserId $userId -AdditionalProperties @{$extensionAttribute = $null} -ErrorAction Stop
                }
                Write-Log "Cleared extension attributes (1-15) for user: $userPrincipalName" Cyan
            }
            catch {
                Write-Log "Failed to clear extension attributes for $userPrincipalName - $($_.Exception.Message)" Yellow
            }

            # Remove user from all groups
            try {
                $groups = Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop | Select-Object Id, DisplayName
                
                if ($groups) {
                    $groupsRemoved = 0
                    $groupsFailed = 0
                    
                    foreach ($group in $groups) {
                        try {
                            # Use Graph API directly to remove member from group
                            $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members/$userId/`$ref"
                            Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop | Out-Null
                            Write-Log "  - Removed from group: $($group.DisplayName)" Cyan
                            $groupsRemoved++
                        }
                        catch {
                            Write-Log "  - Failed to remove from group $($group.DisplayName) - $($_.Exception.Message)" Yellow
                            $groupsFailed++
                        }
                    }
                    Write-Log "Group removal: Removed from $groupsRemoved group(s), Failed: $groupsFailed" Cyan
                }
                else {
                    Write-Log "No groups found for user: $userPrincipalName" Yellow
                }
            }
            catch {
                Write-Log "Failed to retrieve groups for $userPrincipalName - $($_.Exception.Message)" Yellow
            }

            # Remove all assigned licenses
            try {
                $assignedLicenses = (Get-MgUserLicenseDetail -UserId $userId -ErrorAction Stop).SkuId
                if ($assignedLicenses) {
                    $removeLicenses = @()
                    foreach ($license in $assignedLicenses) {
                        $removeLicenses += $license
                    }
                    
                    $assignBody = @{ addLicenses = @(); removeLicenses = $removeLicenses } | ConvertTo-Json -Depth 5
                    try {
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$userId/assignLicense" -Body $assignBody -ErrorAction Stop
                        Write-Log "  - Removed all licenses ($($removeLicenses.Count) license(s))" Cyan
                    }
                    catch {
                        Write-Log "  - Failed to remove licenses - $($_.Exception.Message)" Yellow
                    }
                } else {
                    Write-Log "No direct licenses found for user: $userPrincipalName" Yellow
                }
            }
            catch {
                Write-Log "Failed to retrieve licenses for $userPrincipalName - $($_.Exception.Message)" Yellow
            }

            # Delete the user (optional, uncomment if needed)
            #try {
            #    Remove-MgUser -UserId $userId -ErrorAction Stop
            #    Write-Log "Deleted user: $userPrincipalName" Green
            #}
            #catch {
            #    Write-Log "Failed to delete user $userPrincipalName - $($_.Exception.Message)" Red
            #}

        }
        catch {
            Write-Log "Error processing user $($user.'User logon name') - $($_.Exception.Message)" Red
        }
    }

    Write-Log "Leaver processing completed. Actions logged to $logFile" Green
}
catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
}

Pause