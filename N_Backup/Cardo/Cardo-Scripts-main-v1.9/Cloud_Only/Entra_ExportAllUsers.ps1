<#
.SYNOPSIS
    Export all users from Entra ID with all properties to a CSV file.
.DESCRIPTION
    This script exports all users from Entra ID with comprehensive properties including:
    - Basic identity (UPN, DisplayName, Mail, etc.)
    - Contact information (phone, address, city, etc.)
    - Organization (Department, JobTitle, Company, Manager)
    - License and group information
    - Account status and creation date
.PARAMETER OutputPath
    Path where the CSV file will be saved. If not provided, defaults to "EntraUsers_Export.csv" in the script directory.
.EXAMPLE
    .\Entra_ExportAllUsers.ps1
    .\Entra_ExportAllUsers.ps1 -OutputPath "C:\Exports\AllUsers.csv"
.NOTES
    Author: Peter James
    Date: January 2026
    Version: 1.0
    Requires: Microsoft.Graph PowerShell SDK
    If you have any questions or need further assistance, please contact Peter James at peter.james@nexusos.co.uk
#>

param(
    [string]$OutputPath
)

try {
    if (-not $OutputPath) {
        $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
        $OutputPath = Join-Path $scriptDir "EntraUsers_Export.csv"
    }

    # Prepare log file
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $logFile = Join-Path $scriptDir "Entra_ExportAllUsers_$timestamp.log"

    function Write-Log {
        param([string]$Message, [ConsoleColor]$Color = 'White')
        $time = (Get-Date).ToString('s')
        $line = "[$time] $Message"
        Write-Host $Message -ForegroundColor $Color
        $line | Out-File -FilePath $logFile -Append -Encoding utf8
    }

    # Ensure Microsoft Graph module is available
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Log "Microsoft.Graph.Users module not found. Installing for current user..." Yellow
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    # Connect to Microsoft Graph
    Write-Log "Connecting to Microsoft Graph..." Cyan
    try {
        Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All' | Out-Null
        Write-Log "Connected to Microsoft Graph successfully" Green
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" Red
        return
    }

    # Get all users with selected properties
    Write-Log "Retrieving all users from Entra ID..." Cyan
    try {
        $users = Get-MgUser -All -Property @(
            'id',
            'userPrincipalName',
            'displayName',
            'givenName',
            'surname',
            'mail',
            'mailNickname',
            'mobilePhone',
            'businessPhones',
            'officeLocation',
            'city',
            'state',
            'postalCode',
            'streetAddress',
            'country',
            'department',
            'jobTitle',
            'companyName',
            'manager',
            'accountEnabled',
            'createdDateTime',
            'lastPasswordChangeDateTime',
            'onPremisesDistinguishedName',
            'onPremisesImmutableId',
            'userType',
            'employeeId',
            'employeeOrgManager',
            'extensionAttribute1',
            'extensionAttribute2',
            'extensionAttribute3',
            'extensionAttribute4',
            'extensionAttribute5',
            'extensionAttribute6',
            'extensionAttribute7',
            'extensionAttribute8',
            'extensionAttribute9',
            'extensionAttribute10',
            'extensionAttribute11',
            'extensionAttribute12',
            'extensionAttribute13',
            'extensionAttribute14',
            'extensionAttribute15'
        ) -ErrorAction Stop

        Write-Log "Retrieved $($users.Count) user(s) from Entra ID" Green
    }
    catch {
        Write-Log "Failed to retrieve users: $($_.Exception.Message)" Red
        return
    }

    # Build export data with all properties
    Write-Log "Processing user data for export..." Cyan
    $exportData = @()
    $processedCount = 0

    foreach ($user in $users) {
        try {
            # Get manager name if manager exists
            $managerName = $null
            if ($user.Manager) {
                try {
                    $manager = Get-MgDirectoryObject -DirectoryObjectId $user.Manager.Id -ErrorAction Stop
                    $managerName = if ($manager.AdditionalProperties['displayName']) { $manager.AdditionalProperties['displayName'] } else { "Unknown" }
                }
                catch {
                    $managerName = "Error retrieving"
                }
            }

            # Get licenses
            $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop
            $licenses = if ($licenseDetails) { ($licenseDetails.SkuPartNumber -join '; ') } else { "None" }

            # Get group memberships
            $groups = Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction Stop | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
            $groupNames = if ($groups) { ($groups.AdditionalProperties['displayName'] -join '; ') } else { "None" }

            # Build row object
            $rowData = [PSCustomObject]@{
                'User Principal Name'           = $user.UserPrincipalName
                'Display Name'                  = $user.DisplayName
                'First Name'                    = $user.GivenName
                'Last Name'                     = $user.Surname
                'Email'                         = $user.Mail
                'Mail Nickname'                 = $user.MailNickname
                'Mobile Phone'                  = $user.MobilePhone
                'Office Phone'                  = if ($user.BusinessPhones) { $user.BusinessPhones[0] } else { $null }
                'Office Location'               = $user.OfficeLocation
                'Street Address'                = $user.StreetAddress
                'City'                          = $user.City
                'State/Province'                = $user.State
                'Postal Code'                   = $user.PostalCode
                'Country'                       = $user.Country
                'Department'                    = $user.Department
                'Job Title'                     = $user.JobTitle
                'Company Name'                  = $user.CompanyName
                'Manager'                       = $managerName
                'Account Enabled'               = $user.AccountEnabled
                'User Type'                     = $user.UserType
                'Employee ID'                   = $user.EmployeeId
                'Created Date'                  = $user.CreatedDateTime
                'Last Password Change'          = $user.LastPasswordChangeDateTime
                'On-Premises DN'                = $user.OnPremisesDistinguishedName
                'Extension Attribute 1'         = $user.AdditionalProperties['extensionAttribute1']
                'Extension Attribute 2'         = $user.AdditionalProperties['extensionAttribute2']
                'Extension Attribute 3'         = $user.AdditionalProperties['extensionAttribute3']
                'Extension Attribute 4'         = $user.AdditionalProperties['extensionAttribute4']
                'Extension Attribute 5'         = $user.AdditionalProperties['extensionAttribute5']
                'Extension Attribute 6'         = $user.AdditionalProperties['extensionAttribute6']
                'Extension Attribute 7'         = $user.AdditionalProperties['extensionAttribute7']
                'Extension Attribute 8'         = $user.AdditionalProperties['extensionAttribute8']
                'Extension Attribute 9'         = $user.AdditionalProperties['extensionAttribute9']
                'Extension Attribute 10'        = $user.AdditionalProperties['extensionAttribute10']
                'Extension Attribute 11'        = $user.AdditionalProperties['extensionAttribute11']
                'Extension Attribute 12'        = $user.AdditionalProperties['extensionAttribute12']
                'Extension Attribute 13'        = $user.AdditionalProperties['extensionAttribute13']
                'Extension Attribute 14'        = $user.AdditionalProperties['extensionAttribute14']
                'Extension Attribute 15'        = $user.AdditionalProperties['extensionAttribute15']
                'Licenses'                      = $licenses
                'Groups'                        = $groupNames
                'User ID'                       = $user.Id
            }

            $exportData += $rowData
            $processedCount++

            if ($processedCount % 50 -eq 0) {
                Write-Log "  Processed $processedCount user(s)..." Cyan
            }
        }
        catch {
            Write-Log "  Warning: Error processing user $($user.UserPrincipalName): $_" Yellow
        }
    }

    # Export to CSV
    Write-Log "Exporting $($exportData.Count) user(s) to CSV..." Cyan
    try {
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "Export completed successfully: $OutputPath" Green
        Write-Log "Total users exported: $($exportData.Count)" Green
    }
    catch {
        Write-Log "Failed to export CSV: $($_.Exception.Message)" Red
        return
    }

    Write-Log "User export completed. Log file: $logFile" Green
    Write-Host "`nExport file: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
}

Pause
