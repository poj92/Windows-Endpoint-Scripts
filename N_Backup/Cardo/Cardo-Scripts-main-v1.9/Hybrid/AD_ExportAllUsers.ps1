<#
.SYNOPSIS
    Export all users from on-premises Active Directory with all properties to a CSV file.
.DESCRIPTION
    This script exports all users from on-premises Active Directory with comprehensive properties including:
    - Basic identity (SAM Account Name, UPN, DisplayName, etc.)
    - Contact information (phone, email, address, city, etc.)
    - Organization (Department, Title, Company, Manager)
    - Account status and dates
    - Extension attributes (1-15)
.PARAMETER OutputPath
    Path where the CSV file will be saved. If not provided, defaults to "AD_Users_Export.csv" in the script directory.
.EXAMPLE
    .\AD_ExportAllUsers.ps1
    .\AD_ExportAllUsers.ps1 -OutputPath "C:\Exports\AllADUsers.csv"
.NOTES
    Author: Peter James
    Date: January 2026
    Version: 1.0
    Requires: Active Directory PowerShell module (RSAT tools)
    Run on: Domain Controller or any machine with RSAT tools installed
    If you have any questions or need further assistance, please contact Peter James at peter.james@nexusos.co.uk
#>

param(
    [string]$OutputPath
)

try {
    if (-not $OutputPath) {
        $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
        $OutputPath = Join-Path $scriptDir "AD_Users_Export.csv"
    }

    # Prepare log file
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $logFile = Join-Path $scriptDir "AD_ExportAllUsers_$timestamp.log"

    function Write-Log {
        param([string]$Message, [ConsoleColor]$Color = 'White')
        $time = (Get-Date).ToString('s')
        $line = "[$time] $Message"
        Write-Host $Message -ForegroundColor $Color
        $line | Out-File -FilePath $logFile -Append -Encoding utf8
    }

    # Ensure Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Log "ERROR: Active Directory module not found. Please ensure RSAT tools are installed." Red
        Write-Log "  Install RSAT: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" Yellow
        return
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "Active Directory module imported successfully" Green

    # Get the domain
    $domain = (Get-ADDomain -ErrorAction Stop).Name
    Write-Log "Connected to domain: $domain" Green

    # Get all users from Active Directory
    Write-Log "Retrieving all users from Active Directory..." Cyan
    try {
        $users = Get-ADUser -Filter * -Properties @(
            'SamAccountName',
            'UserPrincipalName',
            'DisplayName',
            'GivenName',
            'Surname',
            'Mail',
            'MailNickname',
            'Mobile',
            'TelephoneNumber',
            'Office',
            'OfficeLocation',
            'StreetAddress',
            'City',
            'State',
            'PostalCode',
            'Country',
            'Department',
            'Title',
            'Company',
            'Manager',
            'Enabled',
            'Created',
            'LastLogonDate',
            'PasswordLastSet',
            'DistinguishedName',
            'ObjectGUID',
            'UserAccountControl',
            'Description',
            'EmployeeID',
            'EmployeeNumber',
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

        Write-Log "Retrieved $($users.Count) user(s) from Active Directory" Green
    }
    catch {
        Write-Log "Failed to retrieve users: $($_.Exception.Message)" Red
        return
    }

    # Build export data
    Write-Log "Processing user data for export..." Cyan
    $exportData = @()
    $processedCount = 0

    foreach ($user in $users) {
        try {
            # Get manager name if manager exists
            $managerName = $null
            if ($user.Manager) {
                try {
                    $manager = Get-ADUser -Identity $user.Manager -Properties DisplayName -ErrorAction Stop
                    $managerName = $manager.DisplayName
                }
                catch {
                    $managerName = "Error retrieving"
                }
            }

            # Determine if account is locked out
            $isLocked = if ($user.UserAccountControl -band 16) { $true } else { $false }

            # Build row object
            $rowData = [PSCustomObject]@{
                'SAM Account Name'              = $user.SamAccountName
                'User Principal Name'           = $user.UserPrincipalName
                'Display Name'                  = $user.DisplayName
                'First Name'                    = $user.GivenName
                'Last Name'                     = $user.Surname
                'Email'                         = $user.Mail
                'Mail Nickname'                 = $user.MailNickname
                'Mobile Phone'                  = $user.Mobile
                'Office Phone'                  = $user.TelephoneNumber
                'Office Location'               = $user.OfficeLocation
                'Street Address'                = $user.StreetAddress
                'City'                          = $user.City
                'State/Province'                = $user.State
                'Postal Code'                   = $user.PostalCode
                'Country'                       = $user.Country
                'Department'                    = $user.Department
                'Job Title'                     = $user.Title
                'Company'                       = $user.Company
                'Manager'                       = $managerName
                'Description'                   = $user.Description
                'Employee ID'                   = $user.EmployeeID
                'Employee Number'               = $user.EmployeeNumber
                'Account Enabled'               = $user.Enabled
                'Account Locked Out'            = $isLocked
                'Created Date'                  = $user.Created
                'Password Last Set'             = $user.PasswordLastSet
                'Last Logon Date'               = $user.LastLogonDate
                'Distinguished Name'            = $user.DistinguishedName
                'Extension Attribute 1'         = $user.extensionAttribute1
                'Extension Attribute 2'         = $user.extensionAttribute2
                'Extension Attribute 3'         = $user.extensionAttribute3
                'Extension Attribute 4'         = $user.extensionAttribute4
                'Extension Attribute 5'         = $user.extensionAttribute5
                'Extension Attribute 6'         = $user.extensionAttribute6
                'Extension Attribute 7'         = $user.extensionAttribute7
                'Extension Attribute 8'         = $user.extensionAttribute8
                'Extension Attribute 9'         = $user.extensionAttribute9
                'Extension Attribute 10'        = $user.extensionAttribute10
                'Extension Attribute 11'        = $user.extensionAttribute11
                'Extension Attribute 12'        = $user.extensionAttribute12
                'Extension Attribute 13'        = $user.extensionAttribute13
                'Extension Attribute 14'        = $user.extensionAttribute14
                'Extension Attribute 15'        = $user.extensionAttribute15
                'User ID (GUID)'                = $user.ObjectGUID
            }

            $exportData += $rowData
            $processedCount++

            if ($processedCount % 100 -eq 0) {
                Write-Log "  Processed $processedCount user(s)..." Cyan
            }
        }
        catch {
            Write-Log "  Warning: Error processing user $($user.SamAccountName): $_" Yellow
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

    Write-Log "Active Directory user export completed. Log file: $logFile" Green
    Write-Host "`nExport file: $OutputPath" -ForegroundColor Green
    Write-Host "Log file: $logFile" -ForegroundColor Green
}
catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
}

Pause
