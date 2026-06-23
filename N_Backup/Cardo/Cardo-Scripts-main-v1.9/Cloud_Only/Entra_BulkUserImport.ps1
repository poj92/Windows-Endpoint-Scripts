<#
    Entra_BulkCreateUsers.ps1
    Bulk create Entra (Azure AD) users from a CSV using Microsoft Graph PowerShell

    Usage:
      - Place a `UserImport.csv` next to this script (same column names as existing Onprem import).
      - Run from PowerShell with an account that has `User.ReadWrite.All` and `Directory.ReadWrite.All` consent.
      - The script will call `Connect-MgGraph` to authenticate.

    Notes:
      - Requires the Microsoft Graph PowerShell SDK (install-module Microsoft.Graph -Scope CurrentUser).
      - The script uses `Invoke-MgGraphRequest` for some operations (manager and extension attributes).
      - Test in a non-production tenant first.

    CSV columns expected (examples from existing CSV):
      First name,Last name,Display name,User logon name,User principal name,Password,Account status,Job Title,Department,Company,Telephone number,Mobile,E-mail,Manager Email,IsFieldOperative,D365FieldService
    CSV columns expected (examples from existing CSV):
      First name,Last name,Display name,User logon name,User principal name,Password,Account status,Job Title,Department,Company,Telephone number,Mobile,E-mail,Manager Email,IsFieldOperative,D365FieldService,Licenses,Groups

    Additional optional CSV columns:
    - `Licenses`: optional, comma-separated skuPartNumber values (e.g. "OFFICE365_BUSINESS_PREMIUM") or skuId GUIDs.
    - `Groups`: optional, comma-separated group display names to add the user to after creation.

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

    # Ensure Microsoft Graph module is available
        # Prepare log file
        $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
        $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $logFile = Join-Path $scriptDir "Entra_BulkCreateUsers_$timestamp.log"

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

    # Connect to Microsoft Graph
    try {
        Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All' | Out-Null
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $Users = Import-Csv -Path $CsvPath
        # Build SKU map for license assignment
        $skuMap = @{}
        try {
            $skus = Get-MgSubscribedSku -ErrorAction SilentlyContinue
            if ($skus) {
                foreach ($s in $skus) {
                    if ($s.SkuPartNumber) { $skuMap[$s.SkuPartNumber.ToUpper()] = $s.SkuId }
                }
            }
        }
        catch {
            Write-Log "Warning: Could not retrieve subscribed SKUs: $($_.Exception.Message)" Yellow
        }

        $Users = Import-Csv -Path $CsvPath

    foreach ($User in $Users) {
        try {
            $upn = $User.'User principal name'
            # Check if user already exists, write log and skip
            $existingUser = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
            if ($existingUser -and $existingUser.Id) {
                Write-Log "User $upn already exists, skipping." Yellow
                continue
            }   

            # Skip if UPN is missing
            if (-not $upn) {
                 Write-Log "Skipping row: missing User principal name." Yellow
                continue
            }

            # Skip if extension attributes are missing
            if (-not $User.'isFieldOperative' -or -not $User.'D365FieldService') {
                 Write-Log "Skipping $upn missing required field (operatives and dynamics 365 field service fields). Please check the csv" Yellow
                continue
            }

            # Skip if company name is missing
            if (-not $User.'Company') {
                 Write-Log "Skipping $upn missing Company name." Yellow
                continue
            }

            # Prepare body for Graph create user
            $mailNickname = if ($User.'User logon name') { ($User.'User logon name' -split '@')[0] } else { ($User.'User principal name' -split '@')[0] }

            $body = @{ }
            $body.accountEnabled = if ($User.'Account status' -eq 'Enabled') { $true } else { $false }
            $body.displayName = $User.'Display name'
            $body.mailNickname = $mailNickname
            $body.userPrincipalName = $upn
            $body.passwordProfile = @{ forceChangePasswordNextSignIn = $false; password = $User.'Password' }
            $body.givenName = $User.'First name'
            $body.surname = $User.'Last name'
            if ($User.'Job Title') { $body.jobTitle = $User.'Job Title' }
            if ($User.'Department') { $body.department = $User.'Department' }
            if ($User.'Office') { $body.officeLocation = $User.'Office' }
            if ($User.'Telephone number') { $body.businessPhones = @($User.'Telephone number') }
            if ($User.'Mobile') { $body.mobilePhone = $User.'Mobile' }
            if ($User.'E-mail') { $body.mail = $User.'E-mail' }
            if ($User.'Company') { $body.companyName = $User.'Company' }
            if ($User.'StreetAddress') { $body.StreetAddress = $User.'StreetAddress' }
            if ($User.'Office') { $body.City = $User.'Office' }
            if ($User.'Post code') { $body.PostalCode = $User.'Post code' }
            if ($User.'Country') { $body.Country = $User.'Country' }

            # Create user via Graph using New-MgUser
            $newUserParams = @{
                AccountEnabled    = $body.accountEnabled
                DisplayName       = $body.displayName
                MailNickname      = $body.mailNickname
                UserPrincipalName = $body.userPrincipalName
                PasswordProfile   = $body.passwordProfile
                GivenName         = $body.givenName
                Surname           = $body.surname
            }
            if ($body.jobTitle) { $newUserParams.JobTitle = $body.jobTitle }
            if ($body.department) { $newUserParams.Department = $body.department }
            if ($body.officeLocation) { $newUserParams.OfficeLocation = $body.officeLocation }
            if ($body.businessPhones) { $newUserParams.BusinessPhones = $body.businessPhones }
            if ($body.mobilePhone) { $newUserParams.MobilePhone = $body.mobilePhone }
            if ($body.mail) { $newUserParams.Mail = $body.mail }
            if ($body.companyName) { $newUserParams.CompanyName = $body.companyName }
            if ($body.streetAddress) { $newUserParams.StreetAddress = $body.streetAddress }
            if ($body.city) { $newUserParams.City = $body.city }
            if ($body.postalCode) { $newUserParams.PostalCode = $body.postalCode }
            if ($body.country) { $newUserParams.Country = $body.country }

            try {
                $newUser = New-MgUser @newUserParams -ErrorAction Stop
            }
            catch {
                Write-Log "Failed to create $upn - $($_.Exception.Message)" Red
                continue
            }

            if ($newUser.Id) {
                Write-Log "Created user $upn (id: $($newUser.Id))" Green

                # Set ExtensionAttributes if provided
                $extAttrs = @{ }
                if ($User.'IsFieldOperative') { $extAttrs.extensionAttribute1 = $User.'IsFieldOperative' }
                if ($User.'D365FieldService') { $extAttrs.extensionAttribute2 = $User.'D365FieldService' }

                if ($extAttrs.Count -gt 0) {
                    try {
                        Update-MgUser -UserId $newUser.Id -OnPremisesExtensionAttributes $extAttrs -ErrorAction Stop
                        Write-Log "  - Extension attributes for $upn have been set" Cyan
                    }
                    catch {
                        Write-Log "  - Failed to set extension attributes: $($_.Exception.Message)" Yellow
                    }
                }

                # Assign licenses if provided
                if ($User.'Licenses') {
                    $licenseTokens = $User.'Licenses' -split ',' | ForEach-Object { ($_).Trim() } | Where-Object { $_ -ne '' }
                    $addLicenses = @()
                    foreach ($lic in $licenseTokens) {
                        $skuId = $null
                        if ($skuMap.ContainsKey($lic.ToUpper())) { $skuId = $skuMap[$lic.ToUpper()] }
                        else {
                            # If value looks like a GUID, accept as skuId
                            try { [guid]$lic; $skuId = $lic } catch { }
                        }

                        if ($skuId) { $addLicenses += @{ skuId = $skuId } }
                        else { Write-Log "  - License not found: $lic" Yellow }
                    }

                    if ($addLicenses.Count -gt 0) {
                        $assignBody = @{ addLicenses = $addLicenses; removeLicenses = @() } | ConvertTo-Json -Depth 5
                        try {
                            Invoke-MgGraphRequest -Method POST -Uri "/users/$($newUser.Id)/assignLicense" -Body $assignBody -ErrorAction Stop
                            Write-Log "  - Assigned licenses: $($licenseTokens -join ', ')" Cyan
                        }
                        catch {
                            Write-Log "  - Failed to assign licenses: $($_.Exception.Message)" Yellow
                        }
                    }
                }

                # Add to groups if provided
                if ($User.'Groups') {
                    $groupTokens = $User.'Groups' -split ',' | ForEach-Object { ($_).Trim() } | Where-Object { $_ -ne '' }
                    foreach ($gname in $groupTokens) {
                        try {
                            $group = Get-MgGroup -Filter "displayName eq '$gname'" -ErrorAction SilentlyContinue | Select-Object -First 1
                            if (-not $group) {
                                # try find by mail or mailNickname
                                $group = Get-MgGroup -Filter "mailNickname eq '$gname'" -ErrorAction SilentlyContinue | Select-Object -First 1
                            }

                            if ($group -and $group.Id) {
                                $memberRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)" } | ConvertTo-Json
                                Invoke-MgGraphRequest -Method POST -Uri "/groups/$($group.Id)/members/\$ref" -Body $memberRef -ErrorAction Stop
                                Write-Log "  - Added to group: $gname" Cyan
                            }
                            else {
                                Write-Log "  - Group not found: $gname" Yellow
                            }
                        }
                        catch {
                            Write-Log "  - Failed to add to group $gname $($_.Exception.Message)" Yellow
                        }
                    }
                }

                # Set manager relationship if Manager Email provided
                if ($User.'Manager Email') {
                    try {
                        $manager = Get-MgUser -UserId $User.'Manager Email' -ErrorAction SilentlyContinue
                        if ($manager -and $manager.Id) {
                            try {
                                # Use Graph API directly to set manager
                                $managerRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)" } | ConvertTo-Json
                                Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/users/$($newUser.Id)/manager/`$ref" -Body $managerRef -ErrorAction Stop
                                Write-Log "  - Manager set to $($User.'Manager Email')" Cyan
                            }
                            catch {
                                Write-Log "  - Failed to set manager: $($_.Exception.Message)" Yellow
                            }
                        }
                        else {
                            Write-Log "  - Manager not found: $($User.'Manager Email')" Yellow
                        }
                    }
                    catch {
                        Write-Log "  - Failed to set manager: $($_.Exception.Message)" Yellow
                    }
                }
            }
            else {
                Write-Log "Failed to create $upn - no id returned" Red
            }

            

        }
        catch {
            Write-Host "Failed to create $($User.'User principal name') - $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "Done processing CSV: $CsvPath" -ForegroundColor Green
}
catch {
    Write-Host "Fatal error: $($_.Exception.Message)" -ForegroundColor Red
}

Pause
