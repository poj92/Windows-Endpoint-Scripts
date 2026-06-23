<#
.DESCRIPTION
    This script updates existing users in Entra ID by performing the following actions:
    - Imports user details from a CSV file, automatically uses "UserImport.csv" in the script directory if not specified
    - Updates specified user properties such as mobile number, address, etc.
    - Logs all actions taken to a timestamped log file
.PARAMETER CsvPath
    Path to the CSV file containing user details. If not provided, defaults to "UserImport.csv" in the script directory.
.PARAMETER PropertyToUpdate
    Specifies which property to update. Accepts values like "All", "Mobile", "Address", etc.
.EXAMPLE
    .\Entra_Update_UserProfile.ps1 -CsvPath "C:\Path\To\UserImport.csv" -PropertyToUpdate "All"
.NOTES
This script requires the Microsoft Graph PowerShell module.
Ensure you have the necessary permissions to update user profiles in Entra ID.  
Update all available properties
    .\Entra_Update_UserProfile.ps1 -CsvPath ".\UserImport.csv" -PropertyToUpdate "All"

Update only mobile numbers
    .\Entra_Update_UserProfile.ps1 -CsvPath ".\UserImport.csv" -PropertyToUpdate "Mobile"

Update address information
    .\Entra_Update_UserProfile.ps1 -CsvPath ".\UserImport.csv" -PropertyToUpdate "Address"

#>

param(
    [string]$CsvPath,
    [string]$PropertyToUpdate = "All"
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
    $logFile = Join-Path $scriptDir "Entra_Update_User_$timestamp.log"

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
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
    Connect-MgGraph -Scopes "User.ReadWrite.All"
    Write-Log "Connected to Microsoft Graph." -Color Green
    $users = Import-Csv -Path $CsvPath
    foreach ($user in $users) {
        # Support both "UserPrincipalName" and "User principal name" column headers
        $userPrincipalName = if ($user.UserPrincipalName) { $user.UserPrincipalName } elseif ($user.'User principal name') { $user.'User principal name' } else { $null }
        if (-not $userPrincipalName) {
            Write-Log "UserPrincipalName is missing in CSV entry. Skipping entry." -Color Yellow
            continue
        }
        try {
            $mgUser = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
            $updateParams = @{}
            switch ($PropertyToUpdate.ToLower()) {
                "all" {
                    if ($user.Mobile) { $updateParams["MobilePhone"] = $user.Mobile }
                    if ($user.StreetAddress) { $updateParams["StreetAddress"] = $user.StreetAddress }
                    if ($user.Office) { $updateParams["City"] = $user.Office }
                    if ($user.State) { $updateParams["State"] = $user.State }
                    if ($user.'Post Code') { $updateParams["PostalCode"] = $user.'Post Code' }
                    if ($user.Country) { $updateParams["Country"] = $user.Country }
                    if ($user.'Job Title') { $updateParams["JobTitle"] = $user.'Job Title' }
                    if ($user.Department) { $updateParams["Department"] = $user.Department }
                    if ($user.Office) { $updateParams["OfficeLocation"] = $user.Office }
                    if ($user.Company) { $updateParams["CompanyName"] = $user.Company }
                    if ($user.'Telephone number') { $updateParams["BusinessPhones"] = @($user.'Telephone number') }
                }
                "mobile" {
                    if ($user.Mobile) { $updateParams["MobilePhone"] = $user.Mobile }
                }
                "address" {
                    if ($user.StreetAddress) { $updateParams["StreetAddress"] = $user.StreetAddress }
                    if ($user.Office) { $updateParams["City"] = $user.Office }
                    if ($user.State) { $updateParams["State"] = $user.State }
                    if ($user.'Post Code') { $updateParams["PostalCode"] = $user.'Post Code' }
                    if ($user.Country) { $updateParams["Country"] = $user.Country }
                }
                default {
                    Write-Log ("Unknown PropertyToUpdate value: {0}. Skipping user {1}." -f $PropertyToUpdate, $userPrincipalName) -Color Yellow
                    continue
                }
            }
            if ($updateParams.Count -gt 0) {
                # Validation: log which properties will be updated (robust string building)
                $propsPairs = @()
                foreach ($kv in $updateParams.GetEnumerator()) {
                    $val = $kv.Value
                    if ($null -ne $val -and $val -is [System.Array]) { $val = ($val -join ';') }
                    $propsPairs += ("{0}='{1}'" -f $kv.Key, $val)
                }
                $propsSummary = ($propsPairs -join ", ")
                Write-Log ("Planned updates for {0}: {1}" -f $userPrincipalName, $propsSummary) -Color Cyan
                Update-MgUser -UserId $userPrincipalName @updateParams -ErrorAction Stop
                Write-Log ("Updated user {0} successfully." -f $userPrincipalName) -Color Green
            }
            else {
                Write-Log ("No valid properties to update for user {0}. Skipping." -f $userPrincipalName) -Color Yellow
            }

            # Set manager relationship if Manager Email provided (follows Entra_BulkUserImport.ps1 pattern)
            if ($user.'Manager Email') {
                try {
                    $manager = Get-MgUser -UserId $user.'Manager Email' -ErrorAction SilentlyContinue
                    if ($manager -and $manager.Id) {
                        try {
                            $managerRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)" } | ConvertTo-Json
                            Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/users/$userPrincipalName/manager/`$ref" -Body $managerRef -ErrorAction Stop
                            Write-Log ("  - Manager set to {0}" -f $user.'Manager Email') -Color Cyan
                        }
                        catch {
                            Write-Log ("  - Failed to set manager: {0}" -f $_.Exception.Message) -Color Yellow
                        }
                    }
                    else {
                        Write-Log ("  - Manager not found: {0}" -f $user.'Manager Email') -Color Yellow
                    }
                }
                catch {
                    Write-Log ("  - Failed to set manager: {0}" -f $_.Exception.Message) -Color Yellow
                }
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log ("Failed to update user {0}: {1}" -f $userPrincipalName, $errorMsg) -Color Red
        }
    }
    Write-Log "User update process completed." -Color Green
}
catch {
    $outerError = $_.Exception.Message
    Write-Host ("An unexpected error occurred: {0}" -f $outerError) -ForegroundColor Red
}
    
Pause