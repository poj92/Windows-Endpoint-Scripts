<#
.DESCRIPTION
    This script sets the phone authentication number for users in Entra ID by performing the following actions:
    - Imports user details from a CSV file, automatically uses "UserImport.csv" in the script directory if not specified
    - Updates the user's phone authentication number based on the "PhoneNumber" field in the CSV
    - Logs all actions taken to a timestamped log file
.PARAMETER CsvPath
    Path to the CSV file containing user details. If not provided, defaults to "UserImport.csv" in the script directory.
.EXAMPLE
    .\Entra_PhoneAuthentication.ps1
    .\Entra_PhoneAuthentication.ps1 -CsvPath "C:\Path\To\UserImport.csv"
.NOTES
    This script requires the Microsoft Graph PowerShell module.
    Ensure you have the necessary permissions to update user profiles in Entra ID.  
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
    $logFile = Join-Path $scriptDir "Entra_PhoneAuthentication_$timestamp.log"
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
    # Connect to Microsoft Graph
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","UserAuthenticationMethod.ReadWrite.All" -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph." -Color Green
    # Import user data from CSV
    $users = Import-Csv -Path $CsvPath
    foreach ($user in $users) {
        try {
            $userPrincipalName = $user.'User Principal Name'
            $newPhoneNumber = $user.'AuthenticationPhone'
            if (-not $userPrincipalName) {
                Write-Log "Skipping entry with missing User Principal Name." -Color Yellow
                continue
            }
            if (-not $newPhoneNumber) {
                Write-Log "Skipping user $userPrincipalName due to missing PhoneNumber." -Color Yellow
                continue
            }
            $mgUser = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
            if (-not $mgUser) {
                Write-Log "User not found: $userPrincipalName" -Color Red
                continue
            }
            # Check existing phone authentication methods
            $authMethods = Get-MgUserAuthenticationPhoneMethod -UserId $userPrincipalName -ErrorAction Stop
            $phoneMethod = $authMethods | Where-Object { $_.PhoneNumber -eq $newPhoneNumber }
            if ($phoneMethod) {
                Write-Log "Authentication phone number for $UserPrincipalName is already set to $newPhoneNumber. No update needed." -Color Green                Write-Log "Phone number for user $userPrincipalName is already up to date." -Color Green
            }
            else {
                # Update phone authentication number
                Write-Log ("Updating phone authentication number for user {0} to {1}..." -f $userPrincipalName, $newPhoneNumber)
                # Remove existing phone methods
                foreach ($method in $authMethods) {
                    Remove-MgUserAuthenticationPhoneMethod -UserId $userPrincipalName -AuthenticationPhoneMethodId $method.Id -ErrorAction Stop
                }
                # Add new phone method
                New-MgUserAuthenticationPhoneMethod -UserId $userPrincipalName -PhoneNumber $newPhoneNumber -PhoneType "mobile" -ErrorAction Stop
                Write-Log ("Successfully updated phone authentication number for user {0}." -f $userPrincipalName) -Color Green
            }
        }   
        catch {
            # Format to avoid colon parsing issues
            Write-Log ("Error updating user {0}: {1}" -f $userPrincipalName, $_.Exception.Message) -Color Red
        }
    }
    Write-Log "Phone authentication update process completed." -Color Green
}
catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
}

Pause