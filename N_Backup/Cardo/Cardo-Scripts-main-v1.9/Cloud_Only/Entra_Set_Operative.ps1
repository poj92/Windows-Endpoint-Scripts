<#
.SYNOPSIS
    Updates user properties in Entra ID based on details provided in a CSV file.
.DESCRIPTION
    This script updates existing users in Entra ID by performing the following actions:
    - Imports user details from a CSV file, automatically uses "UserImport.csv" in the script directory if not specified
    - Updates extensionattribute1 to "Operative" or "Non-Operative" based on the "IsFieldOperative" column  in the CSV
    - Logs all actions taken to a timestamped log file
.PARAMETER CsvPath
    Path to the CSV file containing user details. If not provided, defaults to "UserImport.csv" in the script directory.
.EXAMPLE
    .\Entra_Set_Operative.ps1
    .\Entra_Set_Operative.ps1 -CsvPath "C:\Path\To\UserImport.csv"
.NOTES
    This script requires the Microsoft Graph PowerShell module.
    Ensure you have the necessary permissions to update user profiles in Entra ID.  
    If you have any questions or need further assistance, please contact Peter James at peter.james @nexusos.co.uk
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
    $logFile = Join-Path $scriptDir "Entra_Set_Operative_$timestamp.log"

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
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph." -Color Green

    # Import user data from CSV
    $users = Import-Csv -Path $CsvPath

    foreach ($user in $users) {
        try {
            $userPrincipalName = $user.'User Principal Name'
            $isFieldOperative = $user.'IsFieldOperative'

            if (-not $userPrincipalName) {
                Write-Log "Skipping entry with missing User Principal Name." -Color Yellow
                continue
            }

            if (-not $isFieldOperative) {
                Write-Log "Skipping user $userPrincipalName due to missing IsFieldOperative field." -Color Yellow
                continue    
            }
            # Update user's extensionattribute1
            Write-Log ("Updating extensionattribute1 for user {0} to {1}..." -f $userPrincipalName, $isFieldOperative)
            Update-MgUser -UserId $userPrincipalName -OnPremisesExtensionAttributes @{extensionattribute1 = $isFieldOperative} -ErrorAction Stop
            Write-Log ("Successfully updated extensionattribute1 for user {0}." -f $userPrincipalName) -Color Green
        }   
        catch {
            # Format to avoid colon parsing issues
            Write-Log ("Error updating user {0}: {1}" -f $userPrincipalName, $_.Exception.Message) -Color Red
        }
    }
    Write-Log "User operative status update process completed." -Color Green
}
catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
Pause