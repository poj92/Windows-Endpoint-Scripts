# This updated version attempts to search user based on user principal name

# This script updates user's StreetAddress, City and PostalCode

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Update_Address_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Update_Address script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName
    if($null -eq $ADUser){
        Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" Red
    }
    elseif("" -eq $User.'StreetAddress' -or "" -eq $User.'Post Code'){
        Write-Log "Error: The address or post code field for $($User.'User principal name') is empty, please check the csv file." Yellow
        break
    }
    
    # Set Address
    else{
        Set-ADUser $ADUser.SamAccountName -StreetAddress $User.'StreetAddress' -PostalCode $User.'Post Code' -City $User.'Office'
        Write-Log "Success: The address and post code for $($User.'User principal name') has been updated" Green
    }
}

Write-Log "Update_Address script completed" Cyan
Pause
