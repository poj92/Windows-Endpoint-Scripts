# This updated version attempts to search user based on user principal name

# This script sets Dynamics Field service license for Cardo users

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Set_D365_Field_License_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Set_D365_Field_License script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName
    if($null -eq $ADUser){
        Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" Red
    }
    elseif($User.'D365FieldService' -eq ""){
        Write-Log "Error: The Dynamics 365 Field Service field for $($User.'User principal name') is empty, please check the csv file." Yellow
    }
    
    # Adding license
    elseif($User.'D365FieldService' -eq 'Yes'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute2' = $User.'D365FieldService'}
        Write-Log "Success: A Dynamics 365 Field Service license has now been assigned to $($User.'User principal name')" Green
    }
    
    # Removing license
    elseif($User.'D365FieldService' -eq 'No'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute2' = $User.'D365FieldService'}
        Write-Log "Success: A Dynamics 365 Field Service license has now been removed from $($User.'User principal name')" Yellow
    }
    else{
        Write-Log "The Dynamics 365 Field Service field for $($User.'User principal name') is not a valid entry, no changes were made to this user's account" Red
    }
}

Write-Log "Set_D365_Field_License script completed" Cyan
Pause
