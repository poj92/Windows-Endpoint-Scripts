# This updated version attempts to search user based on user principal name

# This script sets E3 (no Teams) license for Cardo users

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Set_E3NoTeams_License_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Set_E3NoTeams_License script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName
    if($null -eq $ADUser){
        Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" Red
    }
    elseif($User.'E3Only' -eq ""){
        Write-Log "Error: The E3Only field for $($User.'User principal name') is empty, please check the csv file." Yellow
    }
    
    # Adding license
    elseif($User.'E3Only' -eq 'Yes'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute3' = $User.'E3Only'}
        Write-Log "Success: An E3 (no Teams) license will be assigned to $($User.'User principal name')" Green
    }
    
    # Removing license
    elseif($User.'E3Only' -eq 'No'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute3' = $User.'E3Only'}
        Write-Log "Success: An E3 (no Teams) license will be removed from $($User.'User principal name')" Yellow
    }
    else{
        Write-Log "The E3Only field for $($User.'User principal name') is not a valid entry, no changes were made to this user's account" Red
    }
}

Write-Log "Set_E3NoTeams_License script completed" Cyan
Pause
