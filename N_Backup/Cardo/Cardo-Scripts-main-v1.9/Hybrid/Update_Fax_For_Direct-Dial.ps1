# This updated version attempts to search user based on user principal name

# This script updates user's fax number

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Update_Fax_For_Direct-Dial_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Update_Fax_For_Direct-Dial script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName
    if($null -eq $ADUser){
        Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username (email) is correct" Red
    }
    elseif("" -eq $User.'Fax'){
        Write-Log "Error: The Fax field for $($User.'User principal name') is empty, please check the template or csv file." Yellow
        break
    }
    
    # Set Fax Number
    else{
        Set-ADUser $ADUser.SamAccountName -Fax $User.'Fax'
        Write-Log "Success: The Fax for $($User.'User principal name') has been set to $($User.'Fax')" Green
    }
}

Write-Log "Update_Fax_For_Direct-Dial script completed" Cyan
Pause
