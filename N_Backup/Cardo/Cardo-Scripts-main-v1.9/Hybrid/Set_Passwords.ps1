<#
     Import users to Cardo Dc v1.4
     Updated on 08/08/2024

#>



# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 15 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Set_Passwords_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Set_Passwords script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan



# Import the Active Directory module if module is not already installed
# Import-Module ActiveDirectory



# Run through each user
foreach ($User in $Users) {
        Set-ADAccountPassword -Identity $User.'User logon name' -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $User.'Password' -Force)

        Write-Log "The user password for $($User.'User principal name') was updated successfully." Green
            

}

Write-Log "Set_Passwords script completed" Cyan
Pause