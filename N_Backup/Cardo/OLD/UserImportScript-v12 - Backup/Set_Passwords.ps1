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
$Users = Import-Csv $FilePath



# Import the Active Directory module if module is not already installed
# Import-Module ActiveDirectory



# Run through each user
foreach ($User in $Users) {
        Set-ADAccountPassword -Identity $User.'User logon name' -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $User.'Password' -Force)

        Write-Host "The user password for $($User.'User principal name') was updated successfully." -ForegroundColor Green
            

}

Pause