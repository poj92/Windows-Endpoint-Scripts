# This updated version attempts to search user based on user principal name

# This script updates user's fax number

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
    $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
    if($null -eq $ADUser){
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username (email) is correct" -ForegroundColor Red
                         }
    elseif("" -eq $User.'Fax'){

    Write-Host "Error: The Fax field for $($User.'User principal name') is empty, please check the template or csv file." -ForegroundColor Yellow
    break
    }
    
    # Set Address
    else{
        Set-ADUser $SamAccountName -Fax $User.'Fax'
        Write-Host "Success: The Fax for $($User.'User principal name') has been set to $($User.'Fax')" -ForegroundColor Green
        }
}

Pause
