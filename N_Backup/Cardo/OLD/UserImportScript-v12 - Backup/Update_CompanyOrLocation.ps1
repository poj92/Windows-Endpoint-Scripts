# This updated version attempts to search user based on user principal name

# This script updates user's Company, office, street address and post code at the same time!

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
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
                         }
    elseif("" -eq $User.'Company' -or "" -eq $User.'Office' -or "" -eq $User.'StreetAddress' -or "" -eq $User.'Post Code'){

    Write-Host "Error: Please check the csv file to make sure that the company, office, address and post code fields of $($User.'User principal name') are properly filled." -ForegroundColor Yellow
    break
    }
    
    # Set Company 
    else{
        Set-ADUser $SamAccountName -Company $User.'Company' -StreetAddress $User.'StreetAddress' -PostalCode $User.'Post Code' -City $User.'Office'
        Write-Host "Success: The comapany name for $($User.'User principal name') has been updated to $($User.'Company')" -ForegroundColor Green
        }
}

Pause
