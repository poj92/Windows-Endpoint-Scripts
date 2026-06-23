# This updated version attempts to search user based on user principal name

# This script sets Dynamics Field service license for Cardo users

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath


foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
    if($null -eq $ADUser){
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
                         }
    elseif($User.'D365FieldService' -eq ""){
    Write-Host "Error: The Dynamics 365 Field Service field for $($User.'User principal name') is empty, please check the csv file." -ForegroundColor Yellow
    }
    
    # Adding license
    elseif($User.'D365FieldService' -eq 'Yes'){
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        Set-ADUser $SamAccountName -Replace @{'extensionAttribute2' = $User.'D365FieldService'}
        Write-Host "Success: A Dynamics 365 Field Service license has now been assigned to $($User.'User principal name')" -ForegroundColor Green
        }
    
    # Removing license
    elseif($User.'D365FieldService' -eq 'No'){
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        Set-ADUser $SamAccountName -Replace @{'extensionAttribute2' = $User.'D365FieldService'}

        Write-Host "Success: A Dynamics 365 Field Service license has now been removed from $($User.'User principal name')" -ForegroundColor Yellow
        }
    else{
    Write-Host "The Dynamics 365 Field Service field for $($User.'User principal name') is not a valid entry, no changes were made to this user's account" -ForegroundColor Red
        }
}

Pause
