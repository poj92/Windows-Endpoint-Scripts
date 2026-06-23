# This updated version attempts to search user based on user principal name

# Disables or Enables an Account
# Use the scrupt to process leavers automatically.
# This script will ensure that users are moved to the _Leavers OU


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
    
    if($User.'Account status' -eq ""){
        Write-Host "Account status for $($User.'User principal name') is empty." -ForegroundColor Yellow
    }

    elseif($null -eq $ADUser){
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
                         }
    # Disable an account
    elseif($User.'Account status' -eq 'Disabled'){
        
        # Set account status
        Set-ADUser $SamAccountName -Enabled $false
        
        # Move disabled account to leavers OU
        $DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        Move-ADObject -Identity $DistName -TargetPath 'OU=_Leavers,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
        Write-Host "$($User.'User principal name') has been disabled and moved to _Leavers OU. Licenses have been removed" -ForegroundColor Green
        }
    
    # Enable an account
    elseif($User.'Account status' -eq 'Enabled'){
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        
        # Enable account
        Set-ADUser $SamAccountName -Enabled $true
        
        # Move enabled account to correct OU
        $ext1 = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName, extensionAttribute1 | Select-Object -ExpandProperty extensionAttribute1
        $DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        if($ext1 -eq 'operative'){
            Move-ADObject -Identity $DistName -TargetPath 'OU=operatives,OU=Property Services,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
            
            }
        else{
            Move-ADObject -Identity $DistName -TargetPath 'OU=Property Services,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
            
            }
        Write-Host "$($User.'User principal name') has been enabled and moved to the correct OU. Licenses too have been applied" -ForegroundColor Yellow
       
        }

    else{
    Write-Host "Operation was not performed on $($User.'User principal name'). Please check CSV file." -ForegroundColor Yellow
        }
}

Pause
