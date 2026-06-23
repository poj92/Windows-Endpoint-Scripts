# This updated version attempts to search user based on user principal name


# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath


foreach($User in $Users){
		        
	$Department = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Department | Select-Object -ExpandProperty Department
		
	$Company = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Company | Select-Object -ExpandProperty Company
		
	$DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName

	# Set OU Paths 

	$Company = $User.'Company'
	$Department = $User.'Department'
	$OU = if ($User.'IsFieldOperative' -eq 'operative'){
              "OU=Operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
            }
          else{
              "OU=$Department,OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
            }
	
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
    if($null -eq $ADUser){
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
                         }
    elseif($User.'IsFieldOperative' -eq ""){
    Write-Host "The field operative field for $($User.'User principal name') is empty." -ForegroundColor Yellow
    }
    
    # Settings for operatives
    elseif($User.'IsFieldOperative' -eq 'operative'){
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        Set-ADUser $SamAccountName -Replace @{'extensionAttribute1' = $User.'IsFieldOperative'}
        
        $DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        Move-ADObject -Identity $DistName -TargetPath $OU
        Write-Host "$($User.'User principal name') has now been set as an operative. The user has been moved to the correct OU successfully." -ForegroundColor Green
        }
    
    # settings for non-operatives
    elseif($User.'IsFieldOperative' -eq 'no'){
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        Set-ADUser $SamAccountName -Replace @{'extensionAttribute1' = $User.'IsFieldOperative'}

        $DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        Move-ADObject -Identity $DistName -TargetPath $OU
        Write-Host "$($User.'User principal name') has now been set as Non-operative. The user has been moved to the correct OU successfully.." -ForegroundColor Yellow
        }
    else{
    Write-Host "The field operative field for $($User.'User principal name') is not a valid entry, no changes were made to this user's account" -ForegroundColor Red
        }
}

Pause
