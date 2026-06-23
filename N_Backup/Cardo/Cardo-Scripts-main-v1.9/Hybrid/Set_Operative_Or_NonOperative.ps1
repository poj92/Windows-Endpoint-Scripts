# This updated version attempts to search user based on user principal name


# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Set_Operative_Or_NonOperative_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Set_Operative_Or_NonOperative script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan


foreach($User in $Users){
		        
	# Set OU Paths 

	$Company = $User.'Company'
	$Department = $User.'Department'
	$OUdomain = "OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
	$DiName = $User.'Display name'
		
	$OU = if ($User.'IsFieldOperative' -eq 'operative'){
               "OU=Operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
               }
			elseif(($User.'IsFieldOperative' -ne 'operative') -and ($Department -eq "")){
				Write-Log "Warning: $DiName's department is empty, user will be added to non-operative OU under $Company" Yellow
				"OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
			}
			elseif(($User.'IsFieldOperative' -ne 'operative') -and ($Department -ne "")){
				$TestOU = Get-ADOrganizationalUnit -Filter "Name -eq '$Department'" -SearchBase $OUdomain -ErrorAction SilentlyContinue
				if(-not $TestOU){
					Write-Log "Warning: $DiName's department does not match existing OU, user will be added to non-operative OU under $Company" Yellow
					"OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
						}
					else{
						Write-Log "Warning: $DiName user will be added to $Department OU under $Company" Cyan
						"OU=$Department,OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
						}
			}
<#				else{
					"OU=$Department,OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
                } 
#>
	
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName, DistinguishedName
    if($null -eq $ADUser){
        Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" Red
    }
    elseif($User.'IsFieldOperative' -eq ""){
        Write-Log "The field operative field for $($User.'User principal name') is empty." Yellow
    }
    
    # Settings for operatives
    elseif($User.'IsFieldOperative' -eq 'operative'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute1' = $User.'IsFieldOperative'}
        
        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $OU
        Write-Log "$($User.'User principal name') has now been set as an operative. The user has been moved to the correct OU successfully." Green
    }
    
    # settings for non-operatives
    elseif($User.'IsFieldOperative' -eq 'no'){
        Set-ADUser $ADUser.SamAccountName -Replace @{'extensionAttribute1' = $User.'IsFieldOperative'}

        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $OU
        Write-Log "$($User.'User principal name') has now been set as Non-operative. The user has been moved to the correct OU successfully.." Yellow
    }
    else{
        Write-Log "The field operative field for $($User.'User principal name') is not a valid entry, no changes were made to this user's account" Red
    }
}

Write-Log "Set_Operative_Or_NonOperative script completed" Cyan
Pause
