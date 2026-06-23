# This script moves exiting users to the new OU structure

# Tell Powershell the location of the csv file
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath

# Base OU Path
$domain = "OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"

# Loop through every user in the file
foreach($User in $Users){
	# Check if the user exists
	$ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
	if($null -eq $ADUser){
		write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
	}
	else{
		
		#store the parameters
		$SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        
		$Department = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Department | Select-Object -ExpandProperty Department
		
		$Company = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Company | Select-Object -ExpandProperty Company
		
		$DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        
		$ext1 = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties extensionAttribute1 | Select-Object -ExpandProperty extensionAttribute1

		# Set OU Paths 
		$OUPath = if($ext1 -eq "operative"){
					"OU=Operative,OU=$Company,$domain"
					}
				else{
					"OU=Non-operative,OU=$Company,$domain" 
					}
		
		#Move objects to new OU
		Move-ADObject -Identity $DistName -TargetPath $OUPath
		
		Write-Host "'$($User.'User principal name')' has now been moved to the new OU structure" -ForegroundColor Blue
	}
}

Pause
