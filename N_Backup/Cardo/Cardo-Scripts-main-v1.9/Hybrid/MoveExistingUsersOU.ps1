# This script moves exiting users to the new OU structure

# Tell Powershell the location of the csv file
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "MoveExistingUsersOU_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting MoveExistingUsersOU script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan

# Base OU Path
$domain = "OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"

# Loop through every user in the file
foreach($User in $Users){
	# Check if the user exists
	$ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName, Department, Company, DistinguishedName, extensionAttribute1
	if($null -eq $ADUser){
		Write-Log "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" Red
	}
	else{
		
		#store the parameters
		$Department = $ADUser.Department
		$Company = $ADUser.Company
		$ext1 = $ADUser.extensionAttribute1

		# Set OU Paths 
		$OUPath = if($ext1 -eq "operative"){
					"OU=Operative,OU=$Company,$domain"
					}
				else{
					"OU=$Department,OU=Non-operative,OU=$Company,$domain" 
					}
		
		#Move objects to new OU
		Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $OUPath
		
		Write-Log "'$($User.'User principal name')' has now been moved to the new OU structure" Blue
	}
}

Write-Log "MoveExistingUsersOU script completed" Cyan
Pause
