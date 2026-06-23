# This updated version attempts to search user based on user principal name

# Converts an email to sharedmailbox, blocks sign-in 
# Use the script to process leavers automatically.
# This script will ensure that users are moved to the _Leavers OU


# Connect to Exchange Online
try {
    if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Install-Module -Name ExchangeOnlineManagement -Force
    }
    Write-Host "Please sign in to Cardo's Exchange Online with an Admin account to continue" -Foregroundcolor Yellow
	Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit
}

# Load the csv file that is containing user information
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath
Write-Host "csv file loaded okay" -Foregroundcolor Green


foreach($User in $Users){
	
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
    $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
    $UserPrincipalName = $User.'User principal name'
    if($User.'Account status' -eq ""){
        Write-Host "Account status for $($User.'User principal name') is empty." -ForegroundColor Yellow
    }

    elseif($null -eq $ADUser){
    write-host "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct" -ForegroundColor Red
                         }
    
	
    elseif($User.'Account status' -eq 'Disabled'){
        
		# --- Convert Mailbox to Shared Mailbox ---
		try {
			Set-Mailbox -Identity $UserPrincipalName -Type Shared
			Write-Host "Successfully converted $UserPrincipalName to a shared mailbox."
		}
		catch {
			Write-Error "Failed to convert mailbox to shared: $($_.Exception.Message)"
			break
		}
		
		
        # Set account status to disabled 
        Set-ADUser $SamAccountName -Enabled $false
        
        # Move disabled account to leavers OU
        $DistName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties SamAccountName | Select-Object -ExpandProperty DistinguishedName
        Move-ADObject -Identity $DistName -TargetPath 'OU=_Leavers,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
        Write-Host "$($User.'User principal name') has been disabled and moved to _Leavers OU. Licenses will be removed" -ForegroundColor Green
		
        }


    else{
    Write-Host "Operation was not performed on $($User.'User principal name'). Please check CSV file." -ForegroundColor Yellow
        }
}

Pause
