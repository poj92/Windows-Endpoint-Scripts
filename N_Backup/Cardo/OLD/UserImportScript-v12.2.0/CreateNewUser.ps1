<#
     Import users to Cardo DC v12
     Updated on 16/06/2025

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
    try {
        # Retrieve the Manager distinguished name - this allows user to be setup under the right manager
        $managerDN = if ($User.'Manager Email') {
            Get-ADUser -Filter "UserPrincipalName -eq '$($User.'Manager Email')'" -Properties DisplayName |
            Select-Object -ExpandProperty DistinguishedName
        }

        
       #Check if correct company name is entered and operative srarus in csv file
        if ($User.'Company' -eq "" -or $User.'IsFieldOperative' -eq "") {
            Write-Host "Error: Company name or field operative property has not been set for one or more users. `n This process will now be aborted (no user account has been created). `n Please close PowerShell window and set company name and operative status for all users and re-run the script. `n If you need help, please reach out to Nexus Service Desk."  -ForegroundColor Red
            Pause
		return
                }
		
		# Set OU
		$Company = $User.'Company'
		$Department = $User.'Department'
		$OUdomain = "OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
		$DiName = $User.'Display name'
		
		$OU = if ($User.'IsFieldOperative' -eq 'operative'){
                "OU=Operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
                }
				elseif(($User.'IsFieldOperative' -ne 'operative') -and ($Department -eq "")){
					Write-Host "Warning: $DiName's department is empty, user will be added to non-operative OU under $Company" -ForegroundColor Yellow
					"OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
				}
				elseif(($User.'IsFieldOperative' -ne 'operative') -and ($Department -ne "")){
					$TestOU = Get-ADOrganizationalUnit -Filter "Name -eq '$Department'" -SearchBase $OUdomain -ErrorAction SilentlyContinue
					if(-not $TestOU){
						Write-Host "Warning: $DiName's department does not match existing OU, user will be added to non-operative OU under $Company" -ForegroundColor Yellow
						"OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
							}
						else{
							Write-Host "Warning: $DiName user will be added to $Department OU under $Company" -ForegroundColor Cyan
							"OU=$Department,OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
							}
				}
<# 				else{
					"OU=$Department,OU=Non-operative,OU=$Company,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"
                } 
#>


        # Define the parameters using a hashtable
        $NewUserParams = @{
            Name                  = "$($User.'First name') $($User.'Last name')"
            GivenName             = $User.'First name'
            Surname               = $User.'Last name'
            DisplayName           = $User.'Display name'
            OtherName             = $User.'Other Name'
            SamAccountName        = $User.'User logon name'
            UserPrincipalName     = $User.'User principal name'
            Title                 = $User.'Job Title'
            Department            = $User.'Department'
            Company               = $User.'Company'
            Manager               = $managerDN
            Path                  = $OU
            Description           = $User.'Description'
            Office                = $User.'Office'
            OfficePhone           = $User.'Telephone number'
            EmailAddress          = $User.'E-mail'
            MobilePhone           = $User.'Mobile'
            AccountPassword       = (ConvertTo-SecureString $User.'Password' -AsPlainText -Force)
            Enabled               = if ($User.'Account status' -eq "Enabled") { $true } else { $false }
            ChangePasswordAtLogon = $false # 
            City                  = $User.'Office'
            StreetAddress         = $User.'StreetAddress'
            PostalCode            = $User.'Post Code'
        }

        # Check to see if the user already exists in AD
        if (Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'") {
            # Give a warning if user exists and terminate process
            Write-Host "$($User.'User principal name') already exists in Active Directory." -ForegroundColor Yellow
                                                                                          }

            else {
            # User does not exist then proceed to create the new user account
            
                New-ADUser @NewUserParams
                Write-Host "The user $($User.'User principal name') was created successfully." -ForegroundColor Green
            
                if($User.'IsFieldOperative' -ne ""){
                    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
                    Set-ADUser $ADUser -Replace @{'extensionAttribute1' = $User.'IsFieldOperative'}
                                                }

                if($User.'D365FieldService' -ne ""){
                    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
                    Set-ADUser $User.'User logon name' -Replace @{'extensionAttribute2' = $User.'D365FieldService'}
                                                }
                }
    }
    catch {
        # Handle any errors that occur during account creation
        Write-Host "Failed to create $($User.'User principal name') - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Pause