<#
     Import users to Cardo Dc v1.2

#>



# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 15 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "C:\UserScript\UserImport.csv" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
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

        $OU = if ($User.'IsFieldOperative' -eq 'operative'){
                'OU=Operatives,OU=Property Services,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
                }
            else {
                    'OU=Property Services,OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk'
                }
        $IsOperative = if ($User.'IsFieldOperative' -eq 'operative'){
                'operative'
                }
            else {
                    ''
                }
        if ($User.'Company' -ne "Cardo South") {
            Write-Host "Error: Company name not properly set for one or more users. `n This process will now be aborted (no user account has been created). `n Please close PowerShell window and set company name for all users to Cardo South and re-run the script. `n If you need hep, please reach out to Nexus Service Desk."  -ForegroundColor Red
            Pause
		return
                }

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
        }

        # Check to see if the user already exists in AD
        if (Get-ADUser -Filter "SamAccountName -eq '$($User.'User logon name')'") {

            # Give a warning if user exists
            Write-Host "A user with username $($User.'User logon name') already exists in Active Directory." -ForegroundColor Yellow
        }
        else {
            # User does not exist then proceed to create the new user account
            # Account will be created in the OU provided by the $User.OU variable read from the CSV file
            New-ADUser @NewUserParams
            Write-Host "The user $($User.'User logon name') is created successfully." -ForegroundColor Green
        }
    }
    catch {
        # Handle any errors that occur during account creation
        Write-Host "Failed to create user $($User.'User logon name') - $($_.Exception.Message)" -ForegroundColor Red
    }
}


# Now run the following to replace update the extension attribute section
Foreach($User in $Users){
    if($User.'IsOperative' -ne $null){
    Set-ADUser $User.'User logon name' -Replace @{'extensionAttribute1' = $IsOperative}
    }
   }


<#
# To remove users - for testing only
Foreach($User in $Users){
Remove-ADUser $User.'User logon name'
}
#>

Pause