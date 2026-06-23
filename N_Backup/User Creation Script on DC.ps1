<#
     Import users to Cardo Dc

#>

# Setup the user import csv file location
$FilePath = "C:\UserImport.csv"
$Users = Import-Csv $FilePath

# Single password for all new users
$Password = "Go4Live@L0g1ntoCloud"

# Import the Active Directory module if module is not already installed
# Import-Module ActiveDirectory

# Run through each user
foreach ($User in $Users) {
    try {
        # Retrieve the Manager distinguished name - this allows user to be setup under the right manager
        $managerDN = if ($User.'Manager') {
            Get-ADUser -Filter "DisplayName -eq '$($User.'Manager')'" -Properties DisplayName |
            Select-Object -ExpandProperty DistinguishedName
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
            StreetAddress         = $User.'Street'
            City                  = $User.'City'
            State                 = $User.'State/province'
            PostalCode            = $User.'Zip/Postal Code'
            Country               = $User.'Country/region'
            Title                 = $User.'Job Title'
            Department            = $User.'Department'
            Company               = $User.'Company'
            Manager               = $managerDN
            Path                  = $User.'OU'
            Description           = $User.'Description'
            Office                = $User.'Office'
            OfficePhone           = $User.'Telephone number'
            EmailAddress          = $User.'E-mail'
            MobilePhone           = $User.'Mobile'
            AccountPassword       = (ConvertTo-SecureString "$Password" -AsPlainText -Force)
            Enabled               = if ($User.'Account status' -eq "Enabled") { $true } else { $false }
            ChangePasswordAtLogon = $true # 
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
    Set-ADUser $User.'User logon name' -Replace @{'extensionAttribute1'=$User.'IsFieldOperative'}
}


<#
# To remove users - for testing only
Foreach($User in $Users){
Remove-ADUser $User.'User logon name'
}
#>