<#
    .CHANGELOG
    V1.00, 22/10/2025 

#>

# Connect to Microsoft Graph with user read/write permissions
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Specify the path of the CSV file
$CSVFilePath = "C:\temp\NewUsers.csv"

# Create password profile
$PasswordProfile = @{
    Password                             = "P@ssw0rd!"
    ForceChangePasswordNextSignIn        = $true
    ForceChangePasswordNextSignInWithMfa = $true
}

# Import data from CSV file
$Users = Import-Csv -Path $CSVFilePath

# Loop through each row containing user details in the CSV file
foreach ($User in $Users) {
    $UserParams = @{
        DisplayName       = $User.DisplayName
        MailNickName      = $User.MailNickName
        UserPrincipalName = $User.UserPrincipalName
        Department        = $User.Department
        JobTitle          = $User.JobTitle
        Mobile            = $User.Mobile
        Country           = $User.Country
        EmployeeId        = $User.EmployeeId
        PasswordProfile   = $PasswordProfile
        AccountEnabled    = $true
    }

    try {
        $null = New-MgUser @UserParams -ErrorAction Stop
        Write-Host ("Successfully created the account for {0}" -f $User.DisplayName) -ForegroundColor Green
    }
    catch {
        Write-Host ("Failed to create the account for {0}. Error: {1}" -f $User.DisplayName, $_.Exception.Message) -ForegroundColor Red
    }
}