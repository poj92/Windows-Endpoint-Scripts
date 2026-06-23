Connect-MsolService
$users = Import-csv -Path "C:\Users\peterj\OneDrive - Nexus Open Systems Ltd\Desktop\Cardo\CardoTitles.csv"

foreach($user in $users){Set-MsolUser -UserPrincipalName $user.UserPrincipalName -MobilePhone $user.Mobile}


foreach($user in $users){Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select UserPrinciPalName, JobTitle}


# Get Azure-AD
Connect-AzureAD
$users = Import-csv -Path "C:\Users\peterj\OneDrive - Nexus Open Systems Ltd\Desktop\Cardo\update.csv"
Get-AzureADUSer -UserPrincipalName Emily.Light@cardogroup.co.uk | Select JobTitle


# Get information about users based on csv file
foreach($user in $users){Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select UserPrinciPalName, Title}