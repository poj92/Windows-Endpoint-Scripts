# This script connects to Azure AD and pulls infromation about users in a group

Connect-AzureAD

$Group = Get-AzureADGroup -SearchString "Exclaimer Users" 

Get-AzureADGroupMember -ObjectId $Group.ObjectId -All $true | Select DisplayName, UserPrincipalName, ObjectId, JobTitle, Mobile  | 

Export-CSV "C:\Users\peterj\OneDrive - Nexus Open Systems Ltd\Desktop\Cardo\Users.csv" -NoTypeInformation -Encoding UTF8 