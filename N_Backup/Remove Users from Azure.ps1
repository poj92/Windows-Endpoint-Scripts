$Users = Import-Csv "C:\Users\peterj\Downloads\exportUsers_2024-5-7.csv"

foreach($user in $Users){
Get-AzureADUser -ObjectId $User.'userPrincipalName'
}



foreach($user in $Users){
Remove-AzureADUser -ObjectId $User.'userPrincipalName'
}
