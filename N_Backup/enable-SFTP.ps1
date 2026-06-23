Connect-AzureAD

$resourceGroupName = "Cardo-Orchard-Interface-Test-RG"
$storageAccountName = "cardoorchardinterfaad4a"

Set-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -EnableSftp $true