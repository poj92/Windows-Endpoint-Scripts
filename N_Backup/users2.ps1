$RG1 = "Cardo-CoreServers-RG"
$Connection16 = "VNet1toSite6"
$connection6 = Get-AzVirtualNetworkGatewayConnection -Name $Connection16 -ResourceGroupName $RG1
$connection6.IpsecPolicies