# This script creates new OUs with the correct structure for the company as per 
# Cardo's template

$CompanyName = Read-Host "Please enter the Name of the Company"

$domain = "OU=ADUsers,DC=ad,DC=cardogroup,DC=co,DC=uk"

$TestOU = Get-ADOrganizationalUnit -Filter "Name -eq '$CompanyName'" -SearchBase $domain -ErrorAction SilentlyContinue

if(-not $TestOU){
    New-ADOrganizationalUnit -Name $CompanyName -Path $domain
    #$CompanyOU = "$"
    Write-Host "The parent OU for $CompanyName has been created" -ForegroundColor Yellow
	
	# Create operative OU
	$CompanyOU = "OU=$CompanyName,$domain"
	New-ADOrganizationalUnit -Name "Operative" -Path $CompanyOU
	
	#Create Non-operative OU
	New-ADOrganizationalUnit -Name "Non-operative" -Path $CompanyOU
	
	# Create sub OUs for non-operatives
	$NonOperativesOU = "OU=Non-operative,$CompanyOU"
	$OUList = @("Finance", "Maintenance", "IT", "Procurement")
	foreach($OU in $OUList){
		New-ADOrganizationalUnit -Name "$OU" -Path $NonOperativesOU
	}
	
	Write-Host "All OUs for $CompanyName have been created" -ForegroundColor Magenta
}
else{
    Write-Host "The parent OU for $CompanyName already exists. OUs were not created" -ForegroundColor Cyan
}


Pause
