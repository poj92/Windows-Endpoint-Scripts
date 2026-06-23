# This updated version attempts to search user based on user principal name

# This script sets the SMTP address for users

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath

$Results = @()

foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'"
    $Status = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Enabled | Select-Object -ExpandProperty Enabled

    if($null -eq $ADUser){
	$msg = "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct"
        Write-Host $msg -ForegroundColor Red
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}
                      }
    
    # Setting primary email
    elseif($Status -eq $False){
	
	$msg = "$($User.'User principal name') is disabled"
        Write-Host $msg -ForegroundColor DarkRed
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}
        }   

    else{
        $SamAccountName = Get-AdUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties UserPrincipalName | Select-Object -ExpandProperty SamAccountName
        $SMTPAddress = "SMTP:$($User.'User principal name')"
        Set-ADUser $SamAccountName -Replace @{'proxyAddresses' = "$SMTPAddress"}
	
	$msg = "Success: Primary email address for $($User.'User principal name') has been set to $SMTPAddress"
        Write-Host $msg -ForegroundColor Yellow
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}

        }
}

$Results | Export-Csv -Path "$executingScriptDirectory\results.csv" -NoTypeInformation

Write-Host "`nResults have written to $executingScriptDirectory\results.csv"
Pause
