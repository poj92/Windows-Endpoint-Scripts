# This updated version attempts to search user based on user principal name

# This script sets the SMTP address for users

# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually

# Prepare log file
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$logFile = Join-Path $executingScriptDirectory "Set_SMTP_Address_$timestamp.log"

function Write-Log {
    param([string]$Message, [ConsoleColor]$Color = 'White')
    $time = (Get-Date).ToString('s')
    $line = "[$time] $Message"
    Write-Host $Message -ForegroundColor $Color
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

Write-Log "Starting Set_SMTP_Address script" Cyan
Write-Log "Log file: $logFile" Cyan

$Users = Import-Csv $FilePath
Write-Log "Loaded $($Users.Count) user(s) from CSV" Cyan

$Results = @()

foreach($User in $Users){
    $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($User.'User principal name')'" -Properties Enabled, SamAccountName

    if($null -eq $ADUser){
	$msg = "'$($User.'User principal name')' does not exist in active directory, please check that the username is correct"
        Write-Log $msg Red
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}
                      }
    
    # Setting primary email
    elseif($ADUser.Enabled -eq $False){
	
	$msg = "$($User.'User principal name') is disabled"
        Write-Log $msg DarkRed
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}
        }   

    else{
        $SMTPAddress = "SMTP:$($User.'User principal name')"
        Set-ADUser $ADUser.SamAccountName -Replace @{'proxyAddresses' = "$SMTPAddress"}
	
	$msg = "Success: Primary email address for $($User.'User principal name') has been set to $SMTPAddress"
        Write-Log $msg Yellow
	
	$Results += [PSCustomObject]@{
            UserPrincipalName = $User.'User principal name'
            Output           	= $msg
        	}

        }
}

$Results | Export-Csv -Path "$executingScriptDirectory\results.csv" -NoTypeInformation

Write-Log "`nResults have written to $executingScriptDirectory\results.csv" Cyan
Write-Log "Set_SMTP_Address script completed" Cyan
Pause
