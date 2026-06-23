# This updated version attempts to search user based on user principal name

# This script updates authentication mobile number




Install-module Microsoft.Graph.Identity.Signins
Connect-MgGraph -Scopes "User.Read.all","UserAuthenticationMethod.Read.All","UserAuthenticationMethod.ReadWrite.All"
# Select-MgProfile -Name beta


# Setup the user import csv file location

# Instruct script to search for CSV file from the same location as itself
# If running from PowerShell Window, manually, uncomment line 11 and specify csv filepath accordingly
$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$FilePath = Join-Path $executingScriptDirectory "UserImport.csv"

# $FilePath = "Path to Csv file" #only uncommet this line if running script from within PowerShell ISE and need to specify the file path manually
$Users = Import-Csv $FilePath


# Set Phone number for authentication
foreach($User in $Users){
        
        $UserParams = @{
            UserId      = $User.'User principal name'
            phoneType   = "mobile"
            phoneNumber = $User.'AuthenticationPhone'

        }
        New-MgUserAuthenticationPhoneMethod @UserParams
                        }

Pause
