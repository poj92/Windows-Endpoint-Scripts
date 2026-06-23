# This updated version attempts to search user based on user principal name

# This script shows the phone authentication method of users




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


# Remove current phone authenticator methods
foreach($User in $Users){
        
        $UserParams = @{
            UserId     = $User.'User principal name'
            PhoneAuthenticationMethodId = '3179e48a-750b-4051-897c-87b9720928f7' #remember to set correct ID
        }
         
        Remove-MgUserAuthenticationPhoneMethod @UserParams
                        }
Pause
