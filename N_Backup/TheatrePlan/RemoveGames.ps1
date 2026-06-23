
Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like "*Microsoft.Xbox*" | Remove-AppxProvisionedPackage -Online

Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like "*Microsoft.MicrosoftSolitaire*" | Remove-AppxProvisionedPackage -Online