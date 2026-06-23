$path1 = "$env:LOCALAPPDATA\Kingsoft\WPS Office"
$path2 = "C:\Kingsoft\WPS Office"
 
md "$path1" -Force | Out-Null
md "$path2" -Force | Out-Null
 
$denyRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
  "Everyone",
  "FullControl",
  "ContainerInherit,ObjectInherit",
  "None",
  "Deny"
)
 
foreach ($folder in @($path1, $path2)) {
  $acl = Get-Acl $folder
  $acl.AddAccessRule($denyRule)
  Set-Acl -Path $folder -AclObject $acl
  Write-Host "Access restricted to $folder"
}
