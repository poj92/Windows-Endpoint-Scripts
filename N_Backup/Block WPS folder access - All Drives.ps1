# Original path1
$path1 = "$env:LOCALAPPDATA\Kingsoft\WPS Office"

# Get all available drives that are fixed or removable
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match "^[A-Z]:\\$" }

# Build path2 equivalents for each drive
$path2List = @()
foreach ($drive in $drives) {
    $path2List += Join-Path $drive.Root "Kingsoft\WPS Office"
}

# Create all directories
md $path1 -Force | Out-Null
foreach ($path2 in $path2List) {
    md $path2 -Force | Out-Null
}

# Define deny rule
$denyRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Everyone",
    "FullControl",
    "ContainerInherit,ObjectInherit",
    "None",
    "Deny"
)

# Apply deny rule to all folders
foreach ($folder in @($path1) + $path2List) {
    try {
        $acl = Get-Acl $folder
        $acl.AddAccessRule($denyRule)
        Set-Acl -Path $folder -AclObject $acl
        Write-Host "Access restricted to $folder"
    }
    catch {
        Write-Warning "Failed to set ACL on $folder: $_"
    }
}