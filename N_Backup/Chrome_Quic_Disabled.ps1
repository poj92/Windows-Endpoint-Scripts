# Define the registry path for Local Machine (applies to all users)
$regPath = "HKLM:\Software\Policies\Google\Chrome"

# Define the registry keys
$regKeyECH = "EncryptedClientHelloEnabled"
$regKeyQUIC = "QuicAllowed"

# Function to create registry key and set value
function Set-RegistryKey {
    param (
        [string]$key,
        [int]$value
    )

    # Check if the parent registry path exists
    if (-not (Test-Path "$regPath")) {
        Write-Host "Registry path '$regPath' does not exist. Creating the path."
        # Create the full registry path if it doesn't exist
        New-Item -Path $regPath -Force
    }

    # Check if the registry key exists
    if (Test-Path "$regPath\$key") {
        Write-Host "Registry key '$key' already exists. Setting value to $value."
    } else {
        Write-Host "Registry key '$key' does not exist. Creating key."
        # Set the registry value (create the key)
        Set-ItemProperty -Path $regPath -Name $key -Value $value -Type DWord
    }

    # Set the registry value (0 = disabled)
    Set-ItemProperty -Path $regPath -Name $key -Value $value -Type DWord
    Write-Host "Registry key '$key' is set to $value."
}

# Disable Encrypted Client Hello and QUIC by setting them to 0
Set-RegistryKey -key $regKeyECH -value 0
Set-RegistryKey -key $regKeyQUIC -value 0
