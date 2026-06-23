$AccessPath = "HKLM:\Software\Autodesk\ODIS"
$KeyName = "AllowSystemContextInstall"

if (Test-Path $AccessPath) {
    Write-Host "Path exists"

    $existingKey = Get-ItemProperty -Path $AccessPath -Name $KeyName -ErrorAction SilentlyContinue

    if ($existingKey) {
        Set-ItemProperty -Path $AccessPath -Name $KeyName -Value "1"
        Write-Host "Key exists. Value set to '1' (REG_SZ)."
    }
    else {
        New-ItemProperty -Path $AccessPath -Name $KeyName -PropertyType String -Value "1" -Force
        Write-Host "Key did not exist. Created with value '1' (REG_SZ)."
    }
}
else {
    New-Item -Path $AccessPath -Force | Out-Null
    New-ItemProperty -Path $AccessPath -Name $KeyName -PropertyType String -Value "1" -Force
    Write-Host "Path and key created with value '1' (REG_SZ)."
}
