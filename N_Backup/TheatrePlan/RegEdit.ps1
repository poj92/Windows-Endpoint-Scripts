# Set the shared Office templates path
# $templatePath = "C:\Users\'+$env:USERNAME+'\Theatreplan\Templates - Documents\THP Office 365 Templates"
# $templatePathEncoded = $templatePath -replace '\\','\\'  # Just to make sure for logging (not used in binary reg write)

# Convert string path to REG_EXPAND_SZ binary format (hex(2) in .reg file)
# $bytes = ([System.Text.Encoding]::Unicode.GetBytes($templatePath + [char]0))

# New-Item -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common' -Name 'General' -Force
# Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\General' -Name 'SharedTemplates' -Value $bytes -Type ExpandString

# Set officestartdefaulttab = 1 for Word, Excel, PowerPoint
$officeApps = @("Word", "Excel", "PowerPoint")
foreach ($app in $officeApps) {
    $path = "HKCU:\Software\Microsoft\Office\16.0\$app\Options"
    New-Item -Path $path -Force
    New-ItemProperty -Path $path -Name 'officestartdefaulttab' -PropertyType DWord -Value 1 -Force
}

# Set file extensions visible and show hidden files in Explorer
$explorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
New-Item -Path $explorerPath -Force
Set-ItemProperty -Path $explorerPath -Name "HideFileExt" -Value 0
Set-ItemProperty -Path $explorerPath -Name "Hidden" -Value 1

# Revert Windows 11 right-click menu to Windows 10 style
$clsidPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
New-Item -Path $clsidPath -Force
Set-ItemProperty -Path $clsidPath -Name '(default)' -Value ''

Write-Host "Registry settings applied successfully."