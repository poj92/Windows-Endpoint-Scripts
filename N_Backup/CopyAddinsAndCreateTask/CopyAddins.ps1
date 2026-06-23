# PowerShell script to copy addins

# Define source and destination paths for the files
$sourcePathWord = "C:\Users\$env:USERNAME\Theatreplan\Templates - Documents\THP Office 365 Templates\Word Startup\TP_Tools.dotm"
$destinationPathWord = "C:\Program Files\Microsoft Office\root\Office16\STARTUP"

$sourcePathExcel = "C:\Users\$env:USERNAME\Theatreplan\Templates - Documents\THP Office 365 Templates\Excel Add-ins\TP_Tools.xlam"
$destinationPathExcel = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\AddIns"

# Check if the Word template file exists and copy it using Copy-Item if it does
if (Test-Path $sourcePathWord) {
    # Copy the file, overwrite if exists
    Copy-Item -Path $sourcePathWord -Destination $destinationPathWord -Force
    Write-Host "Copied Word template to startup folder (replaced if exists)."
} else {
    Write-Host "Word Template file does not exist: $sourcePathWord"
}

# Check if the Excel Add-in file exists and copy it using Copy-Item if it does
if (Test-Path $sourcePathExcel) {
    # Copy the file, overwrite if exists
    Copy-Item -Path $sourcePathExcel -Destination $destinationPathExcel -Force
    Write-Host "Copied Excel add-in to AddIns folder (replaced if exists)."
} else {
    Write-Host "Excel Add-in file does not exist: $sourcePathExcel"
}
