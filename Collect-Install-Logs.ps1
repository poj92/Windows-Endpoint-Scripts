#Author header added
<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-05
Email: Peter.James@nexusos.co.uk
#>

# Number of hours to look back
# Example: 24 = last 24 hours, 48 = last 48 hours
# Set Datto variable StartTime
[int]$StartTime = 24

# Convert hours into an actual DateTime for Get-WinEvent
$EventStartTime = (Get-Date).AddHours(-$StartTime)

$EventIds = 11707,11708,11724,11725,1033,1034

$Events = Get-WinEvent -FilterHashtable @{
    LogName      = 'Application'
    ProviderName = 'MsiInstaller'
    StartTime    = $EventStartTime
    Id           = $EventIds
} -ErrorAction SilentlyContinue

if (-not $Events) {
    Write-Output "No application install or uninstall events found in the last $StartTime hours."
    exit 0
}

$Events |
    Select-Object `
        TimeCreated,
        Id,
        ProviderName,
        MachineName,
        @{Name='Action';Expression={
            switch ($_.Id) {
                11707 { 'Install successful' }
                11708 { 'Install failed' }
                11724 { 'Uninstall successful' }
                11725 { 'Uninstall failed' }
                1033  { 'Product installed' }
                1034  { 'Product removed' }
                default { 'Other' }
            }
        }},
        @{Name='Message';Expression={$_.Message}} |
    Sort-Object TimeCreated -Descending |
    Format-List

exit 0
