<#
    Powershell script to check whether a Windows Update is pending
    Notifies the user saying the update will be applied and the system will reboot in 5 minutes
    If the user clicks "OK", the update will be applied and the system will reboot in 5 minutes
    The user can choose when to appky the update from a set of dropdown options (e.g. "Apply now", " Apply in 10 minutes", "Apply in 1 hour")

    This script follows similar logic as in Browser-Force-Restart.ps1, but is adapted to check for pending Windows Updates and apply them accordingly.
#>

# Function to check for pending Windows Updates
function Check-PendingUpdates {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $pendingUpdates = $updateSearcher.Search("IsInstalled=0").Updates
    return $pendingUpdates.Count -gt 0
}
# Function to show a notification to the user
function Show-Notification {
    $message = "A Windows Update is pending. The system will reboot in 5 minutes. Please save your work."
    $title = "Windows Update Pending"
    $options = [System.Windows.Forms.MessageBoxButtons]::OKCancel