$CopyScriptPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Policies\Scripts"

Copy-Item -Path .\CopyAddins.ps1 -Destination $CopyScriptPath

# Define the task name and script path
$taskName = "CopyAddins"
$scriptPath = "$CopyScriptPath\CopyAddins.ps1"  # Get the full path of this script

# Define the trigger to run at logo
$trigger = New-ScheduledTaskTrigger -AtLogon

# Define the action to run PowerShell with the execution policy bypassed
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""

# Set the task to run with the logged-on user's privileges (instead of SYSTEM)
$principal = New-ScheduledTaskPrincipal -UserId "CurrentUser" -LogonType Interactive

# Register the task with the Task Scheduler
Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName $taskName -Description "Checks and copies addins when user logs on"
