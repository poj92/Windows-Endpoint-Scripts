<#
.SYNOPSIS
    This script is used to uninstall the ConnectSecure Agent from a Windows machine.
    Datto RMM will require a variable named "cyberKey" to be set with the appropriate uninstallation key for the agent.
#>

cd "C:\Program Files (x86)\CyberCNSAgent"
.\cybercnsagent.exe -r -y $env:cyberKey