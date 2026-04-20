# Datto RMM - Force Microsoft 365 Apps Update
# For Click-to-Run Microsoft 365 Apps / Office 365 Apps on Windows
# Exit codes:
#   0 = Triggered successfully
#   1 = Click-to-Run Office not found
#   2 = Office update scheduled task not found
#   3 = Failed to trigger scheduled task

$ErrorActionPreference = "Stop"

# Set to $true only during a maintenance window if you want to close Office apps first.
$ForceCloseOfficeApps = $false

$TaskName   = "\Microsoft\Office\Office Automatic Updates 2.0"
$ConfigKey  = 'HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$UpdatesKey = 'HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Updates'

$OfficeProcesses = @(
    "WINWORD",
    "EXCEL",
    "POWERPNT",
    "OUTLOOK",
    "ONENOTE",
    "MSACCESS",
    "VISIO",
    "LYNC",
    "MSPUB"
)

function Write-Log {
    param([string]$Message)
    Write-Output ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}

function Get-RegValue64 {
    param(
        [string]$Key,
        [string]$ValueName
    )

    $output = & reg.exe query $Key /v $ValueName /reg:64 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $output) {
        return $null
    }

    foreach ($line in $output) {
        if ($line -match "^\s*$([regex]::Escape($ValueName))\s+REG_\w+\s+(.*)$") {
            return $matches[1].Trim()
        }
    }

    return $null
}

try {
    $svc = Get-Service -Name "ClickToRunSvc" -ErrorAction Stop
}
catch {
    Write-Log "ClickToRunSvc not found. This device does not appear to have Click-to-Run Microsoft 365 Apps installed."
    exit 1
}

Write-Log "Found ClickToRunSvc. Current state: $($svc.Status)"

if ($svc.Status -ne "Running") {
    Write-Log "Starting ClickToRunSvc..."
    Start-Service -Name "ClickToRunSvc"
    Start-Sleep -Seconds 5
}

$versionBefore    = Get-RegValue64 -Key $ConfigKey  -ValueName "ClientVersionToReport"
$channelBefore    = Get-RegValue64 -Key $ConfigKey  -ValueName "UpdateChannel"
$cdnBaseUrl       = Get-RegValue64 -Key $ConfigKey  -ValueName "CDNBaseUrl"
$lastDetectBefore = Get-RegValue64 -Key $UpdatesKey -ValueName "UpdateDetectionLastRunTime"

if ($versionBefore)    { Write-Log "Installed Office version: $versionBefore" }
if ($channelBefore)    { Write-Log "Update channel: $channelBefore" }
if ($cdnBaseUrl)       { Write-Log "CDN base URL: $cdnBaseUrl" }
if ($lastDetectBefore) { Write-Log "Previous detection timestamp: $lastDetectBefore" }

# Confirm the Office scheduled task exists
& schtasks.exe /Query /TN $TaskName 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Log "Scheduled task not found: $TaskName"
    exit 2
}

if ($ForceCloseOfficeApps) {
    Write-Log "ForceCloseOfficeApps is enabled. Stopping running Office apps..."
    foreach ($procName in $OfficeProcesses) {
        $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
        foreach ($proc in $procs) {
            Write-Log "Stopping $($proc.ProcessName) (PID $($proc.Id))"
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
    }
}
else {
    Write-Log "ForceCloseOfficeApps is disabled. If Office apps are open, the update may finish later."
}

# Force a fresh detection cycle
Write-Log "Clearing UpdateDetectionLastRunTime..."
& reg.exe add $UpdatesKey /v UpdateDetectionLastRunTime /t REG_SZ /d "" /f /reg:64 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Log "Warning: Could not clear UpdateDetectionLastRunTime. Continuing anyway."
}

# Trigger the Office automatic update task
Write-Log "Triggering scheduled task: $TaskName"
$runOutput = & schtasks.exe /Run /TN $TaskName 2>&1
$runExit   = $LASTEXITCODE

foreach ($line in $runOutput) {
    if ($line) { Write-Log $line }
}

if ($runExit -ne 0) {
    Write-Log "Failed to trigger scheduled task."
    exit 3
}

Start-Sleep -Seconds 15

$lastDetectAfter = Get-RegValue64 -Key $UpdatesKey -ValueName "UpdateDetectionLastRunTime"
$versionAfter    = Get-RegValue64 -Key $ConfigKey  -ValueName "ClientVersionToReport"

if ($lastDetectAfter) {
    Write-Log "Current detection timestamp: $lastDetectAfter"
}
else {
    Write-Log "Detection timestamp is still blank. Office may still be processing in the background."
}

if ($versionAfter) {
    Write-Log "Current reported Office version: $versionAfter"
}

Write-Log "Office update task triggered successfully. Microsoft 365 Apps may continue downloading and applying updates in the background."
exit 0