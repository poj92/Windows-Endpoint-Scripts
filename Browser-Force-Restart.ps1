#Requires -Version 5.1
param(
    [string]$ScheduledTaskName
)

$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================
$BasePath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = Join-Path $BasePath "Remediation.log"
$TrackingFile = Join-Path $BasePath "BrowserUsageTracking.json"
$QueueFile = Join-Path $BasePath "ReloadQueue.json"
$WarningTimeSeconds = 300
$CompanyName = "Nexus Open Systems Ltd"
$RemediationScriptPath = "C:\ProgramData\Datto\BrowserUpdateCheck\BrowserReloadRemediation.ps1"

# =========================
# Bootstrap
# =========================
if (-not (Test-Path $BasePath)) {
    New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
}

function Write-LogEntry {
    param(
        [string]$Message,
        [ValidateSet("Information","Warning","Error")]
        [string]$Level = "Information"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    try {
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    }
    catch {}

    Write-Host $line
}

function Save-CurrentScriptToStablePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        $currentScriptPath = $PSCommandPath

        if ([string]::IsNullOrWhiteSpace($currentScriptPath) -or -not (Test-Path $currentScriptPath)) {
            throw "Unable to determine current script path."
        }

        $destinationFolder = Split-Path -Path $DestinationPath -Parent
        if (-not (Test-Path $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
        }

        Copy-Item -Path $currentScriptPath -Destination $DestinationPath -Force
        Write-LogEntry "Copied remediation script from '$currentScriptPath' to '$DestinationPath'"
        return $true
    }
    catch {
        Write-LogEntry "Failed to copy remediation script to stable path: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Remove-ScheduledTaskIfRequested {
    param(
        [string]$TaskName
    )

    if ([string]::IsNullOrWhiteSpace($TaskName)) {
        return
    }

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-LogEntry "Deleted scheduled task '$TaskName' after launch"
    }
    catch {
        Write-LogEntry "Failed to delete scheduled task '$TaskName' : $($_.Exception.Message)" "Warning"
    }
}

function Test-IsSystem {
    try {
        return ([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')
    }
    catch {
        return $false
    }
}

function Get-ExecutionContextInfo {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $username = $identity.Name
        $sid = $identity.User.Value
        $sessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
        $processName = (Get-Process -Id $PID).ProcessName
        Write-LogEntry "Execution context: User=$username | SID=$sid | SessionId=$sessionId | Process=$processName"
    }
    catch {
        Write-LogEntry "Failed to collect execution context info: $($_.Exception.Message)" "Warning"
    }
}

function Get-CurrentUserSam {
    try {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        return $env:USERNAME
    }
}

function Get-BrowserProcessName {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    switch ($BrowserName) {
        "Chrome"  { "chrome" }
        "Firefox" { "firefox" }
        "Edge"    { "msedge" }
    }
}

function Test-BrowserRunning {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    $processName = Get-BrowserProcessName -BrowserName $BrowserName
    return [bool](Get-Process -Name $processName -ErrorAction SilentlyContinue)
}

function Get-ReloadQueue {
    if (-not (Test-Path $QueueFile)) {
        return $null
    }

    try {
        $queue = Get-Content -Path $QueueFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        if (-not ($queue.PSObject.Properties.Name -contains 'Browsers')) {
            $queue | Add-Member -MemberType NoteProperty -Name Browsers -Value @() -Force
        }

        foreach ($item in $queue.Browsers) {
            if (-not ($item.PSObject.Properties.Name -contains 'PostponeUntilUtc')) {
                $item | Add-Member -MemberType NoteProperty -Name PostponeUntilUtc -Value $null -Force
            }
            if (-not ($item.PSObject.Properties.Name -contains 'PostponeChoice')) {
                $item | Add-Member -MemberType NoteProperty -Name PostponeChoice -Value $null -Force
            }
            if (-not ($item.PSObject.Properties.Name -contains 'ScheduledTaskName')) {
                $item | Add-Member -MemberType NoteProperty -Name ScheduledTaskName -Value $null -Force
            }
        }

        return $queue
    }
    catch {
        Write-LogEntry "Queue file is unreadable." "Error"
        return $null
    }
}

function Save-ReloadQueue {
    param([array]$Browsers)

    $queue = [PSCustomObject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        Browsers   = $Browsers
    }

    $queue | ConvertTo-Json -Depth 8 | Set-Content -Path $QueueFile -Force -Encoding UTF8
}

function Get-BrowserUsageTracking {
    if (Test-Path $TrackingFile) {
        try {
            $loaded = Get-Content -Path $TrackingFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            foreach ($browser in @('Chrome','Firefox','Edge')) {
                if (-not ($loaded.PSObject.Properties.Name -contains $browser)) {
                    Add-Member -InputObject $loaded -MemberType NoteProperty -Name $browser -Value ([PSCustomObject]@{
                        LastStart = $null
                        LastStop  = $null
                        IsRunning = $false
                    }) -Force
                }

                $browserObj = $loaded.$browser

                if (-not ($browserObj.PSObject.Properties.Name -contains 'LastStart')) {
                    $migratedLastStart = $null
                    if ($browserObj.PSObject.Properties.Name -contains 'SessionStart') {
                        $migratedLastStart = $browserObj.SessionStart
                    }
                    Add-Member -InputObject $browserObj -MemberType NoteProperty -Name LastStart -Value $migratedLastStart -Force
                }

                if (-not ($browserObj.PSObject.Properties.Name -contains 'LastStop')) {
                    Add-Member -InputObject $browserObj -MemberType NoteProperty -Name LastStop -Value $null -Force
                }

                if (-not ($browserObj.PSObject.Properties.Name -contains 'IsRunning')) {
                    Add-Member -InputObject $browserObj -MemberType NoteProperty -Name IsRunning -Value $false -Force
                }
            }

            return $loaded
        }
        catch {
            Write-LogEntry "Tracking file is invalid. Creating minimal structure." "Warning"
        }
    }

    return [PSCustomObject]@{
        Chrome  = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
        Firefox = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
        Edge    = [PSCustomObject]@{ LastStart = $null; LastStop = $null; IsRunning = $false }
    }
}

function Save-BrowserUsageTracking {
    param([object]$TrackingData)
    $TrackingData | ConvertTo-Json -Depth 5 | Set-Content -Path $TrackingFile -Force -Encoding UTF8
}

function Get-ChromeInstallPath {
    @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-EdgeInstallPath {
    @(
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-FirefoxInstallPath {
    @(
        "C:\Program Files\Mozilla Firefox\firefox.exe",
        "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-BrowserExecutablePath {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    switch ($BrowserName) {
        "Chrome"  { Get-ChromeInstallPath }
        "Firefox" { Get-FirefoxInstallPath }
        "Edge"    { Get-EdgeInstallPath }
    }
}

function New-UniqueTaskName {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    "NexusBrowserReload_{0}_{1}" -f $BrowserName, ([guid]::NewGuid().ToString("N").Substring(0,8))
}

function Register-PostponeScheduledTask {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName,
        [datetime]$RunAtLocal,
        [string]$ScriptPath
    )

    try {
        if (-not (Test-Path $ScriptPath)) {
            throw "Remediation script path not found: $ScriptPath"
        }

        $taskName = New-UniqueTaskName -BrowserName $BrowserName
        $currentUser = Get-CurrentUserSam

        $action = New-ScheduledTaskAction `
            -Execute "powershell.exe" `
            -Argument "-ExecutionPolicy Bypass -WindowStyle Normal -File `"$ScriptPath`" -ScheduledTaskName `"$taskName`""

        $trigger = New-ScheduledTaskTrigger -Once -At $RunAtLocal

        $principal = New-ScheduledTaskPrincipal `
            -UserId $currentUser `
            -LogonType Interactive `
            -RunLevel Limited

        $settings = New-ScheduledTaskSettingsSet `
            -StartWhenAvailable `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 1)

        $task = New-ScheduledTask `
            -Action $action `
            -Trigger $trigger `
            -Principal $principal `
            -Settings $settings

        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

        try {
            $null = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
            Write-LogEntry "Verified scheduled task '$taskName' exists"
        }
        catch {
            Write-LogEntry "Scheduled task '$taskName' could not be verified after creation" "Warning"
        }

        Write-LogEntry "Scheduled task '$taskName' created for $BrowserName at $($RunAtLocal.ToString('yyyy-MM-dd HH:mm:ss')) as $currentUser"
        return $taskName
    }
    catch {
        Write-LogEntry "Failed to create scheduled task for $BrowserName : $($_.Exception.Message)" "Error"
        return $null
    }
}

function Show-CountdownWarning {
    param(
        [string]$BrowserName,
        [int]$CountdownSeconds,
        [string]$CompanyName
    )

    try {
        Write-LogEntry "Preparing popup for $BrowserName"

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $script:UserDecision = "Restart"
        $script:SelectedPostponeMinutes = 60
        $script:SecondsLeft = $CountdownSeconds

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "$CompanyName - $BrowserName restart required"
        $form.Size = New-Object System.Drawing.Size(680,360)
        $form.StartPosition = 'CenterScreen'
        $form.TopMost = $true
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.ShowInTaskbar = $true

        $headerLabel = New-Object System.Windows.Forms.Label
        $headerLabel.Location = New-Object System.Drawing.Point(20,20)
        $headerLabel.Size = New-Object System.Drawing.Size(620,25)
        $headerLabel.Text = "Message from $CompanyName"
        $headerLabel.Font = New-Object System.Drawing.Font('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
        $form.Controls.Add($headerLabel)

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(20,55)
        $label.Size = New-Object System.Drawing.Size(620,70)
        $label.Text = "$CompanyName needs to restart $BrowserName to complete a pending security and stability update. The restart will happen automatically in:"
        $label.Font = New-Object System.Drawing.Font('Segoe UI',10)
        $form.Controls.Add($label)

        $countdownLabel = New-Object System.Windows.Forms.Label
        $countdownLabel.Location = New-Object System.Drawing.Point(20,125)
        $countdownLabel.Size = New-Object System.Drawing.Size(620,40)
        $countdownLabel.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
        $countdownLabel.TextAlign = 'MiddleCenter'
        $form.Controls.Add($countdownLabel)

        $postponeLabel = New-Object System.Windows.Forms.Label
        $postponeLabel.Location = New-Object System.Drawing.Point(115,185)
        $postponeLabel.Size = New-Object System.Drawing.Size(190,25)
        $postponeLabel.Text = "Postpone restart for:"
        $postponeLabel.Font = New-Object System.Drawing.Font('Segoe UI',10)
        $form.Controls.Add($postponeLabel)

        $comboBox = New-Object System.Windows.Forms.ComboBox
        $comboBox.Location = New-Object System.Drawing.Point(305,182)
        $comboBox.Size = New-Object System.Drawing.Size(180,25)
        $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$comboBox.Items.Add("5 minutes")
        [void]$comboBox.Items.Add("30 minutes")
        [void]$comboBox.Items.Add("1 hour")
        [void]$comboBox.Items.Add("2 hours")
        [void]$comboBox.Items.Add("4 hours")
        $comboBox.SelectedIndex = 2
        $form.Controls.Add($comboBox)

        $restartNowButton = New-Object System.Windows.Forms.Button
        $restartNowButton.Location = New-Object System.Drawing.Point(150,235)
        $restartNowButton.Size = New-Object System.Drawing.Size(140,35)
        $restartNowButton.Text = 'Restart Now'
        $restartNowButton.Add_Click({
            $script:UserDecision = "RestartNow"
            $form.Close()
        })
        $form.Controls.Add($restartNowButton)

        $postponeButton = New-Object System.Windows.Forms.Button
        $postponeButton.Location = New-Object System.Drawing.Point(330,235)
        $postponeButton.Size = New-Object System.Drawing.Size(160,35)
        $postponeButton.Text = 'Postpone Restart'
        $postponeButton.Add_Click({
            switch ($comboBox.SelectedItem) {
                "5 minutes"  { $script:SelectedPostponeMinutes = 5 }
                "30 minutes" { $script:SelectedPostponeMinutes = 30 }
                "1 hour"     { $script:SelectedPostponeMinutes = 60 }
                "2 hours"    { $script:SelectedPostponeMinutes = 120 }
                "4 hours"    { $script:SelectedPostponeMinutes = 240 }
                default      { $script:SelectedPostponeMinutes = 60 }
            }
            $script:UserDecision = "Postpone"
            $form.Close()
        })
        $form.Controls.Add($postponeButton)

        $footerLabel = New-Object System.Windows.Forms.Label
        $footerLabel.Location = New-Object System.Drawing.Point(20,295)
        $footerLabel.Size = New-Object System.Drawing.Size(620,20)
        $footerLabel.Text = "You can restart now or postpone this restart for a limited time."
        $footerLabel.Font = New-Object System.Drawing.Font('Segoe UI',9)
        $form.Controls.Add($footerLabel)

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 1000
        $timer.Add_Tick({
            $minutes = [math]::Floor($script:SecondsLeft / 60)
            $seconds = $script:SecondsLeft % 60
            $countdownLabel.Text = "{0}:{1:D2}" -f $minutes, $seconds
            $script:SecondsLeft--

            if ($script:SecondsLeft -lt 0) {
                $timer.Stop()
                $script:UserDecision = "CountdownRestart"
                $form.Close()
            }
        })

        $form.Add_Shown({
            Write-LogEntry "Popup shown for $BrowserName"
            $form.Activate()
        })

        $timer.Start()
        [void]$form.ShowDialog()

        [PSCustomObject]@{
            Action          = $script:UserDecision
            PostponeMinutes = $script:SelectedPostponeMinutes
        }
    }
    catch {
        Write-LogEntry "Popup failed for $BrowserName : $($_.Exception.Message)" "Error"
        [PSCustomObject]@{
            Action          = "Restart"
            PostponeMinutes = 60
        }
    }
}

function Invoke-BrowserReload {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    try {
        $exe = Get-BrowserExecutablePath -BrowserName $BrowserName
        if (-not $exe) {
            throw "$BrowserName executable not found"
        }

        $processName = Get-BrowserProcessName -BrowserName $BrowserName

        if (-not (Test-BrowserRunning -BrowserName $BrowserName)) {
            Write-LogEntry "$BrowserName is no longer running. Marking queue item complete without restart."
            return $true
        }

        Write-LogEntry "Attempting graceful close of $BrowserName processes"
        Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $_.CloseMainWindow() | Out-Null
            }
            catch {}
        }

        Start-Sleep -Seconds 10

        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            Write-LogEntry "$BrowserName still has running processes after graceful close; forcing termination" "Warning"
            Get-Process -Name $processName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }

        Write-LogEntry "Starting $BrowserName from '$exe'"
        Start-Process -FilePath $exe | Out-Null

        Start-Sleep -Seconds 2

        if (Test-BrowserRunning -BrowserName $BrowserName) {
            Write-LogEntry "$BrowserName restarted successfully"
            return $true
        }

        throw "$BrowserName did not appear to restart successfully"
    }
    catch {
        Write-LogEntry "$BrowserName restart failed: $($_.Exception.Message)" "Error"
        return $false
    }
}

try {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload remediation started"
    Write-LogEntry "======================================"

    Get-ExecutionContextInfo
    Remove-ScheduledTaskIfRequested -TaskName $ScheduledTaskName

    if (-not (Save-CurrentScriptToStablePath -DestinationPath $RemediationScriptPath)) {
        Write-LogEntry "Continuing, but postponed scheduled tasks may fail because the stable script copy is missing." "Warning"
    }

    if (Test-IsSystem) {
        Write-LogEntry "Remediation script is running as SYSTEM. This script must run as the logged-in user." "Error"
        exit 2
    }

    $queue = Get-ReloadQueue
    if (-not $queue -or -not $queue.Browsers -or $queue.Browsers.Count -eq 0) {
        Write-LogEntry "No queued browsers found. Nothing to do."
        exit 0
    }

    Write-LogEntry "Found $($queue.Browsers.Count) queued browser(s) for restart"

    $remainingQueue = @()
    $nowUtc = (Get-Date).ToUniversalTime()

    foreach ($item in $queue.Browsers) {
        $browser = $item.Browser

        if (-not ($item.PSObject.Properties.Name -contains 'PostponeUntilUtc')) {
            $item | Add-Member -MemberType NoteProperty -Name PostponeUntilUtc -Value $null -Force
        }
        if (-not ($item.PSObject.Properties.Name -contains 'PostponeChoice')) {
            $item | Add-Member -MemberType NoteProperty -Name PostponeChoice -Value $null -Force
        }
        if (-not ($item.PSObject.Properties.Name -contains 'ScheduledTaskName')) {
            $item | Add-Member -MemberType NoteProperty -Name ScheduledTaskName -Value $null -Force
        }

        if ($item.PostponeUntilUtc) {
            try {
                $postponeUntil = [datetime]::Parse($item.PostponeUntilUtc).ToUniversalTime()
                if ($postponeUntil -gt $nowUtc) {
                    Write-LogEntry "$browser is postponed until $($postponeUntil.ToString("o")); skipping for now"
                    $remainingQueue += $item
                    continue
                }
            }
            catch {
                Write-LogEntry "Invalid PostponeUntilUtc for $browser; ignoring postpone value" "Warning"
                $item.PostponeUntilUtc = $null
                $item.PostponeChoice = $null
                $item.ScheduledTaskName = $null
            }
        }

        Write-LogEntry "Processing $browser. Reason: $($item.Reason)"
        Write-LogEntry "Displaying popup for $browser with $WarningTimeSeconds second countdown"

        $decision = Show-CountdownWarning -BrowserName $browser -CountdownSeconds $WarningTimeSeconds -CompanyName $CompanyName
        Write-LogEntry "Popup closed for $browser. Action=$($decision.Action) PostponeMinutes=$($decision.PostponeMinutes)"

        if ($decision.Action -eq "Postpone") {
            $runAtLocal = (Get-Date).AddMinutes([int]$decision.PostponeMinutes)

            $updatedItem = [PSCustomObject]@{
                Browser           = $item.Browser
                Reason            = $item.Reason
                PostponeUntilUtc  = $runAtLocal.ToUniversalTime().ToString("o")
                PostponeChoice    = "$($decision.PostponeMinutes) minutes"
                ScheduledTaskName = $null
            }

            $taskName = Register-PostponeScheduledTask -BrowserName $browser -RunAtLocal $runAtLocal -ScriptPath $RemediationScriptPath

            if ($taskName) {
                $updatedItem.ScheduledTaskName = $taskName
                Write-LogEntry "$browser postponed until $($updatedItem.PostponeUntilUtc) ($($updatedItem.PostponeChoice)); scheduled task '$taskName' created"
            }
            else {
                Write-LogEntry "Scheduled task creation failed for $browser; leaving item queued for retry" "Warning"
            }

            $remainingQueue += $updatedItem
            continue
        }

        if (Invoke-BrowserReload -BrowserName $browser) {
            $tracking = Get-BrowserUsageTracking
            $tracking.$browser.LastStart = (Get-Date).ToString("o")
            $tracking.$browser.IsRunning = $true
            Save-BrowserUsageTracking -TrackingData $tracking
        }
        else {
            $remainingQueue += $item
        }
    }

    Save-ReloadQueue -Browsers $remainingQueue

    try {
        $savedQueueCheck = Get-Content -Path $QueueFile -Raw -ErrorAction Stop
        Write-LogEntry "Queue file saved successfully: $savedQueueCheck"
    }
    catch {
        Write-LogEntry "Unable to re-read queue file after saving: $($_.Exception.Message)" "Warning"
    }

    if ($remainingQueue.Count -gt 0) {
        Write-LogEntry "$($remainingQueue.Count) browser(s) remain queued for retry"
        exit 3
    }

    Write-LogEntry "Queue cleared"
    exit 0
}
catch {
    Write-LogEntry "Fatal remediation error: $($_.Exception.Message)" "Error"
    exit 1
}
finally {
    Write-LogEntry "======================================"
    Write-LogEntry "Browser reload remediation finished"
    Write-LogEntry "======================================"
}