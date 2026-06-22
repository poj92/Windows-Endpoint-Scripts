#Requires -Version 5.1

<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-22
Email: Peter.James@nexusos.co.uk
#>


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

$ClosedBrowserUpdateWaitSeconds = 60
$ClosedBrowserUpdateWaitSecondsWarning = $null
$rawClosedBrowserUpdateWaitSeconds = [Environment]::GetEnvironmentVariable("ClosedBrowserUpdateWaitSeconds")
if (-not [string]::IsNullOrWhiteSpace($rawClosedBrowserUpdateWaitSeconds)) {
    $parsedClosedBrowserUpdateWaitSeconds = 0
    if ([int]::TryParse($rawClosedBrowserUpdateWaitSeconds.Trim(), [ref]$parsedClosedBrowserUpdateWaitSeconds) -and $parsedClosedBrowserUpdateWaitSeconds -gt 0) {
        $ClosedBrowserUpdateWaitSeconds = $parsedClosedBrowserUpdateWaitSeconds
    }
    else {
        $ClosedBrowserUpdateWaitSecondsWarning = "Environment variable 'ClosedBrowserUpdateWaitSeconds' has invalid value '$rawClosedBrowserUpdateWaitSeconds'. Using default value $ClosedBrowserUpdateWaitSeconds."
    }
}

$BrowserUpdateEngineWaitSeconds = 180
$BrowserUpdateEngineWaitSecondsWarning = $null
$rawBrowserUpdateEngineWaitSeconds = [Environment]::GetEnvironmentVariable("BrowserUpdateEngineWaitSeconds")
if (-not [string]::IsNullOrWhiteSpace($rawBrowserUpdateEngineWaitSeconds)) {
    $parsedBrowserUpdateEngineWaitSeconds = 0
    if ([int]::TryParse($rawBrowserUpdateEngineWaitSeconds.Trim(), [ref]$parsedBrowserUpdateEngineWaitSeconds) -and $parsedBrowserUpdateEngineWaitSeconds -gt 0) {
        $BrowserUpdateEngineWaitSeconds = $parsedBrowserUpdateEngineWaitSeconds
    }
    else {
        $BrowserUpdateEngineWaitSecondsWarning = "Environment variable 'BrowserUpdateEngineWaitSeconds' has invalid value '$rawBrowserUpdateEngineWaitSeconds'. Using default value $BrowserUpdateEngineWaitSeconds."
    }
}

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
            if (-not ($item.PSObject.Properties.Name -contains 'RemediationMode')) {
                $item | Add-Member -MemberType NoteProperty -Name RemediationMode -Value $null -Force
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

    try {
        $queue = [PSCustomObject]@{
            CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
            Browsers   = $Browsers
        }

        $queue | ConvertTo-Json -Depth 8 | Set-Content -Path $QueueFile -Force -Encoding UTF8
        return $true
    }
    catch {
        Write-LogEntry "Failed to save reload queue '$QueueFile': $($_.Exception.Message)" "Error"
        return $false
    }
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

    try {
        $TrackingData | ConvertTo-Json -Depth 5 | Set-Content -Path $TrackingFile -Force -Encoding UTF8
        return $true
    }
    catch {
        Write-LogEntry "Failed to save tracking file '$TrackingFile': $($_.Exception.Message)" "Warning"
        return $false
    }
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


function Get-FileVersionStringSafe {
    param([string]$Path)

    try {
        if (Test-Path $Path) {
            $version = (Get-Item -Path $Path -ErrorAction Stop).VersionInfo.ProductVersion
            if ([string]::IsNullOrWhiteSpace($version)) {
                $version = (Get-Item -Path $Path -ErrorAction Stop).VersionInfo.FileVersion
            }
            if (-not [string]::IsNullOrWhiteSpace($version)) {
                return ($version -replace '[^0-9\.]', '').Trim('.')
            }
        }
    }
    catch {}

    return $null
}

function Get-RegistryValueSafe {
    param(
        [string]$Path,
        [string]$Name
    )

    try {
        if (Test-Path $Path) {
            $value = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                return [string]$value
            }
        }
    }
    catch {}

    return $null
}

function Get-HighestVersionFolderSafe {
    param([string[]]$Paths)

    $versions = @()
    foreach ($path in $Paths) {
        try {
            if (Test-Path $path) {
                Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $parsed = $null
                    if ([version]::TryParse($_.Name, [ref]$parsed)) {
                        $versions += [PSCustomObject]@{
                            Version = $parsed
                            Text    = $_.Name
                        }
                    }
                }
            }
        }
        catch {}
    }

    if ($versions.Count -gt 0) {
        return ($versions | Sort-Object Version -Descending | Select-Object -First 1).Text
    }

    return $null
}

function Get-BrowserInstalledVersion {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    $candidates = @()

    switch ($BrowserName) {
        "Chrome" {
            $registryCandidates = @(
                @{ Path = 'HKLM:\SOFTWARE\Google\Chrome\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon'; Name = 'version' },
                @{ Path = 'HKCU:\SOFTWARE\Google\Chrome\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome'; Name = 'DisplayVersion' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome'; Name = 'DisplayVersion' },
                @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome'; Name = 'DisplayVersion' }
            )

            foreach ($candidate in $registryCandidates) {
                $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via registry $($candidate.Path):$($candidate.Name): $value"
                    return $value
                }
            }

            foreach ($exe in @(
                'C:\Program Files\Google\Chrome\Application\chrome.exe',
                'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
            )) {
                $value = Get-FileVersionStringSafe -Path $exe
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via executable '$exe': $value"
                    return $value
                }
            }

            $folderVersion = Get-HighestVersionFolderSafe -Paths @(
                'C:\Program Files\Google\Chrome\Application',
                'C:\Program Files (x86)\Google\Chrome\Application'
            )
            if ($folderVersion) {
                Write-LogEntry "$BrowserName installed version detected via application folder: $folderVersion"
                return $folderVersion
            }
        }

        "Edge" {
            $registryCandidates = @(
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon'; Name = 'version' },
                @{ Path = 'HKCU:\SOFTWARE\Microsoft\Edge\BLBeacon'; Name = 'version' },
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge'; Name = 'DisplayVersion' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge'; Name = 'DisplayVersion' },
                @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge'; Name = 'DisplayVersion' }
            )

            foreach ($candidate in $registryCandidates) {
                $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via registry $($candidate.Path):$($candidate.Name): $value"
                    return $value
                }
            }

            foreach ($exe in @(
                'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
                'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
            )) {
                $value = Get-FileVersionStringSafe -Path $exe
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via executable '$exe': $value"
                    return $value
                }
            }

            $folderVersion = Get-HighestVersionFolderSafe -Paths @(
                'C:\Program Files\Microsoft\Edge\Application',
                'C:\Program Files (x86)\Microsoft\Edge\Application'
            )
            if ($folderVersion) {
                Write-LogEntry "$BrowserName installed version detected via application folder: $folderVersion"
                return $folderVersion
            }
        }

        "Firefox" {
            $registryCandidates = @(
                @{ Path = 'HKLM:\SOFTWARE\Mozilla\Mozilla Firefox'; Name = 'CurrentVersion' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Mozilla\Mozilla Firefox'; Name = 'CurrentVersion' },
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox'; Name = 'DisplayVersion' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox'; Name = 'DisplayVersion' },
                @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox'; Name = 'DisplayVersion' }
            )

            foreach ($candidate in $registryCandidates) {
                $value = Get-RegistryValueSafe -Path $candidate.Path -Name $candidate.Name
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via registry $($candidate.Path):$($candidate.Name): $value"
                    return $value
                }
            }

            foreach ($exe in @(
                'C:\Program Files\Mozilla Firefox\firefox.exe',
                'C:\Program Files (x86)\Mozilla Firefox\firefox.exe'
            )) {
                $value = Get-FileVersionStringSafe -Path $exe
                if ($value) {
                    Write-LogEntry "$BrowserName installed version detected via executable '$exe': $value"
                    return $value
                }
            }
        }
    }

    Write-LogEntry "$BrowserName installed version could not be determined by remediation script" "Warning"
    return $null
}

function Compare-VersionString {
    param(
        [string]$Left,
        [string]$Right
    )

    try {
        $leftVersion = $null
        $rightVersion = $null

        if (-not [version]::TryParse(($Left -replace '[^0-9\.]','').Trim('.'), [ref]$leftVersion)) {
            return $null
        }

        if (-not [version]::TryParse(($Right -replace '[^0-9\.]','').Trim('.'), [ref]$rightVersion)) {
            return $null
        }

        return $leftVersion.CompareTo($rightVersion)
    }
    catch {
        return $null
    }
}

function Invoke-ProcessWithTimeout {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string]$Arguments = '',

        [int]$TimeoutSeconds = 120,

        [string]$Description = $FilePath
    )

    try {
        if (-not (Test-Path $FilePath)) {
            return $false
        }

        Write-LogEntry "Starting update trigger: $Description $Arguments"

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $FilePath
        $startInfo.Arguments = $Arguments
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo

        [void]$process.Start()
        $completed = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $completed) {
            Write-LogEntry "Update trigger timed out after $TimeoutSeconds seconds: $Description" "Warning"
            try { $process.Kill() } catch {}
            return $true
        }

        Write-LogEntry "Update trigger exited with code $($process.ExitCode): $Description"
        return $true
    }
    catch {
        Write-LogEntry "Failed to run update trigger '$Description': $($_.Exception.Message)" "Warning"
        return $false
    }
}

function Start-BrowserUpdateScheduledTasks {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    $started = 0

    try {
        $taskNamePatterns = switch ($BrowserName) {
            "Chrome"  { @('*GoogleUpdate*UA*', '*GoogleUpdate*Core*') }
            "Edge"    { @('*MicrosoftEdgeUpdate*UA*', '*MicrosoftEdgeUpdate*Core*') }
            "Firefox" { @('*Firefox*Update*', '*Mozilla*Update*') }
        }

        $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $task = $_
            $matched = $false
            foreach ($pattern in $taskNamePatterns) {
                if ($task.TaskName -like $pattern -or $task.TaskPath -like $pattern) {
                    $matched = $true
                    break
                }
            }
            $matched
        })

        foreach ($task in $tasks) {
            try {
                Start-ScheduledTask -InputObject $task -ErrorAction Stop
                Write-LogEntry "Started scheduled update task '$($task.TaskPath)$($task.TaskName)' for $BrowserName"
                $started++
            }
            catch {
                Write-LogEntry "Could not start scheduled update task '$($task.TaskPath)$($task.TaskName)' for ${BrowserName}: $($_.Exception.Message)" "Warning"
            }
        }
    }
    catch {
        Write-LogEntry "Failed while enumerating update scheduled tasks for ${BrowserName}: $($_.Exception.Message)" "Warning"
    }

    if ($started -eq 0) {
        Write-LogEntry "No scheduled update tasks were started for $BrowserName" "Warning"
    }

    return ($started -gt 0)
}

function Invoke-BrowserUpdateEngine {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    $triggered = $false

    if (Start-BrowserUpdateScheduledTasks -BrowserName $BrowserName) {
        $triggered = $true
    }

    $updateExecutables = @()

    switch ($BrowserName) {
        "Chrome" {
            $updateExecutables = @(
                @{ Path = 'C:\Program Files (x86)\Google\Update\GoogleUpdate.exe'; Args = '/ua /installsource scheduler' },
                @{ Path = 'C:\Program Files\Google\Update\GoogleUpdate.exe'; Args = '/ua /installsource scheduler' },
                @{ Path = (Join-Path $env:LOCALAPPDATA 'Google\Update\GoogleUpdate.exe'); Args = '/ua /installsource scheduler' }
            )
        }
        "Edge" {
            $updateExecutables = @(
                @{ Path = 'C:\Program Files (x86)\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'; Args = '/ua /installsource scheduler' },
                @{ Path = 'C:\Program Files\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'; Args = '/ua /installsource scheduler' },
                @{ Path = (Join-Path $env:LOCALAPPDATA 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'); Args = '/ua /installsource scheduler' }
            )
        }
        "Firefox" {
            $updateExecutables = @()
        }
    }

    foreach ($entry in $updateExecutables) {
        if ($entry.Path -and (Test-Path $entry.Path)) {
            if (Invoke-ProcessWithTimeout -FilePath $entry.Path -Arguments $entry.Args -TimeoutSeconds 120 -Description "$BrowserName native update engine") {
                $triggered = $true
            }
        }
    }

    if (-not $triggered) {
        Write-LogEntry "No native update engine trigger was available or successful for $BrowserName" "Warning"
    }

    return $triggered
}

function Wait-BrowserVersionChange {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName,

        [string]$OriginalVersion,

        [int]$TimeoutSeconds,

        [string]$Reason
    )

    if ([string]::IsNullOrWhiteSpace($OriginalVersion)) {
        Write-LogEntry "Cannot wait for $BrowserName version change because original version is unknown" "Warning"
        return $false
    }

    if ($TimeoutSeconds -le 0) {
        return $false
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 10
        $currentVersion = Get-BrowserInstalledVersion -BrowserName $BrowserName
        $comparison = Compare-VersionString -Left $currentVersion -Right $OriginalVersion

        if ($comparison -ne $null -and $comparison -gt 0) {
            Write-LogEntry "$BrowserName version changed from $OriginalVersion to $currentVersion while waiting: $Reason"
            return $true
        }
    }

    $finalVersion = Get-BrowserInstalledVersion -BrowserName $BrowserName
    Write-LogEntry "$BrowserName version did not change within $TimeoutSeconds seconds while waiting: $Reason. Original=$OriginalVersion Current=$finalVersion" "Warning"
    return $false
}

function Get-BrowserLaunchArguments {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    switch ($BrowserName) {
        "Chrome"  { return @("--no-first-run", "--disable-session-crashed-bubble", "--new-window", "chrome://settings/help") }
        "Edge"    { return @("--no-first-run", "--disable-session-crashed-bubble", "--new-window", "edge://settings/help") }
        "Firefox" { return @("-new-window", "about:preferences") }
    }
}

function Invoke-ClosedBrowserUpdateCycle {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName,

        [int]$WaitSeconds = 60,

        [int]$EngineWaitSeconds = 180
    )

    try {
        $exe = Get-BrowserExecutablePath -BrowserName $BrowserName
        if (-not $exe) {
            throw "$BrowserName executable not found"
        }

        $processName = Get-BrowserProcessName -BrowserName $BrowserName
        $runningAsSystem = Test-IsSystem

        if (Test-BrowserRunning -BrowserName $BrowserName) {
            Write-LogEntry "$BrowserName is now running. Closed-browser update cycle is not safe; interactive restart flow is required instead." "Warning"
            return $false
        }

        $beforeVersion = Get-BrowserInstalledVersion -BrowserName $BrowserName
        Write-LogEntry "$BrowserName closed-browser update verification started. BeforeVersion=$beforeVersion"

        Write-LogEntry "Triggering native update engine for closed $BrowserName before opening the browser"
        [void](Invoke-BrowserUpdateEngine -BrowserName $BrowserName)

        if (Wait-BrowserVersionChange -BrowserName $BrowserName -OriginalVersion $beforeVersion -TimeoutSeconds ([math]::Min(60, [math]::Max(10, [int]($EngineWaitSeconds / 3)))) -Reason "native update engine pre-open phase") {
            return $true
        }

        if ($runningAsSystem) {
            Write-LogEntry "$BrowserName is closed and script is running as SYSTEM. Browser UI launch is skipped to avoid creating a SYSTEM browser profile. Leaving item queued if version has not changed." "Warning"
            return $false
        }

        $arguments = Get-BrowserLaunchArguments -BrowserName $BrowserName
        Write-LogEntry "Starting closed-browser update cycle for $BrowserName from '$exe'. Browser will be opened to its update/help page, allowed $WaitSeconds seconds to check/apply updates, then closed."

        if ($arguments -and $arguments.Count -gt 0) {
            Start-Process -FilePath $exe -ArgumentList $arguments -WindowStyle Minimized | Out-Null
        }
        else {
            Start-Process -FilePath $exe -WindowStyle Minimized | Out-Null
        }

        Start-Sleep -Seconds 3

        if (-not (Test-BrowserRunning -BrowserName $BrowserName)) {
            throw "$BrowserName did not start during closed-browser update cycle"
        }

        if ($WaitSeconds -gt 3) {
            Start-Sleep -Seconds ($WaitSeconds - 3)
        }

        Write-LogEntry "Closing $BrowserName after closed-browser update cycle"
        Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $_.CloseMainWindow() | Out-Null
            }
            catch {}
        }

        Start-Sleep -Seconds 10

        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            Write-LogEntry "$BrowserName still has running processes after closed-browser update cycle graceful close; forcing termination" "Warning"
            Get-Process -Name $processName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }

        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            throw "$BrowserName processes remain after closed-browser update cycle"
        }

        Write-LogEntry "Triggering native update engine for closed $BrowserName after browser close"
        [void](Invoke-BrowserUpdateEngine -BrowserName $BrowserName)

        $remainingWait = [math]::Max(30, $EngineWaitSeconds - 60)
        if (Wait-BrowserVersionChange -BrowserName $BrowserName -OriginalVersion $beforeVersion -TimeoutSeconds $remainingWait -Reason "native update engine post-close phase") {
            Write-LogEntry "$BrowserName closed-browser update cycle completed successfully and version change was verified"
            return $true
        }

        $afterVersion = Get-BrowserInstalledVersion -BrowserName $BrowserName
        $comparison = Compare-VersionString -Left $afterVersion -Right $beforeVersion

        if ([string]::IsNullOrWhiteSpace($beforeVersion) -or [string]::IsNullOrWhiteSpace($afterVersion)) {
            Write-LogEntry "$BrowserName closed-browser update cycle completed, but version verification was inconclusive. Before=$beforeVersion After=$afterVersion" "Warning"
            return $true
        }

        if ($comparison -ne $null -and $comparison -gt 0) {
            Write-LogEntry "$BrowserName closed-browser update cycle completed successfully. Version changed from $beforeVersion to $afterVersion"
            return $true
        }

        Write-LogEntry "$BrowserName closed-browser update cycle opened and closed the browser, but version did not change. Before=$beforeVersion After=$afterVersion. Leaving queue item for retry rather than marking it complete." "Warning"
        return $false
    }
    catch {
        Write-LogEntry "$BrowserName closed-browser update cycle failed: $($_.Exception.Message)" "Error"
        return $false
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
    Write-LogEntry "Closed browser update cycle wait set to $ClosedBrowserUpdateWaitSeconds seconds"
    Write-LogEntry "Browser update engine verification wait set to $BrowserUpdateEngineWaitSeconds seconds"
    if ($ClosedBrowserUpdateWaitSecondsWarning) {
        Write-LogEntry $ClosedBrowserUpdateWaitSecondsWarning "Warning"
    }
    if ($BrowserUpdateEngineWaitSecondsWarning) {
        Write-LogEntry $BrowserUpdateEngineWaitSecondsWarning "Warning"
    }
    Write-LogEntry "======================================"

    Get-ExecutionContextInfo
    Remove-ScheduledTaskIfRequested -TaskName $ScheduledTaskName

    if (-not (Save-CurrentScriptToStablePath -DestinationPath $RemediationScriptPath)) {
        Write-LogEntry "Continuing, but postponed scheduled tasks may fail because the stable script copy is missing." "Warning"
    }

    $runningAsSystem = Test-IsSystem
    if ($runningAsSystem) {
        Write-LogEntry "Remediation script is running as SYSTEM. Closed-browser update cycles will be attempted without UI where possible; browsers that are running or require user interaction will be left queued." "Warning"
    }

    $queue = Get-ReloadQueue
    if (-not $queue -or -not $queue.Browsers -or $queue.Browsers.Count -eq 0) {
        Write-LogEntry "No queued browsers found. Nothing to do."
        exit 0
    }

    Write-LogEntry "Found $($queue.Browsers.Count) queued browser(s) for remediation"

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
        if (-not ($item.PSObject.Properties.Name -contains 'RemediationMode')) {
            $item | Add-Member -MemberType NoteProperty -Name RemediationMode -Value $null -Force
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

        $isBrowserRunningNow = Test-BrowserRunning -BrowserName $browser
        if (-not $isBrowserRunningNow) {
            Write-LogEntry "$browser is currently closed. Running immediate open/update/close cycle without user prompt. RemediationMode=$($item.RemediationMode)"

            if (Invoke-ClosedBrowserUpdateCycle -BrowserName $browser -WaitSeconds $ClosedBrowserUpdateWaitSeconds -EngineWaitSeconds $BrowserUpdateEngineWaitSeconds) {
                $tracking = Get-BrowserUsageTracking
                $tracking.$browser.LastStop = (Get-Date).ToString("o")
                $tracking.$browser.IsRunning = $false

                if (-not (Save-BrowserUsageTracking -TrackingData $tracking)) {
                    Write-LogEntry "$browser closed-browser update cycle completed, but usage tracking could not be updated. Continuing so the queue can still be cleared." "Warning"
                }
            }
            else {
                $remainingQueue += $item
            }

            continue
        }

        if ($item.RemediationMode -eq "ClosedUpdateCycle") {
            Write-LogEntry "$browser was queued for a closed-browser update cycle, but it is now running. Falling back to interactive restart flow to avoid closing active work without notice." "Warning"
        }

        if ($runningAsSystem) {
            Write-LogEntry "$browser is running and remediation is executing as SYSTEM, so no user popup can be shown. Leaving item queued for a logged-in user remediation run." "Warning"
            $remainingQueue += $item
            continue
        }

        Write-LogEntry "Displaying popup for $browser with $WarningTimeSeconds second countdown"

        $decision = Show-CountdownWarning -BrowserName $browser -CountdownSeconds $WarningTimeSeconds -CompanyName $CompanyName
        if ($decision.Action -eq "Postpone") {
            Write-LogEntry "Popup closed for $browser. Action=Postpone PostponeMinutes=$($decision.PostponeMinutes)"
        }
        else {
            Write-LogEntry "Popup closed for $browser. Action=$($decision.Action)"
        }

        if ($decision.Action -eq "Postpone") {
            $runAtLocal = (Get-Date).AddMinutes([int]$decision.PostponeMinutes)

            $updatedItem = [PSCustomObject]@{
                Browser           = $item.Browser
                Reason            = $item.Reason
                PostponeUntilUtc  = $runAtLocal.ToUniversalTime().ToString("o")
                PostponeChoice    = "$($decision.PostponeMinutes) minutes"
                ScheduledTaskName = $null
                RemediationMode   = $item.RemediationMode
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

            if (-not (Save-BrowserUsageTracking -TrackingData $tracking)) {
                Write-LogEntry "$browser was restarted, but usage tracking could not be updated. Continuing so the queue can still be cleared." "Warning"
            }
        }
        else {
            $remainingQueue += $item
        }
    }

    if (-not (Save-ReloadQueue -Browsers $remainingQueue)) {
        Write-LogEntry "Queue could not be saved after remediation. Browser action may repeat on the next detection/remediation cycle." "Error"
        exit 1
    }

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