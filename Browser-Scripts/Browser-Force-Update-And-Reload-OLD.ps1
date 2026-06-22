#Requires -Version 5.1

<#
Author: Peter Opeyemi James
Company: Nexus Open Systems Ltd
Date: 2026-05-05
Email: Peter.James@nexusos.co.uk
#>

<#
        Browser-Force-Update-And-Reload.ps1

        Purpose
        - Force Google Chrome and Microsoft Edge to check/apply updates immediately.
        - If a browser relaunch is required, prompt the logged-in user with only one action: Reload now.
        - If the user does nothing, reload automatically after the countdown expires.

        Recommended Datto RMM execution context
        - Run as SYSTEM.
            The script will perform the update work in SYSTEM context and, if needed, create a one-time
            scheduled task to show the popup in the logged-in user's session.

        Notes
        - Browsers are only relaunched if they were running when the reload phase starts.
        - Existing tabs/sessions rely on Chrome/Edge session restore/user settings.
        - The script uses multiple update triggers because endpoint builds vary.
#>

param(
    [switch]$PromptOnly,
    [string]$ScheduledTaskName,
    [string]$BasePath = "C:\ProgramData\Datto\BrowserForceUpdate",
    [string]$CompanyName = "Nexus Open Systems Ltd",
    [int]$CountdownSeconds = 300,
    [int]$UpdateWaitSeconds = 90,
    [int]$PollIntervalSeconds = 5,
    [int]$GracefulCloseWaitSeconds = 10,
    [int]$RelaunchDelaySeconds = 2,
    [bool]$PromptOnlyWhenBrowserIsRunning = $true,
    [bool]$ForceKillRemainingProcesses = $true
)

$ErrorActionPreference = 'Stop'

$LogFile = Join-Path $BasePath 'ForceUpdate.log'
$QueueFile = Join-Path $BasePath 'ForceReloadQueue.json'
$StableScriptPath = Join-Path $BasePath 'Browser-Force-Update-And-Reload.ps1'

if (-not (Test-Path $BasePath)) {
    New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
}

function Write-LogEntry {
    param(
        [string]$Message,
        [ValidateSet('Information','Warning','Error')]
        [string]$Level = 'Information'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"

    try {
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    }
    catch {}

    Write-Host $line
}

function Save-CurrentScriptToStablePath {
    param([string]$DestinationPath)

    try {
        $currentScriptPath = $PSCommandPath
        if ([string]::IsNullOrWhiteSpace($currentScriptPath) -or -not (Test-Path $currentScriptPath)) {
            throw 'Unable to determine current script path.'
        }

        $folder = Split-Path -Path $DestinationPath -Parent
        if (-not (Test-Path $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }

        Copy-Item -Path $currentScriptPath -Destination $DestinationPath -Force
        Write-LogEntry "Copied script to stable path '$DestinationPath'"
        return $true
    }
    catch {
        Write-LogEntry "Failed to copy script to stable path: $($_.Exception.Message)" 'Error'
        return $false
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

function Remove-ScheduledTaskIfRequested {
    param([string]$TaskName)

    if ([string]::IsNullOrWhiteSpace($TaskName)) {
        return
    }

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-LogEntry "Deleted scheduled task '$TaskName' after launch"
    }
    catch {
        Write-LogEntry "Failed to delete scheduled task '$TaskName' : $($_.Exception.Message)" 'Warning'
    }
}

function Get-CurrentInteractiveUser {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($cs.UserName)) {
            return $cs.UserName
        }
    }
    catch {
        Write-LogEntry "Unable to query Win32_ComputerSystem for logged-in user: $($_.Exception.Message)" 'Warning'
    }

    try {
        $explorer = Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop |
            Select-Object -First 1
        if ($explorer) {
            $owner = Invoke-CimMethod -InputObject $explorer -MethodName GetOwner -ErrorAction Stop
            if ($owner.User) {
                return "{0}\{1}" -f $owner.Domain, $owner.User
            }
        }
    }
    catch {
        Write-LogEntry "Unable to infer logged-in user from explorer.exe: $($_.Exception.Message)" 'Warning'
    }

    return $null
}

function Get-BrowserDefinition {
    param(
        [ValidateSet('Chrome','Edge')]
        [string]$BrowserName
    )

    switch ($BrowserName) {
        'Chrome' {
            return [PSCustomObject]@{
                Name = 'Chrome'
                ProcessName = 'chrome'
                ExePaths = @(
                    'C:\Program Files\Google\Chrome\Application\chrome.exe',
                    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
                )
                UpdateExePaths = @(
                    'C:\Program Files (x86)\Google\Update\GoogleUpdate.exe',
                    'C:\Program Files\Google\Update\GoogleUpdate.exe'
                )
                CoreTask = 'GoogleUpdateTaskMachineCore'
                UATask = 'GoogleUpdateTaskMachineUA'
                PendingType = 'Chrome'
            }
        }
        'Edge' {
            return [PSCustomObject]@{
                Name = 'Edge'
                ProcessName = 'msedge'
                ExePaths = @(
                    'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
                    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
                )
                UpdateExePaths = @(
                    'C:\Program Files (x86)\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe',
                    'C:\Program Files\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'
                )
                CoreTask = 'MicrosoftEdgeUpdateTaskMachineCore'
                UATask = 'MicrosoftEdgeUpdateTaskMachineUA'
                PendingType = 'Edge'
            }
        }
    }
}

function Get-BrowserExecutablePath {
    param([string]$BrowserName)

    $def = Get-BrowserDefinition -BrowserName $BrowserName
    return ($def.ExePaths | Where-Object { Test-Path $_ } | Select-Object -First 1)
}

function Test-BrowserInstalled {
    param([string]$BrowserName)
    return [bool](Get-BrowserExecutablePath -BrowserName $BrowserName)
}

function Test-BrowserRunning {
    param([string]$BrowserName)

    $def = Get-BrowserDefinition -BrowserName $BrowserName
    return [bool](Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue)
}

function Get-BrowserRunningState {
    $result = @{}
    foreach ($browser in @('Chrome','Edge')) {
        $result[$browser] = Test-BrowserRunning -BrowserName $browser
    }
    return $result
}

function Get-BrowserOldestProcessStartTimeUtc {
    param([string]$BrowserName)

    try {
        $def = Get-BrowserDefinition -BrowserName $BrowserName
        $procs = Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue
        if (-not $procs) {
            return $null
        }

        $startTimes = @()
        foreach ($proc in $procs) {
            try {
                if ($proc.StartTime) {
                    $startTimes += $proc.StartTime.ToUniversalTime()
                }
            }
            catch {}
        }

        if ($startTimes.Count -gt 0) {
            return ($startTimes | Sort-Object | Select-Object -First 1)
        }

        return $null
    }
    catch {
        Write-LogEntry "Failed to read browser process start time for $BrowserName : $($_.Exception.Message)" 'Warning'
        return $null
    }
}

function Test-BrowserRunningInstanceOlderThanInstalledBinary {
    param([string]$BrowserName)

    try {
        if (-not (Test-BrowserRunning -BrowserName $BrowserName)) {
            return $false
        }

        $exe = Get-BrowserExecutablePath -BrowserName $BrowserName
        if (-not $exe -or -not (Test-Path $exe)) {
            return $false
        }

        $oldestProcessStartUtc = Get-BrowserOldestProcessStartTimeUtc -BrowserName $BrowserName
        if (-not $oldestProcessStartUtc) {
            return $false
        }

        $exeWriteUtc = (Get-Item -Path $exe -ErrorAction Stop).LastWriteTimeUtc

        if ($oldestProcessStartUtc.AddSeconds(5) -lt $exeWriteUtc) {
            Write-LogEntry "$BrowserName pending update detected via process-start/binary-write mismatch. OldestProcessStartUtc=$oldestProcessStartUtc BinaryWriteUtc=$exeWriteUtc"
            return $true
        }

        return $false
    }
    catch {
        Write-LogEntry "$BrowserName process-vs-binary update check failed: $($_.Exception.Message)" 'Warning'
        return $false
    }
}

function Invoke-ScheduledTaskNow {
    param([string]$TaskName)

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        Start-ScheduledTask -InputObject $task -ErrorAction Stop
        Write-LogEntry "Started scheduled task '$TaskName'"
        return $true
    }
    catch {
        Write-LogEntry "Unable to start scheduled task '$TaskName': $($_.Exception.Message)" 'Warning'
        return $false
    }
}

function Invoke-ProcessSilently {
    param(
        [string]$FilePath,
        [string]$Arguments
    )

    try {
        if (-not (Test-Path $FilePath)) {
            return $false
        }

        $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -WindowStyle Hidden -PassThru -ErrorAction Stop
        Write-LogEntry "Started process '$FilePath' $Arguments"
        try {
            $null = $proc.WaitForExit(30000)
            if ($proc.HasExited) {
                Write-LogEntry "Process '$FilePath' exited with code $($proc.ExitCode)"
            }
        }
        catch {}
        return $true
    }
    catch {
        Write-LogEntry "Failed to start '$FilePath' $Arguments : $($_.Exception.Message)" 'Warning'
        return $false
    }
}

function Invoke-ChromeUpdateTrigger {
    if (-not (Test-BrowserInstalled -BrowserName 'Chrome')) {
        Write-LogEntry 'Chrome is not installed. Skipping update trigger.'
        return
    }

    $def = Get-BrowserDefinition -BrowserName 'Chrome'

    $null = Invoke-ScheduledTaskNow -TaskName $def.CoreTask
    $null = Invoke-ScheduledTaskNow -TaskName $def.UATask

    foreach ($updater in $def.UpdateExePaths) {
        if (Test-Path $updater) {
            # Common Google Update invocations seen on managed Windows endpoints.
            $null = Invoke-ProcessSilently -FilePath $updater -Arguments '/ua /installsource scheduler'
            $null = Invoke-ProcessSilently -FilePath $updater -Arguments '/c'
        }
    }
}

function Invoke-EdgeUpdateTrigger {
    if (-not (Test-BrowserInstalled -BrowserName 'Edge')) {
        Write-LogEntry 'Edge is not installed. Skipping update trigger.'
        return
    }

    $def = Get-BrowserDefinition -BrowserName 'Edge'

    try {
        $edgeServices = @('edgeupdate','edgeupdatem')
        foreach ($svcName in $edgeServices) {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Running') {
                Start-Service -Name $svcName -ErrorAction SilentlyContinue
                Write-LogEntry "Attempted to start service '$svcName'"
            }
        }
    }
    catch {
        Write-LogEntry "Unable to start one or more Edge update services: $($_.Exception.Message)" 'Warning'
    }

    $null = Invoke-ScheduledTaskNow -TaskName $def.CoreTask
    $null = Invoke-ScheduledTaskNow -TaskName $def.UATask

    foreach ($updater in $def.UpdateExePaths) {
        if (Test-Path $updater) {
            $null = Invoke-ProcessSilently -FilePath $updater -Arguments '/ua /installsource scheduler'
            $null = Invoke-ProcessSilently -FilePath $updater -Arguments '/c'
        }
    }
}

function Test-ChromePendingUpdate {
    try {
        $chromeVersion = $null
        $chromeRegPath = 'HKLM:\SOFTWARE\Google\Chrome\BLBeacon'

        if (Test-Path $chromeRegPath) {
            $chromeVersion = (Get-ItemProperty -Path $chromeRegPath -ErrorAction SilentlyContinue).version
        }

        $updateRegPath = 'HKLM:\SOFTWARE\Google\Update\ClientState\{8A69D345-D564-463C-AFF1-A69D9E530F96}'
        if (Test-Path $updateRegPath) {
            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry 'Chrome pending update detected via UpdateAvailable flag'
                return $true
            }

            if ($props.opv -and $chromeVersion -and $props.opv -ne $chromeVersion) {
                Write-LogEntry 'Chrome pending update detected via staged/current version mismatch'
                return $true
            }
        }

        $chromeExe = Get-BrowserExecutablePath -BrowserName 'Chrome'
        if ($chromeExe) {
            $basePath = Split-Path $chromeExe -Parent

            if (Test-Path (Join-Path $basePath 'new_chrome.exe')) {
                Write-LogEntry 'Chrome pending update detected via new_chrome.exe'
                return $true
            }

            $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }

            if ($versionFolders.Count -gt 1) {
                Write-LogEntry 'Chrome possible pending update detected via multiple version folders'
                return $true
            }
        }

        if (Test-BrowserRunningInstanceOlderThanInstalledBinary -BrowserName 'Chrome') {
            return $true
        }

        Write-LogEntry 'No pending Chrome update detected'
        return $false
    }
    catch {
        Write-LogEntry "Chrome pending update check failed: $($_.Exception.Message)" 'Warning'
        return $false
    }
}

function Test-EdgePendingUpdate {
    try {
        $edgeVersion = $null
        $edgeRegPath = 'HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon'
        if (Test-Path $edgeRegPath) {
            $edgeVersion = (Get-ItemProperty -Path $edgeRegPath -ErrorAction SilentlyContinue).version
        }

        $updateRegPath = 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\ClientState\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
        if (Test-Path $updateRegPath) {
            $props = Get-ItemProperty -Path $updateRegPath -ErrorAction SilentlyContinue

            if ($props.UpdateAvailable -eq 1) {
                Write-LogEntry 'Edge pending update detected via UpdateAvailable flag'
                return $true
            }

            if ($props.opv -and $edgeVersion -and $props.opv -ne $edgeVersion) {
                Write-LogEntry 'Edge pending update detected via staged/current version mismatch'
                return $true
            }
        }

        $edgeExe = Get-BrowserExecutablePath -BrowserName 'Edge'
        if ($edgeExe) {
            $basePath = Split-Path $edgeExe -Parent

            if (Test-Path (Join-Path $basePath 'new_msedge.exe')) {
                Write-LogEntry 'Edge pending update detected via new_msedge.exe'
                return $true
            }

            $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' }

            if ($versionFolders.Count -gt 1) {
                Write-LogEntry 'Edge possible pending update detected via multiple version folders'
                return $true
            }
        }

        if (Test-BrowserRunningInstanceOlderThanInstalledBinary -BrowserName 'Edge') {
            return $true
        }

        Write-LogEntry 'No pending Edge update detected'
        return $false
    }
    catch {
        Write-LogEntry "Edge pending update check failed: $($_.Exception.Message)" 'Warning'
        return $false
    }
}

function Get-ReloadCandidates {
    param([hashtable]$InitialRunningState)

    $candidates = @()

    foreach ($browser in @('Chrome','Edge')) {
        $installed = Test-BrowserInstalled -BrowserName $browser
        if (-not $installed) {
            continue
        }

        $running = Test-BrowserRunning -BrowserName $browser
        $pending = switch ($browser) {
            'Chrome' { Test-ChromePendingUpdate }
            'Edge'   { Test-EdgePendingUpdate }
        }

        $shouldPrompt = $pending
        if ($PromptOnlyWhenBrowserIsRunning -and -not $running) {
            $shouldPrompt = $false
        }

        $candidates += [PSCustomObject]@{
            Browser = $browser
            WasRunningAtStart = [bool]$InitialRunningState[$browser]
            IsRunningNow = $running
            PendingUpdate = $pending
            NeedsReloadPrompt = $shouldPrompt
        }
    }

    return $candidates
}

function Save-ReloadQueue {
    param([array]$Items)

    $payload = [PSCustomObject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        Items = $Items
    }

    $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $QueueFile -Encoding UTF8 -Force
}

function Get-ReloadQueue {
    if (-not (Test-Path $QueueFile)) {
        return $null
    }

    try {
        return (Get-Content -Path $QueueFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        Write-LogEntry "Queue file is unreadable: $($_.Exception.Message)" 'Error'
        return $null
    }
}

function Register-InteractivePromptTask {
    param(
        [string]$InteractiveUser,
        [string]$ScriptPath
    )

    try {
        $taskName = 'BrowserForceUpdatePrompt_{0}' -f ([guid]::NewGuid().ToString('N').Substring(0,8))

        $actionArgs = @(
            '-ExecutionPolicy Bypass'
            '-WindowStyle Normal'
            ('-File "{0}"' -f $ScriptPath)
            '-PromptOnly'
            ('-ScheduledTaskName "{0}"' -f $taskName)
            ('-BasePath "{0}"' -f $BasePath)
            ('-CompanyName "{0}"' -f $CompanyName)
            ('-CountdownSeconds {0}' -f $CountdownSeconds)
            ('-GracefulCloseWaitSeconds {0}' -f $GracefulCloseWaitSeconds)
            ('-RelaunchDelaySeconds {0}' -f $RelaunchDelaySeconds)
        ) -join ' '

        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $actionArgs
        $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
        $principal = New-ScheduledTaskPrincipal -UserId $InteractiveUser -LogonType Interactive -RunLevel Limited
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName | Out-Null
        Write-LogEntry "Created and started interactive prompt task '$taskName' for user '$InteractiveUser'"
        return $taskName
    }
    catch {
        Write-LogEntry "Failed to create interactive prompt task: $($_.Exception.Message)" 'Error'
        return $null
    }
}

function Show-ReloadPrompt {
    param(
        [string[]]$BrowserNames,
        [int]$Seconds,
        [string]$OrgName
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $script:UserDecision = 'CountdownReload'
        $script:SecondsLeft = $Seconds
        $browserList = ($BrowserNames | Sort-Object -Unique) -join ' and '

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "$OrgName - Browser reload required"
        $form.Size = New-Object System.Drawing.Size(700,320)
        $form.StartPosition = 'CenterScreen'
        $form.TopMost = $true
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.ShowInTaskbar = $true

        $headerLabel = New-Object System.Windows.Forms.Label
        $headerLabel.Location = New-Object System.Drawing.Point(20,20)
        $headerLabel.Size = New-Object System.Drawing.Size(640,25)
        $headerLabel.Text = "Message from $OrgName"
        $headerLabel.Font = New-Object System.Drawing.Font('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
        $form.Controls.Add($headerLabel)

        $messageLabel = New-Object System.Windows.Forms.Label
        $messageLabel.Location = New-Object System.Drawing.Point(20,55)
        $messageLabel.Size = New-Object System.Drawing.Size(640,90)
        $messageLabel.Font = New-Object System.Drawing.Font('Segoe UI',10)
        $messageLabel.Text = "$browserList has installed an update and must reload to complete it.`r`n`r`nPlease save your work. The browser will reload automatically when the timer expires."
        $form.Controls.Add($messageLabel)

        $countdownLabel = New-Object System.Windows.Forms.Label
        $countdownLabel.Location = New-Object System.Drawing.Point(20,150)
        $countdownLabel.Size = New-Object System.Drawing.Size(640,40)
        $countdownLabel.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
        $countdownLabel.TextAlign = 'MiddleCenter'
        $form.Controls.Add($countdownLabel)

        $reloadNowButton = New-Object System.Windows.Forms.Button
        $reloadNowButton.Location = New-Object System.Drawing.Point(275,210)
        $reloadNowButton.Size = New-Object System.Drawing.Size(140,35)
        $reloadNowButton.Text = 'Reload now'
        $reloadNowButton.Add_Click({
            $script:UserDecision = 'ReloadNow'
            $form.Close()
        })
        $form.Controls.Add($reloadNowButton)

        $footerLabel = New-Object System.Windows.Forms.Label
        $footerLabel.Location = New-Object System.Drawing.Point(20,260)
        $footerLabel.Size = New-Object System.Drawing.Size(640,20)
        $footerLabel.Font = New-Object System.Drawing.Font('Segoe UI',9)
        $footerLabel.Text = 'If no action is taken, reload starts automatically.'
        $form.Controls.Add($footerLabel)

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 1000
        $timer.Add_Tick({
            $minutes = [math]::Floor($script:SecondsLeft / 60)
            $seconds = $script:SecondsLeft % 60
            $countdownLabel.Text = '{0}:{1:D2}' -f $minutes, $seconds
            $script:SecondsLeft--

            if ($script:SecondsLeft -lt 0) {
                $timer.Stop()
                $script:UserDecision = 'CountdownReload'
                $form.Close()
            }
        })

        $form.Add_Shown({
            $form.Activate()
        })

        $timer.Start()
        [void]$form.ShowDialog()
        $timer.Stop()
        $form.Dispose()

        return $script:UserDecision
    }
    catch {
        Write-LogEntry "Popup failed: $($_.Exception.Message)" 'Error'
        return 'CountdownReload'
    }
}

function Invoke-BrowserReload {
    param([string]$BrowserName)

    try {
        $def = Get-BrowserDefinition -BrowserName $BrowserName
        $exe = Get-BrowserExecutablePath -BrowserName $BrowserName
        if (-not $exe) {
            throw "$BrowserName executable not found"
        }

        $procs = Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue
        if (-not $procs) {
            Write-LogEntry "$BrowserName is no longer running. No reload required."
            return $true
        }

        Write-LogEntry "Attempting graceful close of $BrowserName"
        foreach ($proc in $procs) {
            try {
                $null = $proc.CloseMainWindow()
            }
            catch {}
        }

        Start-Sleep -Seconds $GracefulCloseWaitSeconds

        if ($ForceKillRemainingProcesses -and (Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue)) {
            Write-LogEntry "$BrowserName still has running processes after graceful close; forcing termination" 'Warning'
            Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }

        Write-LogEntry "Relaunching $BrowserName from '$exe'"
        Start-Process -FilePath $exe | Out-Null
        Start-Sleep -Seconds $RelaunchDelaySeconds

        if (Get-Process -Name $def.ProcessName -ErrorAction SilentlyContinue) {
            Write-LogEntry "$BrowserName relaunched successfully"
            return $true
        }

        throw "$BrowserName did not appear to relaunch successfully"
    }
    catch {
        Write-LogEntry "$BrowserName reload failed: $($_.Exception.Message)" 'Error'
        return $false
    }
}

function Invoke-PromptPhase {
    Remove-ScheduledTaskIfRequested -TaskName $ScheduledTaskName

    if (Test-IsSystem) {
        Write-LogEntry 'PromptOnly mode was launched as SYSTEM. Popup requires user context.' 'Error'
        exit 2
    }

    $queue = Get-ReloadQueue
    if (-not $queue -or -not $queue.Items -or $queue.Items.Count -eq 0) {
        Write-LogEntry 'No queued browsers found for prompt phase.'
        exit 0
    }

    $items = @($queue.Items | Where-Object { $_.NeedsReloadPrompt -eq $true })
    if ($items.Count -eq 0) {
        Write-LogEntry 'Queue exists, but no browser currently requires a reload prompt.'
        exit 0
    }

    $browserNames = @($items.Browser | Sort-Object -Unique)
    Write-LogEntry "Displaying popup for: $($browserNames -join ', ')"

    $decision = Show-ReloadPrompt -BrowserNames $browserNames -Seconds $CountdownSeconds -OrgName $CompanyName
    Write-LogEntry "Popup closed. Decision=$decision"

    $anyFailure = $false
    foreach ($browser in $browserNames) {
        $item = $items | Where-Object { $_.Browser -eq $browser } | Select-Object -First 1
        if (-not $item) {
            continue
        }

        if (-not [bool]$item.IsRunningNow -and -not (Test-BrowserRunning -BrowserName $browser)) {
            Write-LogEntry "$browser is no longer running. Skipping reload."
            continue
        }

        if (-not (Invoke-BrowserReload -BrowserName $browser)) {
            $anyFailure = $true
        }
    }

    try {
        Remove-Item -Path $QueueFile -Force -ErrorAction SilentlyContinue
    }
    catch {}

    if ($anyFailure) {
        exit 3
    }

    exit 0
}

function Invoke-SystemPhase {
    if (-not (Test-IsSystem)) {
        Write-LogEntry 'Primary mode is designed to run as SYSTEM. Continuing in current context, but update triggers may be limited.' 'Warning'
    }

    $null = Save-CurrentScriptToStablePath -DestinationPath $StableScriptPath

    $initialRunning = Get-BrowserRunningState
    Write-LogEntry ("Initial running state: Chrome={0} Edge={1}" -f $initialRunning['Chrome'], $initialRunning['Edge'])

    foreach ($browser in @('Chrome','Edge')) {
        $exe = Get-BrowserExecutablePath -BrowserName $browser
        if ($exe) {
            try {
                $item = Get-Item -Path $exe -ErrorAction Stop
                $ver = $item.VersionInfo.ProductVersion
                $writeUtc = $item.LastWriteTimeUtc
                Write-LogEntry ("Pre-update installed binary state: Browser={0} Version={1} BinaryWriteUtc={2}" -f $browser, $ver, $writeUtc)
            }
            catch {
                Write-LogEntry ("Unable to read pre-update binary state for {0}: {1}" -f $browser, $_.Exception.Message) 'Warning'
            }
        }
    }

    Write-LogEntry 'Triggering Chrome update'
    Invoke-ChromeUpdateTrigger

    Write-LogEntry 'Triggering Edge update'
    Invoke-EdgeUpdateTrigger

    $deadline = (Get-Date).AddSeconds($UpdateWaitSeconds)
    do {
        Start-Sleep -Seconds $PollIntervalSeconds
        $chromePending = if (Test-BrowserInstalled -BrowserName 'Chrome') { Test-ChromePendingUpdate } else { $false }
        $edgePending = if (Test-BrowserInstalled -BrowserName 'Edge') { Test-EdgePendingUpdate } else { $false }
        Write-LogEntry "Polling update state: ChromePending=$chromePending EdgePending=$edgePending"
    }
    while ((Get-Date) -lt $deadline)

    foreach ($browser in @('Chrome','Edge')) {
        $exe = Get-BrowserExecutablePath -BrowserName $browser
        if ($exe) {
            try {
                $item = Get-Item -Path $exe -ErrorAction Stop
                $ver = $item.VersionInfo.ProductVersion
                $writeUtc = $item.LastWriteTimeUtc
                Write-LogEntry ("Post-update installed binary state: Browser={0} Version={1} BinaryWriteUtc={2}" -f $browser, $ver, $writeUtc)
            }
            catch {
                Write-LogEntry ("Unable to read post-update binary state for {0}: {1}" -f $browser, $_.Exception.Message) 'Warning'
            }
        }
    }

    $candidates = Get-ReloadCandidates -InitialRunningState $initialRunning
    foreach ($candidate in $candidates) {
        Write-LogEntry ("Candidate: Browser={0} Installed={1} RunningNow={2} PendingUpdate={3} NeedsReloadPrompt={4}" -f $candidate.Browser, (Test-BrowserInstalled -BrowserName $candidate.Browser), $candidate.IsRunningNow, $candidate.PendingUpdate, $candidate.NeedsReloadPrompt)
    }

    $needsPrompt = @($candidates | Where-Object { $_.NeedsReloadPrompt -eq $true })
    if ($needsPrompt.Count -eq 0) {
        Write-LogEntry 'No browser reload prompt required after update attempt.'
        try {
            Remove-Item -Path $QueueFile -Force -ErrorAction SilentlyContinue
        }
        catch {}
        exit 0
    }

    Save-ReloadQueue -Items $needsPrompt
    Write-LogEntry "Saved reload queue with $($needsPrompt.Count) item(s)"

    $interactiveUser = Get-CurrentInteractiveUser
    if ([string]::IsNullOrWhiteSpace($interactiveUser)) {
        Write-LogEntry 'A browser reload is required, but no interactive user session was found. Leaving queue file for later retry.' 'Warning'
        exit 4
    }

    $taskName = Register-InteractivePromptTask -InteractiveUser $interactiveUser -ScriptPath $StableScriptPath
    if (-not $taskName) {
        Write-LogEntry 'Failed to launch interactive prompt task. Queue file retained for retry.' 'Error'
        exit 5
    }

    Write-LogEntry "Interactive prompt launched successfully using task '$taskName'"
    exit 0
}

try {
    Write-LogEntry '======================================'
    Write-LogEntry 'Browser force update script started'
    Write-LogEntry '======================================'
    Write-LogEntry ("Mode: {0}" -f ($(if ($PromptOnly) { 'PromptOnly' } else { 'SystemOrPrimary' })))
    Write-LogEntry ("Execution identity: {0}" -f ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name))

    if ($PromptOnly) {
        Invoke-PromptPhase
    }
    else {
        Invoke-SystemPhase
    }
}
catch {
    Write-LogEntry "Fatal error: $($_.Exception.Message)" 'Error'
    exit 1
}
finally {
    Write-LogEntry '======================================'
    Write-LogEntry 'Browser force update script finished'
    Write-LogEntry '======================================'
}
