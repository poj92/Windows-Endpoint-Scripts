# PowerShell Script to Detect Browser Update Conditions and warnn 
# users of pending updates with session/inactivity thresholds.
# Script must be run as logged in user, not SYSTEM.

#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

# =========================
# Configuration
# =========================
$BasePath = "C:\ProgramData\Datto\BrowserUpdateCheck"
$LogFile = Join-Path $BasePath "Remediation.log"
$TrackingFile = Join-Path $BasePath "BrowserUsageTracking.json"
$QueueFile = Join-Path $BasePath "ReloadQueue.json"
$WarningTimeSeconds = 300

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

function Test-IsSystem {
    try {
        return ([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')
    }
    catch {
        return $false
    }
}

function Get-ReloadQueue {
    if (-not (Test-Path $QueueFile)) {
        return $null
    }

    try {
        return (Get-Content -Path $QueueFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
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

    $queue | ConvertTo-Json -Depth 5 | Set-Content -Path $QueueFile -Force -Encoding UTF8
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

function Show-CountdownWarning {
    param(
        [string]$BrowserName,
        [int]$CountdownSeconds
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $script:Cancelled = $false
    $script:SecondsLeft = $CountdownSeconds

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$BrowserName update requires restart"
    $form.Size = New-Object System.Drawing.Size(540,250)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(490,70)
    $label.Text = "$BrowserName has a pending update and has exceeded the allowed open/inactive time. It will be restarted automatically in:"
    $label.Font = New-Object System.Drawing.Font('Segoe UI',10)
    $form.Controls.Add($label)

    $countdownLabel = New-Object System.Windows.Forms.Label
    $countdownLabel.Location = New-Object System.Drawing.Point(20,90)
    $countdownLabel.Size = New-Object System.Drawing.Size(490,40)
    $countdownLabel.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
    $countdownLabel.TextAlign = 'MiddleCenter'
    $form.Controls.Add($countdownLabel)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(185,150)
    $cancelButton.Size = New-Object System.Drawing.Size(160,35)
    $cancelButton.Text = 'Cancel Restart'
    $cancelButton.Add_Click({
        $script:Cancelled = $true
        $form.Close()
    })
    $form.Controls.Add($cancelButton)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $minutes = [math]::Floor($script:SecondsLeft / 60)
        $seconds = $script:SecondsLeft % 60
        $countdownLabel.Text = "{0}:{1:D2}" -f $minutes, $seconds
        $script:SecondsLeft--

        if ($script:SecondsLeft -lt 0) {
            $timer.Stop()
            $form.Close()
        }
    })

    $timer.Start()
    [void]$form.ShowDialog()

    return (-not $script:Cancelled)
}

function Invoke-BrowserReload {
    param(
        [ValidateSet("Chrome","Firefox","Edge")]
        [string]$BrowserName
    )

    try {
        switch ($BrowserName) {
            "Chrome" {
                $exe = Get-ChromeInstallPath
                if (-not $exe) { throw "Chrome executable not found" }

                Get-Process chrome -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 10
                Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Process -FilePath $exe | Out-Null
            }

            "Firefox" {
                $exe = Get-FirefoxInstallPath
                if (-not $exe) { throw "Firefox executable not found" }

                Get-Process firefox -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 10
                Get-Process firefox -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Process -FilePath $exe | Out-Null
            }

            "Edge" {
                $exe = Get-EdgeInstallPath
                if (-not $exe) { throw "Edge executable not found" }

                Get-Process msedge -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 10
                Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Process -FilePath $exe | Out-Null
            }
        }

        Write-LogEntry "$BrowserName restarted successfully"
        return $true
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

    foreach ($item in $queue.Browsers) {
        $browser = $item.Browser
        Write-LogEntry "Processing $browser. Reason: $($item.Reason)"

        $proceed = Show-CountdownWarning -BrowserName $browser -CountdownSeconds $WarningTimeSeconds

        if (-not $proceed) {
            Write-LogEntry "$browser restart cancelled by user" "Warning"
            $remainingQueue += $item
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