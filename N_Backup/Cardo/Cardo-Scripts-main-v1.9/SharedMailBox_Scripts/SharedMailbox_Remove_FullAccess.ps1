<#
.SYNOPSIS
    Removes users from shared mailbox permissions based on two CSV files.

.DESCRIPTION
    This script reads two CSV files: one containing user email addresses and another containing
    shared mailbox addresses. It then iterates through all users and all shared mailboxes,
    checks if each user has any permissions on each mailbox, and removes those permissions if found.

.PARAMETER UsersCsvPath
    Path to the CSV file containing users to process.
    CSV should have a column named: Email (or UserEmail)

.PARAMETER MailboxesCsvPath
    Path to the CSV file containing shared mailboxes to process.
    CSV should have a column named: Email (or MailboxEmail)

.EXAMPLE
    .\SharedMailbox_Remove_Users.ps1 -UsersCsvPath "C:\users.csv" -MailboxesCsvPath "C:\mailboxes.csv"

.NOTES
    Requires: Exchange Online Management module (EXO v2 or later)
    Install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$UsersCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$MailboxesCsvPath
)

# Resolve a file path relative to the current directory or the script directory
function Resolve-FilePath {
    param([string]$InputPath)

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        return $null
    }

    $Candidates = @($InputPath)

    if ($PSScriptRoot) {
        $Candidates += Join-Path -Path $PSScriptRoot -ChildPath $InputPath
    }

    foreach ($Candidate in $Candidates) {
        if (Test-Path $Candidate -PathType Leaf) {
            return (Resolve-Path $Candidate).Path
        }
    }

    return $null
}

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
catch {
    Write-Error "Failed to import ExchangeOnlineManagement module. Please install it using: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
    Pause
    exit 1
}

# Function to prompt for file path
function Get-FilePath {
    param([string]$FileType)
    
    Write-Host "`nPlease provide the path to the $FileType CSV file:" -ForegroundColor Yellow
    $Path = Read-Host "Enter file path"
    $ResolvedPath = Resolve-FilePath $Path
    
    while (-not $ResolvedPath) {
        Write-Host "File not found in current directory or script directory: $Path" -ForegroundColor Red
        $Path = Read-Host "Please enter a valid file path"
        $ResolvedPath = Resolve-FilePath $Path
    }
    
    return $ResolvedPath
}

# Get CSV file paths
if ([string]::IsNullOrWhiteSpace($UsersCsvPath)) {
    $UsersCsvPath = Get-FilePath "users"
}
else {
    $ResolvedUsersPath = Resolve-FilePath $UsersCsvPath
    if (-not $ResolvedUsersPath) {
        Write-Error "Users CSV file not found: $UsersCsvPath (checked current directory and script directory)"
        Pause
        exit 1
    }
    $UsersCsvPath = $ResolvedUsersPath
}

if ([string]::IsNullOrWhiteSpace($MailboxesCsvPath)) {
    $MailboxesCsvPath = Get-FilePath "shared mailboxes"
}
else {
    $ResolvedMailboxesPath = Resolve-FilePath $MailboxesCsvPath
    if (-not $ResolvedMailboxesPath) {
        Write-Error "Mailboxes CSV file not found: $MailboxesCsvPath (checked current directory and script directory)"
        Pause
        exit 1
    }
    $MailboxesCsvPath = $ResolvedMailboxesPath
}

# Logging setup
$LogFile = "SharedMailbox_Removal_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogFile -Value $LogMessage
}

# Connect to Exchange Online (required for this script)
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    Pause
    exit 1
}

# Import users CSV file
try {
    $Users = Import-Csv -Path $UsersCsvPath -ErrorAction Stop
    Write-Log "Successfully imported users CSV file: $UsersCsvPath"
    Write-Log "Found $($Users.Count) users to process"
}
catch {
    Write-Error "Failed to import users CSV file: $_"
    Pause
    exit 1
}

# Import mailboxes CSV file
try {
    $Mailboxes = Import-Csv -Path $MailboxesCsvPath -ErrorAction Stop
    Write-Log "Successfully imported mailboxes CSV file: $MailboxesCsvPath"
    Write-Log "Found $($Mailboxes.Count) shared mailboxes to process"
}
catch {
    Write-Error "Failed to import mailboxes CSV file: $_"
    Pause
    exit 1
}

# Validate and extract email columns
$UserEmailColumn = $null
$MailboxEmailColumn = $null

$UserColumns = $Users[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
foreach ($Column in @('Email', 'UserEmail', 'user_email', 'email_address')) {
    if ($Column -in $UserColumns) {
        $UserEmailColumn = $Column
        break
    }
}

if (-not $UserEmailColumn) {
    Write-Error "Users CSV file must contain one of these columns: Email, UserEmail, user_email, email_address"
    Write-Host "Available columns: $($UserColumns -join ', ')"
    Pause
    exit 1
}

$MailboxColumns = $Mailboxes[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
foreach ($Column in @('Email', 'MailboxEmail', 'mailbox_email', 'SharedMailboxEmail', 'email_address')) {
    if ($Column -in $MailboxColumns) {
        $MailboxEmailColumn = $Column
        break
    }
}

if (-not $MailboxEmailColumn) {
    Write-Error "Mailboxes CSV file must contain one of these columns: Email, MailboxEmail, mailbox_email, SharedMailboxEmail, email_address"
    Write-Host "Available columns: $($MailboxColumns -join ', ')"
    Pause
    exit 1
}

Write-Log "User email column: $UserEmailColumn"
Write-Log "Mailbox email column: $MailboxEmailColumn"

$ProcessedCount = 0
$RemovedCount = 0
$ErrorCount = 0

# Process each user against each mailbox
foreach ($User in $Users) {
    $UserEmail = $User.$UserEmailColumn.Trim()

    # Skip empty rows
    if ([string]::IsNullOrWhiteSpace($UserEmail)) {
        Write-Log "Skipping empty user entry" "WARNING"
        continue
    }

    foreach ($Mailbox in $Mailboxes) {
        $SharedMailboxEmail = $Mailbox.$MailboxEmailColumn.Trim()

        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($SharedMailboxEmail)) {
            Write-Log "Skipping empty mailbox entry" "WARNING"
            continue
        }

        $ProcessedCount++
        Write-Log "Processing: User=$UserEmail, SharedMailbox=$SharedMailboxEmail"

        try {
            # Get current mailbox permissions
            $CurrentPermissions = Get-MailboxPermission -Identity $SharedMailboxEmail -User $UserEmail -ErrorAction SilentlyContinue

            if ($CurrentPermissions) {
                Write-Log "Found permissions for $UserEmail on $SharedMailboxEmail. Removing access..." "WARNING"

                # Remove user from the shared mailbox
                Remove-MailboxPermission -Identity $SharedMailboxEmail -User $UserEmail -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop

                Write-Log "Successfully removed $UserEmail from $SharedMailboxEmail" "SUCCESS"
                $RemovedCount++
            }
            else {
                Write-Log "No permissions found for $UserEmail on $SharedMailboxEmail"
            }
        }
        catch {
            Write-Log "Error processing $UserEmail on $SharedMailboxEmail : $_" "ERROR"
            $ErrorCount++
        }
    }
}

# Summary report
Write-Log "========== PROCESS SUMMARY =========="
Write-Log "Total entries processed: $ProcessedCount"
Write-Log "Users removed from mailboxes: $RemovedCount"
Write-Log "Errors encountered: $ErrorCount"
Write-Log "Log file saved to: $LogFile"

Write-Host "`nProcess completed. Check log file for details: $LogFile" -ForegroundColor Green

Pause