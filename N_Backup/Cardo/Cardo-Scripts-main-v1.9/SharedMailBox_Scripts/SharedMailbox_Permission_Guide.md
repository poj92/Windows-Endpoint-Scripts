# How to Remove Shared Mailbox Permissions - User Guide

## What You Need

**1. Install Exchange Online Management Module (One-time setup)**
```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
```

**2. Choose Your Script**
- **SharedMailbox_Remove_FullAccess.ps1** - Removes only FullAccess permissions
- **SharedMailbox_Remove_All_Permissions.ps1** - Removes ALL permissions (FullAccess, Send As, Send On Behalf)

---

## Step-by-Step Instructions

### Step 1: Prepare Your CSV Files

Create two CSV files in the **same folder** as the PowerShell script:

**File 1: users.csv** (Users to remove)
```csv
Email
john.doe@company.com
jane.smith@company.com
```

**File 2: mailboxes.csv** (Shared mailboxes to process)
```csv
Email
sales@company.com
support@company.com
```

**Important:** 
- Column name must be: `Email`, `UserEmail`, `user_email`, or `email_address`
- Save files in the same folder as the script (e.g., `C:\Scripts\`)
- Use `.csv` format (Save As → CSV in Excel)

---

### Step 2: Run the Script

**Option A: Quick Run (Files in Same Folder)**
1. Open PowerShell
2. Navigate to script folder: `cd C:\Scripts`
3. Run the command:

**For FullAccess removal only:**
```powershell
.\SharedMailbox_Remove_FullAccess.ps1 -UsersCsvPath "users.csv" -MailboxesCsvPath "mailboxes.csv"
```

**For ALL permissions removal:**
```powershell
.\SharedMailbox_Remove_All_Permissions.ps1 -UsersCsvPath "users.csv" -MailboxesCsvPath "mailboxes.csv"
```

**Option B: Interactive Mode (Script Prompts You)**
1. Open PowerShell
2. Navigate to script folder: `cd C:\Scripts`
3. Run: `.\SharedMailbox_Remove_FullAccess.ps1` (or the All Permissions script)
4. When prompted, enter the file paths

---

### Step 3: Authenticate to Exchange Online

- A login window will appear
- Sign in with your **admin credentials**
- Wait for "Successfully connected to Exchange Online" message

---

### Step 4: Monitor Progress

The script will display:
- Which user/mailbox combination is being processed
- Permissions found and removed
- Any errors encountered

---

### Step 5: Review Results

When complete:
- Review the on-screen summary (total processed, removed, errors)
- Check the log file created in the same folder:
  - `SharedMailbox_Removal_YYYYMMDD_HHMMSS.log` (FullAccess script)
  - `SharedMailbox_AllPermissions_Removal_YYYYMMDD_HHMMSS.log` (All Permissions script)
- Press any key to close the window

---

## Recommended File Organization

```
C:\Scripts\SharedMailboxPermissions\
├── SharedMailbox_Remove_FullAccess.ps1
├── SharedMailbox_Remove_All_Permissions.ps1
├── users.csv
├── mailboxes.csv
└── Logs\
    ├── SharedMailbox_Removal_20260112_103045.log
    └── SharedMailbox_AllPermissions_Removal_20260112_143022.log
```

---

## Quick Tips

✅ **DO:**
- Save CSV files in the same folder as the script
- Use simple filenames: `users.csv` and `mailboxes.csv`
- Check the log file after completion
- Test with one user/mailbox first

❌ **DON'T:**
- Don't close the PowerShell window while script is running
- Don't use special characters in CSV filenames
- Don't forget the header row in CSV files
