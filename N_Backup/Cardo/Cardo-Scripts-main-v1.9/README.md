# Cardo-Scripts

A comprehensive collection of PowerShell scripts for managing users in Microsoft Entra ID (Azure AD) and on-premises Active Directory. These scripts automate common administrative tasks including bulk user creation, offboarding, profile updates, and authentication configuration.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Cloud_Only](#cloud_only)
  - [Entra_BulkUserImport.ps1](#entra_bulkuserimportps1)
  - [Entra_Process_Leaver.ps1](#entra_process_leaverps1)
  - [Entra_UpdateUser.ps1](#entra_updateuserps1)
  - [Entra_SetSMTPAddress.ps1](#entra_setsmtpaddressps1)
  - [Entra_PhoneAuthentication.ps1](#entra_phoneauthenticationps1)
  - [Entra_Set_Operative.ps1](#entra_set_operativeps1)
- [Hybrid](#hybrid)
  - [Hybrid Scripts](#hybrid-scripts)
- [Getting Started](#getting-started)
- [Troubleshooting](#troubleshooting)
- [Support](#support)

---

## Overview

This script collection provides enterprise-grade user management automation for Microsoft Entra ID environments. The scripts handle complex operations such as:

- **Bulk User Creation** - Create hundreds of users with properties, licenses, and group assignments
- **User Offboarding** - Automatically disable accounts, convert mailboxes, remove licenses and group memberships
- **Profile Management** - Update user properties including contact information and job details
- **Email Management** - Configure default SMTP addresses for users
- **Authentication Setup** - Configure phone numbers for multi-factor authentication

All scripts include:
- **Comprehensive Logging** - Timestamped logs with color-coded console output for easy monitoring
- **Error Handling** - Non-fatal error handling to continue processing when individual operations fail
- **Automatic Module Installation** - Required PowerShell modules are installed automatically
- **CSV Input** - Easy-to-use CSV-based input format for bulk operations
- **Microsoft Graph Integration** - Uses the modern Microsoft Graph PowerShell SDK

---

## Prerequisites

### System Requirements

- **PowerShell** 5.1 or later (PowerShell 7.x recommended)
- **Windows**, macOS, or Linux with PowerShell installed
- **Internet connectivity** for module installation and Graph API access

### Required Modules

The scripts automatically install these modules if not present:

- **Microsoft.Graph** - Modern Microsoft Graph PowerShell SDK
- **ExchangeOnlineManagement** - Required for mailbox operations (Entra_Process_Leaver.ps1 only)

Manual installation (optional):
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Permissions Required

You must have the following permissions in your Entra ID tenant:

**For Entra_BulkUserImport.ps1:**
- `User.ReadWrite.All` - Create and modify users
- `Directory.ReadWrite.All` - Manage groups and licenses

**For Entra_Process_Leaver.ps1:**
- `User.ReadWrite.All` - Disable users and manage properties
- `Directory.ReadWrite.All` - Manage groups and licenses
- Exchange Online admin permissions (for mailbox conversion)

**For Entra_UpdateUser.ps1:**
- `User.ReadWrite.All` - Update user properties

**For Entra_SetSMTPAddress.ps1:**
- `User.ReadWrite.All` - Update email properties

**For Entra_PhoneAuthentication.ps1:**
- `User.ReadWrite.All` - Read user information
- `UserAuthenticationMethod.ReadWrite.All` - Manage phone authentication methods

### Network & Access

- Admin account with appropriate tenant permissions
- No IP restrictions blocking Microsoft Graph API access
- For Exchange Online operations: Exchange Online admin account

---

## Cloud_Only

Scripts for managing users in cloud-only Microsoft Entra ID environments (without on-premises AD sync). These scripts work with Entra ID exclusively and do not require on-premises Active Directory infrastructure.

Download the concise Cloud-Only Scripts Guide (PDF): [Cloud_Only_Scripts_Guide.pdf](Cloud_Only_Scripts_Guide.pdf)

### Entra_BulkUserImport.ps1

**Purpose:** Bulk creates users in Entra ID from a CSV file with complete profile setup including licenses, groups, and manager relationships.

#### Features

✓ Creates multiple users from CSV in a single operation
✓ Sets comprehensive user properties (name, email, phone, address, job title, department, company, country)
✓ Assigns licenses with flexible SKU matching (by name or GUID)
✓ Adds users to security groups
✓ Configures manager relationships
✓ Sets custom extension attributes (1-2)
✓ Configures phone and email properties
✓ Skips existing users automatically
✓ Validates required fields (extension attributes, company name)
✓ Timestamped logging with color-coded output
✓ Detailed error messages with non-fatal error handling

#### Usage

```powershell
# Basic usage (looks for UserImport.csv in the script directory)
.\Entra_BulkUserImport.ps1

# Specify custom CSV file path
.\Entra_BulkUserImport.ps1 -CsvPath "C:\Users\Admin\Desktop\NewEmployees.csv"
```

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| CsvPath | string | No | UserImport.csv | Path to the CSV file containing user details |

#### CSV Format

Create a `UserImport.csv` file with the following columns:

```csv
First name,Last name,Display name,User logon name,User principal name,Password,Account status,Job Title,Department,Company,Telephone number,Mobile,E-mail,Manager Email,IsFieldOperative,D365FieldService,Licenses,Groups,StreetAddress,City,Post code,Country
John,Doe,John Doe,johndoe,john.doe@contoso.com,P@ssw0rd!2024,Enabled,Software Engineer,Engineering,Contoso,555-1234,555-5678,john.doe@contoso.com,manager@contoso.com,Yes,No,"OFFICE365_BUSINESS_PREMIUM,TEAMS_ESSENTIALS","Engineering Team,All Users",123 Main St,New York,10001,United States
Jane,Smith,Jane Smith,janesmith,jane.smith@contoso.com,SecurePass123!,Enabled,Project Manager,Engineering,Contoso,555-4321,555-8765,jane.smith@contoso.com,manager@contoso.com,No,Yes,"OFFICE365_ENTERPRISE_E5","Engineering Team",456 Oak Ave,San Francisco,94102,United States
```

#### CSV Column Reference

| Column | Required | Type | Description | Example |
|--------|----------|------|-------------|---------|
| **Basic Identity** | | | | |
| First name | Yes | string | User's first name | John |
| Last name | Yes | string | User's last name | Doe |
| Display name | Yes | string | Full name as displayed in directory | John Doe |
| User logon name | No | string | On-premises login (legacy) | johndoe |
| User principal name | Yes | string | Cloud UPN - email format (MUST be unique) | john.doe@contoso.com |
| Password | Yes | string | Initial password (user must change on first login) | P@ssw0rd!2024 |
| Account status | Yes | string | "Enabled" or "Disabled" | Enabled |
| **Contact Information** | | | | |
| Telephone number | No | string | Office/desk phone number | 555-1234 |
| Mobile | No | string | Mobile/cell phone number | 555-5678 |
| E-mail | No | string | Email address (primary SMTP) | john.doe@contoso.com |
| **Organization** | | | | |
| Job Title | No | string | Job title/position | Software Engineer |
| Department | No | string | Department name | Engineering |
| Company | No | string | Company name (REQUIRED for creation to proceed) | Contoso |
| **Address** | | | | |
| StreetAddress | No | string | Street address | 123 Main St |
| City | No | string | City name | New York |
| Post code | No | string | ZIP/postal code | 10001 |
| Country | No | string | Country name | United States |
| **Management** | | | | |
| Manager Email | No | string | Manager's UPN/email for hierarchy | manager@contoso.com |
| **Custom Fields** | | | | |
| IsFieldOperative | No | string | Extension attribute 1 (custom field) | Yes |
| D365FieldService | No | string | Extension attribute 2 (custom field) | No |
| **Licensing & Groups** | | | | |
| Licenses | No | string (comma-separated) | Licenses to assign - skuPartNumber or GUID | OFFICE365_BUSINESS_PREMIUM,TEAMS_ESSENTIALS |
| Groups | No | string (comma-separated) | Security groups to add user to (groups must exist) | Engineering Team,All Users |

#### Column Notes

- **User principal name (UPN)** - Must be unique within the tenant and follow email format
- **Password** - Should meet your organization's complexity requirements
- **Company** - This field is mandatory; users without a company will be skipped
- **Licenses** - Can use either the friendly name (e.g., OFFICE365_BUSINESS_PREMIUM) or the SKU GUID; script looks up both automatically
- **Groups** - Groups must exist in Entra ID before the script runs; script will skip groups that don't exist
- **Manager Email** - Must be an existing user's UPN; script will handle if manager doesn't exist
- **Extension Attributes** - Custom fields that must be filled; if either is missing, user creation is skipped

#### Processing Flow

For each user in the CSV:

1. **Validation** - Check for required fields (UPN, Company, Extension Attributes)
2. **Duplicate Check** - Skip if user already exists in Entra ID
3. **User Creation** - Create user account with basic properties
4. **Extension Attributes** - Set custom fields (IsFieldOperative, D365FieldService)
5. **License Assignment** - Assign all specified licenses
6. **Group Membership** - Add user to all specified groups
7. **Manager Setup** - Set manager relationship in organizational hierarchy

#### Output

Creates a timestamped log file: `Entra_BulkCreateUsers_YYYYMMDD_HHMMSS.log`

Log includes:
- Successful user creations with object IDs
- Skipped users with reasons
- License assignment results
- Group membership changes
- Manager relationship configuration
- Any errors encountered

#### Example Log Output

```
[2024-12-15T10:23:45] Connecting to Microsoft Graph...
[2024-12-15T10:23:47] Connected to Microsoft Graph.
[2024-12-15T10:23:48] Created user john.doe@contoso.com (id: a1b2c3d4-e5f6-7890-abcd-ef1234567890)
[2024-12-15T10:23:49]   - Extension attributes for john.doe@contoso.com have been set
[2024-12-15T10:23:50]   - Assigned licenses: OFFICE365_BUSINESS_PREMIUM
[2024-12-15T10:23:51]   - Added to group: Engineering Team
[2024-12-15T10:23:52]   - Manager set to manager@contoso.com
```

#### Troubleshooting

| Issue | Cause | Solution |
|-------|-------|----------|
| "CSV file not found" | CSV path is incorrect or file doesn't exist | Verify CSV path and filename |
| "User already exists" | User with that UPN already in tenant | Update CSV with unique UPNs or delete existing users |
| "Company name missing" | Required field is empty | Add Company value for all users |
| "License not found" | SKU name is misspelled or doesn't exist in tenant | Verify license names or use SKU GUIDs |
| "Group not found" | Group doesn't exist in Entra ID | Create groups first or verify group names |
| "Manager not found" | Manager's UPN doesn't exist | Verify manager email is correct |
| Authentication errors | Credentials lack required permissions | Ensure account has User.ReadWrite.All and Directory.ReadWrite.All scopes |

---

### Entra_Process_Leaver.ps1

**Purpose:** Automates the offboarding process for departing employees by disabling accounts, converting mailboxes, removing licenses, and clearing group memberships.

#### Features

✓ Disables user accounts in Entra ID (prevents sign-in)
✓ Converts user mailboxes to shared mailboxes (with Exchange Online connection)
✓ Removes all assigned licenses (frees up license seats)
✓ Removes user from all security groups (maintains group hygiene)
✓ Clears all extension attributes (1-15)
✓ Non-fatal error handling (continues if one step fails)
✓ Optional Exchange Online integration
✓ Timestamped logging with color-coded output
✓ Detailed processing logs for compliance/audit

#### Usage

```powershell
# Basic usage (prompts for Exchange Online admin UPN)
.\Entra_Process_Leaver.ps1

# Specify CSV file (prompts for Exchange Online admin)
.\Entra_Process_Leaver.ps1 -CsvPath "C:\Leavers\December2024.csv"

# Specify both CSV and Exchange Online admin
.\Entra_Process_Leaver.ps1 -CsvPath "C:\Leavers\December2024.csv" -ExchangeUserPrincipalName "exchangeadmin@contoso.com"

# Skip mailbox conversion (Graph only)
.\Entra_Process_Leaver.ps1 -CsvPath "C:\Leavers\December2024.csv" -ExchangeUserPrincipalName ""
```

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| CsvPath | string | No | UserImport.csv | Path to CSV file containing users to offboard |
| ExchangeUserPrincipalName | string | No | (prompted) | Exchange Online admin UPN. Leave empty to skip mailbox conversion |

#### CSV Format

Minimum required columns:

```csv
User principal name
john.doe@contoso.com
jane.smith@contoso.com
bob.wilson@contoso.com
```

#### CSV Column Reference

| Column | Required | Description |
|--------|----------|-------------|
| User principal name | Yes | UPN of user to offboard (must match existing user) |

#### Processing Steps

For each user in the CSV, the script performs:

1. **Disable Account**
   - Sets `AccountEnabled` to `False`
   - Prevents user sign-in immediately
   - Preserves user object for compliance/audit

2. **Convert Mailbox** (if Exchange Online connected)
   - Changes user mailbox to shared mailbox
   - Allows other users to access archived messages
   - Requires Exchange Online admin permissions

3. **Clear Extension Attributes**
   - Removes values from extensionAttribute1-15
   - Clears custom organizational metadata
   - Resets to default state

4. **Remove from Groups**
   - Removes user from all security groups
   - Uses Graph API for clean member removal
   - Prevents access through group permissions

5. **Remove Licenses**
   - Unassigns all licenses from user
   - Frees license seats for new assignments
   - Removes access to licensed services (Office, Teams, etc.)

Each step is logged independently. If one step fails, processing continues to the next step.

#### Output

Creates a timestamped log file: `Entra_Process_Leaver_YYYYMMDD_HHMMSS.log`

Log includes:
- Account disable confirmation
- Mailbox conversion status
- Extension attribute clearing results
- Group removal details
- License removal confirmations
- Any warnings or errors encountered

#### Example Log Output

```
[2024-12-15T14:30:22] Connecting to Microsoft Graph...
[2024-12-15T14:30:24] Connected to Microsoft Graph.
[2024-12-15T14:30:25] Processing user john.doe@contoso.com...
[2024-12-15T14:30:26] User account disabled successfully.
[2024-12-15T14:30:27] Mailbox converted to shared mailbox.
[2024-12-15T14:30:28] Extension attributes cleared.
[2024-12-15T14:30:29] Removed from Engineering Team
[2024-12-15T14:30:29] Removed from All Users
[2024-12-15T14:30:30] Licenses removed: OFFICE365_BUSINESS_PREMIUM
[2024-12-15T14:30:31] Successfully processed john.doe@contoso.com
```

#### Important Notes

- **PowerShell ISE Warning** - This script may fail to connect to Exchange Online if run from PowerShell ISE. Use PowerShell Console or VS Code instead.
- **Exchange Online Optional** - If Exchange connection fails, the script continues with Graph-only operations (account disable, group removal, license removal).
- **Audit Trail** - All actions are logged for compliance purposes. Logs are retained in the script directory with timestamps.
- **Data Preservation** - Disabled accounts can be re-enabled if offboarding was made in error. Mailbox conversion is reversible with Exchange admin.
- **Permissions** - Ensure the account running the script has the required scopes, especially for mailbox operations.

#### Troubleshooting

| Issue | Cause | Solution |
|-------|-------|----------|
| "User not found" | UPN doesn't match any user in tenant | Verify correct spelling of UPN in CSV |
| "Exchange connection failed" | Exchange Online credentials invalid or missing | Provide correct admin UPN or skip mailbox conversion |
| "Group removal failed" | User not member of group or group deleted | Script logs the failure and continues |
| "License removal failed" | License already removed or in use | Script logs and continues |
| Partial processing | Exchange connection lost mid-process | Check Exchange Online status; rerun script for failed users |



### Entra_UpdateUser.ps1

**Purpose:** Updates existing user properties in Entra ID from a CSV file. Supports targeted updates to specific properties or comprehensive profile updates.

#### Features

✓ Updates address information (street, city, postal code, country)
✓ Updates phone numbers (mobile, office)
✓ Selective updates via PropertyToUpdate parameter
✓ Validates changes before applying updates
✓ Logs pending changes for audit/review
✓ Handles flexible CSV column naming
✓ Non-fatal error handling
✓ Detailed validation output

#### Usage

```powershell
# Update all properties
.\Entra_UpdateUser.ps1

# Update with custom CSV path
.\Entra_UpdateUser.ps1 -CsvPath "C:\Users\Admin\Desktop\UserUpdates.csv"

# Update mobile numbers only
.\Entra_UpdateUser.ps1 -PropertyToUpdate "mobile"

# Update address information only
.\Entra_UpdateUser.ps1 -PropertyToUpdate "address"
```

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| CsvPath | string | No | UserImport.csv | Path to CSV file with user updates |
| PropertyToUpdate | string | No | "all" | Properties to update: "all", "mobile", "address" |

#### CSV Format

Minimum required columns:

```csv
User principal name,Mobile,StreetAddress,City,Post Code,Country
john.doe@contoso.com,555-1234,123 Main St,New York,10001,United States
jane.smith@contoso.com,555-5678,456 Oak Ave,San Francisco,94102,United States
```

Alternative column names (script handles both):
- "User principal name" or "UserPrincipalName"
- "Post Code" or "PostalCode"
- "StreetAddress" or "Street address"

#### CSV Column Reference

| Column | Required | Description |
|--------|----------|-------------|
| User principal name | Yes | UPN of user to update |
| Mobile | No | Mobile phone number to set |
| StreetAddress | No | Street address for user |
| City | No | City for user's address |
| Post Code | No | Postal/ZIP code |
| Country | No | Country name |

#### Processing Flow

For each user in the CSV:

1. **Validation** - Verify user exists in Entra ID
2. **Check Updates** - Determine which properties to update based on parameters
3. **Build Update Parameters** - Prepare Graph API parameters
4. **Log Changes** - Display what will be updated (validation logging)
5. **Apply Updates** - Execute Update-MgUser with pending changes
6. **Log Results** - Record success or failure

#### Output

Creates a timestamped log file: `Entra_UpdateUser_YYYYMMDD_HHMMSS.log`

#### Example Log Output

```
[2024-12-15T11:15:30] Updating user john.doe@contoso.com...
[2024-12-15T11:15:30] Validation: Planning to update properties:
[2024-12-15T11:15:30]   - MobilePhone: 555-1234
[2024-12-15T11:15:30]   - StreetAddress: 123 Main St
[2024-12-15T11:15:30]   - City: New York
[2024-12-15T11:15:30]   - PostalCode: 10001
[2024-12-15T11:15:30]   - Country: United States
[2024-12-15T11:15:31] Successfully updated user john.doe@contoso.com
```

---

### Entra_SetSMTPAddress.ps1

**Purpose:** Sets the default SMTP address (primary email) for users in Entra ID from a CSV file.

#### Features

✓ Updates primary email address (Mail property)
✓ Sets default SMTP address in Entra ID
✓ Validates email addresses before updating
✓ Handles missing or invalid emails gracefully
✓ Timestamped logging
✓ Automatic module installation

#### Usage

```powershell
# Use default UserImport.csv in script directory
.\Entra_SetSMTPAddress.ps1

# Specify custom CSV path
.\Entra_SetSMTPAddress.ps1 -CsvPath "C:\EmailUpdates\NewAddresses.csv"
```

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| CsvPath | string | No | UserImport.csv | Path to CSV file with email updates |

#### CSV Format

Required columns:

```csv
User Principal Name,E-mail
john.doe@contoso.com,john.newemail@contoso.com
jane.smith@contoso.com,jane.newemail@contoso.com
```

#### CSV Column Reference

| Column | Required | Description |
|--------|----------|-------------|
| User Principal Name | Yes | UPN of user (may differ from old email) |
| E-mail | Yes | New email address to set as primary SMTP |

#### Processing Flow

1. **Validation** - Check for required fields (UPN, E-mail)
2. **User Lookup** - Retrieve user from Entra ID
3. **Update Email** - Set Mail property via Graph API
4. **Confirmation** - Log successful update

#### Output

Creates a timestamped log file: `Entra_SetSMTPAddress_YYYYMMDD_HHMMSS.log`

#### Example Log Output

```
[2024-12-15T12:45:20] Connecting to Microsoft Graph...
[2024-12-15T12:45:22] Connected to Microsoft Graph.
[2024-12-15T12:45:23] Updating default SMTP address for user john.doe@contoso.com to john.newemail@contoso.com...
[2024-12-15T12:45:24] Successfully updated default SMTP address for user john.doe@contoso.com.
```

---

### Entra_PhoneAuthentication.ps1

**Purpose:** Configures phone authentication numbers for users in Entra ID. Typically used to set up phone numbers for MFA (Multi-Factor Authentication).

#### Features

✓ Sets phone numbers for authentication/MFA
✓ Removes old phone methods and adds new ones
✓ Skips users with already-configured matching numbers
✓ Validates phone format before updating
✓ Timestamped logging
✓ Automatic module installation

#### Usage

```powershell
# Use default UserImport.csv in script directory
.\Entra_PhoneAuthentication.ps1

# Specify custom CSV path
.\Entra_PhoneAuthentication.ps1 -CsvPath "C:\PhoneSetup\UserPhones.csv"
```

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| CsvPath | string | No | UserImport.csv | Path to CSV file with phone numbers |

#### CSV Format

Required columns:

```csv
User Principal Name,Mobile
john.doe@contoso.com,+1-555-1234
jane.smith@contoso.com,+1-555-5678
bob.wilson@contoso.com,+44-20-1234-5678
```

#### CSV Column Reference

| Column | Required | Description | Format |
|--------|----------|-------------|--------|
| User Principal Name | Yes | UPN of user to configure | user@contoso.com |
| Mobile | Yes | Phone number for authentication | +1-555-1234 or +44-201-1234 |

Phone numbers should include country code in international format (+1 for US, +44 for UK, etc.)

#### Processing Flow

1. **Validation** - Check for required fields (UPN, Mobile)
2. **User Lookup** - Retrieve user from Entra ID
3. **Check Existing Methods** - Get current phone authentication methods
4. **Compare** - Check if phone number already configured
5. **Update** (if needed):
   - Remove old phone methods
   - Add new phone method
6. **Confirmation** - Log result (skipped or updated)

#### Output

Creates a timestamped log file: `Entra_PhoneAuthentication_YYYYMMDD_HHMMSS.log`

#### Example Log Output

```
[2024-12-15T13:30:15] Connecting to Microsoft Graph...
[2024-12-15T13:30:17] Connected to Microsoft Graph.
[2024-12-15T13:30:18] Updating phone authentication number for user john.doe@contoso.com to +1-555-1234...
[2024-12-15T13:30:19] Successfully updated phone authentication number for user john.doe@contoso.com.
[2024-12-15T13:30:20] Phone number for user jane.smith@contoso.com is already up to date.
```

#### Required Permissions

This script requires these Microsoft Graph scopes:
- `User.ReadWrite.All` - Read user information
- `UserAuthenticationMethod.ReadWrite.All` - Manage authentication methods

The script will prompt you to grant these permissions during first use.

---

## Hybrid

Scripts for managing users in hybrid environments with both on-premises Active Directory and Microsoft Entra ID. These scripts work with synchronized users and support both AD and cloud properties.

Download the comprehensive Hybrid Scripts Guide: [Markdown](Hybrid_Scripts_Guide.md) | [PDF](Hybrid_Scripts_Guide.pdf)

### Hybrid Scripts

The `Hybrid/` folder contains 18 PowerShell scripts for managing users in hybrid environments. All scripts use Active Directory cmdlets and automatically sync to Entra ID via AAD Connect.

- **User Creation & Management**
  - `CreateNewUser.ps1` - Create users in on-premises AD and sync to Entra ID
  - `MoveExistingUsersOU.ps1` - Move users between organizational units
  - `CreateCompanyOUs.ps1` - Create company-based OU structures

- **License Management**
  - `Set_E3NoTeams_License.ps1` - Assign E3 license without Teams
  - `Set_D365_Field_License.ps1` - Assign Dynamics 365 Field Service licenses
  - `Set_Operative_Or_NonOperative.ps1` - Set user type (operative or non-operative)

- **Contact Information Updates**
  - `Update_MobileNumber.ps1` - Update mobile phone numbers
  - `Update_Telephone.ps1` - Update office telephone numbers
  - `Update_Address.ps1` - Update address information
  - `Update_CompanyOrLocation.ps1` - Update company and location details
  - `Update_Fax_For_Direct-Dial.ps1` - Update fax numbers
  - `Set_SMTP_Address.ps1` - Set SMTP addresses

- **Phone Authentication**
  - `Add_PhoneAuthenticationMethod.ps1` - Add phone authentication for MFA
  - `Get_PhoneAuthenticationMethod.ps1` - Retrieve phone authentication details
  - `Remove_PhoneAuthenticationMethod.ps1` - Remove phone authentication methods
  - `Remove_MobileNumber.ps1` - Remove mobile numbers

- **Offboarding**
  - `Leaver_Script.ps1` - Comprehensive user offboarding automation

- **Password Management**
  - `Set_Passwords.ps1` - Set initial or reset user passwords
---

## Getting Started

### Quick Start Guide

#### 1. Setup (First Time Only)

```powershell
# Navigate to Cloud_Only script directory
cd C:\Scripts\Cardo-Scripts\Cloud_Only

# The scripts will automatically:
# - Install required PowerShell modules
# - Prompt you to authenticate to Microsoft Graph
# - Grant necessary permissions
```

#### 2. Prepare CSV File

Create your `UserImport.csv` with required columns for your task:

**For bulk user creation:** Use all columns from [Entra_BulkUserImport.ps1](#csv-column-reference)

**For updating users:** Use columns from [Entra_UpdateUser.ps1](#csv-column-reference-1)

**For offboarding:** Use single column (User principal name) from [Entra_Process_Leaver.ps1](#csv-format-1)

#### 3. Run Script

```powershell
# Run the script
.\ScriptName.ps1

# Or with custom CSV
.\ScriptName.ps1 -CsvPath "C:\Path\To\YourFile.csv"
```

#### 4. Review Logs

Logs are created in the script directory with timestamps:
- `Entra_BulkCreateUsers_YYYYMMDD_HHMMSS.log`
- `Entra_Process_Leaver_YYYYMMDD_HHMMSS.log`
- etc.

### Common Workflows

#### Workflow 1: Onboarding New Employees

1. Create `NewHires.csv` with user details
2. Run `.\Entra_BulkUserImport.ps1 -CsvPath "NewHires.csv"`
3. Users are created with licenses and group memberships
4. Run `.\Entra_PhoneAuthentication.ps1 -CsvPath "NewHires.csv"` to set MFA phones
5. Review logs in script directory

#### Workflow 2: Offboarding Departing Employee

1. Create `Leavers.csv` with departing user UPNs
2. Run `.\Entra_Process_Leaver.ps1 -CsvPath "Leavers.csv"`
3. Script disables account, converts mailbox, removes licenses/groups
4. Review the log to confirm all steps completed
5. Manually archive mailbox if needed

#### Workflow 3: Updating User Contact Information

1. Create `ContactUpdates.csv` with UPNs and new phone numbers
2. Run `.\Entra_UpdateUser.ps1 -CsvPath "ContactUpdates.csv" -PropertyToUpdate "mobile"`
3. Review log to verify updates applied
4. Repeat for address information with `-PropertyToUpdate "address"`

---

## Troubleshooting

### General Issues

#### "No Module Named Microsoft.Graph"

```powershell
# Solution: Install the module manually
Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
```

#### "Connect-MgGraph: Request_BadRequest"

**Cause:** Authentication failed or invalid credentials
**Solution:**
1. Ensure your account has admin permissions in the tenant
2. Run script again and provide correct credentials when prompted
3. Check for MFA requirements on your admin account

#### "Insufficient privileges to complete the operation"

**Cause:** Account lacks required permissions
**Solution:**
1. Verify account has correct admin role (Global Admin or delegated permissions)
2. Check Graph scopes - they may need to be re-consented
3. Run script with account that has appropriate permissions

### Module Issues

#### Module Installation Fails

```powershell
# Try with elevated permissions
Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -Verbose

# If still fails, check execution policy
Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### "ExchangeOnlineManagement Module Not Found"

**Applies to:** Entra_Process_Leaver.ps1 (mailbox conversion)
**Solution:**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

Or skip mailbox conversion by running with `-ExchangeUserPrincipalName ""`

### CSV & Data Issues

#### "CSV file not found: UserImport.csv"

**Cause:** CSV is not in the script directory
**Solution:**
1. Place CSV in same directory as script, OR
2. Provide full path: `.\Script.ps1 -CsvPath "C:\Path\To\File.csv"`

#### "User already exists, skipping"

**Cause:** UPN already exists in Entra ID
**Solution:**
1. Use different UPN in CSV, OR
2. Delete the existing user first, OR
3. Update the existing user with Entra_UpdateUser.ps1 instead

#### Column Name Mismatches

**Cause:** CSV column headers don't match expected names
**Solution:**
- Script accepts both "User principal name" and "UserPrincipalName"
- Check exact spacing and capitalization in CSV headers
- See CSV Column Reference section for each script

### Bulk Operation Issues

#### Only Some Users Created Successfully

**Cause:** Individual user records have issues
**Solution:**
1. Review the log file for specific error messages
2. Fix the problematic records in CSV
3. Re-run script - it will skip existing users and continue
4. Script is designed to handle this gracefully

#### License Assignment Failed

**Cause:** SKU name incorrect or doesn't exist in tenant
**Solution:**
1. Run this in PowerShell to list available licenses:
   ```powershell
   Connect-MgGraph
   Get-MgSubscribedSku | Select -Property SkuPartNumber, SkuId
   ```
2. Update CSV with correct license SKU names or GUIDs
3. Re-run script

#### Group Assignment Failed

**Cause:** Groups don't exist in Entra ID
**Solution:**
1. Create groups first: `New-MgGroup -DisplayName "GroupName" -MailEnabled $false -SecurityEnabled $true`
2. Or verify group names are spelled correctly in CSV
3. Re-run script

### Exchange Online Issues

#### "Failed to connect to Exchange Online"

**Cause:** Invalid Exchange admin credentials or Exchange connection issues
**Solution:**
1. Verify admin UPN is correct
2. Ensure account has Exchange Online admin role
3. Run PowerShell as Administrator
4. Try again with different credentials
5. Or skip mailbox conversion with `-ExchangeUserPrincipalName ""`

#### "This script may fail to connect to exchange online if run from PowerShell ISE"

**Solution:** Use PowerShell Console instead
```powershell
# Open PowerShell (not ISE)
# Navigate to script directory
# Run the script
```

#### Mailbox Conversion Skipped but Other Operations Succeeded

**This is normal behavior.** The script continues processing even if Exchange connection fails:
- Account disabled: ✓
- Mailbox converted: ✗ (Connection failed - that's OK)
- Licenses removed: ✓
- Groups removed: ✓

You can manually convert the mailbox later with Exchange admin.

---

## Support

### For Issues or Questions

Contact: **Peter James** at peter.james@nexusos.co.uk

### Information to Include in Support Request

- Script name and version
- PowerShell version: `$PSVersionTable.PSVersion`
- Relevant log file (attached)
- CSV sample data (sanitized)
- Error message or unexpected behavior
- Steps to reproduce the issue

---

## License

**Proprietary** - Nexus Open Systems (for Cardo Group Ltd)

## Change Log

### Version 1.0 (December 2024)

**Initial Release**
- Entra_BulkUserImport.ps1 - Bulk user creation with licenses and groups
- Entra_Process_Leaver.ps1 - User offboarding automation
- Entra_UpdateUser.ps1 - Profile updates
- Entra_SetSMTPAddress.ps1 - Email address management
- Entra_PhoneAuthentication.ps1 - Phone authentication setup
- Comprehensive logging and error handling
- Full documentation

---

## Additional Resources

### Microsoft Documentation

- [Microsoft Graph PowerShell SDK Documentation](https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview)
- [Microsoft Entra ID (Azure AD) Documentation](https://learn.microsoft.com/en-us/entra/)
- [Exchange Online Management Documentation](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2)

### PowerShell Resources

- [PowerShell Documentation](https://docs.microsoft.com/en-us/powershell/)
- [PowerShell Scripting Guide](https://docs.microsoft.com/en-us/powershell/scripting/overview)

---

*Last Updated: December 16, 2024*
