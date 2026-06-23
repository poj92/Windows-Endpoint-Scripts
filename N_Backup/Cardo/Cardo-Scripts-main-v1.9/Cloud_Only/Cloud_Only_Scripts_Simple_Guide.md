# Cloud-Only Scripts – Quick Guide

## Prerequisites
- PowerShell 5.1+ (PowerShell 7 recommended)
- Modules: Microsoft.Graph (all), ExchangeOnlineManagement (leaver only)
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```
- Sign in when prompted (admin account). Default CSV is `UserImport.csv` in the script folder; override with `-CsvPath "C:\Path\file.csv"`.

---

## Scripts at a Glance

**Entra_BulkUserImport.ps1** – Create users from CSV
- Run: `./Entra_BulkUserImport.ps1 [-CsvPath "C:\NewEmployees.csv"]`
- CSV (required): First name, Last name, Display name, User principal name, Password, Account status, Company
- Optional columns: Job Title, Department, Telephone number, Mobile, E-mail, Manager Email, IsFieldOperative, D365FieldService, Licenses, Groups, StreetAddress, City, Post code, Country
- Logs: `Entra_BulkCreateUsers_YYYYMMDD_HHMMSS.log`

**Entra_Process_Leaver.ps1** – Offboard users
- Run: `./Entra_Process_Leaver.ps1 [-CsvPath "C:\Leavers.csv"] [-ExchangeUserPrincipalName "exchangeadmin@contoso.com"]`
- CSV: User principal name
- Exchange step is optional (leave `-ExchangeUserPrincipalName` blank to skip)
- Logs: `Entra_Process_Leaver_YYYYMMDD_HHMMSS.log`

**Entra_UpdateUser.ps1** – Update addresses/phones
- Run: `./Entra_UpdateUser.ps1 [-CsvPath "C:\Updates.csv"] [-PropertyToUpdate mobile|address|all]`
- CSV: User principal name + columns to change (Mobile, Telephone number, StreetAddress, City, Post Code, Country)
- Logs: `Entra_UpdateUser_YYYYMMDD_HHMMSS.log`

**Entra_SetSMTPAddress.ps1** – Set primary email
- Run: `./Entra_SetSMTPAddress.ps1 [-CsvPath "C:\EmailUpdates.csv"]`
- CSV: User principal name, E-mail (new primary)

**Entra_PhoneAuthentication.ps1** – Set MFA phone
- Run: `./Entra_PhoneAuthentication.ps1 [-CsvPath "C:\Phones.csv"]`
- CSV: User principal name, Mobile (international, e.g., +1-555-1234)

**Entra_Set_Operative.ps1** – Set operative/non-operative
- Run: `./Entra_Set_Operative.ps1 [-CsvPath "C:\Operative.csv"]`
- CSV: User principal name, IsFieldOperative (Operative | Non-Operative)

**Entra_ExportAllUsers.ps1** – Export all users
- Run: `./Entra_ExportAllUsers.ps1 [-OutputPath "C:\Exports\AllUsers.csv"]`
- No input CSV; exports 30+ columns. Logs: `Entra_ExportAllUsers_YYYYMMDD_HHMMSS.log`

---

## Quick Tips
- Keep CSV headers exact; use UTF-8 without BOM
- Test with 1–2 rows before bulk runs
- Run from the script folder so default paths work
- Check the generated log after each run for errors
