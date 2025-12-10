# Azure SQL Multi-Server Export Script

## Overview

This PowerShell script connects to Azure using interactive login (MFA supported), executes a SQL query against multiple Azure SQL Servers, and exports the results into an Excel file. Each SQL server's output is written into a separate worksheet within the same Excel file.

The script:
- Ensures the **ImportExcel** module is installed.
- Loads **Az.Accounts** and **SQLServer** modules.
- Authenticates the user using Azure AD.
- Accepts user input for **Database Name** and **SQL Query**.
- Runs the SQL query on multiple SQL servers.
- Exports all results into `<DatabaseName>.xlsx`, with unique worksheets per server.

---

## Prerequisites

- PowerShell 5.1+ or PowerShell 7+
- Access permissions to Azure SQL + MFA-enabled Azure AD login
- Internet access to install modules and authenticate
- SQLServer, Az.Accounts, and ImportExcel modules (script auto-installs ImportExcel)

---

## Modules Used

| Module        | Purpose                                           |
|---------------|---------------------------------------------------|
| ImportExcel   | Export output to Excel without Excel installed    |
| SQLServer     | Executes queries using Invoke-Sqlcmd              |
| Az.Accounts   | Azure authentication (MFA supported)              |

---

## How the Script Works

1. Validates and installs ImportExcel if missing.
2. Imports all required PowerShell modules.
3. Authenticates using `Connect-AzAccount`.
4. Prompts the user for:
   - Database name
   - SQL query
5. Iterates through a list of SQL Server hostnames defined in `$serverMap`.
6. Connects using Azure AD Interactive authentication.
7. Executes the provided SQL query on each SQL server.
8. Creates `<DatabaseName>.xlsx` or appends to it.
9. Exports each server's result to a dedicated worksheet.

---

## Output Details

- Output file name:
