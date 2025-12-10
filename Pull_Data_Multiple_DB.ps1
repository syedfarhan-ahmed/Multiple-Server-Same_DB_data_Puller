$moduleName = 'ImportExcel'

# Check if ImportExcel module is installed
$module = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue
if ($null -eq $module) {
    Write-Host "$moduleName not found. Installing..."
    Install-Module -Name $moduleName -Force -Scope CurrentUser
} else {
    Write-Host "$moduleName is already installed. Skipping installation."
}

Import-Module SQLServer
Import-Module Az.Accounts
Import-Module ImportExcel

# Authenticate using Azure AD (with MFA)
Write-Host "Authenticating to Azure AD..."
$aadCredential = Connect-AzAccount

if ($aadCredential -eq $null) {
    Write-Host "Authentication failed. Please ensure you are logged in with your Azure account."
    exit
}

# Extract the Azure AD tenant ID and username (typically the signed-in user)
$tenantId = (Get-AzContext).Tenant.Id
$username = (Get-AzContext).Account.Id

# Get the database name and SQL query from user input
$databaseName = Read-Host "Enter the database name"
$query = Read-Host "Enter the SQL query"

# Mapping of server hostnames to friendly region names
$serverMap = @{
    "xyz.database.windows.net" = "XYZ"
    "server2.database.windows.net" = "S2"
}

# Get the location where the script is being executed
$scriptDirectory = $PSScriptRoot
if (-not $scriptDirectory) {
    # If running interactively, use the current directory
    $scriptDirectory = Get-Location
}

# Define the file path based on the database name and the script's location
$filePath = Join-Path -Path $scriptDirectory -ChildPath "$databaseName.xlsx"

Write-Host "Exporting data to: $filePath"

foreach ($server in $serverMap.Keys) {
    $sheetName = $serverMap[$server]
    Write-Host "Exporting data for DB: $server as $sheetName"

    # Use Azure AD authentication to connect to the SQL Server
    $connectionString = "Server=$server;Database=$databaseName;Authentication=ActiveDirectoryInteractive;"

    # Run the SQL query using Azure AD authentication
    $data = Invoke-Sqlcmd -Query $query -ConnectionString $connectionString | 
            Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors

    # Export the data to an Excel file
    if (-not (Test-Path $filePath)) {
        $data | Export-Excel -Path $filePath -WorksheetName $sheetName -AutoSize
    } else {
        $data | Export-Excel -Path $filePath -WorksheetName $sheetName -AutoSize -Append
    }
}

Write-Host "Process completed successfully."
