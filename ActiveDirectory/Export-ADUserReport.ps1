<#
.SYNOPSIS
    Generates a comprehensive report of Active Directory users including account information, group memberships, and last logon time.

.DESCRIPTION
    This script generates a detailed report of Active Directory users based on specified criteria.
    It can include account information, group memberships, and last logon time, and export the report
    in various formats. It provides detailed logging and error handling.

.PARAMETER SearchBase
    The OU to search for users.

.PARAMETER Filter
    LDAP filter to apply to the search.

.PARAMETER Properties
    Array of user properties to include in the report.

.PARAMETER IncludeGroups
    Whether to include group memberships in the report.

.PARAMETER IncludeLastLogon
    Whether to include last logon information in the report.

.PARAMETER ExportPath
    Path where the report will be saved.

.PARAMETER ExportFormat
    Format of the export file (CSV, JSON, Excel, HTML).

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Export-ADUserReport.ps1 -SearchBase "OU=Users,DC=contoso,DC=com" -Filter "Department -eq 'IT'" -Properties @("Name", "Title", "Department", "Manager", "EmailAddress", "Enabled") -IncludeGroups $true -IncludeLastLogon $true -ExportPath "C:\Reports\ADUsers.xlsx" -ExportFormat "Excel"

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string]$Filter = "*",

    [Parameter(Mandatory = $false)]
    [string[]]$Properties = @("Name", "SamAccountName", "UserPrincipalName", "Enabled", "PasswordLastSet", "PasswordNeverExpires", "PasswordExpired", "LockedOut", "AccountExpirationDate", "Department", "Title", "Manager", "EmailAddress", "WhenCreated", "WhenChanged"),

    [Parameter(Mandatory = $false)]
    [bool]$IncludeGroups = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeLastLogon = $true,

    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Excel", "HTML")]
    [string]$ExportFormat = "CSV",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ADUserReport_$(Get-Date -Format 'yyyyMMdd').log"
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path -Path $LogPath -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    Add-Content -Path $LogPath -Value $logMessage
    
    switch ($Level) {
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
    }
}

function Test-ADModule {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Log -Message "ActiveDirectory module not found. Installing..." -Level "WARNING"
        try {
            Import-Module ServerManager
            Add-WindowsFeature RSAT-AD-PowerShell
            Import-Module ActiveDirectory
            Write-Log -Message "ActiveDirectory module installed successfully." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to install ActiveDirectory module: $_" -Level "ERROR"
            return $false
        }
    }
    else {
        try {
            Import-Module ActiveDirectory
            Write-Log -Message "ActiveDirectory module imported successfully." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to import ActiveDirectory module: $_" -Level "ERROR"
            return $false
        }
    }
    return $true
}

function Get-LastLogonTime {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName
    )
    
    try {
        $domainControllers = Get-ADDomainController -Filter *
        $lastLogonTimes = @()
        
        foreach ($dc in $domainControllers) {
            $user = Get-ADUser -Identity $SamAccountName -Server $dc.HostName -Properties lastLogon
            if ($user.lastLogon -gt 0) {
                $lastLogonTimes += [DateTime]::FromFileTime($user.lastLogon)
            }
        }
        
        if ($lastLogonTimes.Count -gt 0) {
            return ($lastLogonTimes | Sort-Object -Descending)[0]
        }
        else {
            return "Never logged on"
        }
    }
    catch {
        Write-Log -Message "Failed to get last logon time for user '$SamAccountName': $_" -Level "WARNING"
        return "Error retrieving last logon"
    }
}

function Get-UserGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName
    )
    
    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $SamAccountName | Select-Object -ExpandProperty Name
        return ($groups -join "; ")
    }
    catch {
        Write-Log -Message "Failed to get group memberships for user '$SamAccountName': $_" -Level "WARNING"
        return "Error retrieving groups"
    }
}

function Export-ReportToCSV {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Report exported to CSV successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export report to CSV: $_" -Level "ERROR"
        return $false
    }
}

function Export-ReportToJSON {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $Data | ConvertTo-Json -Depth 4 | Out-File -FilePath $Path -Encoding UTF8
        Write-Log -Message "Report exported to JSON successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export report to JSON: $_" -Level "ERROR"
        return $false
    }
}

function Export-ReportToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        # Check if ImportExcel module is available
        if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
            Write-Log -Message "ImportExcel module not found. Installing..." -Level "WARNING"
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        
        Import-Module ImportExcel
        $Data | Export-Excel -Path $Path -AutoSize -TableName "ADUsers" -WorksheetName "AD Users Report"
        Write-Log -Message "Report exported to Excel successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export report to Excel: $_" -Level "ERROR"
        return $false
    }
}

function Export-ReportToHTML {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Active Directory User Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
    </style>
</head>
<body>
    <h1>Active Directory User Report</h1>
    <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <table>
        <tr>
"@
        
        $htmlColumns = ""
        $properties = $Data[0].PSObject.Properties.Name
        foreach ($prop in $properties) {
            $htmlColumns += "            <th>$prop</th>`n"
        }
        
        $htmlRows = ""
        foreach ($row in $Data) {
            $htmlRows += "        <tr>`n"
            foreach ($prop in $properties) {
                $value = $row.$prop
                if ($null -eq $value) { $value = "" }
                $htmlRows += "            <td>$value</td>`n"
            }
            $htmlRows += "        </tr>`n"
        }
        
        $htmlFooter = @"
    </table>
</body>
</html>
"@
        
        $html = $htmlHeader + $htmlColumns + "</tr>`n" + $htmlRows + $htmlFooter
        $html | Out-File -FilePath $Path -Encoding UTF8
        
        Write-Log -Message "Report exported to HTML successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export report to HTML: $_" -Level "ERROR"
        return $false
    }
}

# Main script execution
try {
    Write-Log -Message "Starting AD user report generation process." -Level "INFO"
    
    # Check if ActiveDirectory module is available
    if (-not (Test-ADModule)) {
        Write-Log -Message "Exiting script due to missing ActiveDirectory module." -Level "ERROR"
        exit 1
    }
    
    # Check if search base exists
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$SearchBase'" -ErrorAction SilentlyContinue)) {
        Write-Log -Message "Search base '$SearchBase' does not exist." -Level "ERROR"
        exit 1
    }
    
    # Create export directory if it doesn't exist
    $exportDir = Split-Path -Path $ExportPath -Parent
    if (-not (Test-Path -Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
        Write-Log -Message "Created export directory '$exportDir'." -Level "INFO"
    }
    
    # Get AD users
    Write-Log -Message "Retrieving AD users from '$SearchBase' with filter '$Filter'." -Level "INFO"
    $adUsers = Get-ADUser -Filter $Filter -SearchBase $SearchBase -Properties $Properties
    Write-Log -Message "Retrieved $($adUsers.Count) users." -Level "INFO"
    
    # Create report data
    $reportData = @()
    $counter = 0
    $total = $adUsers.Count
    
    foreach ($user in $adUsers) {
        $counter++
        Write-Progress -Activity "Processing AD Users" -Status "Processing $counter of $total" -PercentComplete (($counter / $total) * 100)
        
        $userData = [ordered]@{}
        
        # Add standard properties
        foreach ($prop in $Properties) {
            if ($prop -eq "Manager" -and $user.Manager) {
                try {
                    $managerName = (Get-ADUser -Identity $user.Manager).Name
                    $userData[$prop] = $managerName
                }
                catch {
                    $userData[$prop] = $user.Manager
                }
            }
            else {
                $userData[$prop] = $user.$prop
            }
        }
        
        # Add group memberships if requested
        if ($IncludeGroups) {
            $userData["GroupMemberships"] = Get-UserGroups -SamAccountName $user.SamAccountName
        }
        
        # Add last logon time if requested
        if ($IncludeLastLogon) {
            $userData["LastLogon"] = Get-LastLogonTime -SamAccountName $user.SamAccountName
        }
        
        $reportData += [PSCustomObject]$userData
    }
    
    # Export report in the specified format
    $exportSuccess = $false
    
    switch ($ExportFormat) {
        "CSV" {
            $exportSuccess = Export-ReportToCSV -Data $reportData -Path $ExportPath
        }
        "JSON" {
            $exportSuccess = Export-ReportToJSON -Data $reportData -Path $ExportPath
        }
        "Excel" {
            $exportSuccess = Export-ReportToExcel -Data $reportData -Path $ExportPath
        }
        "HTML" {
            $exportSuccess = Export-ReportToHTML -Data $reportData -Path $ExportPath
        }
    }
    
    if ($exportSuccess) {
        Write-Log -Message "AD user report generation completed successfully." -Level "INFO"
    }
    else {
        Write-Log -Message "AD user report generation completed with errors." -Level "WARNING"
    }
}
catch {
    Write-Log -Message "An error occurred during AD user report generation: $_" -Level "ERROR"
    exit 1
}
