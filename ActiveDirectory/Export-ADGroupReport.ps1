<#
.SYNOPSIS
    Generates a report of Active Directory groups and their members.

.DESCRIPTION
    This script generates a detailed report of Active Directory groups based on specified criteria.
    It can include group members and nested group memberships, and export the report in various formats.
    It provides detailed logging and error handling.

.PARAMETER SearchBase
    The OU to search for groups.

.PARAMETER Filter
    LDAP filter to apply to the search.

.PARAMETER IncludeMembers
    Whether to include group members in the report.

.PARAMETER IncludeNestedGroups
    Whether to include nested group memberships.

.PARAMETER ExportPath
    Path where the report will be saved.

.PARAMETER ExportFormat
    Format of the export file (CSV, JSON, Excel, HTML).

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Export-ADGroupReport.ps1 -SearchBase "OU=Groups,DC=contoso,DC=com" -Filter "*" -IncludeMembers $true -IncludeNestedGroups $true -ExportPath "C:\Reports\ADGroups.xlsx" -ExportFormat "Excel"

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
    [bool]$IncludeMembers = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeNestedGroups = $false,

    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Excel", "HTML")]
    [string]$ExportFormat = "CSV",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ADGroupReport_$(Get-Date -Format 'yyyyMMdd').log"
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

function Get-GroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        
        [Parameter(Mandatory = $false)]
        [bool]$Recursive = $false
    )
    
    try {
        $members = @()
        $group = Get-ADGroup -Identity $GroupName
        
        $directMembers = Get-ADGroupMember -Identity $GroupName -ErrorAction SilentlyContinue
        
        foreach ($member in $directMembers) {
            $memberInfo = [PSCustomObject]@{
                GroupName = $group.Name
                MemberName = $member.Name
                MemberSamAccountName = $member.SamAccountName
                MemberType = $member.objectClass
                IsDirect = $true
                ParentGroup = $null
            }
            
            $members += $memberInfo
            
            if ($Recursive -and $member.objectClass -eq "group") {
                $nestedMembers = Get-GroupMembers -GroupName $member.SamAccountName -Recursive $true
                
                foreach ($nestedMember in $nestedMembers) {
                    $nestedMemberInfo = [PSCustomObject]@{
                        GroupName = $group.Name
                        MemberName = $nestedMember.MemberName
                        MemberSamAccountName = $nestedMember.MemberSamAccountName
                        MemberType = $nestedMember.MemberType
                        IsDirect = $false
                        ParentGroup = $member.Name
                    }
                    
                    $members += $nestedMemberInfo
                }
            }
        }
        
        return $members
    }
    catch {
        Write-Log -Message "Failed to get members for group '$GroupName': $_" -Level "WARNING"
        return @()
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
        $Data | Export-Excel -Path $Path -AutoSize -TableName "ADGroups" -WorksheetName "AD Groups Report"
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
    <title>Active Directory Group Report</title>
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
    <h1>Active Directory Group Report</h1>
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
    Write-Log -Message "Starting AD group report generation process." -Level "INFO"
    
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
    
    # Get AD groups
    Write-Log -Message "Retrieving AD groups from '$SearchBase' with filter '$Filter'." -Level "INFO"
    $adGroups = Get-ADGroup -Filter $Filter -SearchBase $SearchBase -Properties Description, GroupCategory, GroupScope, whenCreated, whenChanged, managedBy
    Write-Log -Message "Retrieved $($adGroups.Count) groups." -Level "INFO"
    
    # Create report data
    $reportData = @()
    $memberData = @()
    $counter = 0
    $total = $adGroups.Count
    
    foreach ($group in $adGroups) {
        $counter++
        Write-Progress -Activity "Processing AD Groups" -Status "Processing $counter of $total" -PercentComplete (($counter / $total) * 100)
        
        $groupData = [ordered]@{
            Name = $group.Name
            SamAccountName = $group.SamAccountName
            Description = $group.Description
            GroupCategory = $group.GroupCategory
            GroupScope = $group.GroupScope
            DistinguishedName = $group.DistinguishedName
            WhenCreated = $group.whenCreated
            WhenChanged = $group.whenChanged
        }
        
        # Add manager if available
        if ($group.managedBy) {
            try {
                $managerName = (Get-ADObject -Identity $group.managedBy).Name
                $groupData["ManagedBy"] = $managerName
            }
            catch {
                $groupData["ManagedBy"] = $group.managedBy
            }
        }
        else {
            $groupData["ManagedBy"] = ""
        }
        
        # Add member count
        if ($IncludeMembers) {
            $members = Get-GroupMembers -GroupName $group.SamAccountName -Recursive $IncludeNestedGroups
            $groupData["MemberCount"] = ($members | Where-Object { $_.IsDirect -eq $true }).Count
            
            if ($members.Count -gt 0) {
                $memberData += $members
            }
        }
        
        $reportData += [PSCustomObject]$groupData
    }
    
    # Export group report in the specified format
    $exportSuccess = $false
    
    switch ($ExportFormat) {
        "CSV" {
            $exportSuccess = Export-ReportToCSV -Data $reportData -Path $ExportPath
            
            if ($IncludeMembers -and $memberData.Count -gt 0) {
                $memberExportPath = [System.IO.Path]::ChangeExtension($ExportPath, $null) + "_Members" + [System.IO.Path]::GetExtension($ExportPath)
                Export-ReportToCSV -Data $memberData -Path $memberExportPath
            }
        }
        "JSON" {
            $exportSuccess = Export-ReportToJSON -Data $reportData -Path $ExportPath
            
            if ($IncludeMembers -and $memberData.Count -gt 0) {
                $memberExportPath = [System.IO.Path]::ChangeExtension($ExportPath, $null) + "_Members" + [System.IO.Path]::GetExtension($ExportPath)
                Export-ReportToJSON -Data $memberData -Path $memberExportPath
            }
        }
        "Excel" {
            # Check if ImportExcel module is available
            if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
                Write-Log -Message "ImportExcel module not found. Installing..." -Level "WARNING"
                Install-Module -Name ImportExcel -Force -Scope CurrentUser
            }
            
            Import-Module ImportExcel
            
            $reportData | Export-Excel -Path $ExportPath -AutoSize -TableName "ADGroups" -WorksheetName "Groups"
            
            if ($IncludeMembers -and $memberData.Count -gt 0) {
                $memberData | Export-Excel -Path $ExportPath -AutoSize -TableName "ADGroupMembers" -WorksheetName "Members"
            }
            
            $exportSuccess = $true
            Write-Log -Message "Report exported to Excel successfully at '$ExportPath'." -Level "INFO"
        }
        "HTML" {
            $exportSuccess = Export-ReportToHTML -Data $reportData -Path $ExportPath
            
            if ($IncludeMembers -and $memberData.Count -gt 0) {
                $memberExportPath = [System.IO.Path]::ChangeExtension($ExportPath, $null) + "_Members" + [System.IO.Path]::GetExtension($ExportPath)
                Export-ReportToHTML -Data $memberData -Path $memberExportPath
            }
        }
    }
    
    if ($exportSuccess) {
        Write-Log -Message "AD group report generation completed successfully." -Level "INFO"
    }
    else {
        Write-Log -Message "AD group report generation completed with errors." -Level "WARNING"
    }
}
catch {
    Write-Log -Message "An error occurred during AD group report generation: $_" -Level "ERROR"
    exit 1
}
