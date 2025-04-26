<#
.SYNOPSIS
    Analyzes Windows event logs and generates reports based on specified criteria.

.DESCRIPTION
    This script analyzes Windows event logs on local or remote computers and generates reports
    based on specified criteria such as event ID, source, level, and time range. It can export
    the results in various formats and provides detailed logging and error handling.

.PARAMETER ComputerName
    Name of the target computer(s).

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER LogName
    Name of the event log to analyze (e.g., System, Application, Security).

.PARAMETER StartTime
    Start time for the event search.

.PARAMETER EndTime
    End time for the event search.

.PARAMETER EventID
    Event ID(s) to filter by.

.PARAMETER Source
    Event source(s) to filter by.

.PARAMETER Level
    Event level(s) to filter by (e.g., Error, Warning, Information).

.PARAMETER MaxEvents
    Maximum number of events to retrieve.

.PARAMETER ExportPath
    Path where the report will be saved.

.PARAMETER ExportFormat
    Format of the export file (CSV, JSON, Excel, HTML).

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Get-EventLogAnalysis.ps1 -ComputerName "DC01" -Credential (Get-Credential) -LogName "Security" -StartTime (Get-Date).AddDays(-1) -EventID 4625 -Level "Error" -MaxEvents 1000 -ExportPath "C:\Reports\FailedLogins.xlsx" -ExportFormat "Excel"

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string[]]$ComputerName = @($env:COMPUTERNAME),

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $true)]
    [string]$LogName,

    [Parameter(Mandatory = $false)]
    [DateTime]$StartTime,

    [Parameter(Mandatory = $false)]
    [DateTime]$EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]]$EventID,

    [Parameter(Mandatory = $false)]
    [string[]]$Source,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Error", "Warning", "Information", "SuccessAudit", "FailureAudit")]
    [string[]]$Level,

    [Parameter(Mandatory = $false)]
    [int]$MaxEvents = 1000,

    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Excel", "HTML")]
    [string]$ExportFormat = "CSV",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\EventLogAnalysis_$(Get-Date -Format 'yyyyMMdd').log"
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

function Test-PSRemoting {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )
    
    try {
        $result = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Get-EventLogData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$LogName,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$StartTime,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$EndTime,
        
        [Parameter(Mandatory = $false)]
        [int[]]$EventID,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Source,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Level,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxEvents
    )
    
    try {
        Write-Log -Message "Retrieving event log data from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$LogName,
                [DateTime]$StartTime,
                [DateTime]$EndTime,
                [int[]]$EventID,
                [string[]]$Source,
                [string[]]$Level,
                [int]$MaxEvents
            )
            
            # Build filter XPath query
            $filterXPath = "*"
            $filterParts = @()
            
            # Add time filter
            if ($StartTime -ne $null) {
                $filterParts += "TimeCreated[@SystemTime>='" + $StartTime.ToUniversalTime().ToString("o") + "']"
            }
            
            if ($EndTime -ne $null) {
                $filterParts += "TimeCreated[@SystemTime<='" + $EndTime.ToUniversalTime().ToString("o") + "']"
            }
            
            # Add event ID filter
            if ($EventID -and $EventID.Count -gt 0) {
                $eventIDFilter = "("
                for ($i = 0; $i -lt $EventID.Count; $i++) {
                    if ($i -gt 0) {
                        $eventIDFilter += " or "
                    }
                    $eventIDFilter += "EventID=" + $EventID[$i]
                }
                $eventIDFilter += ")"
                $filterParts += $eventIDFilter
            }
            
            # Add source filter
            if ($Source -and $Source.Count -gt 0) {
                $sourceFilter = "("
                for ($i = 0; $i -lt $Source.Count; $i++) {
                    if ($i -gt 0) {
                        $sourceFilter += " or "
                    }
                    $sourceFilter += "ProviderName='" + $Source[$i] + "'"
                }
                $sourceFilter += ")"
                $filterParts += $sourceFilter
            }
            
            # Add level filter
            if ($Level -and $Level.Count -gt 0) {
                $levelValues = @{
                    "Error" = 2
                    "Warning" = 3
                    "Information" = 4
                    "SuccessAudit" = 0
                    "FailureAudit" = 1
                }
                
                $levelFilter = "("
                for ($i = 0; $i -lt $Level.Count; $i++) {
                    if ($i -gt 0) {
                        $levelFilter += " or "
                    }
                    $levelFilter += "Level=" + $levelValues[$Level[$i]]
                }
                $levelFilter += ")"
                $filterParts += $levelFilter
            }
            
            # Combine filter parts
            if ($filterParts.Count -gt 0) {
                $filterXPath = "*[" + ($filterParts -join " and ") + "]"
            }
            
            # Get events
            $events = Get-WinEvent -LogName $LogName -FilterXPath $filterXPath -MaxEvents $MaxEvents -ErrorAction SilentlyContinue
            
            if ($events -eq $null) {
                return @{
                    ComputerName = $env:COMPUTERNAME
                    EventCount = 0
                    Events = @()
                    Status = "No events found"
                    ErrorMessage = $null
                }
            }
            
            # Process events
            $processedEvents = $events | ForEach-Object {
                $eventXML = [xml]$_.ToXml()
                $eventData = @{}
                
                # Add event data properties
                if ($eventXML.Event.EventData) {
                    foreach ($data in $eventXML.Event.EventData.Data) {
                        if ($data.Name) {
                            $eventData[$data.Name] = $data.'#text'
                        }
                    }
                }
                
                # Create event object
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    TimeCreated = $_.TimeCreated
                    LogName = $_.LogName
                    ProviderName = $_.ProviderName
                    Id = $_.Id
                    LevelDisplayName = $_.LevelDisplayName
                    Message = $_.Message
                    EventData = $eventData
                }
            }
            
            return @{
                ComputerName = $env:COMPUTERNAME
                EventCount = $processedEvents.Count
                Events = $processedEvents
                Status = "Success"
                ErrorMessage = $null
            }
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $LogName, $StartTime, $EndTime, $EventID, $Source, $Level, $MaxEvents
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $LogName, $StartTime, $EndTime, $EventID, $Source, $Level, $MaxEvents -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $LogName, $StartTime, $EndTime, $EventID, $Source, $Level, $MaxEvents
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Retrieved $($result.EventCount) events from $ComputerName." -Level "INFO"
        }
        elseif ($result.Status -eq "No events found") {
            Write-Log -Message "No events found on $ComputerName matching the specified criteria." -Level "WARNING"
        }
        else {
            Write-Log -Message "Failed to retrieve events from $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to retrieve event log data from $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            EventCount = 0
            Events = @()
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Export-EventsToCSV {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Events,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $Events | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Events exported to CSV successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export events to CSV: $_" -Level "ERROR"
        return $false
    }
}

function Export-EventsToJSON {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Events,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $Events | ConvertTo-Json -Depth 4 | Out-File -FilePath $Path -Encoding UTF8
        Write-Log -Message "Events exported to JSON successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export events to JSON: $_" -Level "ERROR"
        return $false
    }
}

function Export-EventsToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Events,
        
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
        $Events | Export-Excel -Path $Path -AutoSize -TableName "EventLogs" -WorksheetName "Event Log Analysis"
        Write-Log -Message "Events exported to Excel successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export events to Excel: $_" -Level "ERROR"
        return $false
    }
}

function Export-EventsToHTML {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Events,
        
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Event Log Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .error { background-color: #ffcccc; }
        .warning { background-color: #ffffcc; }
        .info { background-color: #e6f2ff; }
        .success { background-color: #ccffcc; }
        .failure { background-color: #ffcccc; }
    </style>
</head>
<body>
    <h1>Event Log Analysis Report</h1>
    <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p>Log: $LogName</p>
    <table>
        <tr>
"@
        
        $htmlColumns = ""
        $properties = $Events[0].PSObject.Properties.Name | Where-Object { $_ -ne "EventData" }
        foreach ($prop in $properties) {
            $htmlColumns += "            <th>$prop</th>`n"
        }
        
        $htmlRows = ""
        foreach ($event in $Events) {
            $rowClass = ""
            
            # Set row class based on level
            switch ($event.LevelDisplayName) {
                "Error" { $rowClass = "error" }
                "Warning" { $rowClass = "warning" }
                "Information" { $rowClass = "info" }
                "Success Audit" { $rowClass = "success" }
                "Failure Audit" { $rowClass = "failure" }
            }
            
            $htmlRows += "        <tr class=`"$rowClass`">`n"
            
            foreach ($prop in $properties) {
                $value = $event.$prop
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
        
        Write-Log -Message "Events exported to HTML successfully at '$Path'." -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Failed to export events to HTML: $_" -Level "ERROR"
        return $false
    }
}

# Main script execution
try {
    Write-Log -Message "Starting event log analysis process." -Level "INFO"
    
    # Create export directory if it doesn't exist
    $exportDir = Split-Path -Path $ExportPath -Parent
    if (-not (Test-Path -Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
        Write-Log -Message "Created export directory '$exportDir'." -Level "INFO"
    }
    
    $allEvents = @()
    
    foreach ($computer in $ComputerName) {
        Write-Log -Message "Processing computer: $computer" -Level "INFO"
        
        # Check if computer is reachable
        if ($computer -ne $env:COMPUTERNAME) {
            if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
                Write-Log -Message "Computer '$computer' is not reachable. Skipping..." -Level "WARNING"
                continue
            }
            
            # Check if PSRemoting is enabled
            if (-not (Test-PSRemoting -ComputerName $computer)) {
                Write-Log -Message "PowerShell Remoting is not enabled on '$computer'. Skipping..." -Level "WARNING"
                continue
            }
        }
        
        # Get event log data
        $result = Get-EventLogData -ComputerName $computer -Credential $Credential -LogName $LogName -StartTime $StartTime -EndTime $EndTime -EventID $EventID -Source $Source -Level $Level -MaxEvents $MaxEvents
        
        if ($result.Status -eq "Success" -and $result.EventCount -gt 0) {
            $allEvents += $result.Events
        }
    }
    
    # Export events in the specified format
    $exportSuccess = $false
    
    if ($allEvents.Count -gt 0) {
        switch ($ExportFormat) {
            "CSV" {
                $exportSuccess = Export-EventsToCSV -Events $allEvents -Path $ExportPath
            }
            "JSON" {
                $exportSuccess = Export-EventsToJSON -Events $allEvents -Path $ExportPath
            }
            "Excel" {
                $exportSuccess = Export-EventsToExcel -Events $allEvents -Path $ExportPath
            }
            "HTML" {
                $exportSuccess = Export-EventsToHTML -Events $allEvents -Path $ExportPath
            }
        }
        
        if ($exportSuccess) {
            Write-Log -Message "Event log analysis completed successfully. Exported $($allEvents.Count) events." -Level "INFO"
        }
        else {
            Write-Log -Message "Event log analysis completed with export errors." -Level "WARNING"
        }
    }
    else {
        Write-Log -Message "No events found matching the specified criteria." -Level "WARNING"
    }
    
    return @{
        EventCount = $allEvents.Count
        ExportPath = $ExportPath
        ExportFormat = $ExportFormat
        ExportSuccess = $exportSuccess
    }
}
catch {
    Write-Log -Message "An error occurred during event log analysis: $_" -Level "ERROR"
    exit 1
}
