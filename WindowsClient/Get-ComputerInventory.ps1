<#
.SYNOPSIS
    Collects comprehensive hardware and software inventory from Windows client computers.

.DESCRIPTION
    This script collects detailed hardware and software inventory from local or remote Windows client computers.
    It can include hardware information, installed software, Windows updates, and running services.
    The inventory can be exported in various formats. It provides detailed logging and error handling.

.PARAMETER ComputerName
    Name of the target computer(s).

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER IncludeHardware
    Whether to include hardware information.

.PARAMETER IncludeSoftware
    Whether to include installed software.

.PARAMETER IncludeUpdates
    Whether to include installed updates.

.PARAMETER IncludeServices
    Whether to include running services.

.PARAMETER ExportPath
    Path where the inventory will be saved.

.PARAMETER ExportFormat
    Format of the export file (CSV, JSON, Excel, HTML).

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Get-ComputerInventory.ps1 -ComputerName @("PC001", "PC002") -Credential (Get-Credential) -IncludeHardware $true -IncludeSoftware $true -IncludeUpdates $true -IncludeServices $true -ExportPath "C:\Reports\Inventory.xlsx" -ExportFormat "Excel"

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

    [Parameter(Mandatory = $false)]
    [bool]$IncludeHardware = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeSoftware = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeUpdates = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeServices = $true,

    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Excel", "HTML")]
    [string]$ExportFormat = "CSV",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ComputerInventory_$(Get-Date -Format 'yyyyMMdd').log"
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

function Get-HardwareInventory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Collecting hardware inventory from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
            $bios = Get-CimInstance -ClassName Win32_BIOS
            $os = Get-CimInstance -ClassName Win32_OperatingSystem
            $processor = Get-CimInstance -ClassName Win32_Processor
            $memory = Get-CimInstance -ClassName Win32_PhysicalMemory
            $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
            $network = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
            
            $totalMemory = ($memory | Measure-Object -Property Capacity -Sum).Sum / 1GB
            
            $hardwareInfo = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Manufacturer = $computerSystem.Manufacturer
                Model = $computerSystem.Model
                SerialNumber = $bios.SerialNumber
                BIOSVersion = $bios.SMBIOSBIOSVersion
                OSName = $os.Caption
                OSVersion = $os.Version
                OSBuild = $os.BuildNumber
                OSArchitecture = $os.OSArchitecture
                LastBootTime = $os.LastBootUpTime
                ProcessorName = $processor.Name
                ProcessorCores = $processor.NumberOfCores
                ProcessorLogicalProcessors = $processor.NumberOfLogicalProcessors
                TotalMemoryGB = [math]::Round($totalMemory, 2)
                DiskInfo = $disk | ForEach-Object {
                    [PSCustomObject]@{
                        Drive = $_.DeviceID
                        SizeGB = [math]::Round($_.Size / 1GB, 2)
                        FreeSpaceGB = [math]::Round($_.FreeSpace / 1GB, 2)
                        PercentFree = [math]::Round(($_.FreeSpace / $_.Size) * 100, 2)
                    }
                }
                NetworkInfo = $network | ForEach-Object {
                    [PSCustomObject]@{
                        AdapterName = $_.Description
                        MACAddress = $_.MACAddress
                        IPAddress = $_.IPAddress -join ', '
                        SubnetMask = $_.IPSubnet -join ', '
                        DefaultGateway = $_.DefaultIPGateway -join ', '
                        DNSServers = $_.DNSServerSearchOrder -join ', '
                    }
                }
            }
            
            return $hardwareInfo
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $hardwareInfo = Invoke-Command -ScriptBlock $scriptBlock
        }
        else {
            if ($Credential) {
                $hardwareInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
            }
            else {
                $hardwareInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        
        Write-Log -Message "Hardware inventory collected successfully from $ComputerName." -Level "INFO"
        return $hardwareInfo
    }
    catch {
        Write-Log -Message "Failed to collect hardware inventory from $ComputerName: $_" -Level "ERROR"
        return $null
    }
}

function Get-SoftwareInventory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Collecting software inventory from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            $uninstallKeys = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            
            $software = Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -and (-not [string]::IsNullOrEmpty($_.DisplayName)) } |
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, @{Name="ComputerName"; Expression={$env:COMPUTERNAME}}
            
            return $software
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $softwareInfo = Invoke-Command -ScriptBlock $scriptBlock
        }
        else {
            if ($Credential) {
                $softwareInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
            }
            else {
                $softwareInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        
        Write-Log -Message "Software inventory collected successfully from $ComputerName." -Level "INFO"
        return $softwareInfo
    }
    catch {
        Write-Log -Message "Failed to collect software inventory from $ComputerName: $_" -Level "ERROR"
        return $null
    }
}

function Get-UpdateInventory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Collecting Windows updates inventory from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $historyCount = $searcher.GetTotalHistoryCount()
            $updates = $searcher.QueryHistory(0, $historyCount) | 
                Select-Object Title, Description, Date, @{Name="Operation"; Expression={
                    switch($_.Operation) {
                        1 {"Installation"}
                        2 {"Uninstallation"}
                        3 {"Other"}
                        default {"Unknown"}
                    }
                }}, @{Name="Status"; Expression={
                    switch($_.ResultCode) {
                        0 {"Not Started"}
                        1 {"In Progress"}
                        2 {"Succeeded"}
                        3 {"Succeeded With Errors"}
                        4 {"Failed"}
                        5 {"Aborted"}
                        default {"Unknown"}
                    }
                }}, @{Name="ComputerName"; Expression={$env:COMPUTERNAME}}
            
            return $updates
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $updateInfo = Invoke-Command -ScriptBlock $scriptBlock
        }
        else {
            if ($Credential) {
                $updateInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
            }
            else {
                $updateInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        
        Write-Log -Message "Windows updates inventory collected successfully from $ComputerName." -Level "INFO"
        return $updateInfo
    }
    catch {
        Write-Log -Message "Failed to collect Windows updates inventory from $ComputerName: $_" -Level "ERROR"
        return $null
    }
}

function Get-ServiceInventory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Collecting services inventory from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            $services = Get-Service | 
                Select-Object Name, DisplayName, Status, StartType, @{Name="ComputerName"; Expression={$env:COMPUTERNAME}}
            
            return $services
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $serviceInfo = Invoke-Command -ScriptBlock $scriptBlock
        }
        else {
            if ($Credential) {
                $serviceInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
            }
            else {
                $serviceInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock
            }
        }
        
        Write-Log -Message "Services inventory collected successfully from $ComputerName." -Level "INFO"
        return $serviceInfo
    }
    catch {
        Write-Log -Message "Failed to collect services inventory from $ComputerName: $_" -Level "ERROR"
        return $null
    }
}

function Export-InventoryToCSV {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Inventory,
        
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )
    
    try {
        foreach ($key in $Inventory.Keys) {
            if ($Inventory[$key]) {
                $path = [System.IO.Path]::ChangeExtension($BasePath, $null) + "_$key.csv"
                $Inventory[$key] | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
                Write-Log -Message "$key inventory exported to CSV successfully at '$path'." -Level "INFO"
            }
        }
        return $true
    }
    catch {
        Write-Log -Message "Failed to export inventory to CSV: $_" -Level "ERROR"
        return $false
    }
}

function Export-InventoryToJSON {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Inventory,
        
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )
    
    try {
        foreach ($key in $Inventory.Keys) {
            if ($Inventory[$key]) {
                $path = [System.IO.Path]::ChangeExtension($BasePath, $null) + "_$key.json"
                $Inventory[$key] | ConvertTo-Json -Depth 4 | Out-File -FilePath $path -Encoding UTF8
                Write-Log -Message "$key inventory exported to JSON successfully at '$path'." -Level "INFO"
            }
        }
        return $true
    }
    catch {
        Write-Log -Message "Failed to export inventory to JSON: $_" -Level "ERROR"
        return $false
    }
}

function Export-InventoryToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Inventory,
        
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
        
        foreach ($key in $Inventory.Keys) {
            if ($Inventory[$key]) {
                $Inventory[$key] | Export-Excel -Path $Path -AutoSize -TableName "Inventory_$key" -WorksheetName $key
                Write-Log -Message "$key inventory exported to Excel successfully at '$Path'." -Level "INFO"
            }
        }
        return $true
    }
    catch {
        Write-Log -Message "Failed to export inventory to Excel: $_" -Level "ERROR"
        return $false
    }
}

function Export-InventoryToHTML {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Inventory,
        
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )
    
    try {
        foreach ($key in $Inventory.Keys) {
            if ($Inventory[$key]) {
                $path = [System.IO.Path]::ChangeExtension($BasePath, $null) + "_$key.html"
                
                $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Computer Inventory Report - $key</title>
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
    <h1>Computer Inventory Report - $key</h1>
    <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <table>
        <tr>
"@
                
                $htmlColumns = ""
                $properties = $Inventory[$key][0].PSObject.Properties.Name
                foreach ($prop in $properties) {
                    $htmlColumns += "            <th>$prop</th>`n"
                }
                
                $htmlRows = ""
                foreach ($row in $Inventory[$key]) {
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
                $html | Out-File -FilePath $path -Encoding UTF8
                
                Write-Log -Message "$key inventory exported to HTML successfully at '$path'." -Level "INFO"
            }
        }
        return $true
    }
    catch {
        Write-Log -Message "Failed to export inventory to HTML: $_" -Level "ERROR"
        return $false
    }
}

# Main script execution
try {
    Write-Log -Message "Starting computer inventory collection process." -Level "INFO"
    
    # Create export directory if it doesn't exist
    $exportDir = Split-Path -Path $ExportPath -Parent
    if (-not (Test-Path -Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
        Write-Log -Message "Created export directory '$exportDir'." -Level "INFO"
    }
    
    $allInventory = @{}
    
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
        
        $computerInventory = @{}
        
        # Collect hardware inventory
        if ($IncludeHardware) {
            $hardwareInfo = Get-HardwareInventory -ComputerName $computer -Credential $Credential
            if ($hardwareInfo) {
                if (-not $allInventory.ContainsKey("Hardware")) {
                    $allInventory["Hardware"] = @()
                }
                $allInventory["Hardware"] += $hardwareInfo
                $computerInventory["Hardware"] = $hardwareInfo
            }
        }
        
        # Collect software inventory
        if ($IncludeSoftware) {
            $softwareInfo = Get-SoftwareInventory -ComputerName $computer -Credential $Credential
            if ($softwareInfo) {
                if (-not $allInventory.ContainsKey("Software")) {
                    $allInventory["Software"] = @()
                }
                $allInventory["Software"] += $softwareInfo
                $computerInventory["Software"] = $softwareInfo
            }
        }
        
        # Collect updates inventory
        if ($IncludeUpdates) {
            $updateInfo = Get-UpdateInventory -ComputerName $computer -Credential $Credential
            if ($updateInfo) {
                if (-not $allInventory.ContainsKey("Updates")) {
                    $allInventory["Updates"] = @()
                }
                $allInventory["Updates"] += $updateInfo
                $computerInventory["Updates"] = $updateInfo
            }
        }
        
        # Collect services inventory
        if ($IncludeServices) {
            $serviceInfo = Get-ServiceInventory -ComputerName $computer -Credential $Credential
            if ($serviceInfo) {
                if (-not $allInventory.ContainsKey("Services")) {
                    $allInventory["Services"] = @()
                }
                $allInventory["Services"] += $serviceInfo
                $computerInventory["Services"] = $serviceInfo
            }
        }
        
        Write-Log -Message "Inventory collection completed for computer: $computer" -Level "INFO"
    }
    
    # Export inventory in the specified format
    $exportSuccess = $false
    
    switch ($ExportFormat) {
        "CSV" {
            $exportSuccess = Export-InventoryToCSV -Inventory $allInventory -BasePath $ExportPath
        }
        "JSON" {
            $exportSuccess = Export-InventoryToJSON -Inventory $allInventory -BasePath $ExportPath
        }
        "Excel" {
            $exportSuccess = Export-InventoryToExcel -Inventory $allInventory -Path $ExportPath
        }
        "HTML" {
            $exportSuccess = Export-InventoryToHTML -Inventory $allInventory -BasePath $ExportPath
        }
    }
    
    if ($exportSuccess) {
        Write-Log -Message "Computer inventory collection and export completed successfully." -Level "INFO"
    }
    else {
        Write-Log -Message "Computer inventory collection completed with export errors." -Level "WARNING"
    }
}
catch {
    Write-Log -Message "An error occurred during computer inventory collection: $_" -Level "ERROR"
    exit 1
}
