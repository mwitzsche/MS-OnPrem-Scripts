<#
.SYNOPSIS
    Manages Group Policy Objects (GPOs) in Active Directory.

.DESCRIPTION
    This script provides comprehensive management of Group Policy Objects (GPOs) in Active Directory,
    including creating, modifying, backing up, restoring, and reporting on GPOs.
    It provides detailed logging and error handling.

.PARAMETER Action
    Action to perform (Create, Modify, Backup, Restore, Report).

.PARAMETER GPOName
    Name of the GPO to manage.

.PARAMETER DomainName
    Name of the domain where the GPO exists or will be created.

.PARAMETER BackupPath
    Path where GPO backups will be stored or restored from.

.PARAMETER BackupId
    ID of the GPO backup to restore.

.PARAMETER LinkPath
    Path where the GPO will be linked.

.PARAMETER RegistrySettings
    Array of registry settings to configure in the GPO.

.PARAMETER SecuritySettings
    Array of security settings to configure in the GPO.

.PARAMETER ReportPath
    Path where the GPO report will be saved.

.PARAMETER ReportFormat
    Format of the GPO report (HTML, XML).

.PARAMETER Credential
    Credentials to use for domain operations.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Manage-GroupPolicy.ps1 -Action Create -GPOName "Security Settings" -DomainName "contoso.com" -LinkPath "OU=Servers,DC=contoso,DC=com" -Credential (Get-Credential)

.EXAMPLE
    .\Manage-GroupPolicy.ps1 -Action Backup -GPOName "Security Settings" -DomainName "contoso.com" -BackupPath "C:\GPOBackups" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Create", "Modify", "Backup", "Restore", "Report")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$GPOName,

    [Parameter(Mandatory = $false)]
    [string]$DomainName = $env:USERDNSDOMAIN,

    [Parameter(Mandatory = $false)]
    [string]$BackupPath,

    [Parameter(Mandatory = $false)]
    [string]$BackupId,

    [Parameter(Mandatory = $false)]
    [string]$LinkPath,

    [Parameter(Mandatory = $false)]
    [array]$RegistrySettings,

    [Parameter(Mandatory = $false)]
    [array]$SecuritySettings,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("HTML", "XML")]
    [string]$ReportFormat = "HTML",

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\GPOManagement_$(Get-Date -Format 'yyyyMMdd').log"
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

function Test-GPOModules {
    try {
        # Check if required modules are available
        $requiredModules = @(
            "GroupPolicy"
        )
        
        $missingModules = @()
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -Name $module -ListAvailable)) {
                $missingModules += $module
            }
        }
        
        if ($missingModules.Count -gt 0) {
            Write-Log -Message "Missing required modules: $($missingModules -join ', ')" -Level "ERROR"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Error checking required modules: $_" -Level "ERROR"
        return $false
    }
}

function Create-GPO {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $false)]
        [string]$LinkPath,
        
        [Parameter(Mandatory = $false)]
        [array]$RegistrySettings,
        
        [Parameter(Mandatory = $false)]
        [array]$SecuritySettings,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Creating GPO '$GPOName' in domain '$DomainName'..." -Level "INFO"
        
        # Create GPO
        $gpoParams = @{
            Name = $GPOName
            Domain = $DomainName
        }
        
        if ($Credential) {
            $gpoParams.Add("Credential", $Credential)
        }
        
        $gpo = New-GPO @gpoParams
        
        if (-not $gpo) {
            throw "Failed to create GPO."
        }
        
        Write-Log -Message "GPO '$GPOName' created successfully." -Level "INFO"
        
        # Configure registry settings
        if ($RegistrySettings -and $RegistrySettings.Count -gt 0) {
            Write-Log -Message "Configuring registry settings..." -Level "INFO"
            
            foreach ($setting in $RegistrySettings) {
                $keyPath = $setting.KeyPath
                $valueName = $setting.ValueName
                $valueData = $setting.ValueData
                $valueType = $setting.ValueType
                
                $setParams = @{
                    Name = $GPOName
                    Domain = $DomainName
                    Key = $keyPath
                    ValueName = $valueName
                    Value = $valueData
                    Type = $valueType
                }
                
                if ($Credential) {
                    $setParams.Add("Credential", $Credential)
                }
                
                Set-GPRegistryValue @setParams | Out-Null
                
                Write-Log -Message "Registry setting configured: $keyPath\$valueName" -Level "INFO"
            }
        }
        
        # Configure security settings
        if ($SecuritySettings -and $SecuritySettings.Count -gt 0) {
            Write-Log -Message "Configuring security settings..." -Level "INFO"
            
            # This is a simplified example. In a real-world scenario, you would use more specific cmdlets
            # based on the type of security setting (e.g., Set-GPPermission, Set-GPRegistryValue, etc.)
            foreach ($setting in $SecuritySettings) {
                $settingType = $setting.Type
                
                switch ($settingType) {
                    "Permission" {
                        $permParams = @{
                            Name = $GPOName
                            Domain = $DomainName
                            PermissionLevel = $setting.PermissionLevel
                            TargetName = $setting.TargetName
                            TargetType = $setting.TargetType
                        }
                        
                        if ($Credential) {
                            $permParams.Add("Credential", $Credential)
                        }
                        
                        Set-GPPermission @permParams | Out-Null
                        
                        Write-Log -Message "Permission setting configured: $($setting.TargetName) - $($setting.PermissionLevel)" -Level "INFO"
                    }
                    # Add more security setting types as needed
                }
            }
        }
        
        # Link GPO
        if ($LinkPath) {
            Write-Log -Message "Linking GPO to '$LinkPath'..." -Level "INFO"
            
            $linkParams = @{
                Name = $GPOName
                Domain = $DomainName
                Target = $LinkPath
            }
            
            if ($Credential) {
                $linkParams.Add("Credential", $Credential)
            }
            
            New-GPLink @linkParams | Out-Null
            
            Write-Log -Message "GPO linked successfully." -Level "INFO"
        }
        
        return @{
            Status = "Success"
            GPO = $gpo
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create GPO: $_" -Level "ERROR"
        return @{
            Status = "Error"
            GPO = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Modify-GPO {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $false)]
        [string]$LinkPath,
        
        [Parameter(Mandatory = $false)]
        [array]$RegistrySettings,
        
        [Parameter(Mandatory = $false)]
        [array]$SecuritySettings,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Modifying GPO '$GPOName' in domain '$DomainName'..." -Level "INFO"
        
        # Get GPO
        $gpoParams = @{
            Name = $GPOName
            Domain = $DomainName
        }
        
        if ($Credential) {
            $gpoParams.Add("Credential", $Credential)
        }
        
        $gpo = Get-GPO @gpoParams
        
        if (-not $gpo) {
            throw "GPO not found."
        }
        
        # Configure registry settings
        if ($RegistrySettings -and $RegistrySettings.Count -gt 0) {
            Write-Log -Message "Configuring registry settings..." -Level "INFO"
            
            foreach ($setting in $RegistrySettings) {
                $keyPath = $setting.KeyPath
                $valueName = $setting.ValueName
                $valueData = $setting.ValueData
                $valueType = $setting.ValueType
                
                $setParams = @{
                    Name = $GPOName
                    Domain = $DomainName
                    Key = $keyPath
                    ValueName = $valueName
                    Value = $valueData
                    Type = $valueType
                }
                
                if ($Credential) {
                    $setParams.Add("Credential", $Credential)
                }
                
                Set-GPRegistryValue @setParams | Out-Null
                
                Write-Log -Message "Registry setting configured: $keyPath\$valueName" -Level "INFO"
            }
        }
        
        # Configure security settings
        if ($SecuritySettings -and $SecuritySettings.Count -gt 0) {
            Write-Log -Message "Configuring security settings..." -Level "INFO"
            
            # This is a simplified example. In a real-world scenario, you would use more specific cmdlets
            # based on the type of security setting (e.g., Set-GPPermission, Set-GPRegistryValue, etc.)
            foreach ($setting in $SecuritySettings) {
                $settingType = $setting.Type
                
                switch ($settingType) {
                    "Permission" {
                        $permParams = @{
                            Name = $GPOName
                            Domain = $DomainName
                            PermissionLevel = $setting.PermissionLevel
                            TargetName = $setting.TargetName
                            TargetType = $setting.TargetType
                        }
                        
                        if ($Credential) {
                            $permParams.Add("Credential", $Credential)
                        }
                        
                        Set-GPPermission @permParams | Out-Null
                        
                        Write-Log -Message "Permission setting configured: $($setting.TargetName) - $($setting.PermissionLevel)" -Level "INFO"
                    }
                    # Add more security setting types as needed
                }
            }
        }
        
        # Link GPO
        if ($LinkPath) {
            Write-Log -Message "Linking GPO to '$LinkPath'..." -Level "INFO"
            
            # Check if GPO is already linked
            $linkParams = @{
                Name = $GPOName
                Domain = $DomainName
                Target = $LinkPath
            }
            
            if ($Credential) {
                $linkParams.Add("Credential", $Credential)
            }
            
            $existingLink = Get-GPLink @linkParams -ErrorAction SilentlyContinue
            
            if (-not $existingLink) {
                New-GPLink @linkParams | Out-Null
                Write-Log -Message "GPO linked successfully." -Level "INFO"
            }
            else {
                Write-Log -Message "GPO is already linked to '$LinkPath'." -Level "INFO"
            }
        }
        
        Write-Log -Message "GPO '$GPOName' modified successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            GPO = $gpo
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to modify GPO: $_" -Level "ERROR"
        return @{
            Status = "Error"
            GPO = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Backup-GPO {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Backing up GPO '$GPOName' from domain '$DomainName' to '$BackupPath'..." -Level "INFO"
        
        # Create backup directory if it doesn't exist
        if (-not (Test-Path -Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created backup directory '$BackupPath'." -Level "INFO"
        }
        
        # Backup GPO
        $backupParams = @{
            Name = $GPOName
            Domain = $DomainName
            Path = $BackupPath
        }
        
        if ($Credential) {
            $backupParams.Add("Credential", $Credential)
        }
        
        $backup = Backup-GPO @backupParams
        
        if (-not $backup) {
            throw "Failed to backup GPO."
        }
        
        Write-Log -Message "GPO '$GPOName' backed up successfully. Backup ID: $($backup.Id)" -Level "INFO"
        
        return @{
            Status = "Success"
            Backup = $backup
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to backup GPO: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Backup = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Restore-GPOFromBackup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [Parameter(Mandatory = $false)]
        [string]$BackupId,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Restoring GPO '$GPOName' to domain '$DomainName' from '$BackupPath'..." -Level "INFO"
        
        # Check if backup path exists
        if (-not (Test-Path -Path $BackupPath)) {
            throw "Backup path '$BackupPath' does not exist."
        }
        
        # Get backup
        $backupParams = @{
            Path = $BackupPath
        }
        
        if ($Credential) {
            $backupParams.Add("Credential", $Credential)
        }
        
        $backups = Get-GPOBackup @backupParams
        
        if (-not $backups) {
            throw "No backups found in '$BackupPath'."
        }
        
        $backup = $null
        
        if ($BackupId) {
            $backup = $backups | Where-Object { $_.Id -eq $BackupId }
            
            if (-not $backup) {
                throw "Backup with ID '$BackupId' not found."
            }
        }
        else {
            # Get the latest backup for the specified GPO
            $backup = $backups | Where-Object { $_.DisplayName -eq $GPOName } | Sort-Object -Property Timestamp -Descending | Select-Object -First 1
            
            if (-not $backup) {
                throw "No backup found for GPO '$GPOName'."
            }
        }
        
        # Restore GPO
        $restoreParams = @{
            BackupId = $backup.Id
            Path = $BackupPath
            Domain = $DomainName
        }
        
        if ($Credential) {
            $restoreParams.Add("Credential", $Credential)
        }
        
        $restored = Restore-GPO @restoreParams
        
        if (-not $restored) {
            throw "Failed to restore GPO."
        }
        
        Write-Log -Message "GPO '$GPOName' restored successfully from backup ID: $($backup.Id)" -Level "INFO"
        
        return @{
            Status = "Success"
            Restored = $restored
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to restore GPO: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Restored = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Generate-GPOReport {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GPOName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportFormat,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Generating report for GPO '$GPOName' in domain '$DomainName'..." -Level "INFO"
        
        # Create report directory if it doesn't exist
        $reportDir = Split-Path -Path $ReportPath -Parent
        if (-not (Test-Path -Path $reportDir)) {
            New-Item -Path $reportDir -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created report directory '$reportDir'." -Level "INFO"
        }
        
        # Generate report
        $reportParams = @{
            Name = $GPOName
            Domain = $DomainName
            Path = $ReportPath
            ReportType = $ReportFormat
        }
        
        if ($Credential) {
            $reportParams.Add("Credential", $Credential)
        }
        
        Get-GPOReport @reportParams
        
        if (-not (Test-Path -Path $ReportPath)) {
            throw "Failed to generate GPO report."
        }
        
        Write-Log -Message "GPO report generated successfully at '$ReportPath'." -Level "INFO"
        
        return @{
            Status = "Success"
            ReportPath = $ReportPath
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to generate GPO report: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ReportPath = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting GPO management process." -Level "INFO"
    
    # Check if required modules are available
    if (-not (Test-GPOModules)) {
        Write-Log -Message "Required modules are missing. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Perform the requested action
    switch ($Action) {
        "Create" {
            # Validate required parameters
            if (-not $DomainName) {
                Write-Log -Message "DomainName parameter is required for Create action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Create-GPO -GPOName $GPOName -DomainName $DomainName -LinkPath $LinkPath -RegistrySettings $RegistrySettings -SecuritySettings $SecuritySettings -Credential $Credential
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to create GPO. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Modify" {
            # Validate required parameters
            if (-not $DomainName) {
                Write-Log -Message "DomainName parameter is required for Modify action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Modify-GPO -GPOName $GPOName -DomainName $DomainName -LinkPath $LinkPath -RegistrySettings $RegistrySettings -SecuritySettings $SecuritySettings -Credential $Credential
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to modify GPO. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Backup" {
            # Validate required parameters
            if (-not $DomainName) {
                Write-Log -Message "DomainName parameter is required for Backup action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $BackupPath) {
                Write-Log -Message "BackupPath parameter is required for Backup action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Backup-GPO -GPOName $GPOName -DomainName $DomainName -BackupPath $BackupPath -Credential $Credential
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to backup GPO. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Restore" {
            # Validate required parameters
            if (-not $DomainName) {
                Write-Log -Message "DomainName parameter is required for Restore action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $BackupPath) {
                Write-Log -Message "BackupPath parameter is required for Restore action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Restore-GPOFromBackup -GPOName $GPOName -DomainName $DomainName -BackupPath $BackupPath -BackupId $BackupId -Credential $Credential
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to restore GPO. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Report" {
            # Validate required parameters
            if (-not $DomainName) {
                Write-Log -Message "DomainName parameter is required for Report action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $ReportPath) {
                Write-Log -Message "ReportPath parameter is required for Report action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Generate-GPOReport -GPOName $GPOName -DomainName $DomainName -ReportPath $ReportPath -ReportFormat $ReportFormat -Credential $Credential
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to generate GPO report. Exiting..." -Level "ERROR"
                exit 1
            }
        }
    }
    
    Write-Log -Message "GPO management process completed successfully." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during GPO management process: $_" -Level "ERROR"
    exit 1
}
