<#
.SYNOPSIS
    Creates and configures a Desired State Configuration (DSC) configuration for Windows Server.

.DESCRIPTION
    This script creates and applies a Desired State Configuration (DSC) for Windows Server,
    allowing for consistent and automated server configuration. It supports configuring
    Windows features, registry settings, files, services, and more.

.PARAMETER ConfigurationName
    Name of the DSC configuration.

.PARAMETER OutputPath
    Path where the DSC configuration MOF files will be saved.

.PARAMETER NodeName
    Name of the target node(s).

.PARAMETER WindowsFeatures
    Hashtable of Windows features to ensure are present or absent.

.PARAMETER RegistrySettings
    Array of registry settings to configure.

.PARAMETER FilesToCopy
    Array of files to copy to the target node.

.PARAMETER Services
    Hashtable of services to configure.

.PARAMETER LocalGroups
    Hashtable of local groups to configure.

.PARAMETER ScheduledTasks
    Hashtable of scheduled tasks to configure.

.PARAMETER ApplyConfiguration
    Whether to apply the configuration after creating it.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    $windowsFeatures = @{
        Present = @("Web-Server", "Web-Mgmt-Tools", "NET-Framework-45-Features")
        Absent = @("Telnet-Client")
    }
    $registrySettings = @(
        @{Path="HKLM:\SOFTWARE\MyApp"; Name="Version"; Value="1.0.0"; Type="String"},
        @{Path="HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"; Name="fDenyTSConnections"; Value=0; Type="DWord"}
    )
    $filesToCopy = @(
        @{Source="C:\Source\config.xml"; Destination="C:\App\config.xml"}
    )
    $services = @{
        Running = @(
            @{Name="BITS"; StartupType="Automatic"}
        )
        Stopped = @(
            @{Name="Telnet"; StartupType="Disabled"}
        )
    }
    .\New-ServerDSCConfiguration.ps1 -ConfigurationName "WebServerConfig" -OutputPath "C:\DSC" -NodeName "WebServer01" -WindowsFeatures $windowsFeatures -RegistrySettings $registrySettings -FilesToCopy $filesToCopy -Services $services -ApplyConfiguration $true

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ConfigurationName,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string[]]$NodeName = @("localhost"),

    [Parameter(Mandatory = $false)]
    [hashtable]$WindowsFeatures,

    [Parameter(Mandatory = $false)]
    [array]$RegistrySettings,

    [Parameter(Mandatory = $false)]
    [array]$FilesToCopy,

    [Parameter(Mandatory = $false)]
    [hashtable]$Services,

    [Parameter(Mandatory = $false)]
    [hashtable]$LocalGroups,

    [Parameter(Mandatory = $false)]
    [hashtable]$ScheduledTasks,

    [Parameter(Mandatory = $false)]
    [bool]$ApplyConfiguration = $false,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ServerDSC_$(Get-Date -Format 'yyyyMMdd').log"
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

function Test-DSCModules {
    try {
        # Check if DSC modules are available
        $requiredModules = @(
            "PSDesiredStateConfiguration"
        )
        
        $missingModules = @()
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -Name $module -ListAvailable)) {
                $missingModules += $module
            }
        }
        
        if ($missingModules.Count -gt 0) {
            Write-Log -Message "Missing required DSC modules: $($missingModules -join ', ')" -Level "ERROR"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Error checking DSC modules: $_" -Level "ERROR"
        return $false
    }
}

function Create-DSCConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigurationName,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [string[]]$NodeName,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$WindowsFeatures,
        
        [Parameter(Mandatory = $false)]
        [array]$RegistrySettings,
        
        [Parameter(Mandatory = $false)]
        [array]$FilesToCopy,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Services,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$LocalGroups,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$ScheduledTasks
    )
    
    try {
        Write-Log -Message "Creating DSC configuration '$ConfigurationName'..." -Level "INFO"
        
        # Create output directory if it doesn't exist
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created output directory '$OutputPath'." -Level "INFO"
        }
        
        # Create configuration script
        $configScript = @"
Configuration $ConfigurationName
{
    Import-DscResource -ModuleName PSDesiredStateConfiguration
    
    Node $($NodeName -join ",")
    {
"@
        
        # Add Windows features
        if ($WindowsFeatures) {
            if ($WindowsFeatures.ContainsKey("Present") -and $WindowsFeatures.Present.Count -gt 0) {
                foreach ($feature in $WindowsFeatures.Present) {
                    $featureId = [Guid]::NewGuid().ToString().Substring(0, 8)
                    $configScript += @"
        
        # Ensure Windows feature $feature is present
        WindowsFeature $($feature.Replace("-", ""))_$featureId
        {
            Name = "$feature"
            Ensure = "Present"
        }
"@
                }
            }
            
            if ($WindowsFeatures.ContainsKey("Absent") -and $WindowsFeatures.Absent.Count -gt 0) {
                foreach ($feature in $WindowsFeatures.Absent) {
                    $featureId = [Guid]::NewGuid().ToString().Substring(0, 8)
                    $configScript += @"
        
        # Ensure Windows feature $feature is absent
        WindowsFeature $($feature.Replace("-", ""))_$featureId
        {
            Name = "$feature"
            Ensure = "Absent"
        }
"@
                }
            }
        }
        
        # Add registry settings
        if ($RegistrySettings -and $RegistrySettings.Count -gt 0) {
            foreach ($setting in $RegistrySettings) {
                $settingId = [Guid]::NewGuid().ToString().Substring(0, 8)
                $path = $setting.Path
                $name = $setting.Name
                $value = $setting.Value
                $type = $setting.Type
                
                $configScript += @"
        
        # Configure registry setting $path\$name
        Registry Registry_$settingId
        {
            Key = "$path"
            ValueName = "$name"
            ValueData = "$value"
            ValueType = "$type"
            Ensure = "Present"
            Force = `$true
        }
"@
            }
        }
        
        # Add files to copy
        if ($FilesToCopy -and $FilesToCopy.Count -gt 0) {
            foreach ($file in $FilesToCopy) {
                $fileId = [Guid]::NewGuid().ToString().Substring(0, 8)
                $source = $file.Source
                $destination = $file.Destination
                
                $configScript += @"
        
        # Copy file from $source to $destination
        File File_$fileId
        {
            SourcePath = "$source"
            DestinationPath = "$destination"
            Ensure = "Present"
            Type = "File"
            Force = `$true
        }
"@
            }
        }
        
        # Add services
        if ($Services) {
            if ($Services.ContainsKey("Running") -and $Services.Running.Count -gt 0) {
                foreach ($service in $Services.Running) {
                    $serviceName = $service.Name
                    $startupType = $service.StartupType
                    
                    $configScript += @"
        
        # Ensure service $serviceName is running
        Service Service_$serviceName
        {
            Name = "$serviceName"
            State = "Running"
            StartupType = "$startupType"
        }
"@
                }
            }
            
            if ($Services.ContainsKey("Stopped") -and $Services.Stopped.Count -gt 0) {
                foreach ($service in $Services.Stopped) {
                    $serviceName = $service.Name
                    $startupType = $service.StartupType
                    
                    $configScript += @"
        
        # Ensure service $serviceName is stopped
        Service Service_$serviceName
        {
            Name = "$serviceName"
            State = "Stopped"
            StartupType = "$startupType"
        }
"@
                }
            }
        }
        
        # Add local groups
        if ($LocalGroups) {
            if ($LocalGroups.ContainsKey("Present") -and $LocalGroups.Present.Count -gt 0) {
                foreach ($group in $LocalGroups.Present) {
                    $groupName = $group.Name
                    $members = $group.Members -join "','"
                    
                    $configScript += @"
        
        # Ensure local group $groupName exists with specified members
        Group Group_$groupName
        {
            GroupName = "$groupName"
            Ensure = "Present"
            MembersToInclude = @('$members')
        }
"@
                }
            }
            
            if ($LocalGroups.ContainsKey("Absent") -and $LocalGroups.Absent.Count -gt 0) {
                foreach ($group in $LocalGroups.Absent) {
                    $groupName = $group.Name
                    
                    $configScript += @"
        
        # Ensure local group $groupName does not exist
        Group Group_$groupName
        {
            GroupName = "$groupName"
            Ensure = "Absent"
        }
"@
                }
            }
        }
        
        # Add scheduled tasks
        if ($ScheduledTasks) {
            # Note: Scheduled tasks require additional modules like ComputerManagementDsc
            # This is a simplified example
            if ($ScheduledTasks.ContainsKey("Present") -and $ScheduledTasks.Present.Count -gt 0) {
                $configScript += @"
        
        # Note: To properly configure scheduled tasks, you need to import the ComputerManagementDsc module
        # Import-DscResource -ModuleName ComputerManagementDsc
"@
            }
        }
        
        # Close configuration
        $configScript += @"
    }
}
"@
        
        # Save configuration script
        $scriptPath = Join-Path -Path $OutputPath -ChildPath "$ConfigurationName.ps1"
        $configScript | Out-File -FilePath $scriptPath -Encoding utf8
        
        Write-Log -Message "DSC configuration script saved to '$scriptPath'." -Level "INFO"
        
        # Dot source the configuration script
        . $scriptPath
        
        # Create the MOF files
        & $ConfigurationName -OutputPath $OutputPath
        
        Write-Log -Message "DSC configuration MOF files created in '$OutputPath'." -Level "INFO"
        
        return @{
            Status = "Success"
            ScriptPath = $scriptPath
            MOFPath = $OutputPath
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create DSC configuration: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ScriptPath = $null
            MOFPath = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Apply-DSCConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        Write-Log -Message "Applying DSC configuration..." -Level "INFO"
        
        # Apply configuration
        $result = Start-DscConfiguration -Path $OutputPath -Wait -Verbose -Force
        
        Write-Log -Message "DSC configuration applied successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            Result = $result
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to apply DSC configuration: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Result = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting DSC configuration process." -Level "INFO"
    
    # Check if DSC modules are available
    if (-not (Test-DSCModules)) {
        Write-Log -Message "Required DSC modules are missing. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Create DSC configuration
    $createResult = Create-DSCConfiguration -ConfigurationName $ConfigurationName -OutputPath $OutputPath -NodeName $NodeName -WindowsFeatures $WindowsFeatures -RegistrySettings $RegistrySettings -FilesToCopy $FilesToCopy -Services $Services -LocalGroups $LocalGroups -ScheduledTasks $ScheduledTasks
    
    if ($createResult.Status -eq "Success") {
        Write-Log -Message "DSC configuration created successfully." -Level "INFO"
        
        # Apply configuration if requested
        if ($ApplyConfiguration) {
            $applyResult = Apply-DSCConfiguration -OutputPath $OutputPath
            
            if ($applyResult.Status -eq "Success") {
                Write-Log -Message "DSC configuration applied successfully." -Level "INFO"
            }
            else {
                Write-Log -Message "Failed to apply DSC configuration: $($applyResult.ErrorMessage)" -Level "ERROR"
                exit 1
            }
        }
        else {
            Write-Log -Message "DSC configuration created but not applied. To apply, use Start-DscConfiguration -Path '$OutputPath' -Wait -Verbose -Force" -Level "INFO"
        }
    }
    else {
        Write-Log -Message "Failed to create DSC configuration: $($createResult.ErrorMessage)" -Level "ERROR"
        exit 1
    }
    
    Write-Log -Message "DSC configuration process completed." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during DSC configuration process: $_" -Level "ERROR"
    exit 1
}
