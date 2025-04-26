<#
.SYNOPSIS
    Configures various Windows settings on local or remote computers.

.DESCRIPTION
    This script configures various Windows settings on local or remote computers, including
    power settings, User Account Control (UAC), Windows features, and registry settings.
    It provides detailed logging and error handling.

.PARAMETER ComputerName
    Name of the target computer(s).

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER PowerSettings
    Power plan settings to configure.

.PARAMETER UAC
    User Account Control settings.

.PARAMETER WindowsFeatures
    Windows features to enable or disable.

.PARAMETER RegistrySettings
    Registry settings to configure.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    $powerSettings = @{
        PlanName = "High Performance"
        TurnOffDisplayMinutes = 15
        SleepMinutes = 30
    }
    $registrySettings = @(
        @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name="EnableLUA"; Value=1; Type="DWord"},
        @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name="ConsentPromptBehaviorAdmin"; Value=2; Type="DWord"}
    )
    .\Set-WindowsConfiguration.ps1 -ComputerName "PC001" -Credential (Get-Credential) -PowerSettings $powerSettings -UAC "Default" -WindowsFeatures @{Enable=@("Telnet-Client"); Disable=@("Internet-Explorer-Optional-amd64")} -RegistrySettings $registrySettings

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
    [hashtable]$PowerSettings,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Off", "Default", "AlwaysNotify", "NeverNotify")]
    [string]$UAC,

    [Parameter(Mandatory = $false)]
    [hashtable]$WindowsFeatures,

    [Parameter(Mandatory = $false)]
    [array]$RegistrySettings,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\WindowsConfiguration_$(Get-Date -Format 'yyyyMMdd').log"
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

function Set-PowerSettings {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$PowerSettings
    )
    
    try {
        Write-Log -Message "Configuring power settings on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [hashtable]$PowerSettings
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Get active power plan
                $activePlan = powercfg /GETACTIVESCHEME
                $activePlanGuid = ($activePlan | Select-String -Pattern "([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})").Matches.Groups[1].Value
                
                # Change power plan if specified
                if ($PowerSettings.ContainsKey("PlanName")) {
                    $planName = $PowerSettings.PlanName
                    
                    switch ($planName) {
                        "Balanced" {
                            $planGuid = "381b4222-f694-41f0-9685-ff5bb260df2e"
                        }
                        "High Performance" {
                            $planGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
                        }
                        "Power Saver" {
                            $planGuid = "a1841308-3541-4fab-bc81-f71556f20b4a"
                        }
                        default {
                            # Try to find the plan by name
                            $allPlans = powercfg /LIST
                            $matchingPlan = $allPlans | Select-String -Pattern "($planName).*?([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})"
                            
                            if ($matchingPlan) {
                                $planGuid = $matchingPlan.Matches.Groups[2].Value
                            }
                            else {
                                throw "Power plan '$planName' not found."
                            }
                        }
                    }
                    
                    # Set the active power plan
                    powercfg /SETACTIVE $planGuid
                    $result.Changes += "Set active power plan to '$planName'"
                    
                    # Update active plan GUID for subsequent settings
                    $activePlanGuid = $planGuid
                }
                
                # Configure display timeout
                if ($PowerSettings.ContainsKey("TurnOffDisplayMinutes")) {
                    $minutes = $PowerSettings.TurnOffDisplayMinutes
                    
                    # Set AC and DC timeouts
                    powercfg /CHANGE $activePlanGuid /monitor-timeout-ac $minutes
                    powercfg /CHANGE $activePlanGuid /monitor-timeout-dc $minutes
                    
                    $result.Changes += "Set monitor timeout to $minutes minutes"
                }
                
                # Configure sleep timeout
                if ($PowerSettings.ContainsKey("SleepMinutes")) {
                    $minutes = $PowerSettings.SleepMinutes
                    
                    # Set AC and DC timeouts
                    powercfg /CHANGE $activePlanGuid /standby-timeout-ac $minutes
                    powercfg /CHANGE $activePlanGuid /standby-timeout-dc $minutes
                    
                    $result.Changes += "Set sleep timeout to $minutes minutes"
                }
                
                # Configure hibernate timeout
                if ($PowerSettings.ContainsKey("HibernateMinutes")) {
                    $minutes = $PowerSettings.HibernateMinutes
                    
                    # Set AC and DC timeouts
                    powercfg /CHANGE $activePlanGuid /hibernate-timeout-ac $minutes
                    powercfg /CHANGE $activePlanGuid /hibernate-timeout-dc $minutes
                    
                    $result.Changes += "Set hibernate timeout to $minutes minutes"
                }
                
                # Enable/disable hibernation
                if ($PowerSettings.ContainsKey("EnableHibernation")) {
                    $enable = $PowerSettings.EnableHibernation
                    
                    if ($enable) {
                        powercfg /HIBERNATE ON
                        $result.Changes += "Enabled hibernation"
                    }
                    else {
                        powercfg /HIBERNATE OFF
                        $result.Changes += "Disabled hibernation"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $PowerSettings
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $PowerSettings -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $PowerSettings
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Power settings configured successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to configure power settings on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to configure power settings on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Set-UACSettings {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$UACLevel
    )
    
    try {
        Write-Log -Message "Configuring UAC settings on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$UACLevel
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                
                # Define UAC settings based on level
                switch ($UACLevel) {
                    "Off" {
                        Set-ItemProperty -Path $regPath -Name "EnableLUA" -Value 0 -Type DWord
                        $result.Changes += "Disabled User Account Control"
                    }
                    "Default" {
                        Set-ItemProperty -Path $regPath -Name "EnableLUA" -Value 1 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorAdmin" -Value 5 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorUser" -Value 3 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "EnableInstallerDetection" -Value 1 -Type DWord
                        $result.Changes += "Set UAC to default level"
                    }
                    "AlwaysNotify" {
                        Set-ItemProperty -Path $regPath -Name "EnableLUA" -Value 1 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorAdmin" -Value 2 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorUser" -Value 3 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "EnableInstallerDetection" -Value 1 -Type DWord
                        $result.Changes += "Set UAC to always notify level"
                    }
                    "NeverNotify" {
                        Set-ItemProperty -Path $regPath -Name "EnableLUA" -Value 1 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorAdmin" -Value 0 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "ConsentPromptBehaviorUser" -Value 0 -Type DWord
                        Set-ItemProperty -Path $regPath -Name "EnableInstallerDetection" -Value 0 -Type DWord
                        $result.Changes += "Set UAC to never notify level"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $UACLevel
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $UACLevel -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $UACLevel
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "UAC settings configured successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to configure UAC settings on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to configure UAC settings on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Set-WindowsFeaturesConfig {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$WindowsFeatures
    )
    
    try {
        Write-Log -Message "Configuring Windows features on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [hashtable]$WindowsFeatures
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Enable features
                if ($WindowsFeatures.ContainsKey("Enable") -and $WindowsFeatures.Enable.Count -gt 0) {
                    foreach ($feature in $WindowsFeatures.Enable) {
                        try {
                            $featureState = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
                            
                            if ($featureState -and $featureState.State -ne "Enabled") {
                                Enable-WindowsOptionalFeature -Online -FeatureName $feature -NoRestart
                                $result.Changes += "Enabled Windows feature: $feature"
                            }
                            elseif (-not $featureState) {
                                # Try as Windows capability
                                $capability = Get-WindowsCapability -Online | Where-Object { $_.Name -like "*$feature*" -and $_.State -ne "Installed" }
                                
                                if ($capability) {
                                    Add-WindowsCapability -Online -Name $capability.Name
                                    $result.Changes += "Installed Windows capability: $($capability.Name)"
                                }
                                else {
                                    # Try as Windows feature (server)
                                    Import-Module ServerManager -ErrorAction SilentlyContinue
                                    $serverFeature = Get-WindowsFeature -Name $feature -ErrorAction SilentlyContinue
                                    
                                    if ($serverFeature -and -not $serverFeature.Installed) {
                                        Install-WindowsFeature -Name $feature -IncludeManagementTools
                                        $result.Changes += "Installed Windows server feature: $feature"
                                    }
                                    else {
                                        $result.Changes += "Feature not found: $feature"
                                    }
                                }
                            }
                        }
                        catch {
                            $result.Changes += "Failed to enable feature $feature: $($_.Exception.Message)"
                        }
                    }
                }
                
                # Disable features
                if ($WindowsFeatures.ContainsKey("Disable") -and $WindowsFeatures.Disable.Count -gt 0) {
                    foreach ($feature in $WindowsFeatures.Disable) {
                        try {
                            $featureState = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
                            
                            if ($featureState -and $featureState.State -eq "Enabled") {
                                Disable-WindowsOptionalFeature -Online -FeatureName $feature -NoRestart
                                $result.Changes += "Disabled Windows feature: $feature"
                            }
                            elseif (-not $featureState) {
                                # Try as Windows capability
                                $capability = Get-WindowsCapability -Online | Where-Object { $_.Name -like "*$feature*" -and $_.State -eq "Installed" }
                                
                                if ($capability) {
                                    Remove-WindowsCapability -Online -Name $capability.Name
                                    $result.Changes += "Removed Windows capability: $($capability.Name)"
                                }
                                else {
                                    # Try as Windows feature (server)
                                    Import-Module ServerManager -ErrorAction SilentlyContinue
                                    $serverFeature = Get-WindowsFeature -Name $feature -ErrorAction SilentlyContinue
                                    
                                    if ($serverFeature -and $serverFeature.Installed) {
                                        Uninstall-WindowsFeature -Name $feature
                                        $result.Changes += "Uninstalled Windows server feature: $feature"
                                    }
                                    else {
                                        $result.Changes += "Feature not found or already disabled: $feature"
                                    }
                                }
                            }
                        }
                        catch {
                            $result.Changes += "Failed to disable feature $feature: $($_.Exception.Message)"
                        }
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $WindowsFeatures
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $WindowsFeatures -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $WindowsFeatures
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows features configured successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to configure Windows features on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to configure Windows features on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Set-RegistryConfig {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [array]$RegistrySettings
    )
    
    try {
        Write-Log -Message "Configuring registry settings on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [array]$RegistrySettings
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                foreach ($setting in $RegistrySettings) {
                    try {
                        $path = $setting.Path
                        $name = $setting.Name
                        $value = $setting.Value
                        $type = $setting.Type
                        
                        # Create registry path if it doesn't exist
                        if (-not (Test-Path -Path $path)) {
                            New-Item -Path $path -Force | Out-Null
                            $result.Changes += "Created registry path: $path"
                        }
                        
                        # Set registry value
                        Set-ItemProperty -Path $path -Name $name -Value $value -Type $type
                        $result.Changes += "Set registry value: $path\$name = $value ($type)"
                    }
                    catch {
                        $result.Changes += "Failed to set registry value $($setting.Path)\$($setting.Name): $($_.Exception.Message)"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $RegistrySettings
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RegistrySettings -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RegistrySettings
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Registry settings configured successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to configure registry settings on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to configure registry settings on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Windows configuration process." -Level "INFO"
    
    $results = @()
    
    foreach ($computer in $ComputerName) {
        Write-Log -Message "Processing computer: $computer" -Level "INFO"
        
        # Check if computer is reachable
        if ($computer -ne $env:COMPUTERNAME) {
            if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
                Write-Log -Message "Computer '$computer' is not reachable. Skipping..." -Level "WARNING"
                $results += @{
                    ComputerName = $computer
                    Status = "Unreachable"
                    ErrorMessage = "Computer is not reachable"
                }
                continue
            }
            
            # Check if PSRemoting is enabled
            if (-not (Test-PSRemoting -ComputerName $computer)) {
                Write-Log -Message "PowerShell Remoting is not enabled on '$computer'. Skipping..." -Level "WARNING"
                $results += @{
                    ComputerName = $computer
                    Status = "PSRemoting Disabled"
                    ErrorMessage = "PowerShell Remoting is not enabled"
                }
                continue
            }
        }
        
        $computerResult = @{
            ComputerName = $computer
            Status = "Success"
            PowerSettings = $null
            UAC = $null
            WindowsFeatures = $null
            RegistrySettings = $null
        }
        
        # Configure power settings
        if ($PowerSettings) {
            $powerResult = Set-PowerSettings -ComputerName $computer -Credential $Credential -PowerSettings $PowerSettings
            $computerResult.PowerSettings = $powerResult
            
            if ($powerResult.Status -ne "Success") {
                $computerResult.Status = "Partial Success"
            }
        }
        
        # Configure UAC settings
        if ($UAC) {
            $uacResult = Set-UACSettings -ComputerName $computer -Credential $Credential -UACLevel $UAC
            $computerResult.UAC = $uacResult
            
            if ($uacResult.Status -ne "Success") {
                $computerResult.Status = "Partial Success"
            }
        }
        
        # Configure Windows features
        if ($WindowsFeatures) {
            $featuresResult = Set-WindowsFeaturesConfig -ComputerName $computer -Credential $Credential -WindowsFeatures $WindowsFeatures
            $computerResult.WindowsFeatures = $featuresResult
            
            if ($featuresResult.Status -ne "Success") {
                $computerResult.Status = "Partial Success"
            }
        }
        
        # Configure registry settings
        if ($RegistrySettings) {
            $registryResult = Set-RegistryConfig -ComputerName $computer -Credential $Credential -RegistrySettings $RegistrySettings
            $computerResult.RegistrySettings = $registryResult
            
            if ($registryResult.Status -ne "Success") {
                $computerResult.Status = "Partial Success"
            }
        }
        
        $results += $computerResult
    }
    
    # Output summary
    Write-Log -Message "Windows configuration process completed." -Level "INFO"
    Write-Log -Message "Summary:" -Level "INFO"
    
    foreach ($result in $results) {
        $status = "Computer: $($result.ComputerName), Status: $($result.Status)"
        
        if ($result.Status -eq "Success" -or $result.Status -eq "Partial Success") {
            Write-Log -Message $status -Level "INFO"
        }
        else {
            Write-Log -Message $status -Level "WARNING"
        }
    }
    
    return $results
}
catch {
    Write-Log -Message "An error occurred during Windows configuration: $_" -Level "ERROR"
    exit 1
}
