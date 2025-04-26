<#
.SYNOPSIS
    Configures and manages Windows Firewall settings on local or remote computers.

.DESCRIPTION
    This script configures and manages Windows Firewall settings on local or remote computers,
    including enabling/disabling firewall profiles, creating firewall rules, and exporting/importing
    firewall configurations. It provides detailed logging and error handling.

.PARAMETER ComputerName
    Name of the target computer(s).

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER Action
    Action to perform (Enable, Disable, CreateRule, RemoveRule, ExportConfig, ImportConfig).

.PARAMETER ProfileType
    Firewall profile type(s) to configure (Domain, Private, Public, All).

.PARAMETER RuleName
    Name of the firewall rule (for CreateRule and RemoveRule actions).

.PARAMETER Direction
    Direction of the firewall rule (Inbound, Outbound).

.PARAMETER Protocol
    Protocol for the firewall rule (TCP, UDP, Any).

.PARAMETER LocalPort
    Local port(s) for the firewall rule.

.PARAMETER RemotePort
    Remote port(s) for the firewall rule.

.PARAMETER LocalAddress
    Local IP address(es) for the firewall rule.

.PARAMETER RemoteAddress
    Remote IP address(es) for the firewall rule.

.PARAMETER Program
    Program path for the firewall rule.

.PARAMETER Action
    Action for the firewall rule (Allow, Block).

.PARAMETER ConfigPath
    Path for exporting or importing firewall configuration.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Set-WindowsFirewall.ps1 -ComputerName "PC001" -Credential (Get-Credential) -Action "Enable" -ProfileType "All"

.EXAMPLE
    .\Set-WindowsFirewall.ps1 -ComputerName "PC001" -Action "CreateRule" -RuleName "Allow RDP" -Direction "Inbound" -Protocol "TCP" -LocalPort 3389 -Action "Allow" -ProfileType "Domain", "Private"

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
    [ValidateSet("Enable", "Disable", "CreateRule", "RemoveRule", "ExportConfig", "ImportConfig")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Domain", "Private", "Public", "All")]
    [string[]]$ProfileType = @("All"),

    [Parameter(Mandatory = $false)]
    [string]$RuleName,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Inbound", "Outbound")]
    [string]$Direction,

    [Parameter(Mandatory = $false)]
    [ValidateSet("TCP", "UDP", "Any")]
    [string]$Protocol = "Any",

    [Parameter(Mandatory = $false)]
    [string[]]$LocalPort,

    [Parameter(Mandatory = $false)]
    [string[]]$RemotePort,

    [Parameter(Mandatory = $false)]
    [string[]]$LocalAddress,

    [Parameter(Mandatory = $false)]
    [string[]]$RemoteAddress,

    [Parameter(Mandatory = $false)]
    [string]$Program,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Allow", "Block")]
    [string]$RuleAction = "Allow",

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\WindowsFirewall_$(Get-Date -Format 'yyyyMMdd').log"
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

function Enable-FirewallProfile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string[]]$ProfileType
    )
    
    try {
        Write-Log -Message "Enabling Windows Firewall profiles on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string[]]$ProfileType
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Get all profiles if "All" is specified
                $profiles = @()
                if ($ProfileType -contains "All") {
                    $profiles = @("Domain", "Private", "Public")
                }
                else {
                    $profiles = $ProfileType
                }
                
                # Enable each profile
                foreach ($profile in $profiles) {
                    $currentState = (Get-NetFirewallProfile -Name $profile).Enabled
                    
                    if ($currentState -eq $false) {
                        Set-NetFirewallProfile -Name $profile -Enabled True
                        $result.Changes += "Enabled $profile firewall profile"
                    }
                    else {
                        $result.Changes += "$profile firewall profile already enabled"
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
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $ProfileType
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfileType -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfileType
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall profiles enabled successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to enable Windows Firewall profiles on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to enable Windows Firewall profiles on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Disable-FirewallProfile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string[]]$ProfileType
    )
    
    try {
        Write-Log -Message "Disabling Windows Firewall profiles on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string[]]$ProfileType
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Get all profiles if "All" is specified
                $profiles = @()
                if ($ProfileType -contains "All") {
                    $profiles = @("Domain", "Private", "Public")
                }
                else {
                    $profiles = $ProfileType
                }
                
                # Disable each profile
                foreach ($profile in $profiles) {
                    $currentState = (Get-NetFirewallProfile -Name $profile).Enabled
                    
                    if ($currentState -eq $true) {
                        Set-NetFirewallProfile -Name $profile -Enabled False
                        $result.Changes += "Disabled $profile firewall profile"
                    }
                    else {
                        $result.Changes += "$profile firewall profile already disabled"
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
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $ProfileType
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfileType -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ProfileType
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall profiles disabled successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to disable Windows Firewall profiles on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to disable Windows Firewall profiles on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-FirewallRule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleName,
        
        [Parameter(Mandatory = $true)]
        [string]$Direction,
        
        [Parameter(Mandatory = $true)]
        [string]$Protocol,
        
        [Parameter(Mandatory = $false)]
        [string[]]$LocalPort,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RemotePort,
        
        [Parameter(Mandatory = $false)]
        [string[]]$LocalAddress,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RemoteAddress,
        
        [Parameter(Mandatory = $false)]
        [string]$Program,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleAction,
        
        [Parameter(Mandatory = $true)]
        [string[]]$ProfileType
    )
    
    try {
        Write-Log -Message "Creating Windows Firewall rule '$RuleName' on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$RuleName,
                [string]$Direction,
                [string]$Protocol,
                [string[]]$LocalPort,
                [string[]]$RemotePort,
                [string[]]$LocalAddress,
                [string[]]$RemoteAddress,
                [string]$Program,
                [string]$RuleAction,
                [string[]]$ProfileType
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Check if rule already exists
                $existingRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
                
                if ($existingRule) {
                    # Remove existing rule
                    Remove-NetFirewallRule -DisplayName $RuleName
                    $result.Changes += "Removed existing rule '$RuleName'"
                }
                
                # Get all profiles if "All" is specified
                $profiles = @()
                if ($ProfileType -contains "All") {
                    $profiles = @("Domain", "Private", "Public")
                }
                else {
                    $profiles = $ProfileType
                }
                
                # Create rule parameters
                $ruleParams = @{
                    DisplayName = $RuleName
                    Direction = $Direction
                    Action = $RuleAction
                    Profile = $profiles
                }
                
                # Add protocol
                if ($Protocol -ne "Any") {
                    $ruleParams.Add("Protocol", $Protocol)
                }
                
                # Add ports if specified
                if ($LocalPort) {
                    $ruleParams.Add("LocalPort", $LocalPort)
                }
                
                if ($RemotePort) {
                    $ruleParams.Add("RemotePort", $RemotePort)
                }
                
                # Add addresses if specified
                if ($LocalAddress) {
                    $ruleParams.Add("LocalAddress", $LocalAddress)
                }
                
                if ($RemoteAddress) {
                    $ruleParams.Add("RemoteAddress", $RemoteAddress)
                }
                
                # Add program if specified
                if ($Program) {
                    $ruleParams.Add("Program", $Program)
                }
                
                # Create the rule
                New-NetFirewallRule @ruleParams | Out-Null
                
                $result.Changes += "Created firewall rule '$RuleName'"
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $RuleName, $Direction, $Protocol, $LocalPort, $RemotePort, $LocalAddress, $RemoteAddress, $Program, $RuleAction, $ProfileType
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RuleName, $Direction, $Protocol, $LocalPort, $RemotePort, $LocalAddress, $RemoteAddress, $Program, $RuleAction, $ProfileType -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RuleName, $Direction, $Protocol, $LocalPort, $RemotePort, $LocalAddress, $RemoteAddress, $Program, $RuleAction, $ProfileType
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall rule created successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to create Windows Firewall rule on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to create Windows Firewall rule on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Remove-FirewallRule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$RuleName
    )
    
    try {
        Write-Log -Message "Removing Windows Firewall rule '$RuleName' on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$RuleName
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Check if rule exists
                $existingRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
                
                if ($existingRule) {
                    # Remove rule
                    Remove-NetFirewallRule -DisplayName $RuleName
                    $result.Changes += "Removed firewall rule '$RuleName'"
                }
                else {
                    $result.Changes += "Firewall rule '$RuleName' not found"
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $RuleName
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RuleName -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RuleName
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall rule removed successfully on $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to remove Windows Firewall rule on $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to remove Windows Firewall rule on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Export-FirewallConfig {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        Write-Log -Message "Exporting Windows Firewall configuration from $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$ConfigPath
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Create directory if it doesn't exist
                $configDir = Split-Path -Path $ConfigPath -Parent
                if (-not (Test-Path -Path $configDir)) {
                    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
                }
                
                # Export firewall configuration
                netsh advfirewall export $ConfigPath
                
                if (Test-Path -Path $ConfigPath) {
                    $result.Changes += "Exported firewall configuration to '$ConfigPath'"
                }
                else {
                    throw "Failed to export firewall configuration. File not created."
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $ConfigPath
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ConfigPath -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ConfigPath
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall configuration exported successfully from $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to export Windows Firewall configuration from $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to export Windows Firewall configuration from $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            Status = "Error"
            Changes = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Import-FirewallConfig {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        Write-Log -Message "Importing Windows Firewall configuration to $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$ConfigPath
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                Status = "Success"
                Changes = @()
                ErrorMessage = $null
            }
            
            try {
                # Check if configuration file exists
                if (-not (Test-Path -Path $ConfigPath)) {
                    throw "Configuration file '$ConfigPath' not found."
                }
                
                # Import firewall configuration
                netsh advfirewall import $ConfigPath
                
                $result.Changes += "Imported firewall configuration from '$ConfigPath'"
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $ConfigPath
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ConfigPath -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $ConfigPath
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Windows Firewall configuration imported successfully to $ComputerName. Changes: $($result.Changes -join ', ')" -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to import Windows Firewall configuration to $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to import Windows Firewall configuration to $ComputerName: $_" -Level "ERROR"
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
    Write-Log -Message "Starting Windows Firewall management process." -Level "INFO"
    
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
        
        # Perform the requested action
        switch ($Action) {
            "Enable" {
                $result = Enable-FirewallProfile -ComputerName $computer -Credential $Credential -ProfileType $ProfileType
            }
            "Disable" {
                $result = Disable-FirewallProfile -ComputerName $computer -Credential $Credential -ProfileType $ProfileType
            }
            "CreateRule" {
                # Validate required parameters
                if (-not $RuleName) {
                    Write-Log -Message "RuleName parameter is required for CreateRule action. Skipping..." -Level "ERROR"
                    $result = @{
                        ComputerName = $computer
                        Status = "Error"
                        Changes = @()
                        ErrorMessage = "RuleName parameter is required for CreateRule action"
                    }
                }
                elseif (-not $Direction) {
                    Write-Log -Message "Direction parameter is required for CreateRule action. Skipping..." -Level "ERROR"
                    $result = @{
                        ComputerName = $computer
                        Status = "Error"
                        Changes = @()
                        ErrorMessage = "Direction parameter is required for CreateRule action"
                    }
                }
                else {
                    $result = Create-FirewallRule -ComputerName $computer -Credential $Credential -RuleName $RuleName -Direction $Direction -Protocol $Protocol -LocalPort $LocalPort -RemotePort $RemotePort -LocalAddress $LocalAddress -RemoteAddress $RemoteAddress -Program $Program -RuleAction $RuleAction -ProfileType $ProfileType
                }
            }
            "RemoveRule" {
                # Validate required parameters
                if (-not $RuleName) {
                    Write-Log -Message "RuleName parameter is required for RemoveRule action. Skipping..." -Level "ERROR"
                    $result = @{
                        ComputerName = $computer
                        Status = "Error"
                        Changes = @()
                        ErrorMessage = "RuleName parameter is required for RemoveRule action"
                    }
                }
                else {
                    $result = Remove-FirewallRule -ComputerName $computer -Credential $Credential -RuleName $RuleName
                }
            }
            "ExportConfig" {
                # Validate required parameters
                if (-not $ConfigPath) {
                    Write-Log -Message "ConfigPath parameter is required for ExportConfig action. Skipping..." -Level "ERROR"
                    $result = @{
                        ComputerName = $computer
                        Status = "Error"
                        Changes = @()
                        ErrorMessage = "ConfigPath parameter is required for ExportConfig action"
                    }
                }
                else {
                    $result = Export-FirewallConfig -ComputerName $computer -Credential $Credential -ConfigPath $ConfigPath
                }
            }
            "ImportConfig" {
                # Validate required parameters
                if (-not $ConfigPath) {
                    Write-Log -Message "ConfigPath parameter is required for ImportConfig action. Skipping..." -Level "ERROR"
                    $result = @{
                        ComputerName = $computer
                        Status = "Error"
                        Changes = @()
                        ErrorMessage = "ConfigPath parameter is required for ImportConfig action"
                    }
                }
                else {
                    $result = Import-FirewallConfig -ComputerName $computer -Credential $Credential -ConfigPath $ConfigPath
                }
            }
        }
        
        $results += $result
    }
    
    # Output summary
    Write-Log -Message "Windows Firewall management process completed." -Level "INFO"
    Write-Log -Message "Summary:" -Level "INFO"
    
    foreach ($result in $results) {
        $status = "Computer: $($result.ComputerName), Status: $($result.Status)"
        
        if ($result.Status -eq "Success") {
            Write-Log -Message $status -Level "INFO"
        }
        else {
            Write-Log -Message "$status, Error: $($result.ErrorMessage)" -Level "WARNING"
        }
    }
    
    return $results
}
catch {
    Write-Log -Message "An error occurred during Windows Firewall management: $_" -Level "ERROR"
    exit 1
}
