<#
.SYNOPSIS
    Renames a computer and optionally joins it to a domain.

.DESCRIPTION
    This script renames a local or remote computer and optionally joins it to a domain.
    It provides detailed logging and error handling, and can schedule a restart after the operation.

.PARAMETER ComputerName
    Current name of the target computer.

.PARAMETER NewName
    New name for the computer.

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER DomainName
    Domain to join (optional).

.PARAMETER DomainCredential
    Credentials to use for domain join.

.PARAMETER RestartComputer
    Whether to restart the computer after renaming.

.PARAMETER RestartTimeout
    Timeout in seconds before restarting.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Rename-WindowsComputer.ps1 -ComputerName "PC001" -NewName "PC-IT-001" -Credential (Get-Credential) -DomainName "contoso.com" -DomainCredential (Get-Credential) -RestartComputer $true -RestartTimeout 60

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [Parameter(Mandatory = $true)]
    [string]$NewName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$DomainName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$DomainCredential,

    [Parameter(Mandatory = $false)]
    [bool]$RestartComputer = $true,

    [Parameter(Mandatory = $false)]
    [int]$RestartTimeout = 30,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ComputerRename_$(Get-Date -Format 'yyyyMMdd').log"
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

function Rename-Computer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$NewName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$DomainCredential,
        
        [Parameter(Mandatory = $false)]
        [bool]$RestartComputer,
        
        [Parameter(Mandatory = $false)]
        [int]$RestartTimeout
    )
    
    try {
        Write-Log -Message "Renaming computer $ComputerName to $NewName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$NewName,
                [string]$DomainName,
                [bool]$RestartComputer,
                [int]$RestartTimeout
            )
            
            $result = @{
                OldName = $env:COMPUTERNAME
                NewName = $NewName
                Status = "Success"
                DomainJoined = $false
                RestartInitiated = $false
                ErrorMessage = $null
            }
            
            try {
                # Check if already in domain
                $currentDomain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
                $isInDomain = $currentDomain -ne "WORKGROUP" -and $currentDomain -ne $env:COMPUTERNAME
                
                # Rename and optionally join domain
                if ($DomainName) {
                    # Domain join with rename
                    if ($isInDomain -and $currentDomain -eq $DomainName) {
                        # Already in the correct domain, just rename
                        Rename-Computer -NewName $NewName -Force
                        $result.DomainJoined = $true
                    }
                    else {
                        # Join new domain with rename
                        Add-Computer -DomainName $DomainName -NewName $NewName -Credential $using:DomainCredential -Force
                        $result.DomainJoined = $true
                    }
                }
                else {
                    # Just rename
                    Rename-Computer -NewName $NewName -Force
                }
                
                # Restart if requested
                if ($RestartComputer) {
                    if ($RestartTimeout -gt 0) {
                        $message = "Computer rename operation completed. Restarting in $RestartTimeout seconds..."
                        Write-Output $message
                        
                        # Schedule restart
                        $shutdownParams = @{
                            ComputerName = "localhost"
                            Force = $true
                            Restart = $true
                            Timeout = $RestartTimeout
                            Comment = "Restarting after computer rename operation"
                        }
                        
                        shutdown.exe /r /t $RestartTimeout /c "Restarting after computer rename operation"
                        $result.RestartInitiated = $true
                    }
                    else {
                        # Immediate restart
                        Write-Output "Computer rename operation completed. Restarting immediately..."
                        shutdown.exe /r /t 0 /c "Restarting after computer rename operation"
                        $result.RestartInitiated = $true
                    }
                }
                else {
                    Write-Output "Computer rename operation completed. Restart is required to apply changes."
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $NewName, $DomainName, $RestartComputer, $RestartTimeout
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $NewName, $DomainName, $RestartComputer, $RestartTimeout -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $NewName, $DomainName, $RestartComputer, $RestartTimeout
            }
        }
        
        if ($result.Status -eq "Success") {
            $message = "Computer renamed successfully from $($result.OldName) to $($result.NewName)."
            if ($result.DomainJoined) {
                $message += " Joined to domain $DomainName."
            }
            if ($result.RestartInitiated) {
                $message += " Restart initiated."
            }
            else {
                $message += " Restart is required to apply changes."
            }
            
            Write-Log -Message $message -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to rename computer $ComputerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to rename computer $ComputerName: $_" -Level "ERROR"
        return @{
            OldName = $ComputerName
            NewName = $NewName
            Status = "Error"
            DomainJoined = $false
            RestartInitiated = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting computer rename process." -Level "INFO"
    
    # Check if computer is reachable
    if ($ComputerName -ne $env:COMPUTERNAME) {
        if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
            Write-Log -Message "Computer '$ComputerName' is not reachable. Exiting..." -Level "ERROR"
            exit 1
        }
        
        # Check if PSRemoting is enabled
        if (-not (Test-PSRemoting -ComputerName $ComputerName)) {
            Write-Log -Message "PowerShell Remoting is not enabled on '$ComputerName'. Exiting..." -Level "ERROR"
            exit 1
        }
    }
    
    # Validate domain credentials if domain join is requested
    if ($DomainName -and -not $DomainCredential) {
        Write-Log -Message "Domain credentials are required for domain join. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Rename computer
    $result = Rename-Computer -ComputerName $ComputerName -NewName $NewName -Credential $Credential -DomainName $DomainName -DomainCredential $DomainCredential -RestartComputer $RestartComputer -RestartTimeout $RestartTimeout
    
    if ($result.Status -eq "Success") {
        Write-Log -Message "Computer rename process completed successfully." -Level "INFO"
    }
    else {
        Write-Log -Message "Computer rename process failed: $($result.ErrorMessage)" -Level "ERROR"
        exit 1
    }
    
    return $result
}
catch {
    Write-Log -Message "An error occurred during computer rename process: $_" -Level "ERROR"
    exit 1
}
