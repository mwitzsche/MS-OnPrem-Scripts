<#
.SYNOPSIS
    Joins a computer to an Active Directory domain.

.DESCRIPTION
    This script joins a local or remote computer to an Active Directory domain.
    It provides detailed logging and error handling, and can schedule a restart after the operation.

.PARAMETER ComputerName
    Name of the target computer.

.PARAMETER DomainName
    Name of the domain to join.

.PARAMETER Credential
    Credentials to use for domain join.

.PARAMETER OUPath
    Organizational Unit path where the computer account will be created.

.PARAMETER LocalCredential
    Credentials to use for remote connection.

.PARAMETER RestartComputer
    Whether to restart the computer after joining the domain.

.PARAMETER RestartTimeout
    Timeout in seconds before restarting.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Join-Domain.ps1 -ComputerName "PC001" -DomainName "contoso.com" -Credential (Get-Credential) -OUPath "OU=Servers,DC=contoso,DC=com" -RestartComputer $true -RestartTimeout 60

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,

    [Parameter(Mandatory = $true)]
    [string]$DomainName,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$OUPath,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$LocalCredential,

    [Parameter(Mandatory = $false)]
    [bool]$RestartComputer = $true,

    [Parameter(Mandatory = $false)]
    [int]$RestartTimeout = 30,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\DomainJoin_$(Get-Date -Format 'yyyyMMdd').log"
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

function Join-ComputerToDomain {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$DomainName,
        
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [string]$OUPath,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$LocalCredential,
        
        [Parameter(Mandatory = $false)]
        [bool]$RestartComputer,
        
        [Parameter(Mandatory = $false)]
        [int]$RestartTimeout
    )
    
    try {
        Write-Log -Message "Joining computer $ComputerName to domain $DomainName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$DomainName,
                [string]$OUPath,
                [bool]$RestartComputer,
                [int]$RestartTimeout
            )
            
            $result = @{
                ComputerName = $env:COMPUTERNAME
                DomainName = $DomainName
                Status = "Success"
                DomainJoined = $false
                RestartInitiated = $false
                ErrorMessage = $null
            }
            
            try {
                # Check if already in domain
                $currentDomain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
                $isInDomain = $currentDomain -ne "WORKGROUP" -and $currentDomain -ne $env:COMPUTERNAME
                
                if ($isInDomain -and $currentDomain -eq $DomainName) {
                    $result.DomainJoined = $true
                    $result.Status = "AlreadyJoined"
                    return $result
                }
                
                # Join domain
                $joinParams = @{
                    DomainName = $DomainName
                    Credential = $using:Credential
                    Force = $true
                }
                
                if ($OUPath) {
                    $joinParams.Add("OUPath", $OUPath)
                }
                
                $joinResult = Add-Computer @joinParams -PassThru
                
                if ($joinResult) {
                    $result.DomainJoined = $true
                    
                    # Restart if requested
                    if ($RestartComputer) {
                        if ($RestartTimeout -gt 0) {
                            $message = "Domain join operation completed. Restarting in $RestartTimeout seconds..."
                            Write-Output $message
                            
                            # Schedule restart
                            $shutdownParams = @{
                                ComputerName = "localhost"
                                Force = $true
                                Restart = $true
                                Timeout = $RestartTimeout
                                Comment = "Restarting after domain join operation"
                            }
                            
                            shutdown.exe /r /t $RestartTimeout /c "Restarting after domain join operation"
                            $result.RestartInitiated = $true
                        }
                        else {
                            # Immediate restart
                            Write-Output "Domain join operation completed. Restarting immediately..."
                            shutdown.exe /r /t 0 /c "Restarting after domain join operation"
                            $result.RestartInitiated = $true
                        }
                    }
                    else {
                        Write-Output "Domain join operation completed. Restart is required to apply changes."
                    }
                }
                else {
                    throw "Failed to join domain. No specific error returned."
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $DomainName, $OUPath, $RestartComputer, $RestartTimeout
        }
        else {
            if ($LocalCredential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $DomainName, $OUPath, $RestartComputer, $RestartTimeout -Credential $LocalCredential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $DomainName, $OUPath, $RestartComputer, $RestartTimeout
            }
        }
        
        if ($result.Status -eq "Success") {
            $message = "Computer $ComputerName joined to domain $DomainName successfully."
            if ($result.RestartInitiated) {
                $message += " Restart initiated."
            }
            else {
                $message += " Restart is required to apply changes."
            }
            
            Write-Log -Message $message -Level "INFO"
        }
        elseif ($result.Status -eq "AlreadyJoined") {
            Write-Log -Message "Computer $ComputerName is already joined to domain $DomainName." -Level "INFO"
        }
        else {
            Write-Log -Message "Failed to join computer $ComputerName to domain $DomainName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to join computer $ComputerName to domain $DomainName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            DomainName = $DomainName
            Status = "Error"
            DomainJoined = $false
            RestartInitiated = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting domain join process." -Level "INFO"
    
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
    
    # Join domain
    $result = Join-ComputerToDomain -ComputerName $ComputerName -DomainName $DomainName -Credential $Credential -OUPath $OUPath -LocalCredential $LocalCredential -RestartComputer $RestartComputer -RestartTimeout $RestartTimeout
    
    if ($result.Status -eq "Success" -or $result.Status -eq "AlreadyJoined") {
        Write-Log -Message "Domain join process completed successfully." -Level "INFO"
    }
    else {
        Write-Log -Message "Domain join process failed: $($result.ErrorMessage)" -Level "ERROR"
        exit 1
    }
    
    return $result
}
catch {
    Write-Log -Message "An error occurred during domain join process: $_" -Level "ERROR"
    exit 1
}
