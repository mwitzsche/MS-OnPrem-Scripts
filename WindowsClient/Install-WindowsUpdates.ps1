<#
.SYNOPSIS
    Installs Windows updates on local or remote computers.

.DESCRIPTION
    This script installs Windows updates on local or remote computers based on specified criteria.
    It can install security updates, critical updates, or all available updates, and optionally
    reboot the computer if required. It provides detailed logging and error handling.

.PARAMETER ComputerName
    Name of the target computer(s).

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER UpdateType
    Type of updates to install (Security, Critical, All).

.PARAMETER RebootIfRequired
    Whether to reboot the computer if required.

.PARAMETER ScheduleReboot
    Time to schedule reboot (if not immediate).

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Install-WindowsUpdates.ps1 -ComputerName @("PC001", "PC002") -Credential (Get-Credential) -UpdateType "Security" -RebootIfRequired $true -ScheduleReboot "22:00"

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
    [ValidateSet("Security", "Critical", "All")]
    [string]$UpdateType = "All",

    [Parameter(Mandatory = $false)]
    [bool]$RebootIfRequired = $false,

    [Parameter(Mandatory = $false)]
    [string]$ScheduleReboot,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\WindowsUpdates_$(Get-Date -Format 'yyyyMMdd').log"
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

function Install-Updates {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string]$UpdateType,
        
        [Parameter(Mandatory = $true)]
        [bool]$RebootIfRequired,
        
        [Parameter(Mandatory = $false)]
        [string]$ScheduleReboot
    )
    
    try {
        Write-Log -Message "Installing Windows updates on $ComputerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string]$UpdateType,
                [bool]$RebootIfRequired,
                [string]$ScheduleReboot
            )
            
            # Create update search criteria based on update type
            $criteria = "IsInstalled=0"
            
            switch ($UpdateType) {
                "Security" {
                    $criteria += " AND CategoryIDs contains '0FA1201D-4330-4FA8-8AE9-B877473B6441'"
                }
                "Critical" {
                    $criteria += " AND IsHidden=0 AND IsAssigned=1"
                }
                "All" {
                    $criteria += " AND IsHidden=0"
                }
            }
            
            # Create update session and searcher
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            
            # Search for updates
            Write-Output "Searching for updates..."
            $searchResult = $searcher.Search($criteria)
            
            if ($searchResult.Updates.Count -eq 0) {
                Write-Output "No updates found."
                return @{
                    ComputerName = $env:COMPUTERNAME
                    UpdatesFound = 0
                    UpdatesInstalled = 0
                    RebootRequired = $false
                    Status = "No updates found"
                    ErrorMessage = $null
                }
            }
            
            Write-Output "Found $($searchResult.Updates.Count) update(s)."
            
            # Create update collection for download
            $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
            
            foreach ($update in $searchResult.Updates) {
                if ($update.EulaAccepted -eq $false) {
                    $update.AcceptEula()
                }
                $updatesToDownload.Add($update) | Out-Null
            }
            
            # Download updates
            Write-Output "Downloading updates..."
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $updatesToDownload
            $downloadResult = $downloader.Download()
            
            # Create update collection for installation
            $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            
            foreach ($update in $searchResult.Updates) {
                if ($update.IsDownloaded) {
                    $updatesToInstall.Add($update) | Out-Null
                }
            }
            
            if ($updatesToInstall.Count -eq 0) {
                Write-Output "No updates were successfully downloaded."
                return @{
                    ComputerName = $env:COMPUTERNAME
                    UpdatesFound = $searchResult.Updates.Count
                    UpdatesInstalled = 0
                    RebootRequired = $false
                    Status = "Download failed"
                    ErrorMessage = "No updates were successfully downloaded"
                }
            }
            
            # Install updates
            Write-Output "Installing updates..."
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $updatesToInstall
            $installResult = $installer.Install()
            
            # Check if reboot is required
            $rebootRequired = $installResult.RebootRequired
            
            # Handle reboot if required
            if ($rebootRequired -and $RebootIfRequired) {
                if ($ScheduleReboot) {
                    try {
                        $scheduledTime = [DateTime]::Parse($ScheduleReboot)
                        $currentTime = Get-Date
                        
                        if ($scheduledTime -gt $currentTime) {
                            $secondsUntilReboot = ($scheduledTime - $currentTime).TotalSeconds
                            Write-Output "Scheduling reboot for $ScheduleReboot (in $([math]::Round($secondsUntilReboot / 60)) minutes)..."
                            shutdown.exe /r /t $secondsUntilReboot /c "Scheduled reboot after Windows Update installation"
                        }
                        else {
                            Write-Output "Scheduled time is in the past. Rebooting immediately..."
                            shutdown.exe /r /t 60 /c "Rebooting after Windows Update installation"
                        }
                    }
                    catch {
                        Write-Output "Invalid scheduled time format. Rebooting immediately..."
                        shutdown.exe /r /t 60 /c "Rebooting after Windows Update installation"
                    }
                }
                else {
                    Write-Output "Rebooting in 60 seconds..."
                    shutdown.exe /r /t 60 /c "Rebooting after Windows Update installation"
                }
            }
            
            # Prepare result
            $installedUpdates = @()
            for ($i = 0; $i -lt $updatesToInstall.Count; $i++) {
                $update = $updatesToInstall.Item($i)
                $installedUpdates += [PSCustomObject]@{
                    Title = $update.Title
                    Result = switch ($installResult.GetUpdateResult($i).ResultCode) {
                        0 { "Not Started" }
                        1 { "In Progress" }
                        2 { "Succeeded" }
                        3 { "Succeeded With Errors" }
                        4 { "Failed" }
                        5 { "Aborted" }
                        default { "Unknown" }
                    }
                }
            }
            
            return @{
                ComputerName = $env:COMPUTERNAME
                UpdatesFound = $searchResult.Updates.Count
                UpdatesInstalled = ($installedUpdates | Where-Object { $_.Result -eq "Succeeded" }).Count
                RebootRequired = $rebootRequired
                Status = "Installation completed"
                InstalledUpdates = $installedUpdates
                ErrorMessage = $null
            }
        }
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $UpdateType, $RebootIfRequired, $ScheduleReboot
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $UpdateType, $RebootIfRequired, $ScheduleReboot -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $UpdateType, $RebootIfRequired, $ScheduleReboot
            }
        }
        
        Write-Log -Message "Windows updates installation completed on $ComputerName. Found: $($result.UpdatesFound), Installed: $($result.UpdatesInstalled), Reboot Required: $($result.RebootRequired)" -Level "INFO"
        return $result
    }
    catch {
        Write-Log -Message "Failed to install Windows updates on $ComputerName: $_" -Level "ERROR"
        return @{
            ComputerName = $ComputerName
            UpdatesFound = 0
            UpdatesInstalled = 0
            RebootRequired = $false
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Windows updates installation process." -Level "INFO"
    
    $results = @()
    
    foreach ($computer in $ComputerName) {
        Write-Log -Message "Processing computer: $computer" -Level "INFO"
        
        # Check if computer is reachable
        if ($computer -ne $env:COMPUTERNAME) {
            if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
                Write-Log -Message "Computer '$computer' is not reachable. Skipping..." -Level "WARNING"
                $results += @{
                    ComputerName = $computer
                    UpdatesFound = 0
                    UpdatesInstalled = 0
                    RebootRequired = $false
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
                    UpdatesFound = 0
                    UpdatesInstalled = 0
                    RebootRequired = $false
                    Status = "PSRemoting Disabled"
                    ErrorMessage = "PowerShell Remoting is not enabled"
                }
                continue
            }
        }
        
        # Install updates
        $result = Install-Updates -ComputerName $computer -Credential $Credential -UpdateType $UpdateType -RebootIfRequired $RebootIfRequired -ScheduleReboot $ScheduleReboot
        $results += $result
    }
    
    # Output summary
    Write-Log -Message "Windows updates installation process completed." -Level "INFO"
    Write-Log -Message "Summary:" -Level "INFO"
    
    foreach ($result in $results) {
        $status = "Computer: $($result.ComputerName), Status: $($result.Status), Updates Found: $($result.UpdatesFound), Updates Installed: $($result.UpdatesInstalled), Reboot Required: $($result.RebootRequired)"
        if ($result.ErrorMessage) {
            $status += ", Error: $($result.ErrorMessage)"
            Write-Log -Message $status -Level "WARNING"
        }
        else {
            Write-Log -Message $status -Level "INFO"
        }
    }
    
    return $results
}
catch {
    Write-Log -Message "An error occurred during Windows updates installation: $_" -Level "ERROR"
    exit 1
}
