<#
.SYNOPSIS
    Installs and configures Internet Information Services (IIS) on Windows Server.

.DESCRIPTION
    This script installs and configures Internet Information Services (IIS) on Windows Server,
    including required features, application pools, websites, virtual directories, and SSL bindings.
    It provides detailed logging and error handling.

.PARAMETER ServerName
    Name of the target server.

.PARAMETER Credential
    Credentials to use for remote connection.

.PARAMETER Features
    Array of IIS features to install.

.PARAMETER ApplicationPools
    Array of application pools to create.

.PARAMETER Websites
    Array of websites to create.

.PARAMETER VirtualDirectories
    Array of virtual directories to create.

.PARAMETER SSLBindings
    Array of SSL bindings to configure.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    $features = @("Web-Server", "Web-Mgmt-Tools", "Web-Asp-Net45", "Web-Net-Ext45")
    $appPools = @(
        @{Name="MyAppPool"; RuntimeVersion="v4.0"; PipelineMode="Integrated"; Identity="ApplicationPoolIdentity"}
    )
    $websites = @(
        @{Name="MyWebsite"; PhysicalPath="C:\inetpub\wwwroot\MyWebsite"; ApplicationPool="MyAppPool"; Port=80; IPAddress="*"; HostHeader="mywebsite.local"}
    )
    $virtualDirs = @(
        @{Name="MyVDir"; Website="MyWebsite"; PhysicalPath="C:\inetpub\wwwroot\MyVDir"; Path="/MyVDir"}
    )
    $sslBindings = @(
        @{Website="MyWebsite"; Port=443; IPAddress="*"; HostHeader="mywebsite.local"; CertificateThumbprint="ABCDEF1234567890ABCDEF1234567890ABCDEF12"}
    )
    .\Install-IISServer.ps1 -ServerName "WebServer01" -Features $features -ApplicationPools $appPools -Websites $websites -VirtualDirectories $virtualDirs -SSLBindings $sslBindings

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ServerName = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string[]]$Features = @("Web-Server", "Web-Mgmt-Tools", "Web-Asp-Net45", "Web-Net-Ext45"),

    [Parameter(Mandatory = $false)]
    [array]$ApplicationPools,

    [Parameter(Mandatory = $false)]
    [array]$Websites,

    [Parameter(Mandatory = $false)]
    [array]$VirtualDirectories,

    [Parameter(Mandatory = $false)]
    [array]$SSLBindings,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\IISInstall_$(Get-Date -Format 'yyyyMMdd').log"
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

function Install-IISFeatures {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Features
    )
    
    try {
        Write-Log -Message "Installing IIS features on $ServerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [string[]]$Features
            )
            
            $result = @{
                ServerName = $env:COMPUTERNAME
                Status = "Success"
                InstalledFeatures = @()
                FailedFeatures = @()
                ErrorMessage = $null
            }
            
            try {
                # Check if server OS
                $osInfo = Get-WmiObject -Class Win32_OperatingSystem
                $isServer = $osInfo.ProductType -eq 3
                
                if (-not $isServer) {
                    throw "This script is designed for Windows Server. Current OS is not a server OS."
                }
                
                # Import required modules
                Import-Module ServerManager
                
                # Get currently installed features
                $installedFeatures = Get-WindowsFeature | Where-Object { $_.Installed -eq $true } | Select-Object -ExpandProperty Name
                
                # Install features
                foreach ($feature in $Features) {
                    try {
                        if ($installedFeatures -contains $feature) {
                            $result.InstalledFeatures += "$feature (already installed)"
                        }
                        else {
                            $installResult = Install-WindowsFeature -Name $feature -IncludeManagementTools
                            
                            if ($installResult.Success) {
                                $result.InstalledFeatures += $feature
                            }
                            else {
                                $result.FailedFeatures += "$feature (installation failed)"
                            }
                        }
                    }
                    catch {
                        $result.FailedFeatures += "$feature (error: $($_.Exception.Message))"
                    }
                }
                
                # Check if IIS is installed
                $iisFeature = Get-WindowsFeature -Name Web-Server
                if (-not $iisFeature.Installed) {
                    throw "IIS (Web-Server) feature is not installed. Installation may have failed."
                }
                
                # Import WebAdministration module
                Import-Module WebAdministration
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ServerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $Features
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $Features -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $Features
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "IIS features installed successfully on $ServerName." -Level "INFO"
            Write-Log -Message "Installed features: $($result.InstalledFeatures -join ', ')" -Level "INFO"
            
            if ($result.FailedFeatures.Count -gt 0) {
                Write-Log -Message "Failed features: $($result.FailedFeatures -join ', ')" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "Failed to install IIS features on $ServerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to install IIS features on $ServerName: $_" -Level "ERROR"
        return @{
            ServerName = $ServerName
            Status = "Error"
            InstalledFeatures = @()
            FailedFeatures = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-ApplicationPools {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [array]$ApplicationPools
    )
    
    try {
        Write-Log -Message "Creating application pools on $ServerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [array]$ApplicationPools
            )
            
            $result = @{
                ServerName = $env:COMPUTERNAME
                Status = "Success"
                CreatedPools = @()
                FailedPools = @()
                ErrorMessage = $null
            }
            
            try {
                # Import WebAdministration module
                Import-Module WebAdministration
                
                # Create application pools
                foreach ($pool in $ApplicationPools) {
                    try {
                        $poolName = $pool.Name
                        $runtimeVersion = $pool.RuntimeVersion
                        $pipelineMode = $pool.PipelineMode
                        $identity = $pool.Identity
                        
                        # Check if pool already exists
                        $existingPool = Get-Item "IIS:\AppPools\$poolName" -ErrorAction SilentlyContinue
                        
                        if ($existingPool) {
                            # Update existing pool
                            Set-ItemProperty "IIS:\AppPools\$poolName" -Name "managedRuntimeVersion" -Value $runtimeVersion
                            Set-ItemProperty "IIS:\AppPools\$poolName" -Name "managedPipelineMode" -Value $pipelineMode
                            
                            # Set identity
                            if ($identity -eq "ApplicationPoolIdentity") {
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.identityType" -Value 4
                            }
                            elseif ($identity -eq "LocalSystem") {
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.identityType" -Value 0
                            }
                            elseif ($identity -eq "LocalService") {
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.identityType" -Value 1
                            }
                            elseif ($identity -eq "NetworkService") {
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.identityType" -Value 2
                            }
                            elseif ($identity -eq "SpecificUser") {
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.identityType" -Value 3
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.userName" -Value $pool.UserName
                                Set-ItemProperty "IIS:\AppPools\$poolName" -Name "processModel.password" -Value $pool.Password
                            }
                            
                            $result.CreatedPools += "$poolName (updated)"
                        }
                        else {
                            # Create new pool
                            $newPool = New-WebAppPool -Name $poolName
                            $newPool | Set-ItemProperty -Name "managedRuntimeVersion" -Value $runtimeVersion
                            $newPool | Set-ItemProperty -Name "managedPipelineMode" -Value $pipelineMode
                            
                            # Set identity
                            if ($identity -eq "ApplicationPoolIdentity") {
                                $newPool | Set-ItemProperty -Name "processModel.identityType" -Value 4
                            }
                            elseif ($identity -eq "LocalSystem") {
                                $newPool | Set-ItemProperty -Name "processModel.identityType" -Value 0
                            }
                            elseif ($identity -eq "LocalService") {
                                $newPool | Set-ItemProperty -Name "processModel.identityType" -Value 1
                            }
                            elseif ($identity -eq "NetworkService") {
                                $newPool | Set-ItemProperty -Name "processModel.identityType" -Value 2
                            }
                            elseif ($identity -eq "SpecificUser") {
                                $newPool | Set-ItemProperty -Name "processModel.identityType" -Value 3
                                $newPool | Set-ItemProperty -Name "processModel.userName" -Value $pool.UserName
                                $newPool | Set-ItemProperty -Name "processModel.password" -Value $pool.Password
                            }
                            
                            $result.CreatedPools += "$poolName (created)"
                        }
                    }
                    catch {
                        $result.FailedPools += "$poolName (error: $($_.Exception.Message))"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ServerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $ApplicationPools
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $ApplicationPools -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $ApplicationPools
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Application pools created successfully on $ServerName." -Level "INFO"
            Write-Log -Message "Created pools: $($result.CreatedPools -join ', ')" -Level "INFO"
            
            if ($result.FailedPools.Count -gt 0) {
                Write-Log -Message "Failed pools: $($result.FailedPools -join ', ')" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "Failed to create application pools on $ServerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to create application pools on $ServerName: $_" -Level "ERROR"
        return @{
            ServerName = $ServerName
            Status = "Error"
            CreatedPools = @()
            FailedPools = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-Websites {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [array]$Websites
    )
    
    try {
        Write-Log -Message "Creating websites on $ServerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [array]$Websites
            )
            
            $result = @{
                ServerName = $env:COMPUTERNAME
                Status = "Success"
                CreatedSites = @()
                FailedSites = @()
                ErrorMessage = $null
            }
            
            try {
                # Import WebAdministration module
                Import-Module WebAdministration
                
                # Create websites
                foreach ($site in $Websites) {
                    try {
                        $siteName = $site.Name
                        $physicalPath = $site.PhysicalPath
                        $appPool = $site.ApplicationPool
                        $port = $site.Port
                        $ipAddress = $site.IPAddress
                        $hostHeader = $site.HostHeader
                        
                        # Create physical path if it doesn't exist
                        if (-not (Test-Path -Path $physicalPath)) {
                            New-Item -Path $physicalPath -ItemType Directory -Force | Out-Null
                        }
                        
                        # Check if site already exists
                        $existingSite = Get-Website -Name $siteName -ErrorAction SilentlyContinue
                        
                        if ($existingSite) {
                            # Update existing site
                            Set-ItemProperty "IIS:\Sites\$siteName" -Name "physicalPath" -Value $physicalPath
                            Set-ItemProperty "IIS:\Sites\$siteName" -Name "applicationPool" -Value $appPool
                            
                            # Update bindings
                            $bindingInfo = "*:$port:$hostHeader"
                            if ($hostHeader -eq "") {
                                $bindingInfo = "*:$port"
                            }
                            
                            Set-WebBinding -Name $siteName -BindingInformation $bindingInfo -PropertyName "bindingInformation"
                            
                            $result.CreatedSites += "$siteName (updated)"
                        }
                        else {
                            # Create new site
                            $bindingInfo = "*:$port:$hostHeader"
                            if ($hostHeader -eq "") {
                                $bindingInfo = "*:$port"
                            }
                            
                            $newSite = New-Website -Name $siteName -PhysicalPath $physicalPath -ApplicationPool $appPool -Port $port -HostHeader $hostHeader -IPAddress $ipAddress -Force
                            
                            $result.CreatedSites += "$siteName (created)"
                        }
                        
                        # Start website
                        Start-Website -Name $siteName
                    }
                    catch {
                        $result.FailedSites += "$siteName (error: $($_.Exception.Message))"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ServerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $Websites
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $Websites -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $Websites
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Websites created successfully on $ServerName." -Level "INFO"
            Write-Log -Message "Created sites: $($result.CreatedSites -join ', ')" -Level "INFO"
            
            if ($result.FailedSites.Count -gt 0) {
                Write-Log -Message "Failed sites: $($result.FailedSites -join ', ')" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "Failed to create websites on $ServerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to create websites on $ServerName: $_" -Level "ERROR"
        return @{
            ServerName = $ServerName
            Status = "Error"
            CreatedSites = @()
            FailedSites = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-VirtualDirectories {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [array]$VirtualDirectories
    )
    
    try {
        Write-Log -Message "Creating virtual directories on $ServerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [array]$VirtualDirectories
            )
            
            $result = @{
                ServerName = $env:COMPUTERNAME
                Status = "Success"
                CreatedVDirs = @()
                FailedVDirs = @()
                ErrorMessage = $null
            }
            
            try {
                # Import WebAdministration module
                Import-Module WebAdministration
                
                # Create virtual directories
                foreach ($vdir in $VirtualDirectories) {
                    try {
                        $vdirName = $vdir.Name
                        $website = $vdir.Website
                        $physicalPath = $vdir.PhysicalPath
                        $path = $vdir.Path
                        
                        # Create physical path if it doesn't exist
                        if (-not (Test-Path -Path $physicalPath)) {
                            New-Item -Path $physicalPath -ItemType Directory -Force | Out-Null
                        }
                        
                        # Check if virtual directory already exists
                        $existingVDir = Get-WebVirtualDirectory -Site $website -Name $vdirName -ErrorAction SilentlyContinue
                        
                        if ($existingVDir) {
                            # Update existing virtual directory
                            Set-ItemProperty "IIS:\Sites\$website\$path" -Name "physicalPath" -Value $physicalPath
                            
                            $result.CreatedVDirs += "$website$path (updated)"
                        }
                        else {
                            # Create new virtual directory
                            $newVDir = New-WebVirtualDirectory -Site $website -Name $vdirName -PhysicalPath $physicalPath
                            
                            $result.CreatedVDirs += "$website$path (created)"
                        }
                    }
                    catch {
                        $result.FailedVDirs += "$website$path (error: $($_.Exception.Message))"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ServerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $VirtualDirectories
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $VirtualDirectories -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $VirtualDirectories
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "Virtual directories created successfully on $ServerName." -Level "INFO"
            Write-Log -Message "Created virtual directories: $($result.CreatedVDirs -join ', ')" -Level "INFO"
            
            if ($result.FailedVDirs.Count -gt 0) {
                Write-Log -Message "Failed virtual directories: $($result.FailedVDirs -join ', ')" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "Failed to create virtual directories on $ServerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to create virtual directories on $ServerName: $_" -Level "ERROR"
        return @{
            ServerName = $ServerName
            Status = "Error"
            CreatedVDirs = @()
            FailedVDirs = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Configure-SSLBindings {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [array]$SSLBindings
    )
    
    try {
        Write-Log -Message "Configuring SSL bindings on $ServerName..." -Level "INFO"
        
        $scriptBlock = {
            param (
                [array]$SSLBindings
            )
            
            $result = @{
                ServerName = $env:COMPUTERNAME
                Status = "Success"
                ConfiguredBindings = @()
                FailedBindings = @()
                ErrorMessage = $null
            }
            
            try {
                # Import WebAdministration module
                Import-Module WebAdministration
                
                # Configure SSL bindings
                foreach ($binding in $SSLBindings) {
                    try {
                        $website = $binding.Website
                        $port = $binding.Port
                        $ipAddress = $binding.IPAddress
                        $hostHeader = $binding.HostHeader
                        $thumbprint = $binding.CertificateThumbprint
                        
                        # Check if website exists
                        $existingSite = Get-Website -Name $website -ErrorAction SilentlyContinue
                        
                        if (-not $existingSite) {
                            throw "Website '$website' does not exist."
                        }
                        
                        # Check if certificate exists
                        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $thumbprint }
                        
                        if (-not $cert) {
                            throw "Certificate with thumbprint '$thumbprint' not found."
                        }
                        
                        # Create binding information
                        $bindingInfo = "${ipAddress}:${port}:${hostHeader}"
                        
                        # Check if binding already exists
                        $existingBinding = Get-WebBinding -Name $website -Protocol "https" -HostHeader $hostHeader -Port $port -IPAddress $ipAddress -ErrorAction SilentlyContinue
                        
                        if ($existingBinding) {
                            # Remove existing binding
                            Remove-WebBinding -Name $website -Protocol "https" -HostHeader $hostHeader -Port $port -IPAddress $ipAddress
                        }
                        
                        # Create new binding
                        New-WebBinding -Name $website -Protocol "https" -HostHeader $hostHeader -Port $port -IPAddress $ipAddress -SslFlags 0
                        
                        # Assign certificate
                        $sslPath = "IIS:\SslBindings\${ipAddress}!${port}!${hostHeader}"
                        if ($hostHeader -eq "") {
                            $sslPath = "IIS:\SslBindings\${ipAddress}!${port}"
                        }
                        
                        if (Test-Path $sslPath) {
                            Remove-Item $sslPath -Force
                        }
                        
                        $cert | New-Item $sslPath -Force | Out-Null
                        
                        $result.ConfiguredBindings += "$website (https://${hostHeader}:${port})"
                    }
                    catch {
                        $result.FailedBindings += "$website (https://${hostHeader}:${port}) (error: $($_.Exception.Message))"
                    }
                }
            }
            catch {
                $result.Status = "Error"
                $result.ErrorMessage = $_.Exception.Message
            }
            
            return $result
        }
        
        if ($ServerName -eq $env:COMPUTERNAME) {
            $result = Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $SSLBindings
        }
        else {
            if ($Credential) {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $SSLBindings -Credential $Credential
            }
            else {
                $result = Invoke-Command -ComputerName $ServerName -ScriptBlock $scriptBlock -ArgumentList $SSLBindings
            }
        }
        
        if ($result.Status -eq "Success") {
            Write-Log -Message "SSL bindings configured successfully on $ServerName." -Level "INFO"
            Write-Log -Message "Configured bindings: $($result.ConfiguredBindings -join ', ')" -Level "INFO"
            
            if ($result.FailedBindings.Count -gt 0) {
                Write-Log -Message "Failed bindings: $($result.FailedBindings -join ', ')" -Level "WARNING"
            }
        }
        else {
            Write-Log -Message "Failed to configure SSL bindings on $ServerName: $($result.ErrorMessage)" -Level "ERROR"
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Failed to configure SSL bindings on $ServerName: $_" -Level "ERROR"
        return @{
            ServerName = $ServerName
            Status = "Error"
            ConfiguredBindings = @()
            FailedBindings = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting IIS installation and configuration process." -Level "INFO"
    
    # Check if server is reachable
    if ($ServerName -ne $env:COMPUTERNAME) {
        if (-not (Test-Connection -ComputerName $ServerName -Count 1 -Quiet)) {
            Write-Log -Message "Server '$ServerName' is not reachable. Exiting..." -Level "ERROR"
            exit 1
        }
        
        # Check if PSRemoting is enabled
        if (-not (Test-PSRemoting -ComputerName $ServerName)) {
            Write-Log -Message "PowerShell Remoting is not enabled on '$ServerName'. Exiting..." -Level "ERROR"
            exit 1
        }
    }
    
    # Install IIS features
    $featuresResult = Install-IISFeatures -ServerName $ServerName -Credential $Credential -Features $Features
    
    if ($featuresResult.Status -ne "Success") {
        Write-Log -Message "Failed to install IIS features. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Create application pools
    if ($ApplicationPools -and $ApplicationPools.Count -gt 0) {
        $appPoolsResult = Create-ApplicationPools -ServerName $ServerName -Credential $Credential -ApplicationPools $ApplicationPools
        
        if ($appPoolsResult.Status -ne "Success") {
            Write-Log -Message "Failed to create application pools. Continuing..." -Level "WARNING"
        }
    }
    
    # Create websites
    if ($Websites -and $Websites.Count -gt 0) {
        $websitesResult = Create-Websites -ServerName $ServerName -Credential $Credential -Websites $Websites
        
        if ($websitesResult.Status -ne "Success") {
            Write-Log -Message "Failed to create websites. Continuing..." -Level "WARNING"
        }
    }
    
    # Create virtual directories
    if ($VirtualDirectories -and $VirtualDirectories.Count -gt 0) {
        $vdirsResult = Create-VirtualDirectories -ServerName $ServerName -Credential $Credential -VirtualDirectories $VirtualDirectories
        
        if ($vdirsResult.Status -ne "Success") {
            Write-Log -Message "Failed to create virtual directories. Continuing..." -Level "WARNING"
        }
    }
    
    # Configure SSL bindings
    if ($SSLBindings -and $SSLBindings.Count -gt 0) {
        $sslResult = Configure-SSLBindings -ServerName $ServerName -Credential $Credential -SSLBindings $SSLBindings
        
        if ($sslResult.Status -ne "Success") {
            Write-Log -Message "Failed to configure SSL bindings. Continuing..." -Level "WARNING"
        }
    }
    
    Write-Log -Message "IIS installation and configuration process completed." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during IIS installation and configuration: $_" -Level "ERROR"
    exit 1
}
