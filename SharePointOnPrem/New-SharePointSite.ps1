<#
.SYNOPSIS
    Creates a new SharePoint site collection in an on-premises SharePoint environment.

.DESCRIPTION
    This script creates a new SharePoint site collection in an on-premises SharePoint environment,
    including configuring site settings, permissions, and features. It provides detailed logging and error handling.

.PARAMETER SiteUrl
    URL for the new site collection.

.PARAMETER Title
    Title for the new site collection.

.PARAMETER Description
    Description for the new site collection.

.PARAMETER OwnerAlias
    User account that will be the primary site collection administrator.

.PARAMETER SecondaryOwnerAlias
    User account that will be the secondary site collection administrator.

.PARAMETER Template
    Site template to use for the new site collection.

.PARAMETER ContentDatabase
    Content database where the site collection will be created.

.PARAMETER Language
    Language ID for the site collection.

.PARAMETER TimeZone
    Time zone ID for the site collection.

.PARAMETER Quota
    Quota template to apply to the site collection.

.PARAMETER SharePointServer
    SharePoint server to connect to.

.PARAMETER Credential
    Credentials to use for SharePoint operations.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\New-SharePointSite.ps1 -SiteUrl "https://sharepoint.contoso.com/sites/ProjectX" -Title "Project X" -Description "Project X Collaboration Site" -OwnerAlias "contoso\john.doe" -SecondaryOwnerAlias "contoso\jane.smith" -Template "STS#0" -ContentDatabase "WSS_Content" -Language 1033 -TimeZone 4 -Quota "Team Site" -SharePointServer "sharepoint.contoso.com" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$Title,

    [Parameter(Mandatory = $false)]
    [string]$Description,

    [Parameter(Mandatory = $true)]
    [string]$OwnerAlias,

    [Parameter(Mandatory = $false)]
    [string]$SecondaryOwnerAlias,

    [Parameter(Mandatory = $false)]
    [string]$Template = "STS#0",

    [Parameter(Mandatory = $false)]
    [string]$ContentDatabase,

    [Parameter(Mandatory = $false)]
    [int]$Language = 1033,

    [Parameter(Mandatory = $false)]
    [int]$TimeZone = 4,

    [Parameter(Mandatory = $false)]
    [string]$Quota,

    [Parameter(Mandatory = $true)]
    [string]$SharePointServer,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\SharePointSite_$(Get-Date -Format 'yyyyMMdd').log"
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

function Connect-SharePointServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SharePointServer,
        
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Connecting to SharePoint server '$SharePointServer'..." -Level "INFO"
        
        # Check if SharePoint PowerShell module is available
        if (-not (Get-Module -Name Microsoft.SharePoint.PowerShell -ListAvailable)) {
            throw "SharePoint PowerShell module is not available. Please run this script on a SharePoint server or a server with SharePoint Management Shell installed."
        }
        
        # Add SharePoint PowerShell snap-in if not already loaded
        if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.SharePoint.PowerShell
        }
        
        # Connect to SharePoint server
        $spServer = Get-SPServer -Identity $SharePointServer -ErrorAction Stop
        
        if (-not $spServer) {
            throw "Failed to connect to SharePoint server."
        }
        
        Write-Log -Message "Connected to SharePoint server successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            Server = $spServer
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to connect to SharePoint server: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Server = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-SharePointSiteCollection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [string]$OwnerAlias,
        
        [Parameter(Mandatory = $false)]
        [string]$SecondaryOwnerAlias,
        
        [Parameter(Mandatory = $false)]
        [string]$Template,
        
        [Parameter(Mandatory = $false)]
        [string]$ContentDatabase,
        
        [Parameter(Mandatory = $false)]
        [int]$Language,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeZone,
        
        [Parameter(Mandatory = $false)]
        [string]$Quota
    )
    
    try {
        Write-Log -Message "Creating SharePoint site collection at '$SiteUrl'..." -Level "INFO"
        
        # Check if site collection already exists
        $existingSite = Get-SPSite -Identity $SiteUrl -ErrorAction SilentlyContinue
        
        if ($existingSite) {
            throw "Site collection already exists at '$SiteUrl'."
        }
        
        # Get web application
        $webAppUrl = [System.Uri]$SiteUrl
        $webAppUrl = $webAppUrl.Scheme + "://" + $webAppUrl.Authority
        
        $webApp = Get-SPWebApplication -Identity $webAppUrl -ErrorAction SilentlyContinue
        
        if (-not $webApp) {
            throw "Web application not found at '$webAppUrl'."
        }
        
        # Get content database if specified
        $contentDb = $null
        
        if ($ContentDatabase) {
            $contentDb = Get-SPContentDatabase -Identity $ContentDatabase -ErrorAction SilentlyContinue
            
            if (-not $contentDb) {
                throw "Content database '$ContentDatabase' not found."
            }
        }
        
        # Get quota template if specified
        $quotaTemplate = $null
        
        if ($Quota) {
            $quotaTemplate = Get-SPQuotaTemplate -Identity $Quota -ErrorAction SilentlyContinue
            
            if (-not $quotaTemplate) {
                throw "Quota template '$Quota' not found."
            }
        }
        
        # Create site collection
        $newSiteParams = @{
            Url = $SiteUrl
            OwnerAlias = $OwnerAlias
            Name = $Title
            Template = $Template
            Language = $Language
            TimeZone = $TimeZone
        }
        
        if ($Description) {
            $newSiteParams.Add("Description", $Description)
        }
        
        if ($SecondaryOwnerAlias) {
            $newSiteParams.Add("SecondaryOwnerAlias", $SecondaryOwnerAlias)
        }
        
        if ($contentDb) {
            $newSiteParams.Add("ContentDatabase", $contentDb)
        }
        
        $site = New-SPSite @newSiteParams
        
        if (-not $site) {
            throw "Failed to create site collection."
        }
        
        # Apply quota template if specified
        if ($quotaTemplate) {
            $site.Quota = $quotaTemplate
            $site.Update()
        }
        
        Write-Log -Message "Site collection created successfully at '$SiteUrl'." -Level "INFO"
        
        return @{
            Status = "Success"
            Site = $site
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create site collection: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Site = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Configure-SharePointSiteCollection {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPSite]$Site
    )
    
    try {
        Write-Log -Message "Configuring SharePoint site collection..." -Level "INFO"
        
        # Get root web
        $rootWeb = $Site.RootWeb
        
        # Enable features
        Write-Log -Message "Enabling features..." -Level "INFO"
        
        # Team Collaboration Features
        Enable-SPFeature -Identity "TeamCollab" -Url $Site.Url -ErrorAction SilentlyContinue
        
        # Wiki Page Home Page
        Enable-SPFeature -Identity "WikiWelcome" -Url $Site.Url -ErrorAction SilentlyContinue
        
        # Document Library
        Enable-SPFeature -Identity "DocumentLibrary" -Url $Site.Url -ErrorAction SilentlyContinue
        
        # Configure navigation
        Write-Log -Message "Configuring navigation..." -Level "INFO"
        
        $rootWeb.Navigation.UseShared = $false
        $rootWeb.Update()
        
        Write-Log -Message "Site collection configured successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to configure site collection: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting SharePoint site collection creation process." -Level "INFO"
    
    # Connect to SharePoint server
    $connectionResult = Connect-SharePointServer -SharePointServer $SharePointServer -Credential $Credential
    
    if ($connectionResult.Status -ne "Success") {
        Write-Log -Message "Failed to connect to SharePoint server. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Create site collection
    $createResult = Create-SharePointSiteCollection -SiteUrl $SiteUrl -Title $Title -Description $Description -OwnerAlias $OwnerAlias -SecondaryOwnerAlias $SecondaryOwnerAlias -Template $Template -ContentDatabase $ContentDatabase -Language $Language -TimeZone $TimeZone -Quota $Quota
    
    if ($createResult.Status -ne "Success") {
        Write-Log -Message "Failed to create site collection. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Configure site collection
    $configureResult = Configure-SharePointSiteCollection -Site $createResult.Site
    
    if ($configureResult.Status -ne "Success") {
        Write-Log -Message "Failed to configure site collection. Exiting..." -Level "ERROR"
        exit 1
    }
    
    Write-Log -Message "SharePoint site collection creation process completed successfully." -Level "INFO"
    Write-Log -Message "Site URL: $SiteUrl" -Level "INFO"
    
    # Return site information
    return @{
        SiteUrl = $SiteUrl
        Title = $Title
        Owner = $OwnerAlias
        Template = $Template
        Status = "Success"
    }
}
catch {
    Write-Log -Message "An error occurred during SharePoint site collection creation process: $_" -Level "ERROR"
    exit 1
}
