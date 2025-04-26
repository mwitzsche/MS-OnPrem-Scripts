<#
.SYNOPSIS
    Removes a SharePoint site collection from an on-premises SharePoint environment.

.DESCRIPTION
    This script removes a SharePoint site collection from an on-premises SharePoint environment,
    with options for backup before deletion and gradual deletion. It provides detailed logging and error handling.

.PARAMETER SiteUrl
    URL of the site collection to remove.

.PARAMETER BackupBeforeDelete
    Whether to backup the site collection before deletion.

.PARAMETER BackupPath
    Path where the site collection backup will be stored.

.PARAMETER GradualDelete
    Whether to use gradual deletion to reduce impact on production environments.

.PARAMETER DeleteADAccounts
    Whether to delete associated AD accounts.

.PARAMETER SharePointServer
    SharePoint server to connect to.

.PARAMETER Credential
    Credentials to use for SharePoint operations.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Remove-SharePointSite.ps1 -SiteUrl "https://sharepoint.contoso.com/sites/ProjectX" -BackupBeforeDelete $true -BackupPath "C:\Backups" -GradualDelete $true -SharePointServer "sharepoint.contoso.com" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [bool]$BackupBeforeDelete = $true,

    [Parameter(Mandatory = $false)]
    [string]$BackupPath,

    [Parameter(Mandatory = $false)]
    [bool]$GradualDelete = $false,

    [Parameter(Mandatory = $false)]
    [bool]$DeleteADAccounts = $false,

    [Parameter(Mandatory = $true)]
    [string]$SharePointServer,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\SharePointSiteRemoval_$(Get-Date -Format 'yyyyMMdd').log"
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

function Backup-SharePointSiteCollection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )
    
    try {
        Write-Log -Message "Backing up SharePoint site collection '$SiteUrl'..." -Level "INFO"
        
        # Create backup directory if it doesn't exist
        if (-not (Test-Path -Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created backup directory '$BackupPath'." -Level "INFO"
        }
        
        # Generate backup filename
        $siteUrlParts = $SiteUrl -split '/'
        $siteName = $siteUrlParts[-1]
        if ([string]::IsNullOrEmpty($siteName)) {
            $siteName = $siteUrlParts[-2]
        }
        
        $backupFileName = "$siteName-$(Get-Date -Format 'yyyyMMdd-HHmmss').bak"
        $backupFilePath = Join-Path -Path $BackupPath -ChildPath $backupFileName
        
        # Backup site collection
        Backup-SPSite -Identity $SiteUrl -Path $backupFilePath -Force
        
        if (-not (Test-Path -Path $backupFilePath)) {
            throw "Backup file was not created."
        }
        
        Write-Log -Message "Site collection backed up successfully to '$backupFilePath'." -Level "INFO"
        
        return @{
            Status = "Success"
            BackupPath = $backupFilePath
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to backup site collection: $_" -Level "ERROR"
        return @{
            Status = "Error"
            BackupPath = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Remove-SharePointSiteCollection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [bool]$GradualDelete,
        
        [Parameter(Mandatory = $false)]
        [bool]$DeleteADAccounts
    )
    
    try {
        Write-Log -Message "Removing SharePoint site collection '$SiteUrl'..." -Level "INFO"
        
        # Check if site collection exists
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction SilentlyContinue
        
        if (-not $site) {
            throw "Site collection not found at '$SiteUrl'."
        }
        
        # Get site collection information for reporting
        $siteInfo = @{
            Url = $site.Url
            Title = $site.RootWeb.Title
            Owner = $site.Owner.UserLogin
            Created = $site.Created
            LastContentModified = $site.LastContentModifiedDate
        }
        
        # Remove site collection
        if ($GradualDelete) {
            Write-Log -Message "Using gradual deletion to reduce impact on production environment..." -Level "INFO"
            
            # Lock the site to prevent access during deletion
            $site.Lock([Microsoft.SharePoint.SPSiteLockType]::NoAccess)
            $site.Update()
            
            # Remove site collection with gradual deletion
            Remove-SPSite -Identity $SiteUrl -GradualDelete -Confirm:$false
            
            Write-Log -Message "Site collection marked for gradual deletion." -Level "INFO"
        }
        else {
            # Remove site collection immediately
            Remove-SPSite -Identity $SiteUrl -Confirm:$false
            
            Write-Log -Message "Site collection removed immediately." -Level "INFO"
        }
        
        # Delete associated AD accounts if requested
        if ($DeleteADAccounts) {
            Write-Log -Message "Deleting associated AD accounts is not implemented in this script." -Level "WARNING"
            # This would require additional AD module and permissions
            # Implementation would depend on specific requirements and AD structure
        }
        
        Write-Log -Message "Site collection removed successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            SiteInfo = $siteInfo
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to remove site collection: $_" -Level "ERROR"
        return @{
            Status = "Error"
            SiteInfo = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting SharePoint site collection removal process." -Level "INFO"
    
    # Connect to SharePoint server
    $connectionResult = Connect-SharePointServer -SharePointServer $SharePointServer -Credential $Credential
    
    if ($connectionResult.Status -ne "Success") {
        Write-Log -Message "Failed to connect to SharePoint server. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Backup site collection if requested
    if ($BackupBeforeDelete) {
        if (-not $BackupPath) {
            Write-Log -Message "BackupPath parameter is required when BackupBeforeDelete is true. Exiting..." -Level "ERROR"
            exit 1
        }
        
        $backupResult = Backup-SharePointSiteCollection -SiteUrl $SiteUrl -BackupPath $BackupPath
        
        if ($backupResult.Status -ne "Success") {
            Write-Log -Message "Failed to backup site collection. Exiting..." -Level "ERROR"
            exit 1
        }
    }
    
    # Remove site collection
    $removeResult = Remove-SharePointSiteCollection -SiteUrl $SiteUrl -GradualDelete $GradualDelete -DeleteADAccounts $DeleteADAccounts
    
    if ($removeResult.Status -ne "Success") {
        Write-Log -Message "Failed to remove site collection. Exiting..." -Level "ERROR"
        exit 1
    }
    
    Write-Log -Message "SharePoint site collection removal process completed successfully." -Level "INFO"
    
    # Return site information
    if ($BackupBeforeDelete) {
        return @{
            SiteUrl = $SiteUrl
            BackupPath = $backupResult.BackupPath
            GradualDelete = $GradualDelete
            Status = "Success"
        }
    }
    else {
        return @{
            SiteUrl = $SiteUrl
            GradualDelete = $GradualDelete
            Status = "Success"
        }
    }
}
catch {
    Write-Log -Message "An error occurred during SharePoint site collection removal process: $_" -Level "ERROR"
    exit 1
}
