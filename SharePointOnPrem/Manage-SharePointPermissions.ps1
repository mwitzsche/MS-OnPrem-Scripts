<#
.SYNOPSIS
    Manages permissions for SharePoint sites in an on-premises SharePoint environment.

.DESCRIPTION
    This script manages permissions for SharePoint sites in an on-premises SharePoint environment,
    including setting permissions, creating permission levels, and managing site access.
    It provides detailed logging and error handling.

.PARAMETER Action
    Action to perform (Get, Set, Add, Remove, CreateLevel, RemoveLevel).

.PARAMETER SiteUrl
    URL of the site to manage permissions for.

.PARAMETER UserOrGroup
    User or group to manage permissions for.

.PARAMETER PermissionLevel
    Permission level to assign (e.g., Full Control, Contribute, Read).

.PARAMETER CustomPermissionLevel
    Name of custom permission level to create.

.PARAMETER BasePermissionLevel
    Base permission level to use when creating a custom permission level.

.PARAMETER Permissions
    Array of permissions to include in a custom permission level.

.PARAMETER Recursive
    Whether to apply permissions recursively to all subsites and lists.

.PARAMETER SharePointServer
    SharePoint server to connect to.

.PARAMETER Credential
    Credentials to use for SharePoint operations.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Manage-SharePointPermissions.ps1 -Action Set -SiteUrl "https://sharepoint.contoso.com/sites/ProjectX" -UserOrGroup "contoso\ProjectX_Members" -PermissionLevel "Contribute" -Recursive $true -SharePointServer "sharepoint.contoso.com" -Credential (Get-Credential)

.EXAMPLE
    .\Manage-SharePointPermissions.ps1 -Action CreateLevel -SiteUrl "https://sharepoint.contoso.com/sites/ProjectX" -CustomPermissionLevel "Custom Editor" -BasePermissionLevel "Contribute" -Permissions @("ManageListPermissions", "ManagePermissions") -SharePointServer "sharepoint.contoso.com" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Get", "Set", "Add", "Remove", "CreateLevel", "RemoveLevel")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$UserOrGroup,

    [Parameter(Mandatory = $false)]
    [string]$PermissionLevel,

    [Parameter(Mandatory = $false)]
    [string]$CustomPermissionLevel,

    [Parameter(Mandatory = $false)]
    [string]$BasePermissionLevel,

    [Parameter(Mandatory = $false)]
    [string[]]$Permissions,

    [Parameter(Mandatory = $false)]
    [bool]$Recursive = $false,

    [Parameter(Mandatory = $true)]
    [string]$SharePointServer,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\SharePointPermissions_$(Get-Date -Format 'yyyyMMdd').log"
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

function Get-SharePointSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log -Message "Getting permissions for site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        $results = @{
            SiteUrl = $SiteUrl
            Status = "Success"
            Permissions = @()
            ErrorMessage = $null
        }
        
        # Get permission levels
        $permissionLevels = $web.RoleDefinitions
        
        Write-Log -Message "Found $($permissionLevels.Count) permission levels." -Level "INFO"
        
        # Get role assignments
        $roleAssignments = $web.RoleAssignments
        
        foreach ($roleAssignment in $roleAssignments) {
            $member = $roleAssignment.Member
            $memberType = $member.GetType().Name
            
            foreach ($roleDefinitionBinding in $roleAssignment.RoleDefinitionBindings) {
                $results.Permissions += @{
                    UserOrGroup = $member.Name
                    UserOrGroupType = $memberType
                    PermissionLevel = $roleDefinitionBinding.Name
                    Description = $roleDefinitionBinding.Description
                }
            }
        }
        
        Write-Log -Message "Found $($results.Permissions.Count) permission assignments." -Level "INFO"
        
        # If recursive, get permissions for subsites
        if ($Recursive) {
            Write-Log -Message "Getting permissions for subsites..." -Level "INFO"
            
            foreach ($subweb in $web.Webs) {
                Write-Log -Message "Processing subsite '$($subweb.Url)'..." -Level "INFO"
                
                $subwebRoleAssignments = $subweb.RoleAssignments
                
                foreach ($roleAssignment in $subwebRoleAssignments) {
                    $member = $roleAssignment.Member
                    $memberType = $member.GetType().Name
                    
                    foreach ($roleDefinitionBinding in $roleAssignment.RoleDefinitionBindings) {
                        $results.Permissions += @{
                            SiteUrl = $subweb.Url
                            UserOrGroup = $member.Name
                            UserOrGroupType = $memberType
                            PermissionLevel = $roleDefinitionBinding.Name
                            Description = $roleDefinitionBinding.Description
                        }
                    }
                }
                
                $subweb.Dispose()
            }
        }
        
        $web.Dispose()
        $site.Dispose()
        
        Write-Log -Message "Permissions retrieved successfully." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to get site permissions: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            Permissions = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Set-SharePointSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$UserOrGroup,
        
        [Parameter(Mandatory = $true)]
        [string]$PermissionLevel,
        
        [Parameter(Mandatory = $false)]
        [bool]$Recursive
    )
    
    try {
        Write-Log -Message "Setting permissions for site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        # Check if permission level exists
        $permLevel = $web.RoleDefinitions[$PermissionLevel]
        
        if (-not $permLevel) {
            throw "Permission level '$PermissionLevel' not found."
        }
        
        # Check if user or group exists
        $user = $null
        
        try {
            $user = $web.EnsureUser($UserOrGroup)
        }
        catch {
            try {
                $user = $web.SiteGroups[$UserOrGroup]
            }
            catch {
                throw "User or group '$UserOrGroup' not found."
            }
        }
        
        if (-not $user) {
            throw "User or group '$UserOrGroup' not found."
        }
        
        # Remove existing permissions
        $assignment = $web.RoleAssignments.GetAssignmentByPrincipal($user)
        
        if ($assignment) {
            $web.RoleAssignments.Remove($user)
        }
        
        # Add new permission
        $assignment = New-Object Microsoft.SharePoint.SPRoleAssignment($user)
        $assignment.RoleDefinitionBindings.Add($permLevel)
        $web.RoleAssignments.Add($assignment)
        
        Write-Log -Message "Permissions set successfully for '$UserOrGroup' with level '$PermissionLevel'." -Level "INFO"
        
        # If recursive, set permissions for subsites
        if ($Recursive) {
            Write-Log -Message "Setting permissions for subsites..." -Level "INFO"
            
            foreach ($subweb in $web.Webs) {
                Write-Log -Message "Processing subsite '$($subweb.Url)'..." -Level "INFO"
                
                # Check if permission level exists in subsite
                $subPermLevel = $subweb.RoleDefinitions[$PermissionLevel]
                
                if (-not $subPermLevel) {
                    Write-Log -Message "Permission level '$PermissionLevel' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                    continue
                }
                
                # Check if user or group exists in subsite
                $subUser = $null
                
                try {
                    $subUser = $subweb.EnsureUser($UserOrGroup)
                }
                catch {
                    try {
                        $subUser = $subweb.SiteGroups[$UserOrGroup]
                    }
                    catch {
                        Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                        continue
                    }
                }
                
                if (-not $subUser) {
                    Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                    continue
                }
                
                # Remove existing permissions
                $subAssignment = $subweb.RoleAssignments.GetAssignmentByPrincipal($subUser)
                
                if ($subAssignment) {
                    $subweb.RoleAssignments.Remove($subUser)
                }
                
                # Add new permission
                $subAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($subUser)
                $subAssignment.RoleDefinitionBindings.Add($subPermLevel)
                $subweb.RoleAssignments.Add($subAssignment)
                
                Write-Log -Message "Permissions set successfully for '$UserOrGroup' with level '$PermissionLevel' in subsite '$($subweb.Url)'." -Level "INFO"
                
                $subweb.Dispose()
            }
        }
        
        $web.Dispose()
        $site.Dispose()
        
        return @{
            SiteUrl = $SiteUrl
            Status = "Success"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to set site permissions: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Add-SharePointSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$UserOrGroup,
        
        [Parameter(Mandatory = $true)]
        [string]$PermissionLevel,
        
        [Parameter(Mandatory = $false)]
        [bool]$Recursive
    )
    
    try {
        Write-Log -Message "Adding permissions for site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        # Check if permission level exists
        $permLevel = $web.RoleDefinitions[$PermissionLevel]
        
        if (-not $permLevel) {
            throw "Permission level '$PermissionLevel' not found."
        }
        
        # Check if user or group exists
        $user = $null
        
        try {
            $user = $web.EnsureUser($UserOrGroup)
        }
        catch {
            try {
                $user = $web.SiteGroups[$UserOrGroup]
            }
            catch {
                throw "User or group '$UserOrGroup' not found."
            }
        }
        
        if (-not $user) {
            throw "User or group '$UserOrGroup' not found."
        }
        
        # Get existing assignment or create new one
        $assignment = $web.RoleAssignments.GetAssignmentByPrincipal($user)
        
        if (-not $assignment) {
            $assignment = New-Object Microsoft.SharePoint.SPRoleAssignment($user)
            $web.RoleAssignments.Add($assignment)
        }
        
        # Add permission level if not already assigned
        if (-not $assignment.RoleDefinitionBindings.Contains($permLevel)) {
            $assignment.RoleDefinitionBindings.Add($permLevel)
            Write-Log -Message "Permission level '$PermissionLevel' added for '$UserOrGroup'." -Level "INFO"
        }
        else {
            Write-Log -Message "Permission level '$PermissionLevel' already assigned to '$UserOrGroup'." -Level "INFO"
        }
        
        # If recursive, add permissions for subsites
        if ($Recursive) {
            Write-Log -Message "Adding permissions for subsites..." -Level "INFO"
            
            foreach ($subweb in $web.Webs) {
                Write-Log -Message "Processing subsite '$($subweb.Url)'..." -Level "INFO"
                
                # Check if permission level exists in subsite
                $subPermLevel = $subweb.RoleDefinitions[$PermissionLevel]
                
                if (-not $subPermLevel) {
                    Write-Log -Message "Permission level '$PermissionLevel' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                    continue
                }
                
                # Check if user or group exists in subsite
                $subUser = $null
                
                try {
                    $subUser = $subweb.EnsureUser($UserOrGroup)
                }
                catch {
                    try {
                        $subUser = $subweb.SiteGroups[$UserOrGroup]
                    }
                    catch {
                        Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                        continue
                    }
                }
                
                if (-not $subUser) {
                    Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                    continue
                }
                
                # Get existing assignment or create new one
                $subAssignment = $subweb.RoleAssignments.GetAssignmentByPrincipal($subUser)
                
                if (-not $subAssignment) {
                    $subAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($subUser)
                    $subweb.RoleAssignments.Add($subAssignment)
                }
                
                # Add permission level if not already assigned
                if (-not $subAssignment.RoleDefinitionBindings.Contains($subPermLevel)) {
                    $subAssignment.RoleDefinitionBindings.Add($subPermLevel)
                    Write-Log -Message "Permission level '$PermissionLevel' added for '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                }
                else {
                    Write-Log -Message "Permission level '$PermissionLevel' already assigned to '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                }
                
                $subweb.Dispose()
            }
        }
        
        $web.Dispose()
        $site.Dispose()
        
        return @{
            SiteUrl = $SiteUrl
            Status = "Success"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to add site permissions: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Remove-SharePointSitePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$UserOrGroup,
        
        [Parameter(Mandatory = $false)]
        [string]$PermissionLevel,
        
        [Parameter(Mandatory = $false)]
        [bool]$Recursive
    )
    
    try {
        Write-Log -Message "Removing permissions for site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        # Check if user or group exists
        $user = $null
        
        try {
            $user = $web.EnsureUser($UserOrGroup)
        }
        catch {
            try {
                $user = $web.SiteGroups[$UserOrGroup]
            }
            catch {
                throw "User or group '$UserOrGroup' not found."
            }
        }
        
        if (-not $user) {
            throw "User or group '$UserOrGroup' not found."
        }
        
        # Get existing assignment
        $assignment = $web.RoleAssignments.GetAssignmentByPrincipal($user)
        
        if (-not $assignment) {
            Write-Log -Message "No permissions found for '$UserOrGroup'." -Level "INFO"
        }
        else {
            if ($PermissionLevel) {
                # Check if permission level exists
                $permLevel = $web.RoleDefinitions[$PermissionLevel]
                
                if (-not $permLevel) {
                    throw "Permission level '$PermissionLevel' not found."
                }
                
                # Remove specific permission level
                if ($assignment.RoleDefinitionBindings.Contains($permLevel)) {
                    $assignment.RoleDefinitionBindings.Remove($permLevel)
                    Write-Log -Message "Permission level '$PermissionLevel' removed for '$UserOrGroup'." -Level "INFO"
                }
                else {
                    Write-Log -Message "Permission level '$PermissionLevel' not assigned to '$UserOrGroup'." -Level "INFO"
                }
                
                # If no permissions left, remove assignment
                if ($assignment.RoleDefinitionBindings.Count -eq 0) {
                    $web.RoleAssignments.Remove($user)
                    Write-Log -Message "All permissions removed for '$UserOrGroup'." -Level "INFO"
                }
            }
            else {
                # Remove all permissions
                $web.RoleAssignments.Remove($user)
                Write-Log -Message "All permissions removed for '$UserOrGroup'." -Level "INFO"
            }
        }
        
        # If recursive, remove permissions for subsites
        if ($Recursive) {
            Write-Log -Message "Removing permissions for subsites..." -Level "INFO"
            
            foreach ($subweb in $web.Webs) {
                Write-Log -Message "Processing subsite '$($subweb.Url)'..." -Level "INFO"
                
                # Check if user or group exists in subsite
                $subUser = $null
                
                try {
                    $subUser = $subweb.EnsureUser($UserOrGroup)
                }
                catch {
                    try {
                        $subUser = $subweb.SiteGroups[$UserOrGroup]
                    }
                    catch {
                        Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                        continue
                    }
                }
                
                if (-not $subUser) {
                    Write-Log -Message "User or group '$UserOrGroup' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                    continue
                }
                
                # Get existing assignment
                $subAssignment = $subweb.RoleAssignments.GetAssignmentByPrincipal($subUser)
                
                if (-not $subAssignment) {
                    Write-Log -Message "No permissions found for '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                }
                else {
                    if ($PermissionLevel) {
                        # Check if permission level exists in subsite
                        $subPermLevel = $subweb.RoleDefinitions[$PermissionLevel]
                        
                        if (-not $subPermLevel) {
                            Write-Log -Message "Permission level '$PermissionLevel' not found in subsite '$($subweb.Url)'. Skipping..." -Level "WARNING"
                            continue
                        }
                        
                        # Remove specific permission level
                        if ($subAssignment.RoleDefinitionBindings.Contains($subPermLevel)) {
                            $subAssignment.RoleDefinitionBindings.Remove($subPermLevel)
                            Write-Log -Message "Permission level '$PermissionLevel' removed for '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                        }
                        else {
                            Write-Log -Message "Permission level '$PermissionLevel' not assigned to '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                        }
                        
                        # If no permissions left, remove assignment
                        if ($subAssignment.RoleDefinitionBindings.Count -eq 0) {
                            $subweb.RoleAssignments.Remove($subUser)
                            Write-Log -Message "All permissions removed for '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                        }
                    }
                    else {
                        # Remove all permissions
                        $subweb.RoleAssignments.Remove($subUser)
                        Write-Log -Message "All permissions removed for '$UserOrGroup' in subsite '$($subweb.Url)'." -Level "INFO"
                    }
                }
                
                $subweb.Dispose()
            }
        }
        
        $web.Dispose()
        $site.Dispose()
        
        return @{
            SiteUrl = $SiteUrl
            Status = "Success"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to remove site permissions: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            UserOrGroup = $UserOrGroup
            PermissionLevel = $PermissionLevel
            Recursive = $Recursive
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Create-SharePointPermissionLevel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$CustomPermissionLevel,
        
        [Parameter(Mandatory = $false)]
        [string]$BasePermissionLevel,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Permissions
    )
    
    try {
        Write-Log -Message "Creating permission level '$CustomPermissionLevel' for site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        # Check if permission level already exists
        if ($web.RoleDefinitions[$CustomPermissionLevel]) {
            throw "Permission level '$CustomPermissionLevel' already exists."
        }
        
        # Create new permission level
        $permLevel = New-Object Microsoft.SharePoint.SPRoleDefinition
        $permLevel.Name = $CustomPermissionLevel
        $permLevel.Description = "Custom permission level created by script"
        
        # Set base permissions
        if ($BasePermissionLevel) {
            $basePermLevel = $web.RoleDefinitions[$BasePermissionLevel]
            
            if (-not $basePermLevel) {
                throw "Base permission level '$BasePermissionLevel' not found."
            }
            
            $permLevel.BasePermissions = $basePermLevel.BasePermissions
            Write-Log -Message "Using base permission level '$BasePermissionLevel'." -Level "INFO"
        }
        else {
            $permLevel.BasePermissions = [Microsoft.SharePoint.SPBasePermissions]::EmptyMask
        }
        
        # Add additional permissions
        if ($Permissions -and $Permissions.Count -gt 0) {
            foreach ($permission in $Permissions) {
                try {
                    $permEnum = [Microsoft.SharePoint.SPBasePermissions]::$permission
                    $permLevel.BasePermissions = $permLevel.BasePermissions -bor $permEnum
                    Write-Log -Message "Added permission '$permission'." -Level "INFO"
                }
                catch {
                    Write-Log -Message "Invalid permission '$permission'. Skipping..." -Level "WARNING"
                }
            }
        }
        
        # Add permission level to site
        $web.RoleDefinitions.Add($permLevel)
        
        Write-Log -Message "Permission level '$CustomPermissionLevel' created successfully." -Level "INFO"
        
        $web.Dispose()
        $site.Dispose()
        
        return @{
            SiteUrl = $SiteUrl
            Status = "Success"
            PermissionLevel = $CustomPermissionLevel
            BasePermissionLevel = $BasePermissionLevel
            Permissions = $Permissions
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create permission level: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            PermissionLevel = $CustomPermissionLevel
            BasePermissionLevel = $BasePermissionLevel
            Permissions = $Permissions
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Remove-SharePointPermissionLevel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$CustomPermissionLevel
    )
    
    try {
        Write-Log -Message "Removing permission level '$CustomPermissionLevel' from site '$SiteUrl'..." -Level "INFO"
        
        # Get site
        $site = Get-SPSite -Identity $SiteUrl -ErrorAction Stop
        $web = $site.RootWeb
        
        # Check if permission level exists
        $permLevel = $web.RoleDefinitions[$CustomPermissionLevel]
        
        if (-not $permLevel) {
            throw "Permission level '$CustomPermissionLevel' not found."
        }
        
        # Check if permission level is in use
        $inUse = $false
        
        foreach ($roleAssignment in $web.RoleAssignments) {
            if ($roleAssignment.RoleDefinitionBindings.Contains($permLevel)) {
                $inUse = $true
                break
            }
        }
        
        if ($inUse) {
            throw "Permission level '$CustomPermissionLevel' is in use and cannot be removed."
        }
        
        # Remove permission level
        $web.RoleDefinitions.Delete($permLevel.Id)
        
        Write-Log -Message "Permission level '$CustomPermissionLevel' removed successfully." -Level "INFO"
        
        $web.Dispose()
        $site.Dispose()
        
        return @{
            SiteUrl = $SiteUrl
            Status = "Success"
            PermissionLevel = $CustomPermissionLevel
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to remove permission level: $_" -Level "ERROR"
        return @{
            SiteUrl = $SiteUrl
            Status = "Error"
            PermissionLevel = $CustomPermissionLevel
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting SharePoint permissions management process." -Level "INFO"
    
    # Connect to SharePoint server
    $connectionResult = Connect-SharePointServer -SharePointServer $SharePointServer -Credential $Credential
    
    if ($connectionResult.Status -ne "Success") {
        Write-Log -Message "Failed to connect to SharePoint server. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Perform the requested action
    switch ($Action) {
        "Get" {
            $result = Get-SharePointSitePermissions -SiteUrl $SiteUrl
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to get site permissions. Exiting..." -Level "ERROR"
                exit 1
            }
            
            # Output permissions
            Write-Log -Message "Permissions for site '$SiteUrl':" -Level "INFO"
            
            foreach ($permission in $result.Permissions) {
                Write-Log -Message "User/Group: $($permission.UserOrGroup), Type: $($permission.UserOrGroupType), Permission: $($permission.PermissionLevel)" -Level "INFO"
            }
        }
        "Set" {
            # Validate required parameters
            if (-not $UserOrGroup) {
                Write-Log -Message "UserOrGroup parameter is required for Set action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $PermissionLevel) {
                Write-Log -Message "PermissionLevel parameter is required for Set action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Set-SharePointSitePermissions -SiteUrl $SiteUrl -UserOrGroup $UserOrGroup -PermissionLevel $PermissionLevel -Recursive $Recursive
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to set site permissions. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Add" {
            # Validate required parameters
            if (-not $UserOrGroup) {
                Write-Log -Message "UserOrGroup parameter is required for Add action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $PermissionLevel) {
                Write-Log -Message "PermissionLevel parameter is required for Add action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Add-SharePointSitePermissions -SiteUrl $SiteUrl -UserOrGroup $UserOrGroup -PermissionLevel $PermissionLevel -Recursive $Recursive
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to add site permissions. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Remove" {
            # Validate required parameters
            if (-not $UserOrGroup) {
                Write-Log -Message "UserOrGroup parameter is required for Remove action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Remove-SharePointSitePermissions -SiteUrl $SiteUrl -UserOrGroup $UserOrGroup -PermissionLevel $PermissionLevel -Recursive $Recursive
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to remove site permissions. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "CreateLevel" {
            # Validate required parameters
            if (-not $CustomPermissionLevel) {
                Write-Log -Message "CustomPermissionLevel parameter is required for CreateLevel action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Create-SharePointPermissionLevel -SiteUrl $SiteUrl -CustomPermissionLevel $CustomPermissionLevel -BasePermissionLevel $BasePermissionLevel -Permissions $Permissions
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to create permission level. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "RemoveLevel" {
            # Validate required parameters
            if (-not $CustomPermissionLevel) {
                Write-Log -Message "CustomPermissionLevel parameter is required for RemoveLevel action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Remove-SharePointPermissionLevel -SiteUrl $SiteUrl -CustomPermissionLevel $CustomPermissionLevel
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to remove permission level. Exiting..." -Level "ERROR"
                exit 1
            }
        }
    }
    
    Write-Log -Message "SharePoint permissions management process completed successfully." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during SharePoint permissions management process: $_" -Level "ERROR"
    exit 1
}
