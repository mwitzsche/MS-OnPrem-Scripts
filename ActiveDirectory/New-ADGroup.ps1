<#
.SYNOPSIS
    Creates a new security or distribution group in Active Directory.

.DESCRIPTION
    This script creates a new security or distribution group in Active Directory with the specified
    attributes and adds the specified members to the group. It provides detailed logging and error handling.

.PARAMETER Name
    Name of the group.

.PARAMETER SamAccountName
    SAM account name for the group.

.PARAMETER GroupScope
    Scope of the group (Global, Universal, DomainLocal).

.PARAMETER GroupCategory
    Category of the group (Security, Distribution).

.PARAMETER Description
    Description of the group.

.PARAMETER Path
    OU path where the group will be created.

.PARAMETER Members
    Array of user SAM account names to add as group members.

.PARAMETER MemberOf
    Array of group names to add this group as a member of.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\New-ADGroup.ps1 -Name "IT Department" -SamAccountName "IT_Dept" -GroupScope "Global" -GroupCategory "Security" -Description "IT Department Security Group" -Path "OU=Groups,DC=contoso,DC=com" -Members @("jdoe", "asmith") -MemberOf @("All Staff")

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$SamAccountName,

    [Parameter(Mandatory = $true)]
    [ValidateSet("Global", "Universal", "DomainLocal")]
    [string]$GroupScope,

    [Parameter(Mandatory = $true)]
    [ValidateSet("Security", "Distribution")]
    [string]$GroupCategory,

    [Parameter(Mandatory = $false)]
    [string]$Description,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [string[]]$Members,

    [Parameter(Mandatory = $false)]
    [string[]]$MemberOf,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ADGroup_$(Get-Date -Format 'yyyyMMdd').log"
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

function Test-ADModule {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Log -Message "ActiveDirectory module not found. Installing..." -Level "WARNING"
        try {
            Import-Module ServerManager
            Add-WindowsFeature RSAT-AD-PowerShell
            Import-Module ActiveDirectory
            Write-Log -Message "ActiveDirectory module installed successfully." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to install ActiveDirectory module: $_" -Level "ERROR"
            return $false
        }
    }
    else {
        try {
            Import-Module ActiveDirectory
            Write-Log -Message "ActiveDirectory module imported successfully." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to import ActiveDirectory module: $_" -Level "ERROR"
            return $false
        }
    }
    return $true
}

# Main script execution
try {
    Write-Log -Message "Starting group creation process for $Name ($SamAccountName)." -Level "INFO"
    
    # Check if ActiveDirectory module is available
    if (-not (Test-ADModule)) {
        Write-Log -Message "Exiting script due to missing ActiveDirectory module." -Level "ERROR"
        exit 1
    }
    
    # Check if group already exists
    if (Get-ADGroup -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue) {
        Write-Log -Message "Group with SamAccountName '$SamAccountName' already exists." -Level "ERROR"
        exit 1
    }
    
    # Check if OU exists
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$Path'" -ErrorAction SilentlyContinue)) {
        Write-Log -Message "Organizational Unit '$Path' does not exist." -Level "ERROR"
        exit 1
    }
    
    # Create group parameters
    $groupParams = @{
        Name = $Name
        SamAccountName = $SamAccountName
        GroupScope = $GroupScope
        GroupCategory = $GroupCategory
        Path = $Path
    }
    
    # Add optional parameters if provided
    if ($Description) { $groupParams.Add("Description", $Description) }
    
    # Create the group
    New-ADGroup @groupParams
    Write-Log -Message "Group '$SamAccountName' created successfully." -Level "INFO"
    
    # Add members to the group
    if ($Members -and $Members.Count -gt 0) {
        foreach ($member in $Members) {
            try {
                if (Get-ADUser -Filter "SamAccountName -eq '$member'" -ErrorAction SilentlyContinue) {
                    Add-ADGroupMember -Identity $SamAccountName -Members $member
                    Write-Log -Message "Added member '$member' to group '$SamAccountName'." -Level "INFO"
                }
                else {
                    Write-Log -Message "User '$member' not found. Cannot add to group." -Level "WARNING"
                }
            }
            catch {
                Write-Log -Message "Failed to add member '$member' to group '$SamAccountName': $_" -Level "WARNING"
            }
        }
    }
    
    # Add group to other groups
    if ($MemberOf -and $MemberOf.Count -gt 0) {
        foreach ($parentGroup in $MemberOf) {
            try {
                if (Get-ADGroup -Filter "Name -eq '$parentGroup'" -ErrorAction SilentlyContinue) {
                    Add-ADGroupMember -Identity $parentGroup -Members $SamAccountName
                    Write-Log -Message "Added group '$SamAccountName' to group '$parentGroup'." -Level "INFO"
                }
                else {
                    Write-Log -Message "Group '$parentGroup' not found. Cannot add as member." -Level "WARNING"
                }
            }
            catch {
                Write-Log -Message "Failed to add group '$SamAccountName' to group '$parentGroup': $_" -Level "WARNING"
            }
        }
    }
    
    Write-Log -Message "Group creation process completed successfully for $Name ($SamAccountName)." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during group creation: $_" -Level "ERROR"
    exit 1
}
