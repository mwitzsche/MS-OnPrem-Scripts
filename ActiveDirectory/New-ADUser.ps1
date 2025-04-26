<#
.SYNOPSIS
    Creates a new user in Active Directory with specified attributes and group memberships.

.DESCRIPTION
    This script creates a new user account in Active Directory with the specified attributes
    and adds the user to the specified groups. It provides detailed logging and error handling.

.PARAMETER FirstName
    First name of the user.

.PARAMETER LastName
    Last name of the user.

.PARAMETER SamAccountName
    SAM account name for the user.

.PARAMETER UserPrincipalName
    User principal name (email format) for the user.

.PARAMETER Password
    Initial password for the user account.

.PARAMETER ChangePasswordAtLogon
    Whether to force password change at next logon.

.PARAMETER Enabled
    Whether the account should be enabled.

.PARAMETER Department
    User's department.

.PARAMETER Title
    User's job title.

.PARAMETER Company
    User's company name.

.PARAMETER Path
    OU path where the user will be created.

.PARAMETER Groups
    Array of group names to add the user to.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\New-ADUser.ps1 -FirstName "John" -LastName "Doe" -SamAccountName "jdoe" -UserPrincipalName "john.doe@contoso.com" -Password (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) -ChangePasswordAtLogon $true -Enabled $true -Department "IT" -Title "System Administrator" -Company "Contoso" -Path "OU=IT,OU=Users,DC=contoso,DC=com" -Groups @("IT Staff", "Domain Admins")

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [string]$LastName,

    [Parameter(Mandatory = $true)]
    [string]$SamAccountName,

    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [SecureString]$Password,

    [Parameter(Mandatory = $false)]
    [bool]$ChangePasswordAtLogon = $true,

    [Parameter(Mandatory = $false)]
    [bool]$Enabled = $true,

    [Parameter(Mandatory = $false)]
    [string]$Department,

    [Parameter(Mandatory = $false)]
    [string]$Title,

    [Parameter(Mandatory = $false)]
    [string]$Company,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [string[]]$Groups,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ADUser_$(Get-Date -Format 'yyyyMMdd').log"
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
    Write-Log -Message "Starting user creation process for $FirstName $LastName ($SamAccountName)." -Level "INFO"
    
    # Check if ActiveDirectory module is available
    if (-not (Test-ADModule)) {
        Write-Log -Message "Exiting script due to missing ActiveDirectory module." -Level "ERROR"
        exit 1
    }
    
    # Check if user already exists
    if (Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -ErrorAction SilentlyContinue) {
        Write-Log -Message "User with SamAccountName '$SamAccountName' already exists." -Level "ERROR"
        exit 1
    }
    
    # Check if OU exists
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$Path'" -ErrorAction SilentlyContinue)) {
        Write-Log -Message "Organizational Unit '$Path' does not exist." -Level "ERROR"
        exit 1
    }
    
    # Create user parameters
    $userParams = @{
        Name = "$FirstName $LastName"
        GivenName = $FirstName
        Surname = $LastName
        SamAccountName = $SamAccountName
        UserPrincipalName = $UserPrincipalName
        AccountPassword = $Password
        ChangePasswordAtLogon = $ChangePasswordAtLogon
        Enabled = $Enabled
        Path = $Path
    }
    
    # Add optional parameters if provided
    if ($Department) { $userParams.Add("Department", $Department) }
    if ($Title) { $userParams.Add("Title", $Title) }
    if ($Company) { $userParams.Add("Company", $Company) }
    
    # Create the user
    New-ADUser @userParams
    Write-Log -Message "User '$SamAccountName' created successfully." -Level "INFO"
    
    # Add user to groups
    if ($Groups -and $Groups.Count -gt 0) {
        foreach ($group in $Groups) {
            try {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
                Write-Log -Message "Added user '$SamAccountName' to group '$group'." -Level "INFO"
            }
            catch {
                Write-Log -Message "Failed to add user '$SamAccountName' to group '$group': $_" -Level "WARNING"
            }
        }
    }
    
    Write-Log -Message "User creation process completed successfully for $FirstName $LastName ($SamAccountName)." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during user creation: $_" -Level "ERROR"
    exit 1
}
