<#
.SYNOPSIS
    Creates a new organizational unit in Active Directory with optional nested OUs.

.DESCRIPTION
    This script creates a new organizational unit in Active Directory with the specified attributes
    and optionally creates nested OUs within it. It provides detailed logging and error handling.

.PARAMETER Name
    Name of the organizational unit.

.PARAMETER Path
    Parent path where the OU will be created.

.PARAMETER Description
    Description of the OU.

.PARAMETER ProtectedFromAccidentalDeletion
    Whether the OU is protected from accidental deletion.

.PARAMETER NestedOUs
    Array of nested OUs to create within this OU. Each nested OU should be a hashtable with Name, Description, and Protected keys.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    $nestedOUs = @(
        @{Name="Users"; Description="Department Users"; Protected=$true},
        @{Name="Computers"; Description="Department Computers"; Protected=$true},
        @{Name="Groups"; Description="Department Groups"; Protected=$true}
    )
    .\New-ADOrganizationalUnit.ps1 -Name "IT" -Path "DC=contoso,DC=com" -Description "IT Department" -ProtectedFromAccidentalDeletion $true -NestedOUs $nestedOUs

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
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [string]$Description,

    [Parameter(Mandatory = $false)]
    [bool]$ProtectedFromAccidentalDeletion = $true,

    [Parameter(Mandatory = $false)]
    [hashtable[]]$NestedOUs,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ADOU_$(Get-Date -Format 'yyyyMMdd').log"
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
    Write-Log -Message "Starting organizational unit creation process for $Name." -Level "INFO"
    
    # Check if ActiveDirectory module is available
    if (-not (Test-ADModule)) {
        Write-Log -Message "Exiting script due to missing ActiveDirectory module." -Level "ERROR"
        exit 1
    }
    
    # Check if parent path exists
    if (-not (Get-ADObject -Filter "DistinguishedName -eq '$Path'" -ErrorAction SilentlyContinue)) {
        Write-Log -Message "Parent path '$Path' does not exist." -Level "ERROR"
        exit 1
    }
    
    # Check if OU already exists
    $ouDN = "OU=$Name,$Path"
    if (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ouDN'" -ErrorAction SilentlyContinue) {
        Write-Log -Message "Organizational Unit with DN '$ouDN' already exists." -Level "ERROR"
        exit 1
    }
    
    # Create OU parameters
    $ouParams = @{
        Name = $Name
        Path = $Path
        ProtectedFromAccidentalDeletion = $ProtectedFromAccidentalDeletion
    }
    
    # Add optional parameters if provided
    if ($Description) { $ouParams.Add("Description", $Description) }
    
    # Create the OU
    New-ADOrganizationalUnit @ouParams
    Write-Log -Message "Organizational Unit '$Name' created successfully at path '$Path'." -Level "INFO"
    
    # Create nested OUs if specified
    if ($NestedOUs -and $NestedOUs.Count -gt 0) {
        foreach ($nestedOU in $NestedOUs) {
            try {
                $nestedOUParams = @{
                    Name = $nestedOU.Name
                    Path = $ouDN
                    ProtectedFromAccidentalDeletion = $nestedOU.Protected -eq $true
                }
                
                if ($nestedOU.Description) { 
                    $nestedOUParams.Add("Description", $nestedOU.Description) 
                }
                
                New-ADOrganizationalUnit @nestedOUParams
                Write-Log -Message "Nested Organizational Unit '$($nestedOU.Name)' created successfully under '$ouDN'." -Level "INFO"
            }
            catch {
                Write-Log -Message "Failed to create nested Organizational Unit '$($nestedOU.Name)': $_" -Level "WARNING"
            }
        }
    }
    
    Write-Log -Message "Organizational Unit creation process completed successfully for $Name." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during organizational unit creation: $_" -Level "ERROR"
    exit 1
}
