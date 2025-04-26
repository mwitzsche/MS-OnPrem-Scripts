<#
.SYNOPSIS
    Creates and manages Exchange mailboxes in an on-premises Exchange environment.

.DESCRIPTION
    This script creates and manages Exchange mailboxes in an on-premises Exchange environment,
    including creating new mailboxes, configuring mailbox settings, and managing mailbox permissions.
    It provides detailed logging and error handling.

.PARAMETER Action
    Action to perform (Create, Configure, Enable, Disable, Remove, SetPermission).

.PARAMETER Identity
    Identity of the mailbox or user.

.PARAMETER Name
    Display name for the new mailbox.

.PARAMETER Alias
    Email alias for the new mailbox.

.PARAMETER Database
    Exchange database where the mailbox will be created.

.PARAMETER Password
    Password for the new mailbox user.

.PARAMETER FirstName
    First name of the mailbox user.

.PARAMETER LastName
    Last name of the mailbox user.

.PARAMETER OrganizationalUnit
    Organizational Unit where the user account will be created.

.PARAMETER EmailAddresses
    Additional email addresses for the mailbox.

.PARAMETER QuotaInGB
    Mailbox quota size in GB.

.PARAMETER ArchiveEnabled
    Whether to enable archive for the mailbox.

.PARAMETER ArchiveQuotaInGB
    Archive mailbox quota size in GB.

.PARAMETER AccessRights
    Access rights to grant to a user.

.PARAMETER User
    User to grant access rights to.

.PARAMETER ExchangeServer
    Exchange server to connect to.

.PARAMETER Credential
    Credentials to use for Exchange operations.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Manage-ExchangeMailbox.ps1 -Action Create -Identity "john.doe" -Name "John Doe" -Alias "john.doe" -Database "Mailbox Database 01" -Password (ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force) -FirstName "John" -LastName "Doe" -OrganizationalUnit "OU=Users,DC=contoso,DC=com" -EmailAddresses "john.doe@contoso.com" -QuotaInGB 2 -ArchiveEnabled $true -ArchiveQuotaInGB 10 -ExchangeServer "exchange01.contoso.com" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Create", "Configure", "Enable", "Disable", "Remove", "SetPermission")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$Identity,

    [Parameter(Mandatory = $false)]
    [string]$Name,

    [Parameter(Mandatory = $false)]
    [string]$Alias,

    [Parameter(Mandatory = $false)]
    [string]$Database,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$Password,

    [Parameter(Mandatory = $false)]
    [string]$FirstName,

    [Parameter(Mandatory = $false)]
    [string]$LastName,

    [Parameter(Mandatory = $false)]
    [string]$OrganizationalUnit,

    [Parameter(Mandatory = $false)]
    [string[]]$EmailAddresses,

    [Parameter(Mandatory = $false)]
    [int]$QuotaInGB,

    [Parameter(Mandatory = $false)]
    [bool]$ArchiveEnabled = $false,

    [Parameter(Mandatory = $false)]
    [int]$ArchiveQuotaInGB,

    [Parameter(Mandatory = $false)]
    [ValidateSet("FullAccess", "SendAs", "SendOnBehalf")]
    [string]$AccessRights,

    [Parameter(Mandatory = $false)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$ExchangeServer,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ExchangeMailbox_$(Get-Date -Format 'yyyyMMdd').log"
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

function Connect-ExchangeServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log -Message "Connecting to Exchange server '$ExchangeServer'..." -Level "INFO"
        
        # Create a new PowerShell session to the Exchange server
        $sessionParams = @{
            ConfigurationName = "Microsoft.Exchange"
            ConnectionUri = "http://$ExchangeServer/PowerShell/"
            Authentication = "Kerberos"
            Credential = $Credential
        }
        
        $session = New-PSSession @sessionParams
        
        # Import the Exchange cmdlets
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        
        Write-Log -Message "Connected to Exchange server successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            Session = $session
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to connect to Exchange server: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Session = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Disconnect-ExchangeServer {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$Session
    )
    
    try {
        Write-Log -Message "Disconnecting from Exchange server..." -Level "INFO"
        
        # Remove the PowerShell session
        Remove-PSSession -Session $Session
        
        Write-Log -Message "Disconnected from Exchange server successfully." -Level "INFO"
        
        return $true
    }
    catch {
        Write-Log -Message "Failed to disconnect from Exchange server: $_" -Level "WARNING"
        return $false
    }
}

function Create-Mailbox {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$Alias,
        
        [Parameter(Mandatory = $true)]
        [string]$Database,
        
        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$FirstName,
        
        [Parameter(Mandatory = $false)]
        [string]$LastName,
        
        [Parameter(Mandatory = $false)]
        [string]$OrganizationalUnit,
        
        [Parameter(Mandatory = $false)]
        [string[]]$EmailAddresses,
        
        [Parameter(Mandatory = $false)]
        [int]$QuotaInGB,
        
        [Parameter(Mandatory = $false)]
        [bool]$ArchiveEnabled,
        
        [Parameter(Mandatory = $false)]
        [int]$ArchiveQuotaInGB
    )
    
    try {
        Write-Log -Message "Creating mailbox for user '$Identity'..." -Level "INFO"
        
        # Create new mailbox
        $mailboxParams = @{
            Name = $Name
            Alias = $Alias
            UserPrincipalName = $Identity
            Database = $Database
            Password = $Password
        }
        
        if ($FirstName) {
            $mailboxParams.Add("FirstName", $FirstName)
        }
        
        if ($LastName) {
            $mailboxParams.Add("LastName", $LastName)
        }
        
        if ($OrganizationalUnit) {
            $mailboxParams.Add("OrganizationalUnit", $OrganizationalUnit)
        }
        
        $mailbox = New-Mailbox @mailboxParams
        
        if (-not $mailbox) {
            throw "Failed to create mailbox."
        }
        
        Write-Log -Message "Mailbox created successfully." -Level "INFO"
        
        # Configure additional settings
        if ($EmailAddresses -and $EmailAddresses.Count -gt 0) {
            Write-Log -Message "Configuring email addresses..." -Level "INFO"
            
            $emailAddressesValue = $EmailAddresses | ForEach-Object { "smtp:$_" }
            Set-Mailbox -Identity $Identity -EmailAddresses @{Add = $emailAddressesValue}
            
            Write-Log -Message "Email addresses configured successfully." -Level "INFO"
        }
        
        if ($QuotaInGB -gt 0) {
            Write-Log -Message "Setting mailbox quota..." -Level "INFO"
            
            $quotaValue = $QuotaInGB * 1GB
            Set-Mailbox -Identity $Identity -ProhibitSendQuota $quotaValue -ProhibitSendReceiveQuota ($quotaValue * 1.1) -IssueWarningQuota ($quotaValue * 0.9)
            
            Write-Log -Message "Mailbox quota set successfully." -Level "INFO"
        }
        
        if ($ArchiveEnabled) {
            Write-Log -Message "Enabling archive mailbox..." -Level "INFO"
            
            Enable-Mailbox -Identity $Identity -Archive
            
            if ($ArchiveQuotaInGB -gt 0) {
                $archiveQuotaValue = $ArchiveQuotaInGB * 1GB
                Set-Mailbox -Identity $Identity -ArchiveQuota $archiveQuotaValue -ArchiveWarningQuota ($archiveQuotaValue * 0.9)
            }
            
            Write-Log -Message "Archive mailbox enabled successfully." -Level "INFO"
        }
        
        return @{
            Status = "Success"
            Mailbox = $mailbox
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create mailbox: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Mailbox = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Configure-Mailbox {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $false)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [string]$Alias,
        
        [Parameter(Mandatory = $false)]
        [string[]]$EmailAddresses,
        
        [Parameter(Mandatory = $false)]
        [int]$QuotaInGB,
        
        [Parameter(Mandatory = $false)]
        [bool]$ArchiveEnabled,
        
        [Parameter(Mandatory = $false)]
        [int]$ArchiveQuotaInGB
    )
    
    try {
        Write-Log -Message "Configuring mailbox for user '$Identity'..." -Level "INFO"
        
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        
        if (-not $mailbox) {
            throw "Mailbox not found."
        }
        
        # Configure mailbox settings
        $mailboxParams = @{
            Identity = $Identity
        }
        
        $settingsChanged = $false
        
        if ($Name) {
            $mailboxParams.Add("Name", $Name)
            $settingsChanged = $true
        }
        
        if ($Alias) {
            $mailboxParams.Add("Alias", $Alias)
            $settingsChanged = $true
        }
        
        if ($settingsChanged) {
            Set-Mailbox @mailboxParams
            Write-Log -Message "Mailbox settings updated successfully." -Level "INFO"
        }
        
        # Configure email addresses
        if ($EmailAddresses -and $EmailAddresses.Count -gt 0) {
            Write-Log -Message "Configuring email addresses..." -Level "INFO"
            
            $emailAddressesValue = $EmailAddresses | ForEach-Object { "smtp:$_" }
            Set-Mailbox -Identity $Identity -EmailAddresses @{Add = $emailAddressesValue}
            
            Write-Log -Message "Email addresses configured successfully." -Level "INFO"
        }
        
        # Configure quota
        if ($QuotaInGB -gt 0) {
            Write-Log -Message "Setting mailbox quota..." -Level "INFO"
            
            $quotaValue = $QuotaInGB * 1GB
            Set-Mailbox -Identity $Identity -ProhibitSendQuota $quotaValue -ProhibitSendReceiveQuota ($quotaValue * 1.1) -IssueWarningQuota ($quotaValue * 0.9)
            
            Write-Log -Message "Mailbox quota set successfully." -Level "INFO"
        }
        
        # Configure archive
        if ($ArchiveEnabled) {
            $archiveMailbox = Get-Mailbox -Identity $Identity -Archive -ErrorAction SilentlyContinue
            
            if (-not $archiveMailbox) {
                Write-Log -Message "Enabling archive mailbox..." -Level "INFO"
                
                Enable-Mailbox -Identity $Identity -Archive
                
                Write-Log -Message "Archive mailbox enabled successfully." -Level "INFO"
            }
            
            if ($ArchiveQuotaInGB -gt 0) {
                Write-Log -Message "Setting archive mailbox quota..." -Level "INFO"
                
                $archiveQuotaValue = $ArchiveQuotaInGB * 1GB
                Set-Mailbox -Identity $Identity -ArchiveQuota $archiveQuotaValue -ArchiveWarningQuota ($archiveQuotaValue * 0.9)
                
                Write-Log -Message "Archive mailbox quota set successfully." -Level "INFO"
            }
        }
        
        return @{
            Status = "Success"
            Mailbox = $mailbox
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to configure mailbox: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Mailbox = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Enable-MailboxUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$Database
    )
    
    try {
        Write-Log -Message "Enabling mailbox for user '$Identity'..." -Level "INFO"
        
        # Enable mailbox
        $mailbox = Enable-Mailbox -Identity $Identity -Database $Database
        
        if (-not $mailbox) {
            throw "Failed to enable mailbox."
        }
        
        Write-Log -Message "Mailbox enabled successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            Mailbox = $mailbox
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to enable mailbox: $_" -Level "ERROR"
        return @{
            Status = "Error"
            Mailbox = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Disable-MailboxUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-Log -Message "Disabling mailbox for user '$Identity'..." -Level "INFO"
        
        # Disable mailbox
        Disable-Mailbox -Identity $Identity -Confirm:$false
        
        Write-Log -Message "Mailbox disabled successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to disable mailbox: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Remove-MailboxUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    
    try {
        Write-Log -Message "Removing mailbox for user '$Identity'..." -Level "INFO"
        
        # Remove mailbox
        Remove-Mailbox -Identity $Identity -Confirm:$false
        
        Write-Log -Message "Mailbox removed successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to remove mailbox: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Set-MailboxPermission {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$User,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    try {
        Write-Log -Message "Setting mailbox permissions for user '$Identity'..." -Level "INFO"
        
        # Set permissions based on access rights
        switch ($AccessRights) {
            "FullAccess" {
                Add-MailboxPermission -Identity $Identity -User $User -AccessRights FullAccess -InheritanceType All
                Write-Log -Message "Full Access permissions granted to $User on mailbox $Identity." -Level "INFO"
            }
            "SendAs" {
                Add-ADPermission -Identity $Identity -User $User -ExtendedRights "Send As"
                Write-Log -Message "Send As permissions granted to $User on mailbox $Identity." -Level "INFO"
            }
            "SendOnBehalf" {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add = $User}
                Write-Log -Message "Send On Behalf permissions granted to $User on mailbox $Identity." -Level "INFO"
            }
        }
        
        return @{
            Status = "Success"
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to set mailbox permissions: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Exchange mailbox management process." -Level "INFO"
    
    # Connect to Exchange server
    $connectionResult = Connect-ExchangeServer -ExchangeServer $ExchangeServer -Credential $Credential
    
    if ($connectionResult.Status -ne "Success") {
        Write-Log -Message "Failed to connect to Exchange server. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Perform the requested action
    switch ($Action) {
        "Create" {
            # Validate required parameters
            if (-not $Name) {
                Write-Log -Message "Name parameter is required for Create action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $Alias) {
                Write-Log -Message "Alias parameter is required for Create action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $Database) {
                Write-Log -Message "Database parameter is required for Create action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $Password) {
                Write-Log -Message "Password parameter is required for Create action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Create-Mailbox -Identity $Identity -Name $Name -Alias $Alias -Database $Database -Password $Password -FirstName $FirstName -LastName $LastName -OrganizationalUnit $OrganizationalUnit -EmailAddresses $EmailAddresses -QuotaInGB $QuotaInGB -ArchiveEnabled $ArchiveEnabled -ArchiveQuotaInGB $ArchiveQuotaInGB
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to create mailbox. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Configure" {
            $result = Configure-Mailbox -Identity $Identity -Name $Name -Alias $Alias -EmailAddresses $EmailAddresses -QuotaInGB $QuotaInGB -ArchiveEnabled $ArchiveEnabled -ArchiveQuotaInGB $ArchiveQuotaInGB
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to configure mailbox. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Enable" {
            # Validate required parameters
            if (-not $Database) {
                Write-Log -Message "Database parameter is required for Enable action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Enable-MailboxUser -Identity $Identity -Database $Database
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to enable mailbox. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Disable" {
            $result = Disable-MailboxUser -Identity $Identity
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to disable mailbox. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "Remove" {
            $result = Remove-MailboxUser -Identity $Identity
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to remove mailbox. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "SetPermission" {
            # Validate required parameters
            if (-not $User) {
                Write-Log -Message "User parameter is required for SetPermission action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            if (-not $AccessRights) {
                Write-Log -Message "AccessRights parameter is required for SetPermission action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Set-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to set mailbox permissions. Exiting..." -Level "ERROR"
                exit 1
            }
        }
    }
    
    # Disconnect from Exchange server
    Disconnect-ExchangeServer -Session $connectionResult.Session
    
    Write-Log -Message "Exchange mailbox management process completed successfully." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during Exchange mailbox management process: $_" -Level "ERROR"
    
    # Attempt to disconnect from Exchange server
    if ($connectionResult -and $connectionResult.Session) {
        Disconnect-ExchangeServer -Session $connectionResult.Session
    }
    
    exit 1
}
