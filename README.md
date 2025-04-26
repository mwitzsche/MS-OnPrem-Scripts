# PowerShell Scripts for On-Premises Windows Environments

## Overview

This manual documents a comprehensive collection of PowerShell scripts for managing all aspects of on-premises Windows environments, including Active Directory, Windows 10/11 clients, Windows Server, Exchange Server, and SharePoint Server. These scripts are designed for system administrators to automate common tasks, configure systems, and generate detailed reports.

**Author:** Michael Witzsche  
**Date:** April 26, 2025  
**Version:** 1.0.0

## Table of Contents

1. [Active Directory](#active-directory)
2. [Windows Client](#windows-client)
3. [Windows Server](#windows-server)
4. [Exchange Server](#exchange-server)
5. [SharePoint Server](#sharepoint-server)

## Installation and Requirements

### Prerequisites

- PowerShell 5.1 or PowerShell 7.x
- Required PowerShell modules:
  - ActiveDirectory
  - GroupPolicy
  - ServerManager
  - DnsServer
  - NetTCPIP
  - ExchangeManagementShell (for Exchange scripts)
  - SharePoint.PowerShell (for SharePoint scripts)
  - ImportExcel (for report export)

### Installation

1. Install required PowerShell modules:

```powershell
# Install general modules
Install-Module -Name ImportExcel -Force

# For Active Directory management
Add-WindowsFeature RSAT-AD-PowerShell

# For Exchange management (run on Exchange server)
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# For SharePoint management (run on SharePoint server)
Add-PSSnapin Microsoft.SharePoint.PowerShell
```

2. Download the scripts to your local machine
3. Ensure execution policy allows running the scripts:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Active Directory

Scripts for managing Active Directory users, groups, organizational units, and other objects.

### New-ADUser.ps1

**Description:** Creates a new user in Active Directory with specified attributes and group memberships.

**Parameters:**
- `FirstName` - First name of the user
- `LastName` - Last name of the user
- `SamAccountName` - SAM account name for the user
- `UserPrincipalName` - User principal name (email format)
- `Password` - Initial password
- `ChangePasswordAtLogon` - Whether to force password change at next logon
- `Enabled` - Whether the account should be enabled
- `Department` - User's department
- `Title` - User's job title
- `Company` - User's company name
- `Path` - OU path where the user will be created
- `Groups` - Array of group names to add the user to
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-ADUser.ps1 -FirstName "John" -LastName "Doe" -SamAccountName "jdoe" -UserPrincipalName "john.doe@contoso.com" -Password (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) -ChangePasswordAtLogon $true -Enabled $true -Department "IT" -Title "System Administrator" -Company "Contoso" -Path "OU=IT,OU=Users,DC=contoso,DC=com" -Groups @("IT Staff", "Domain Admins")
```

### New-ADGroup.ps1

**Description:** Creates a new security or distribution group in Active Directory.

**Parameters:**
- `Name` - Name of the group
- `SamAccountName` - SAM account name for the group
- `GroupScope` - Scope of the group (Global, Universal, DomainLocal)
- `GroupCategory` - Category of the group (Security, Distribution)
- `Description` - Description of the group
- `Path` - OU path where the group will be created
- `Members` - Array of user SAM account names to add as group members
- `MemberOf` - Array of group names to add this group as a member of
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-ADGroup.ps1 -Name "IT Department" -SamAccountName "IT_Dept" -GroupScope "Global" -GroupCategory "Security" -Description "IT Department Security Group" -Path "OU=Groups,DC=contoso,DC=com" -Members @("jdoe", "asmith") -MemberOf @("All Staff")
```

### New-ADOrganizationalUnit.ps1

**Description:** Creates a new organizational unit in Active Directory with optional nested OUs.

**Parameters:**
- `Name` - Name of the organizational unit
- `Path` - Parent path where the OU will be created
- `Description` - Description of the OU
- `ProtectedFromAccidentalDeletion` - Whether the OU is protected from accidental deletion
- `NestedOUs` - Array of nested OUs to create within this OU
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$nestedOUs = @(
    @{Name="Users"; Description="Department Users"; Protected=$true},
    @{Name="Computers"; Description="Department Computers"; Protected=$true},
    @{Name="Groups"; Description="Department Groups"; Protected=$true}
)
.\New-ADOrganizationalUnit.ps1 -Name "IT" -Path "DC=contoso,DC=com" -Description "IT Department" -ProtectedFromAccidentalDeletion $true -NestedOUs $nestedOUs
```

### Export-ADUserReport.ps1

**Description:** Generates a comprehensive report of Active Directory users including account information, group memberships, and last logon time.

**Parameters:**
- `SearchBase` - The OU to search for users
- `Filter` - LDAP filter to apply to the search
- `Properties` - Array of user properties to include in the report
- `IncludeGroups` - Whether to include group memberships in the report
- `IncludeLastLogon` - Whether to include last logon information in the report
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Export-ADUserReport.ps1 -SearchBase "OU=Users,DC=contoso,DC=com" -Filter "Department -eq 'IT'" -Properties @("Name", "Title", "Department", "Manager", "EmailAddress", "Enabled") -IncludeGroups $true -IncludeLastLogon $true -ExportPath "C:\Reports\ADUsers.xlsx" -ExportFormat "Excel"
```

### Export-ADGroupReport.ps1

**Description:** Generates a report of Active Directory groups and their members.

**Parameters:**
- `SearchBase` - The OU to search for groups
- `Filter` - LDAP filter to apply to the search
- `IncludeMembers` - Whether to include group members in the report
- `IncludeNestedGroups` - Whether to include nested group memberships
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Export-ADGroupReport.ps1 -SearchBase "OU=Groups,DC=contoso,DC=com" -Filter "*" -IncludeMembers $true -IncludeNestedGroups $true -ExportPath "C:\Reports\ADGroups.xlsx" -ExportFormat "Excel"
```

## Windows Client

Scripts for managing Windows 10/11 client computers, including inventory, updates, settings, and more.

### Get-ComputerInventory.ps1

**Description:** Collects comprehensive hardware and software inventory from Windows client computers.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `IncludeHardware` - Whether to include hardware information
- `IncludeSoftware` - Whether to include installed software
- `IncludeUpdates` - Whether to include installed updates
- `IncludeServices` - Whether to include running services
- `ExportPath` - Path where the inventory will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Get-ComputerInventory.ps1 -ComputerName @("PC001", "PC002") -Credential (Get-Credential) -IncludeHardware $true -IncludeSoftware $true -IncludeUpdates $true -IncludeServices $true -ExportPath "C:\Reports\Inventory.xlsx" -ExportFormat "Excel"
```

### Install-WindowsUpdates.ps1

**Description:** Installs Windows updates on local or remote computers.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `UpdateType` - Type of updates to install (Security, Critical, All)
- `RebootIfRequired` - Whether to reboot the computer if required
- `ScheduleReboot` - Time to schedule reboot (if not immediate)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Install-WindowsUpdates.ps1 -ComputerName @("PC001", "PC002") -Credential (Get-Credential) -UpdateType "Security" -RebootIfRequired $true -ScheduleReboot "22:00"
```

### Set-WindowsConfiguration.ps1

**Description:** Configures various Windows settings on local or remote computers.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `PowerSettings` - Power plan settings to configure
- `UAC` - User Account Control settings
- `WindowsFeatures` - Windows features to enable or disable
- `RegistrySettings` - Registry settings to configure
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$powerSettings = @{
    PlanName = "High Performance"
    TurnOffDisplayMinutes = 15
    SleepMinutes = 30
}
$registrySettings = @(
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name="EnableLUA"; Value=1; Type="DWord"},
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"; Name="ConsentPromptBehaviorAdmin"; Value=2; Type="DWord"}
)
.\Set-WindowsConfiguration.ps1 -ComputerName "PC001" -Credential (Get-Credential) -PowerSettings $powerSettings -UAC "Default" -WindowsFeatures @{Enable=@("Telnet-Client"); Disable=@("Internet-Explorer-Optional-amd64")} -RegistrySettings $registrySettings
```

### Rename-Computer.ps1

**Description:** Renames a local or remote computer and optionally joins it to a domain.

**Parameters:**
- `ComputerName` - Current name of the computer
- `NewName` - New name for the computer
- `Credential` - Credentials to use for remote connection
- `DomainName` - Domain to join (if not already joined)
- `DomainCredential` - Credentials to use for domain join
- `Restart` - Whether to restart the computer after renaming
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Rename-Computer.ps1 -ComputerName "PC001" -NewName "LAPTOP-SALES01" -Credential (Get-Credential) -DomainName "contoso.com" -DomainCredential (Get-Credential) -Restart $true
```

### Get-EventLogAnalysis.ps1

**Description:** Analyzes Windows event logs for errors, warnings, and specific events.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `LogName` - Name of the event log(s) to analyze
- `StartTime` - Start time for the analysis
- `EndTime` - End time for the analysis
- `EventType` - Type of events to include (Error, Warning, Information, All)
- `EventID` - Specific event IDs to search for
- `ExportPath` - Path where the analysis will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Get-EventLogAnalysis.ps1 -ComputerName "DC01" -Credential (Get-Credential) -LogName @("System", "Application") -StartTime (Get-Date).AddDays(-7) -EndTime (Get-Date) -EventType @("Error", "Warning") -EventID @(1001, 1018, 4624) -ExportPath "C:\Reports\EventLogs.xlsx" -ExportFormat "Excel"
```

### Set-FirewallConfiguration.ps1

**Description:** Configures Windows Firewall settings and rules on local or remote computers.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `ProfileSettings` - Settings for each firewall profile (Domain, Private, Public)
- `Rules` - Firewall rules to create or modify
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$profileSettings = @{
    Domain = @{Enabled=$true; DefaultInboundAction="Block"; DefaultOutboundAction="Allow"; LogAllowed=$true}
    Private = @{Enabled=$true; DefaultInboundAction="Block"; DefaultOutboundAction="Allow"; LogBlocked=$true}
    Public = @{Enabled=$true; DefaultInboundAction="Block"; DefaultOutboundAction="Allow"; LogMaxSizeKB=4096}
}
$rules = @(
    @{Name="Allow RDP"; Direction="Inbound"; Action="Allow"; Protocol="TCP"; LocalPort=3389; Profile=@("Domain", "Private")},
    @{Name="Block Telnet"; Direction="Inbound"; Action="Block"; Protocol="TCP"; LocalPort=23; Profile=@("Domain", "Private", "Public")}
)
.\Set-FirewallConfiguration.ps1 -ComputerName "PC001" -Credential (Get-Credential) -ProfileSettings $profileSettings -Rules $rules
```

## Windows Server

Scripts for managing Windows Server, including Desired State Configuration, IIS, domain joining, and GPO management.

### New-DSCConfiguration.ps1

**Description:** Creates and applies a Desired State Configuration to Windows servers.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `ConfigurationName` - Name of the DSC configuration
- `ConfigurationData` - Configuration data for the DSC configuration
- `Features` - Windows features to install or remove
- `Services` - Services to configure
- `RegistrySettings` - Registry settings to configure
- `Files` - Files to create or modify
- `OutputPath` - Path where the MOF files will be saved
- `ApplyConfiguration` - Whether to apply the configuration immediately
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$features = @(
    @{Name="Web-Server"; Ensure="Present"},
    @{Name="Web-Asp-Net45"; Ensure="Present"},
    @{Name="Telnet-Client"; Ensure="Absent"}
)
$services = @(
    @{Name="BITS"; State="Running"; StartupType="Automatic"},
    @{Name="Spooler"; State="Running"; StartupType="Automatic"}
)
.\New-DSCConfiguration.ps1 -ComputerName "WEB01" -Credential (Get-Credential) -ConfigurationName "WebServerConfig" -Features $features -Services $services -RegistrySettings @() -Files @() -OutputPath "C:\DSC" -ApplyConfiguration $true
```

### Install-IISServer.ps1

**Description:** Installs and configures Internet Information Services (IIS) on Windows Server.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `Features` - IIS features to install
- `WebsiteName` - Name of the website to create
- `WebsitePath` - Physical path for the website
- `AppPoolName` - Name of the application pool
- `AppPoolIdentity` - Identity for the application pool
- `BindingInformation` - Binding information for the website
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$features = @(
    "Web-Server",
    "Web-Common-Http",
    "Web-Default-Doc",
    "Web-Dir-Browsing",
    "Web-Http-Errors",
    "Web-Static-Content",
    "Web-Http-Logging",
    "Web-Stat-Compression",
    "Web-Filtering",
    "Web-Mgmt-Console",
    "Web-Asp-Net45"
)
.\Install-IISServer.ps1 -ComputerName "WEB01" -Credential (Get-Credential) -Features $features -WebsiteName "Corporate Intranet" -WebsitePath "C:\inetpub\wwwroot\intranet" -AppPoolName "IntranetAppPool" -AppPoolIdentity "ApplicationPoolIdentity" -BindingInformation "*:80:intranet.contoso.com"
```

### Join-Domain.ps1

**Description:** Joins a computer to an Active Directory domain.

**Parameters:**
- `ComputerName` - Name of the target computer(s)
- `Credential` - Credentials to use for remote connection
- `DomainName` - Name of the domain to join
- `DomainCredential` - Credentials to use for domain join
- `OrganizationalUnit` - OU path where the computer account will be created
- `Restart` - Whether to restart the computer after joining
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Join-Domain.ps1 -ComputerName "WEB01" -Credential (Get-Credential) -DomainName "contoso.com" -DomainCredential (Get-Credential) -OrganizationalUnit "OU=Servers,DC=contoso,DC=com" -Restart $true
```

### Manage-GroupPolicy.ps1

**Description:** Creates, modifies, and links Group Policy Objects in Active Directory.

**Parameters:**
- `Action` - Action to perform (Create, Modify, Link, Unlink, Remove)
- `GPOName` - Name of the GPO
- `Domain` - Domain where the GPO exists
- `Comment` - Comment for the GPO
- `TargetOU` - OU to link the GPO to
- `LinkEnabled` - Whether the GPO link is enabled
- `Settings` - Group Policy settings to configure
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
$settings = @(
    @{Type="Registry"; Key="HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"; ValueName="NoAutoUpdate"; Value=0; ValueType="DWord"},
    @{Type="Registry"; Key="HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"; ValueName="AUOptions"; Value=4; ValueType="DWord"}
)
.\Manage-GroupPolicy.ps1 -Action "Create" -GPOName "Windows Update Settings" -Domain "contoso.com" -Comment "Configures Windows Update settings" -TargetOU "OU=Workstations,DC=contoso,DC=com" -LinkEnabled $true -Settings $settings
```

### Get-ServerEventLogs.ps1

**Description:** Analyzes Windows Server event logs for critical events and generates a report.

**Parameters:**
- `ComputerName` - Name of the target server(s)
- `Credential` - Credentials to use for remote connection
- `LogName` - Name of the event log(s) to analyze
- `StartTime` - Start time for the analysis
- `EndTime` - End time for the analysis
- `EventType` - Type of events to include (Error, Warning, Critical, All)
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Get-ServerEventLogs.ps1 -ComputerName @("DC01", "WEB01") -Credential (Get-Credential) -LogName @("System", "Application", "Security") -StartTime (Get-Date).AddDays(-1) -EndTime (Get-Date) -EventType @("Error", "Critical") -ExportPath "C:\Reports\ServerEvents.xlsx" -ExportFormat "Excel"
```

## Exchange Server

Scripts for managing Exchange Server, including mailbox creation, user management, settings configuration, and error analysis.

### New-ExchangeMailbox.ps1

**Description:** Creates a new mailbox in Exchange Server.

**Parameters:**
- `Name` - Name for the mailbox
- `Alias` - Email alias for the mailbox
- `FirstName` - First name of the user
- `LastName` - Last name of the user
- `DisplayName` - Display name for the mailbox
- `UserPrincipalName` - User principal name (email format)
- `Password` - Initial password
- `Database` - Exchange database to store the mailbox
- `MailboxType` - Type of mailbox (Regular, Shared, Room, Equipment)
- `OrganizationalUnit` - OU path where the user account will be created
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-ExchangeMailbox.ps1 -Name "John Doe" -Alias "jdoe" -FirstName "John" -LastName "Doe" -DisplayName "John Doe" -UserPrincipalName "john.doe@contoso.com" -Password (ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force) -Database "Mailbox Database 01" -MailboxType "Regular" -OrganizationalUnit "OU=Users,DC=contoso,DC=com"
```

### New-ExchangeDistributionGroup.ps1

**Description:** Creates a new distribution group in Exchange Server.

**Parameters:**
- `Name` - Name for the distribution group
- `Alias` - Email alias for the group
- `DisplayName` - Display name for the group
- `PrimarySmtpAddress` - Primary SMTP address for the group
- `Members` - Array of members to add to the group
- `ManagedBy` - Array of users who can manage the group
- `OrganizationalUnit` - OU path where the group will be created
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-ExchangeDistributionGroup.ps1 -Name "Sales Team" -Alias "sales" -DisplayName "Sales Team" -PrimarySmtpAddress "sales@contoso.com" -Members @("john.doe@contoso.com", "jane.smith@contoso.com") -ManagedBy @("john.doe@contoso.com") -OrganizationalUnit "OU=Groups,DC=contoso,DC=com"
```

### Set-ExchangeMailboxPermissions.ps1

**Description:** Sets permissions on Exchange mailboxes.

**Parameters:**
- `Identity` - Identity of the mailbox
- `User` - User to grant permissions to
- `AccessRights` - Access rights to grant
- `AutoMapping` - Whether to automatically map the mailbox in Outlook
- `InheritanceType` - Type of inheritance for the permissions
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Set-ExchangeMailboxPermissions.ps1 -Identity "john.doe@contoso.com" -User "jane.smith@contoso.com" -AccessRights @("FullAccess", "SendAs") -AutoMapping $true -InheritanceType "All"
```

### Get-ExchangeMailboxReport.ps1

**Description:** Generates a report of Exchange mailboxes including size, item count, and last logon time.

**Parameters:**
- `Database` - Exchange database to report on
- `Filter` - Filter to apply to the mailboxes
- `IncludeSize` - Whether to include mailbox size in the report
- `IncludeItemCount` - Whether to include item count in the report
- `IncludeLastLogon` - Whether to include last logon time in the report
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Get-ExchangeMailboxReport.ps1 -Database "Mailbox Database 01" -Filter "RecipientTypeDetails -eq 'UserMailbox'" -IncludeSize $true -IncludeItemCount $true -IncludeLastLogon $true -ExportPath "C:\Reports\MailboxReport.xlsx" -ExportFormat "Excel"
```

### Test-ExchangeHealth.ps1

**Description:** Tests the health of Exchange Server components and services.

**Parameters:**
- `Server` - Name of the Exchange server(s) to test
- `Components` - Components to test (Transport, ClientAccess, Mailbox, All)
- `IncludeDAG` - Whether to include Database Availability Group tests
- `IncludeMailflow` - Whether to include mail flow tests
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Test-ExchangeHealth.ps1 -Server @("EXCH01", "EXCH02") -Components "All" -IncludeDAG $true -IncludeMailflow $true -ExportPath "C:\Reports\ExchangeHealth.html" -ExportFormat "HTML"
```

## SharePoint Server

Scripts for managing SharePoint Server, including site creation, site deletion, and permission management.

### New-SharePointSite.ps1

**Description:** Creates a new SharePoint site collection.

**Parameters:**
- `URL` - URL for the new site collection
- `Title` - Title for the site collection
- `Description` - Description of the site collection
- `Template` - Template to use for the site collection
- `OwnerAlias` - Owner of the site collection
- `SecondaryOwnerAlias` - Secondary owner of the site collection
- `ContentDatabase` - Content database to store the site collection
- `Language` - Language for the site collection
- `TimeZone` - Time zone for the site collection
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-SharePointSite.ps1 -URL "https://sharepoint.contoso.com/sites/HR" -Title "Human Resources" -Description "Human Resources Department Site" -Template "STS#0" -OwnerAlias "contoso\jdoe" -SecondaryOwnerAlias "contoso\asmith" -ContentDatabase "WSS_Content" -Language 1033 -TimeZone 4
```

### New-SharePointSubsite.ps1

**Description:** Creates a new SharePoint subsite within an existing site collection.

**Parameters:**
- `ParentSiteURL` - URL of the parent site
- `Title` - Title for the subsite
- `URL` - URL name for the subsite
- `Description` - Description of the subsite
- `Template` - Template to use for the subsite
- `InheritPermissions` - Whether to inherit permissions from the parent site
- `InheritNavigation` - Whether to inherit navigation from the parent site
- `Language` - Language for the subsite
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\New-SharePointSubsite.ps1 -ParentSiteURL "https://sharepoint.contoso.com/sites/HR" -Title "Benefits" -URL "Benefits" -Description "Employee Benefits Site" -Template "STS#0" -InheritPermissions $true -InheritNavigation $true -Language 1033
```

### Remove-SharePointSite.ps1

**Description:** Removes a SharePoint site collection.

**Parameters:**
- `URL` - URL of the site collection to remove
- `GradualDelete` - Whether to delete the site gradually
- `Force` - Whether to force deletion of the site
- `DeleteADAccounts` - Whether to delete associated AD accounts
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Remove-SharePointSite.ps1 -URL "https://sharepoint.contoso.com/sites/Archive" -GradualDelete $true -Force $false -DeleteADAccounts $false
```

### Set-SharePointPermissions.ps1

**Description:** Sets permissions on SharePoint sites, lists, or items.

**Parameters:**
- `SiteURL` - URL of the SharePoint site
- `ObjectType` - Type of object to set permissions on (Site, List, Item)
- `ObjectName` - Name or path of the object
- `User` - User or group to grant permissions to
- `PermissionLevel` - Permission level to grant
- `InheritPermissions` - Whether the object should inherit permissions
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Set-SharePointPermissions.ps1 -SiteURL "https://sharepoint.contoso.com/sites/HR" -ObjectType "List" -ObjectName "Shared Documents" -User "contoso\HR Team" -PermissionLevel "Contribute" -InheritPermissions $false
```

### Get-SharePointPermissionReport.ps1

**Description:** Generates a report of permissions on SharePoint sites, lists, and items.

**Parameters:**
- `SiteURL` - URL of the SharePoint site
- `IncludeLists` - Whether to include lists in the report
- `IncludeItems` - Whether to include items in the report
- `IncludeInheritedPermissions` - Whether to include inherited permissions
- `ExportPath` - Path where the report will be saved
- `ExportFormat` - Format of the export file (CSV, JSON, Excel, HTML)
- `LogPath` - Path where logs will be stored

**Example:**
```powershell
.\Get-SharePointPermissionReport.ps1 -SiteURL "https://sharepoint.contoso.com/sites/HR" -IncludeLists $true -IncludeItems $false -IncludeInheritedPermissions $false -ExportPath "C:\Reports\SharePointPermissions.xlsx" -ExportFormat "Excel"
```
