# Demo Tenant Creation Script

This repository contains a PowerShell script for creating a complete demo tenant with a fictional company structure in Active Directory. The script creates organizational units, security groups, and generates 50 fictional employees distributed across different departments.

## Overview

The `New-DemoTenant.ps1` script automates the creation of a fictional company structure in Active Directory with:

- A company organizational unit
- Department organizational units (Management, Marketing, Sales, IT, Production)
- Security groups for each department
- 50 fictional employees with realistic names, job titles, and email addresses
- Proper group memberships based on department and role
- Detailed HTML report of all created objects

## Requirements

- Windows Server with Active Directory Domain Services installed
- PowerShell 5.1 or higher
- Domain Administrator credentials
- Active Directory PowerShell module

## Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| CompanyName | Name of the fictional company to create | Yes |
| DomainName | Domain name for the fictional company | Yes |
| AdminCredential | Credentials with domain admin rights | Yes |
| UserPassword | Default password for all created users | Yes |
| OutputPath | Path where the report will be saved | No |
| LogPath | Path where logs will be stored | No |

## Usage

### Basic Usage

```powershell
.\New-DemoTenant.ps1 -CompanyName "Contoso Technologies" -DomainName "contoso.com" -AdminCredential (Get-Credential) -UserPassword (ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force)
```

### Specifying Output Path

```powershell
.\New-DemoTenant.ps1 -CompanyName "Contoso Technologies" -DomainName "contoso.com" -AdminCredential (Get-Credential) -UserPassword (ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force) -OutputPath "C:\Reports\DemoTenant"
```

## Company Structure

The script creates the following structure:

### Departments

1. **Management**
   - 5 users
   - Groups: Management-Users, Management-Admins

2. **Marketing**
   - 10 users
   - Groups: Marketing-Users, Marketing-Admins

3. **Sales**
   - 15 users
   - Groups: Sales-Users, Sales-Admins

4. **IT**
   - 8 users
   - Groups: IT-Users, IT-Admins, IT-Helpdesk

5. **Production**
   - 12 users
   - Groups: Production-Users, Production-Admins, Production-Operators

### User Distribution

- Each user is assigned a random first name, last name, and job title appropriate for their department
- Usernames are generated as first initial + last name (e.g., jsmith)
- Email addresses follow the format username@domainname
- The first user in each department is added to the department's admin group
- IT users (2-4) are added to the IT-Helpdesk group
- Production users (3-8) are added to the Production-Operators group

## Output

The script generates:

1. **HTML Report** - A detailed report of all created objects including:
   - Organizational Units
   - Security Groups
   - Users with their details

2. **Log File** - A detailed log of all operations performed during script execution

## Examples

### Creating a Demo Tenant for Contoso

```powershell
$adminCred = Get-Credential -Message "Enter Domain Admin credentials"
$userPassword = ConvertTo-SecureString "Demo@Pass123" -AsPlainText -Force

.\New-DemoTenant.ps1 -CompanyName "Contoso Technologies" `
                     -DomainName "contoso.com" `
                     -AdminCredential $adminCred `
                     -UserPassword $userPassword `
                     -OutputPath "C:\Reports\Contoso"
```

### Creating a Demo Tenant for Fabrikam

```powershell
$adminCred = Get-Credential -Message "Enter Domain Admin credentials"
$userPassword = ConvertTo-SecureString "Fabrikam@2025" -AsPlainText -Force

.\New-DemoTenant.ps1 -CompanyName "Fabrikam Industries" `
                     -DomainName "fabrikam.com" `
                     -AdminCredential $adminCred `
                     -UserPassword $userPassword
```

## Notes

- The script includes error handling and will log any issues encountered during execution
- All created users have the same password as specified in the UserPassword parameter
- The script checks for existing objects before creation to avoid duplicates
- The first user in each department is automatically added to the department's admin group

## Author

Michael Witzsche  
Date: April 26, 2025  
Version: 1.0.0
