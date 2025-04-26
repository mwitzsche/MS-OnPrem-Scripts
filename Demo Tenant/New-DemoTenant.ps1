<#
.SYNOPSIS
    Creates a complete demo tenant with fictional company structure, departments, users, and groups.

.DESCRIPTION
    This script creates a fictional company structure in Active Directory with organizational units
    for different departments (Management, Marketing, Sales, IT, Production), security groups,
    and 50 fictional employees distributed across these departments. It provides detailed logging
    and error handling.

.PARAMETER CompanyName
    Name of the fictional company to create.

.PARAMETER DomainName
    Domain name for the fictional company.

.PARAMETER AdminCredential
    Credentials with domain admin rights to create the structure.

.PARAMETER UserPassword
    Default password to set for all created users.

.PARAMETER OutputPath
    Path where the report of created objects will be saved.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\New-DemoTenant.ps1 -CompanyName "Contoso Technologies" -DomainName "contoso.com" -AdminCredential (Get-Credential) -UserPassword (ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force) -OutputPath "C:\Reports"

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$CompanyName,

    [Parameter(Mandatory = $true)]
    [string]$DomainName,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [Parameter(Mandatory = $true)]
    [System.Security.SecureString]$UserPassword,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\Documents\DemoTenant",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Documents\DemoTenant\DemoTenant_$(Get-Date -Format 'yyyyMMdd').log"
)

# Define company structure
$departments = @(
    @{
        Name = "Management"
        Description = "Company Management and Executives"
        Abbreviation = "MGT"
        UserCount = 5
        Groups = @(
            @{
                Name = "Management-Users"
                Description = "All Management Department Users"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "Management-Admins"
                Description = "Management Department Administrators"
                Scope = "Global"
                Category = "Security"
            }
        )
    },
    @{
        Name = "Marketing"
        Description = "Marketing and Public Relations"
        Abbreviation = "MKT"
        UserCount = 10
        Groups = @(
            @{
                Name = "Marketing-Users"
                Description = "All Marketing Department Users"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "Marketing-Admins"
                Description = "Marketing Department Administrators"
                Scope = "Global"
                Category = "Security"
            }
        )
    },
    @{
        Name = "Sales"
        Description = "Sales and Customer Relations"
        Abbreviation = "SLS"
        UserCount = 15
        Groups = @(
            @{
                Name = "Sales-Users"
                Description = "All Sales Department Users"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "Sales-Admins"
                Description = "Sales Department Administrators"
                Scope = "Global"
                Category = "Security"
            }
        )
    },
    @{
        Name = "IT"
        Description = "Information Technology"
        Abbreviation = "IT"
        UserCount = 8
        Groups = @(
            @{
                Name = "IT-Users"
                Description = "All IT Department Users"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "IT-Admins"
                Description = "IT Department Administrators"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "IT-Helpdesk"
                Description = "IT Helpdesk Staff"
                Scope = "Global"
                Category = "Security"
            }
        )
    },
    @{
        Name = "Production"
        Description = "Production and Manufacturing"
        Abbreviation = "PRD"
        UserCount = 12
        Groups = @(
            @{
                Name = "Production-Users"
                Description = "All Production Department Users"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "Production-Admins"
                Description = "Production Department Administrators"
                Scope = "Global"
                Category = "Security"
            },
            @{
                Name = "Production-Operators"
                Description = "Production Machine Operators"
                Scope = "Global"
                Category = "Security"
            }
        )
    }
)

# Define first names for user generation
$firstNames = @(
    "James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph", "Thomas", "Charles",
    "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Barbara", "Susan", "Jessica", "Sarah", "Karen",
    "Christopher", "Daniel", "Matthew", "Anthony", "Mark", "Donald", "Steven", "Paul", "Andrew", "Joshua",
    "Lisa", "Nancy", "Betty", "Margaret", "Sandra", "Ashley", "Kimberly", "Emily", "Donna", "Michelle",
    "Kenneth", "George", "Brian", "Edward", "Ronald", "Timothy", "Jason", "Jeffrey", "Ryan", "Jacob",
    "Carol", "Amanda", "Melissa", "Deborah", "Stephanie", "Rebecca", "Laura", "Sharon", "Cynthia", "Kathleen"
)

# Define last names for user generation
$lastNames = @(
    "Smith", "Johnson", "Williams", "Jones", "Brown", "Davis", "Miller", "Wilson", "Moore", "Taylor",
    "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Robinson",
    "Clark", "Rodriguez", "Lewis", "Lee", "Walker", "Hall", "Allen", "Young", "Hernandez", "King",
    "Wright", "Lopez", "Hill", "Scott", "Green", "Adams", "Baker", "Gonzalez", "Nelson", "Carter",
    "Mitchell", "Perez", "Roberts", "Turner", "Phillips", "Campbell", "Parker", "Evans", "Edwards", "Collins"
)

# Define job titles for user generation
$jobTitles = @{
    "Management" = @(
        "CEO", "CFO", "COO", "CTO", "President", "Vice President", "Director", "Executive Assistant"
    )
    "Marketing" = @(
        "Marketing Manager", "Marketing Specialist", "Digital Marketing Specialist", "Content Creator",
        "Social Media Manager", "Brand Manager", "Marketing Analyst", "Public Relations Specialist"
    )
    "Sales" = @(
        "Sales Manager", "Sales Representative", "Account Executive", "Business Development Manager",
        "Sales Analyst", "Customer Success Manager", "Sales Coordinator", "Regional Sales Manager"
    )
    "IT" = @(
        "IT Manager", "System Administrator", "Network Engineer", "Software Developer",
        "Database Administrator", "IT Support Specialist", "Security Analyst", "DevOps Engineer"
    )
    "Production" = @(
        "Production Manager", "Production Supervisor", "Quality Control Specialist", "Manufacturing Engineer",
        "Production Planner", "Maintenance Technician", "Inventory Specialist", "Machine Operator"
    )
}

# Function to write log messages
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

# Function to generate a random user
function Get-RandomUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Department,
        
        [Parameter(Mandatory = $true)]
        [string]$DepartmentAbbreviation
    )
    
    $firstName = $firstNames | Get-Random
    $lastName = $lastNames | Get-Random
    $jobTitle = $jobTitles[$Department] | Get-Random
    
    $username = ($firstName.Substring(0, 1) + $lastName).ToLower()
    $email = "$username@$DomainName"
    
    return @{
        FirstName = $firstName
        LastName = $lastName
        Username = $username
        Email = $email
        JobTitle = $jobTitle
        Department = $Department
        DepartmentAbbreviation = $DepartmentAbbreviation
    }
}

# Function to create the company structure
function New-CompanyStructure {
    try {
        Write-Log -Message "Starting creation of company structure for '$CompanyName'..." -Level "INFO"
        
        # Create output directory if it doesn't exist
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created output directory '$OutputPath'." -Level "INFO"
        }
        
        # Import Active Directory module
        Import-Module ActiveDirectory
        Write-Log -Message "Imported Active Directory module." -Level "INFO"
        
        # Create company OU
        $companyOUName = $CompanyName.Replace(" ", "")
        $companyOUPath = "DC=" + $DomainName.Replace(".", ",DC=")
        
        try {
            $companyOU = Get-ADOrganizationalUnit -Identity "OU=$companyOUName,$companyOUPath" -ErrorAction Stop
            Write-Log -Message "Company OU already exists: 'OU=$companyOUName,$companyOUPath'" -Level "WARNING"
        }
        catch {
            $companyOU = New-ADOrganizationalUnit -Name $companyOUName -Path $companyOUPath -Description $CompanyName -ProtectedFromAccidentalDeletion $true -Credential $AdminCredential
            Write-Log -Message "Created company OU: 'OU=$companyOUName,$companyOUPath'" -Level "INFO"
        }
        
        # Create departments, groups, and users
        $createdObjects = @{
            OrganizationalUnits = @()
            Groups = @()
            Users = @()
        }
        
        foreach ($dept in $departments) {
            # Create department OU
            $deptOUPath = "OU=$companyOUName,$companyOUPath"
            
            try {
                $deptOU = Get-ADOrganizationalUnit -Identity "OU=$($dept.Name),$deptOUPath" -ErrorAction Stop
                Write-Log -Message "Department OU already exists: 'OU=$($dept.Name),$deptOUPath'" -Level "WARNING"
            }
            catch {
                $deptOU = New-ADOrganizationalUnit -Name $dept.Name -Path $deptOUPath -Description $dept.Description -ProtectedFromAccidentalDeletion $true -Credential $AdminCredential
                Write-Log -Message "Created department OU: 'OU=$($dept.Name),$deptOUPath'" -Level "INFO"
            }
            
            $createdObjects.OrganizationalUnits += @{
                Name = $dept.Name
                Path = "OU=$($dept.Name),$deptOUPath"
                Description = $dept.Description
            }
            
            # Create department groups
            foreach ($group in $dept.Groups) {
                $groupName = "$companyOUName-$($group.Name)"
                
                try {
                    $adGroup = Get-ADGroup -Identity $groupName -ErrorAction Stop
                    Write-Log -Message "Group already exists: '$groupName'" -Level "WARNING"
                }
                catch {
                    $adGroup = New-ADGroup -Name $groupName -SamAccountName $groupName -GroupCategory $group.Category -GroupScope $group.Scope -DisplayName $groupName -Path "OU=$($dept.Name),$deptOUPath" -Description $group.Description -Credential $AdminCredential
                    Write-Log -Message "Created group: '$groupName'" -Level "INFO"
                }
                
                $createdObjects.Groups += @{
                    Name = $groupName
                    Path = "OU=$($dept.Name),$deptOUPath"
                    Category = $group.Category
                    Scope = $group.Scope
                    Description = $group.Description
                }
            }
            
            # Create department users
            for ($i = 1; $i -le $dept.UserCount; $i++) {
                $user = Get-RandomUser -Department $dept.Name -DepartmentAbbreviation $dept.Abbreviation
                
                # Ensure unique username by adding a number if necessary
                $baseUsername = $user.Username
                $counter = 1
                
                while ($true) {
                    try {
                        $existingUser = Get-ADUser -Identity $user.Username -ErrorAction Stop
                        $user.Username = "$baseUsername$counter"
                        $user.Email = "$($user.Username)@$DomainName"
                        $counter++
                    }
                    catch {
                        # Username is available
                        break
                    }
                }
                
                try {
                    $adUser = New-ADUser -Name "$($user.FirstName) $($user.LastName)" `
                        -GivenName $user.FirstName `
                        -Surname $user.LastName `
                        -SamAccountName $user.Username `
                        -UserPrincipalName "$($user.Username)@$DomainName" `
                        -Path "OU=$($dept.Name),$deptOUPath" `
                        -EmailAddress $user.Email `
                        -Title $user.JobTitle `
                        -Department $dept.Name `
                        -Company $CompanyName `
                        -AccountPassword $UserPassword `
                        -Enabled $true `
                        -PasswordNeverExpires $true `
                        -Credential $AdminCredential
                    
                    Write-Log -Message "Created user: '$($user.Username)' ($($user.FirstName) $($user.LastName))" -Level "INFO"
                    
                    # Add user to department users group
                    Add-ADGroupMember -Identity "$companyOUName-$($dept.Name)-Users" -Members $user.Username -Credential $AdminCredential
                    Write-Log -Message "Added user '$($user.Username)' to group '$companyOUName-$($dept.Name)-Users'" -Level "INFO"
                    
                    # Add some users to admin groups (first user in each department)
                    if ($i -eq 1) {
                        Add-ADGroupMember -Identity "$companyOUName-$($dept.Name)-Admins" -Members $user.Username -Credential $AdminCredential
                        Write-Log -Message "Added user '$($user.Username)' to group '$companyOUName-$($dept.Name)-Admins'" -Level "INFO"
                    }
                    
                    # Add IT users to helpdesk group
                    if ($dept.Name -eq "IT" -and $i -gt 1 -and $i -le 4) {
                        Add-ADGroupMember -Identity "$companyOUName-IT-Helpdesk" -Members $user.Username -Credential $AdminCredential
                        Write-Log -Message "Added user '$($user.Username)' to group '$companyOUName-IT-Helpdesk'" -Level "INFO"
                    }
                    
                    # Add Production users to operators group
                    if ($dept.Name -eq "Production" -and $i -gt 2 -and $i -le 8) {
                        Add-ADGroupMember -Identity "$companyOUName-Production-Operators" -Members $user.Username -Credential $AdminCredential
                        Write-Log -Message "Added user '$($user.Username)' to group '$companyOUName-Production-Operators'" -Level "INFO"
                    }
                    
                    $createdObjects.Users += @{
                        Username = $user.Username
                        Name = "$($user.FirstName) $($user.LastName)"
                        Email = $user.Email
                        Department = $dept.Name
                        JobTitle = $user.JobTitle
                    }
                }
                catch {
                    Write-Log -Message "Failed to create user '$($user.Username)': $_" -Level "ERROR"
                }
            }
        }
        
        # Generate report
        $reportPath = Join-Path -Path $OutputPath -ChildPath "DemoTenant_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        $report = @"
<!DOCTYPE html>
<html>
<head>
    <title>Demo Tenant Report - $CompanyName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2, h3 { color: #0066cc; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { background-color: #e6f2ff; padding: 10px; border-radius: 5px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>Demo Tenant Report - $CompanyName</h1>
    <div class="summary">
        <p><strong>Domain:</strong> $DomainName</p>
        <p><strong>Created:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Organizational Units:</strong> $($createdObjects.OrganizationalUnits.Count)</p>
        <p><strong>Groups:</strong> $($createdObjects.Groups.Count)</p>
        <p><strong>Users:</strong> $($createdObjects.Users.Count)</p>
    </div>
    
    <h2>Organizational Units</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Path</th>
            <th>Description</th>
        </tr>
"@
        
        foreach ($ou in $createdObjects.OrganizationalUnits) {
            $report += @"
        <tr>
            <td>$($ou.Name)</td>
            <td>$($ou.Path)</td>
            <td>$($ou.Description)</td>
        </tr>
"@
        }
        
        $report += @"
    </table>
    
    <h2>Groups</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Path</th>
            <th>Category</th>
            <th>Scope</th>
            <th>Description</th>
        </tr>
"@
        
        foreach ($group in $createdObjects.Groups) {
            $report += @"
        <tr>
            <td>$($group.Name)</td>
            <td>$($group.Path)</td>
            <td>$($group.Category)</td>
            <td>$($group.Scope)</td>
            <td>$($group.Description)</td>
        </tr>
"@
        }
        
        $report += @"
    </table>
    
    <h2>Users</h2>
    <table>
        <tr>
            <th>Username</th>
            <th>Name</th>
            <th>Email</th>
            <th>Department</th>
            <th>Job Title</th>
        </tr>
"@
        
        foreach ($user in $createdObjects.Users) {
            $report += @"
        <tr>
            <td>$($user.Username)</td>
            <td>$($user.Name)</td>
            <td>$($user.Email)</td>
            <td>$($user.Department)</td>
            <td>$($user.JobTitle)</td>
        </tr>
"@
        }
        
        $report += @"
    </table>
</body>
</html>
"@
        
        $report | Out-File -FilePath $reportPath -Encoding UTF8
        
        Write-Log -Message "Generated report at '$reportPath'" -Level "INFO"
        Write-Log -Message "Company structure creation completed successfully." -Level "INFO"
        
        return @{
            Status = "Success"
            ReportPath = $reportPath
            CreatedObjects = $createdObjects
            ErrorMessage = $null
        }
    }
    catch {
        Write-Log -Message "Failed to create company structure: $_" -Level "ERROR"
        return @{
            Status = "Error"
            ReportPath = $null
            CreatedObjects = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting demo tenant creation process for '$CompanyName'." -Level "INFO"
    
    # Create company structure
    $result = New-CompanyStructure
    
    if ($result.Status -ne "Success") {
        Write-Log -Message "Failed to create demo tenant. Exiting..." -Level "ERROR"
        exit 1
    }
    
    Write-Log -Message "Demo tenant creation process completed successfully." -Level "INFO"
    Write-Log -Message "Report generated at: $($result.ReportPath)" -Level "INFO"
    Write-Log -Message "Created $($result.CreatedObjects.OrganizationalUnits.Count) OUs, $($result.CreatedObjects.Groups.Count) groups, and $($result.CreatedObjects.Users.Count) users." -Level "INFO"
    
    # Return summary
    return @{
        CompanyName = $CompanyName
        DomainName = $DomainName
        ReportPath = $result.ReportPath
        OrganizationalUnits = $result.CreatedObjects.OrganizationalUnits.Count
        Groups = $result.CreatedObjects.Groups.Count
        Users = $result.CreatedObjects.Users.Count
        Status = "Success"
    }
}
catch {
    Write-Log -Message "An error occurred during demo tenant creation process: $_" -Level "ERROR"
    exit 1
}
