<#
.SYNOPSIS
    Analyzes and troubleshoots Exchange Server issues.

.DESCRIPTION
    This script analyzes and troubleshoots Exchange Server issues, including checking server health,
    analyzing logs, checking mail flow, and generating reports. It provides detailed logging and error handling.

.PARAMETER Action
    Action to perform (CheckHealth, AnalyzeLogs, CheckMailFlow, TestConnectivity, GenerateReport).

.PARAMETER ExchangeServer
    Exchange server to analyze.

.PARAMETER LogPath
    Path where logs will be stored.

.PARAMETER StartDate
    Start date for log analysis.

.PARAMETER EndDate
    End date for log analysis.

.PARAMETER LogLevel
    Log level for analysis (Error, Warning, Information).

.PARAMETER MailboxServer
    Mailbox server to check.

.PARAMETER HubTransportServer
    Hub Transport server to check.

.PARAMETER ClientAccessServer
    Client Access server to check.

.PARAMETER TestMailbox
    Mailbox to use for mail flow testing.

.PARAMETER TestRecipient
    Recipient to use for mail flow testing.

.PARAMETER ReportPath
    Path where the report will be saved.

.PARAMETER ReportType
    Type of report to generate (ServerHealth, MailboxStatistics, DatabaseStatistics, QueueStatistics).

.PARAMETER Credential
    Credentials to use for Exchange operations.

.EXAMPLE
    .\Analyze-ExchangeServer.ps1 -Action CheckHealth -ExchangeServer "exchange01.contoso.com" -Credential (Get-Credential)

.EXAMPLE
    .\Analyze-ExchangeServer.ps1 -Action AnalyzeLogs -ExchangeServer "exchange01.contoso.com" -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -LogLevel "Error" -Credential (Get-Credential)

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("CheckHealth", "AnalyzeLogs", "CheckMailFlow", "TestConnectivity", "GenerateReport")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$ExchangeServer,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\ExchangeAnalysis_$(Get-Date -Format 'yyyyMMdd').log",

    [Parameter(Mandatory = $false)]
    [DateTime]$StartDate = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [DateTime]$EndDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [ValidateSet("Error", "Warning", "Information")]
    [string]$LogLevel = "Error",

    [Parameter(Mandatory = $false)]
    [string]$MailboxServer,

    [Parameter(Mandatory = $false)]
    [string]$HubTransportServer,

    [Parameter(Mandatory = $false)]
    [string]$ClientAccessServer,

    [Parameter(Mandatory = $false)]
    [string]$TestMailbox,

    [Parameter(Mandatory = $false)]
    [string]$TestRecipient,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("ServerHealth", "MailboxStatistics", "DatabaseStatistics", "QueueStatistics")]
    [string]$ReportType = "ServerHealth",

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential
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

function Check-ExchangeServerHealth {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer
    )
    
    try {
        Write-Log -Message "Checking health of Exchange server '$ExchangeServer'..." -Level "INFO"
        
        $results = @{
            ServerName = $ExchangeServer
            Status = "Success"
            Services = @()
            Databases = @()
            Queues = @()
            ErrorMessage = $null
        }
        
        # Check Exchange services
        Write-Log -Message "Checking Exchange services..." -Level "INFO"
        $services = Get-Service -ComputerName $ExchangeServer | Where-Object { $_.DisplayName -like "Microsoft Exchange*" }
        
        foreach ($service in $services) {
            $results.Services += @{
                Name = $service.Name
                DisplayName = $service.DisplayName
                Status = $service.Status
                StartType = $service.StartType
            }
            
            if ($service.Status -ne "Running" -and $service.StartType -eq "Automatic") {
                Write-Log -Message "Service '$($service.DisplayName)' is not running." -Level "WARNING"
            }
        }
        
        # Check Exchange databases
        Write-Log -Message "Checking Exchange databases..." -Level "INFO"
        $databases = Get-MailboxDatabase -Server $ExchangeServer -Status
        
        foreach ($database in $databases) {
            $results.Databases += @{
                Name = $database.Name
                Server = $database.Server
                Status = $database.Mounted
                Size = (Get-MailboxDatabase -Identity $database.Name -Status).DatabaseSize
                BackupDate = $database.LastFullBackup
            }
            
            if (-not $database.Mounted) {
                Write-Log -Message "Database '$($database.Name)' is not mounted." -Level "WARNING"
            }
            
            if ($database.LastFullBackup -lt (Get-Date).AddDays(-1)) {
                Write-Log -Message "Database '$($database.Name)' has not been backed up in the last 24 hours." -Level "WARNING"
            }
        }
        
        # Check Exchange queues
        Write-Log -Message "Checking Exchange queues..." -Level "INFO"
        $queues = Get-Queue -Server $ExchangeServer
        
        foreach ($queue in $queues) {
            $results.Queues += @{
                Name = $queue.Identity
                Status = $queue.Status
                MessageCount = $queue.MessageCount
                NextHopDomain = $queue.NextHopDomain
            }
            
            if ($queue.MessageCount -gt 100) {
                Write-Log -Message "Queue '$($queue.Identity)' has $($queue.MessageCount) messages." -Level "WARNING"
            }
        }
        
        Write-Log -Message "Exchange server health check completed." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to check Exchange server health: $_" -Level "ERROR"
        return @{
            ServerName = $ExchangeServer
            Status = "Error"
            Services = @()
            Databases = @()
            Queues = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Analyze-ExchangeLogs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $true)]
        [DateTime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [DateTime]$EndDate,
        
        [Parameter(Mandatory = $true)]
        [string]$LogLevel
    )
    
    try {
        Write-Log -Message "Analyzing Exchange logs on '$ExchangeServer' from $StartDate to $EndDate..." -Level "INFO"
        
        $results = @{
            ServerName = $ExchangeServer
            Status = "Success"
            EventLogs = @()
            TransportLogs = @()
            ErrorMessage = $null
        }
        
        # Analyze Event Logs
        Write-Log -Message "Analyzing Event Logs..." -Level "INFO"
        
        $eventLogLevel = switch ($LogLevel) {
            "Error" { 1 }
            "Warning" { 2 }
            "Information" { 3 }
        }
        
        $eventLogs = Get-WinEvent -ComputerName $ExchangeServer -FilterHashtable @{
            LogName = "Application", "System"
            StartTime = $StartDate
            EndTime = $EndDate
            Level = 1..$eventLogLevel
            ProviderName = "MSExchange*"
        } -ErrorAction SilentlyContinue
        
        foreach ($event in $eventLogs) {
            $results.EventLogs += @{
                TimeCreated = $event.TimeCreated
                Id = $event.Id
                LevelDisplayName = $event.LevelDisplayName
                ProviderName = $event.ProviderName
                Message = $event.Message
            }
        }
        
        Write-Log -Message "Found $($results.EventLogs.Count) event log entries." -Level "INFO"
        
        # Analyze Transport Logs
        Write-Log -Message "Analyzing Transport Logs..." -Level "INFO"
        
        $transportLogPath = "\\$ExchangeServer\c$\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\MessageTracking"
        
        if (Test-Path $transportLogPath) {
            $transportLogs = Get-MessageTrackingLog -Server $ExchangeServer -Start $StartDate -End $EndDate -ResultSize Unlimited
            
            foreach ($log in $transportLogs) {
                if (($LogLevel -eq "Error" -and $log.EventId -like "*Fail*") -or
                    ($LogLevel -eq "Warning" -and ($log.EventId -like "*Fail*" -or $log.EventId -like "*Defer*")) -or
                    $LogLevel -eq "Information") {
                    
                    $results.TransportLogs += @{
                        Timestamp = $log.Timestamp
                        EventId = $log.EventId
                        Source = $log.Source
                        Sender = $log.Sender
                        Recipients = $log.Recipients
                        MessageSubject = $log.MessageSubject
                    }
                }
            }
            
            Write-Log -Message "Found $($results.TransportLogs.Count) transport log entries." -Level "INFO"
        }
        else {
            Write-Log -Message "Transport log path not found." -Level "WARNING"
        }
        
        Write-Log -Message "Exchange log analysis completed." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to analyze Exchange logs: $_" -Level "ERROR"
        return @{
            ServerName = $ExchangeServer
            Status = "Error"
            EventLogs = @()
            TransportLogs = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Check-ExchangeMailFlow {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $false)]
        [string]$TestMailbox,
        
        [Parameter(Mandatory = $false)]
        [string]$TestRecipient
    )
    
    try {
        Write-Log -Message "Checking mail flow on Exchange server '$ExchangeServer'..." -Level "INFO"
        
        $results = @{
            ServerName = $ExchangeServer
            Status = "Success"
            MailFlowTests = @()
            ErrorMessage = $null
        }
        
        # Check mail flow using Test-Mailflow
        Write-Log -Message "Testing internal mail flow..." -Level "INFO"
        
        $mailFlowTest = Test-Mailflow -TargetMailboxServer $ExchangeServer
        
        $results.MailFlowTests += @{
            TestName = "Internal Mail Flow"
            TestType = "Test-Mailflow"
            Status = $mailFlowTest.TestMailflowResult
            MessageLatency = $mailFlowTest.MessageLatencyTime
            Error = $mailFlowTest.Error
        }
        
        if ($mailFlowTest.TestMailflowResult -ne "Success") {
            Write-Log -Message "Internal mail flow test failed: $($mailFlowTest.Error)" -Level "WARNING"
        }
        else {
            Write-Log -Message "Internal mail flow test succeeded. Latency: $($mailFlowTest.MessageLatencyTime)" -Level "INFO"
        }
        
        # Test mail flow using specific mailboxes if provided
        if ($TestMailbox -and $TestRecipient) {
            Write-Log -Message "Testing mail flow from $TestMailbox to $TestRecipient..." -Level "INFO"
            
            $testSubject = "Mail Flow Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $testBody = "This is an automated mail flow test."
            
            try {
                Send-MailMessage -From $TestMailbox -To $TestRecipient -Subject $testSubject -Body $testBody -SmtpServer $ExchangeServer
                
                # Wait for message to be delivered
                Start-Sleep -Seconds 30
                
                # Check if message was delivered
                $messageTrackingLog = Get-MessageTrackingLog -Server $ExchangeServer -Start (Get-Date).AddMinutes(-5) -End (Get-Date) -Sender $TestMailbox -Recipients $TestRecipient -Subject $testSubject
                
                $delivered = $messageTrackingLog | Where-Object { $_.EventId -eq "DELIVER" }
                
                if ($delivered) {
                    $results.MailFlowTests += @{
                        TestName = "Custom Mail Flow"
                        TestType = "Send-MailMessage"
                        Status = "Success"
                        MessageLatency = ($delivered.Timestamp - $messageTrackingLog[0].Timestamp).TotalSeconds
                        Error = $null
                    }
                    
                    Write-Log -Message "Custom mail flow test succeeded." -Level "INFO"
                }
                else {
                    $results.MailFlowTests += @{
                        TestName = "Custom Mail Flow"
                        TestType = "Send-MailMessage"
                        Status = "Failed"
                        MessageLatency = $null
                        Error = "Message not delivered"
                    }
                    
                    Write-Log -Message "Custom mail flow test failed: Message not delivered" -Level "WARNING"
                }
            }
            catch {
                $results.MailFlowTests += @{
                    TestName = "Custom Mail Flow"
                    TestType = "Send-MailMessage"
                    Status = "Failed"
                    MessageLatency = $null
                    Error = $_.Exception.Message
                }
                
                Write-Log -Message "Custom mail flow test failed: $_" -Level "WARNING"
            }
        }
        
        Write-Log -Message "Mail flow check completed." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to check mail flow: $_" -Level "ERROR"
        return @{
            ServerName = $ExchangeServer
            Status = "Error"
            MailFlowTests = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Test-ExchangeConnectivity {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $false)]
        [string]$MailboxServer,
        
        [Parameter(Mandatory = $false)]
        [string]$HubTransportServer,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientAccessServer
    )
    
    try {
        Write-Log -Message "Testing Exchange connectivity for server '$ExchangeServer'..." -Level "INFO"
        
        $results = @{
            ServerName = $ExchangeServer
            Status = "Success"
            ConnectivityTests = @()
            ErrorMessage = $null
        }
        
        # Test Mailbox server connectivity
        if ($MailboxServer) {
            Write-Log -Message "Testing Mailbox server connectivity..." -Level "INFO"
            
            $mailboxTest = Test-MAPIConnectivity -Server $MailboxServer
            
            foreach ($test in $mailboxTest) {
                $results.ConnectivityTests += @{
                    TestName = "MAPI Connectivity"
                    Server = $test.Server
                    Database = $test.Database
                    Result = $test.Result
                    Error = $test.Error
                }
                
                if ($test.Result -ne "Success") {
                    Write-Log -Message "MAPI connectivity test failed for database $($test.Database): $($test.Error)" -Level "WARNING"
                }
                else {
                    Write-Log -Message "MAPI connectivity test succeeded for database $($test.Database)" -Level "INFO"
                }
            }
        }
        
        # Test Hub Transport server connectivity
        if ($HubTransportServer) {
            Write-Log -Message "Testing Hub Transport server connectivity..." -Level "INFO"
            
            $transportTest = Test-SmtpConnectivity -TargetServer $HubTransportServer
            
            foreach ($test in $transportTest) {
                $results.ConnectivityTests += @{
                    TestName = "SMTP Connectivity"
                    Server = $test.TargetServer
                    Scenario = $test.Scenario
                    Result = $test.Result
                    Error = $test.Error
                }
                
                if ($test.Result -ne "Success") {
                    Write-Log -Message "SMTP connectivity test failed for server $($test.TargetServer): $($test.Error)" -Level "WARNING"
                }
                else {
                    Write-Log -Message "SMTP connectivity test succeeded for server $($test.TargetServer)" -Level "INFO"
                }
            }
        }
        
        # Test Client Access server connectivity
        if ($ClientAccessServer) {
            Write-Log -Message "Testing Client Access server connectivity..." -Level "INFO"
            
            $clientAccessTest = Test-OutlookWebServices -Identity $ClientAccessServer
            
            foreach ($test in $clientAccessTest) {
                $results.ConnectivityTests += @{
                    TestName = "Outlook Web Services"
                    Server = $test.Identity
                    Scenario = $test.Scenario
                    Result = $test.Result
                    Error = $test.Error
                }
                
                if ($test.Result -ne "Success") {
                    Write-Log -Message "Outlook Web Services test failed for server $($test.Identity): $($test.Error)" -Level "WARNING"
                }
                else {
                    Write-Log -Message "Outlook Web Services test succeeded for server $($test.Identity)" -Level "INFO"
                }
            }
        }
        
        Write-Log -Message "Exchange connectivity testing completed." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to test Exchange connectivity: $_" -Level "ERROR"
        return @{
            ServerName = $ExchangeServer
            Status = "Error"
            ConnectivityTests = @()
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Generate-ExchangeReport {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExchangeServer,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportType
    )
    
    try {
        Write-Log -Message "Generating Exchange report for server '$ExchangeServer'..." -Level "INFO"
        
        # Create report directory if it doesn't exist
        $reportDir = Split-Path -Path $ReportPath -Parent
        if (-not (Test-Path -Path $reportDir)) {
            New-Item -Path $reportDir -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created report directory '$reportDir'." -Level "INFO"
        }
        
        $results = @{
            ServerName = $ExchangeServer
            Status = "Success"
            ReportType = $ReportType
            ReportPath = $ReportPath
            ErrorMessage = $null
        }
        
        # Generate report based on report type
        switch ($ReportType) {
            "ServerHealth" {
                Write-Log -Message "Generating Server Health report..." -Level "INFO"
                
                $serverInfo = Get-ExchangeServer -Identity $ExchangeServer
                $services = Get-Service -ComputerName $ExchangeServer | Where-Object { $_.DisplayName -like "Microsoft Exchange*" }
                $databases = Get-MailboxDatabase -Server $ExchangeServer -Status
                $queues = Get-Queue -Server $ExchangeServer
                
                $report = @"
<html>
<head>
    <title>Exchange Server Health Report - $ExchangeServer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0066cc; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .success { color: green; }
        .warning { color: orange; }
        .error { color: red; }
    </style>
</head>
<body>
    <h1>Exchange Server Health Report</h1>
    <p>Server: $ExchangeServer</p>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <h2>Server Information</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Version</th>
            <th>Edition</th>
            <th>Role</th>
            <th>Site</th>
        </tr>
        <tr>
            <td>$($serverInfo.Name)</td>
            <td>$($serverInfo.AdminDisplayVersion)</td>
            <td>$($serverInfo.Edition)</td>
            <td>$($serverInfo.ServerRole)</td>
            <td>$($serverInfo.Site)</td>
        </tr>
    </table>
    
    <h2>Exchange Services</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Display Name</th>
            <th>Status</th>
            <th>Start Type</th>
        </tr>
"@
                
                foreach ($service in $services) {
                    $statusClass = switch ($service.Status) {
                        "Running" { "success" }
                        "Stopped" { if ($service.StartType -eq "Automatic") { "error" } else { "warning" } }
                        default { "warning" }
                    }
                    
                    $report += @"
        <tr>
            <td>$($service.Name)</td>
            <td>$($service.DisplayName)</td>
            <td class="$statusClass">$($service.Status)</td>
            <td>$($service.StartType)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
    
    <h2>Exchange Databases</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Server</th>
            <th>Status</th>
            <th>Size</th>
            <th>Last Backup</th>
        </tr>
"@
                
                foreach ($database in $databases) {
                    $mountedClass = if ($database.Mounted) { "success" } else { "error" }
                    $backupClass = if ($database.LastFullBackup -gt (Get-Date).AddDays(-1)) { "success" } else { "warning" }
                    
                    $report += @"
        <tr>
            <td>$($database.Name)</td>
            <td>$($database.Server)</td>
            <td class="$mountedClass">$($database.Mounted)</td>
            <td>$((Get-MailboxDatabase -Identity $database.Name -Status).DatabaseSize)</td>
            <td class="$backupClass">$($database.LastFullBackup)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
    
    <h2>Exchange Queues</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Status</th>
            <th>Message Count</th>
            <th>Next Hop Domain</th>
        </tr>
"@
                
                foreach ($queue in $queues) {
                    $countClass = if ($queue.MessageCount -gt 100) { "warning" } else { "success" }
                    
                    $report += @"
        <tr>
            <td>$($queue.Identity)</td>
            <td>$($queue.Status)</td>
            <td class="$countClass">$($queue.MessageCount)</td>
            <td>$($queue.NextHopDomain)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
</body>
</html>
"@
                
                $report | Out-File -FilePath $ReportPath -Encoding UTF8
            }
            "MailboxStatistics" {
                Write-Log -Message "Generating Mailbox Statistics report..." -Level "INFO"
                
                $mailboxes = Get-Mailbox -Server $ExchangeServer
                $mailboxStats = $mailboxes | Get-MailboxStatistics
                
                $report = @"
<html>
<head>
    <title>Exchange Mailbox Statistics Report - $ExchangeServer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0066cc; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .success { color: green; }
        .warning { color: orange; }
        .error { color: red; }
    </style>
</head>
<body>
    <h1>Exchange Mailbox Statistics Report</h1>
    <p>Server: $ExchangeServer</p>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <h2>Mailbox Statistics</h2>
    <table>
        <tr>
            <th>Display Name</th>
            <th>Database</th>
            <th>Item Count</th>
            <th>Total Size (MB)</th>
            <th>Last Logon Time</th>
        </tr>
"@
                
                foreach ($stat in $mailboxStats) {
                    $sizeInMB = [math]::Round(($stat.TotalItemSize.Value.ToBytes() / 1MB), 2)
                    $sizeClass = if ($sizeInMB -gt 5000) { "warning" } else { "success" }
                    
                    $report += @"
        <tr>
            <td>$($stat.DisplayName)</td>
            <td>$($stat.Database)</td>
            <td>$($stat.ItemCount)</td>
            <td class="$sizeClass">$sizeInMB</td>
            <td>$($stat.LastLogonTime)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
</body>
</html>
"@
                
                $report | Out-File -FilePath $ReportPath -Encoding UTF8
            }
            "DatabaseStatistics" {
                Write-Log -Message "Generating Database Statistics report..." -Level "INFO"
                
                $databases = Get-MailboxDatabase -Server $ExchangeServer -Status
                
                $report = @"
<html>
<head>
    <title>Exchange Database Statistics Report - $ExchangeServer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0066cc; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .success { color: green; }
        .warning { color: orange; }
        .error { color: red; }
    </style>
</head>
<body>
    <h1>Exchange Database Statistics Report</h1>
    <p>Server: $ExchangeServer</p>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <h2>Database Statistics</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Server</th>
            <th>EDB File Size</th>
            <th>Available New MB</th>
            <th>Mailbox Count</th>
            <th>Last Full Backup</th>
        </tr>
"@
                
                foreach ($database in $databases) {
                    $dbStats = Get-MailboxDatabase -Identity $database.Name -Status
                    $mailboxCount = (Get-Mailbox -Database $database.Name).Count
                    
                    $backupClass = if ($database.LastFullBackup -gt (Get-Date).AddDays(-1)) { "success" } else { "warning" }
                    
                    $report += @"
        <tr>
            <td>$($database.Name)</td>
            <td>$($database.Server)</td>
            <td>$($dbStats.DatabaseSize)</td>
            <td>$($dbStats.AvailableNewMailboxSpace)</td>
            <td>$mailboxCount</td>
            <td class="$backupClass">$($database.LastFullBackup)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
</body>
</html>
"@
                
                $report | Out-File -FilePath $ReportPath -Encoding UTF8
            }
            "QueueStatistics" {
                Write-Log -Message "Generating Queue Statistics report..." -Level "INFO"
                
                $queues = Get-Queue -Server $ExchangeServer
                
                $report = @"
<html>
<head>
    <title>Exchange Queue Statistics Report - $ExchangeServer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0066cc; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
        .success { color: green; }
        .warning { color: orange; }
        .error { color: red; }
    </style>
</head>
<body>
    <h1>Exchange Queue Statistics Report</h1>
    <p>Server: $ExchangeServer</p>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <h2>Queue Statistics</h2>
    <table>
        <tr>
            <th>Identity</th>
            <th>Status</th>
            <th>Message Count</th>
            <th>Next Hop Domain</th>
            <th>Delivery Type</th>
            <th>Last Error</th>
        </tr>
"@
                
                foreach ($queue in $queues) {
                    $countClass = if ($queue.MessageCount -gt 100) { "warning" } else { "success" }
                    
                    $report += @"
        <tr>
            <td>$($queue.Identity)</td>
            <td>$($queue.Status)</td>
            <td class="$countClass">$($queue.MessageCount)</td>
            <td>$($queue.NextHopDomain)</td>
            <td>$($queue.DeliveryType)</td>
            <td>$($queue.LastError)</td>
        </tr>
"@
                }
                
                $report += @"
    </table>
</body>
</html>
"@
                
                $report | Out-File -FilePath $ReportPath -Encoding UTF8
            }
        }
        
        Write-Log -Message "Exchange report generated successfully at '$ReportPath'." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to generate Exchange report: $_" -Level "ERROR"
        return @{
            ServerName = $ExchangeServer
            Status = "Error"
            ReportType = $ReportType
            ReportPath = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Exchange server analysis process." -Level "INFO"
    
    # Connect to Exchange server
    $connectionResult = Connect-ExchangeServer -ExchangeServer $ExchangeServer -Credential $Credential
    
    if ($connectionResult.Status -ne "Success") {
        Write-Log -Message "Failed to connect to Exchange server. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Perform the requested action
    switch ($Action) {
        "CheckHealth" {
            $result = Check-ExchangeServerHealth -ExchangeServer $ExchangeServer
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to check Exchange server health. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "AnalyzeLogs" {
            $result = Analyze-ExchangeLogs -ExchangeServer $ExchangeServer -StartDate $StartDate -EndDate $EndDate -LogLevel $LogLevel
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to analyze Exchange logs. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "CheckMailFlow" {
            $result = Check-ExchangeMailFlow -ExchangeServer $ExchangeServer -TestMailbox $TestMailbox -TestRecipient $TestRecipient
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to check mail flow. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "TestConnectivity" {
            $result = Test-ExchangeConnectivity -ExchangeServer $ExchangeServer -MailboxServer $MailboxServer -HubTransportServer $HubTransportServer -ClientAccessServer $ClientAccessServer
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to test Exchange connectivity. Exiting..." -Level "ERROR"
                exit 1
            }
        }
        "GenerateReport" {
            # Validate required parameters
            if (-not $ReportPath) {
                Write-Log -Message "ReportPath parameter is required for GenerateReport action. Exiting..." -Level "ERROR"
                exit 1
            }
            
            $result = Generate-ExchangeReport -ExchangeServer $ExchangeServer -ReportPath $ReportPath -ReportType $ReportType
            
            if ($result.Status -ne "Success") {
                Write-Log -Message "Failed to generate Exchange report. Exiting..." -Level "ERROR"
                exit 1
            }
        }
    }
    
    # Disconnect from Exchange server
    Disconnect-ExchangeServer -Session $connectionResult.Session
    
    Write-Log -Message "Exchange server analysis process completed successfully." -Level "INFO"
}
catch {
    Write-Log -Message "An error occurred during Exchange server analysis process: $_" -Level "ERROR"
    
    # Attempt to disconnect from Exchange server
    if ($connectionResult -and $connectionResult.Session) {
        Disconnect-ExchangeServer -Session $connectionResult.Session
    }
    
    exit 1
}
