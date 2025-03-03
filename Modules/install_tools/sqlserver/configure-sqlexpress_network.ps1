$ErrorActionPreference = 'Stop'

# Configuration variables
$sqlInstance = "$env:COMPUTERNAME\SQLEXPRESS"
$sqlServiceName = "MSSQL`$SQLEXPRESS"
$tcpPort = "1433"
$logFile = Join-Path -Path $env:TEMP -ChildPath "SQLExpressConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write log messages
Function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with appropriate color
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
    }
    
    # Append to log file
    Add-Content -Path $logFile -Value $logMessage
}

# Function to check if running as Administrator
Function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check if SQL Express is installed
Function Test-SqlExpressInstalled {
    try {
        $service = Get-Service -Name $sqlServiceName -ErrorAction SilentlyContinue
        return ($service -ne $null)
    }
    catch {
        return $false
    }
}

# Self-elevate the script if required
if (-not (Test-Administrator)) {
    Write-Log "Requesting administrative privileges..." -Level Warning
    
    # Build the script path and arguments
    $scriptPath = $MyInvocation.MyCommand.Definition
    $args = $MyInvocation.UnboundArguments
    
    # Construct the arguments string for the elevated process
    $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    if ($args.Count -gt 0) {
        $argString += " " + ($args -join ' ')
    }
    
    # Start PowerShell as administrator with this script
    try {
        Start-Process -FilePath PowerShell.exe -ArgumentList $argString -Verb RunAs
    }
    catch {
        Write-Log "Failed to elevate script: $_" -Level Error
        exit 1
    }
    
    # Exit the non-elevated script
    exit
}

# Create log file
try {
    $null = New-Item -Path $logFile -ItemType File -Force
    Write-Log "SQL Express configuration log started"
    Write-Log "Log file location: $logFile"
}
catch {
    Write-Host "Failed to create log file: $_" -ForegroundColor Red
    exit 1
}

# Verify SQL Express is installed
if (-not (Test-SqlExpressInstalled)) {
    Write-Log "SQL Express is not installed on this machine." -Level Error
    exit 1
}

# Import SqlServer module
Write-Log "Importing SqlServer module..."
try {
    Import-Module -Name "SqlServer" -ErrorAction Stop
    Write-Log "SqlServer module imported successfully"
}
catch {
    Write-Log "Failed to import SqlServer module: $_" -Level Error
    Write-Log "Please ensure the SqlServer module is installed. You can install it with: Install-Module -Name SqlServer" -Level Warning
    exit 1
}

# Configure TCP protocol
Write-Log "Configuring TCP protocol for SQL Express..."
try {
    Push-Location
    
    $wmi = New-Object ('Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer') $env:COMPUTERNAME
    $uri = "ManagedComputer[@Name='$env:COMPUTERNAME']/ServerInstance[@Name='SQLEXPRESS']/ServerProtocol[@Name='Tcp']"
    $Tcp = $wmi.GetSmoObject($uri)
    
    # Enable TCP protocol
    if (-not $Tcp.IsEnabled) {
        $Tcp.IsEnabled = $true
        Write-Log "TCP protocol enabled"
    }
    else {
        Write-Log "TCP protocol was already enabled"
    }
    
    # Set port to 1433
    $ipAllUri = $uri + "/IPAddress[@Name='IPAll']"
    $ipAll = $wmi.GetSmoObject($ipAllUri)
    $currentPort = $ipAll.IPAddressProperties[1].Value
    
    if ($currentPort -ne $tcpPort) {
        $ipAll.IPAddressProperties[1].Value = $tcpPort
        Write-Log "TCP port set to $tcpPort (was: $currentPort)"
    }
    else {
        Write-Log "TCP port was already set to $tcpPort"
    }
    
    # Apply TCP changes
    $Tcp.Alter()
    Write-Log "TCP configuration applied successfully"
}
catch {
    Write-Log "Failed to configure TCP protocol: $_" -Level Error
    if ($null -ne (Get-Location -Stack)) {
        Pop-Location
    }
    exit 1
}

# Configure Mixed Authentication Mode
Write-Log "Configuring Mixed Authentication Mode..."
try {
    $sql = [Microsoft.SqlServer.Management.Smo.Server]::new($sqlInstance)
    $currentMode = $sql.Settings.LoginMode
    
    if ($currentMode -ne 'Mixed') {
        $sql.Settings.LoginMode = 'Mixed'
        $sql.Alter()
        Write-Log "Authentication mode changed to Mixed Mode (was: $currentMode)"
    }
    else {
        Write-Log "Authentication mode was already set to Mixed Mode"
    }
}
catch {
    Write-Log "Failed to configure Mixed Authentication Mode: $_" -Level Error
    if ($null -ne (Get-Location -Stack)) {
        Pop-Location
    }
    exit 1
}

# Restart SQL Service
Write-Log "Restarting SQL Express service to apply changes..."
try {
    Restart-Service -Name $sqlServiceName -Force
    Start-Sleep -Seconds 5
    $service = Get-Service -Name $sqlServiceName
    
    if ($service.Status -eq 'Running') {
        Write-Log "SQL Express service restarted successfully"
    }
    else {
        Write-Log "SQL Express service is not running after restart attempt (Status: $($service.Status))" -Level Warning
    }
}
catch {
    Write-Log "Failed to restart SQL Express service: $_" -Level Error
    if ($null -ne (Get-Location -Stack)) {
        Pop-Location
    }
    exit 1
}

# Clean up
if ($null -ne (Get-Location -Stack)) {
    Pop-Location
}

Write-Log "SQL Express configuration completed successfully"
Write-Log "Instance: $sqlInstance"
Write-Log "TCP Port: $tcpPort"
Write-Log "Authentication Mode: Mixed"