# Create sqlserver login user

$ErrorActionPreference = 'Stop'

# Configuration variables
$sqlInstance = "$env:COMPUTERNAME\SQLEXPRESS"
$defaultDatabase = "master"
$logFile = Join-Path -Path $env:TEMP -ChildPath "SqlLoginCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

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
        $service = Get-Service -Name "MSSQL`$SQLEXPRESS" -ErrorAction SilentlyContinue
        return ($service -ne $null)
    }
    catch {
        return $false
    }
}

# Function to prompt for secure password with confirmation
Function Get-SecurePasswordWithConfirmation {
    $passwordsMatch = $false
    $securePassword = $null
    
    while (-not $passwordsMatch) {
        $securePassword = Read-Host "Enter password for the SQL login" -AsSecureString
        $confirmPassword = Read-Host "Confirm password" -AsSecureString
        
        # Convert SecureString to plain text for comparison only
        $bstr1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $bstr2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmPassword)
        
        try {
            $password1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr1)
            $password2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr2)
            
            if ($password1 -eq $password2) {
                $passwordsMatch = $true
            }
            else {
                Write-Host "Passwords do not match. Please try again." -ForegroundColor Red
            }
        }
        finally {
            # Clean up the unprotected strings from memory
            if ($bstr1 -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1) }
            if ($bstr2 -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2) }
            [System.GC]::Collect()
        }
    }
    
    return $securePassword
}

# Self-elevate the script if required
if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
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
        Write-Host "Failed to elevate script: $_" -ForegroundColor Red
        exit 1
    }
    
    # Exit the non-elevated script
    exit
}

# Create log file
try {
    $null = New-Item -Path $logFile -ItemType File -Force
    Write-Log "SQL login creation log started"
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

# Get SQL login name from user with default value
$defaultLoginName = "wmuser"
$loginPrompt = Read-Host "Enter the SQL login name to create [default: $defaultLoginName]"

# Use default if user just pressed Enter
$loginName = if ([string]::IsNullOrWhiteSpace($loginPrompt)) { 
    Write-Log "Using default login name: '$defaultLoginName'" -Level Info
    $defaultLoginName 
} else { 
    $loginPrompt 
}

# Get secure password with confirmation
Write-Log "Please provide a secure password for the login '$loginName'" -Level Info
$securePassword = Get-SecurePasswordWithConfirmation

# Create credential object
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $loginName, $securePassword

# Create SQL login
Write-Log "Creating SQL login '$loginName' on instance '$sqlInstance'..."
try {
    Add-SqlLogin -ServerInstance $sqlInstance -LoginName $loginName -LoginType "SqlLogin" -DefaultDatabase $defaultDatabase -Enable -GrantConnectSql -LoginPSCredential $credential -ErrorAction Stop
    Write-Log "SQL login '$loginName' created successfully" -Level Info
}
catch {
    Write-Log "Failed to create SQL login: $_" -Level Error
    # Check for specific error conditions
    if ($_.Exception.Message -like "*already exists*") {
        Write-Log "A login with the name '$loginName' already exists." -Level Warning
    }
    exit 1
}

Write-Log "SQL login creation process completed"