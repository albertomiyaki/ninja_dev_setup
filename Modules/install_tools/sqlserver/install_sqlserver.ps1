# Install SQL Server Express
$ErrorActionPreference = 'Stop'

# Check OS version
if ([Version] (Get-CimInstance Win32_OperatingSystem).Version -lt [version] "10.0.0.0") {
    Write-Error "SQL Server Express requires a minimum of Windows 10"
}

# Prompt for SA password securely with confirmation
function Get-SecurePasswordWithConfirmation {
    $passwordsMatch = $false
    $password = $null
    
    while (-not $passwordsMatch) {
        $securePassword1 = Read-Host "Enter SA password for SQL Server" -AsSecureString
        $securePassword2 = Read-Host "Confirm SA password" -AsSecureString
        
        # Convert secure strings to plain text for comparison
        $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword1)
        $password1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR1)
        
        $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword2)
        $password2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR2)
        
        # Check if passwords match
        if ($password1 -eq $password2) {
            $passwordsMatch = $true
            $password = $password1
            
            # Clear second password variable
            $password2 = $null
        } else {
            Write-Host "Passwords do not match. Please try again." -ForegroundColor Yellow
        }
        
        # Clear secure string objects from memory
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR1)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR2)
    }
    
    return $password
}

# Get SA password with confirmation
$saPassword = Get-SecurePasswordWithConfirmation

# Define paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path -Path $scriptPath -ChildPath "sqlserver_configuration.ini"
$networkPath = "Z:\installers\SQLEXPR_x64_ENU.exe"
$tempDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Temp\sqlserverexpress2019"
$extractPath = Join-Path -Path $tempDir -ChildPath "SQLEXPR"
$setupPath = Join-Path -Path $extractPath -ChildPath "setup.exe"
$tempConfigPath = Join-Path -Path $tempDir -ChildPath "configuration.ini"

# Get the installer filename from the network path
$installerFileName = [System.IO.Path]::GetFileName($networkPath)
$fileFullPath = Join-Path -Path $tempDir -ChildPath $installerFileName

# Verify network path is accessible
if (-not (Test-Path -Path $networkPath -PathType Leaf)) {
    Write-Error "Cannot access SQL Server installer at: $networkPath"
    exit 1
}

# Create temp directory if it doesn't exist
if (-not (Test-Path -Path $tempDir -PathType Container)) {
    $null = New-Item -Path $tempDir -ItemType Directory -Force
}

# Copy installer from network to local temp with error handling
Write-Host "Copying SQL Server Express installer from network..."
try {
    Copy-Item -Path $networkPath -Destination $fileFullPath -Force -ErrorAction Stop
    Write-Host "Installer copied successfully to: $fileFullPath"
} catch {
    Write-Error "Failed to copy SQL Server installer: $_"
    exit 1
}

# Verify configuration file exists
if (-not (Test-Path -Path $configPath -PathType Leaf)) {
    Write-Error "Configuration file not found at: $configPath"
    exit 1
}

# Copy and modify the configuration file with the password
Write-Host "Preparing configuration file..."
try {
    $configContent = Get-Content -Path $configPath -Raw -ErrorAction Stop
    $configContent = $configContent -replace '(SAPWD=")[^"]*(".*)', "`$1$saPassword`$2"
    Set-Content -Path $tempConfigPath -Value $configContent -ErrorAction Stop
    Write-Host "Configuration file prepared successfully"
} catch {
    Write-Error "Failed to prepare configuration file: $_"
    exit 1
}

# Clear the password variable for security
$saPassword = $null

# Verify the installer file exists in the temp location
if (-not (Test-Path -Path $fileFullPath -PathType Leaf)) {
    Write-Error "Installer not found at: $fileFullPath"
    exit 1
}

Write-Host "Full Path: $fileFullPath"
Write-Host "Extract Path: $extractPath"

# Extract silently
Write-Host "Extracting..."
try {
    $extractProcess = Start-Process -FilePath $fileFullPath -ArgumentList "/Q", "/x:`"$extractPath`"" -PassThru -Wait -NoNewWindow -ErrorAction Stop
    if ($extractProcess.ExitCode -ne 0) {
        Write-Error "Extraction failed with exit code: $($extractProcess.ExitCode)"
        exit 1
    }
    Write-Host "Extraction completed successfully"
} catch {
    Write-Error "Failed to start extraction process: $_"
    exit 1
}

# Verify setup.exe exists
if (-not (Test-Path -Path $setupPath -PathType Leaf)) {
    Write-Error "Setup.exe not found at: $setupPath"
    exit 1
}

# Install using the configuration file
Write-Host "Installing SQL Server Express..."
try {
    $installProcess = Start-Process -FilePath $setupPath -ArgumentList "/ConfigurationFile=`"$tempConfigPath`"" -PassThru -Wait -NoNewWindow -ErrorAction Stop
    $exitCode = $installProcess.ExitCode
    
    # Check valid exit codes
    if ($exitCode -notin @(0, 3010, 1116)) {
        Write-Error "Installation failed with exit code: $exitCode"
        exit 1
    } else {
        if ($exitCode -eq 3010) {
            Write-Host "Installation successful. A system restart is required."
        } elseif ($exitCode -eq 1116) {
            Write-Host "Installation successful, but the SQL Browser service failed to start."
        } else {
            Write-Host "Installation successful."
        }
    }
} catch {
    Write-Error "Failed to start installation process: $_"
    exit 1
}

# Cleanup
Write-Host "Removing temporary files..."
try {
    if (Test-Path -Path $tempDir -PathType Container) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction Stop
        Write-Host "Temporary files removed successfully"
    }
} catch {
    Write-Warning "Failed to remove some temporary files: $_"
}

Write-Host "SQL Server Express installation complete"