# Install SQL Server Management Studio (SSMS)
$ErrorActionPreference = 'Stop'

# Configuration variables - easy to modify
$mediaPath = "Z:\installers\SSMS-Setup-ENU.exe"
$tempFolderName = "sql-server-management-studio\20.2"
$installerFileName = "SSMS-Setup-ENU.exe"
$logFileName = "SSMS.MsiInstall.log"

# Derived paths
$tempDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Temp\$tempFolderName"
$installerPath = Join-Path -Path $tempDir -ChildPath $installerFileName
$logFile = Join-Path -Path $tempDir -ChildPath $logFileName

# Create temp directory if it doesn't exist
if (-not (Test-Path -Path $tempDir -PathType Container)) {
    Write-Host "Creating temporary directory: $tempDir"
    $null = New-Item -Path $tempDir -ItemType Directory -Force
}

# Verify source path is accessible
if (-not (Test-Path -Path $mediaPath -PathType Leaf)) {
    Write-Error "Cannot access SSMS installer at: $mediaPath"
    exit 1
}

# Copy installer from media location to temporary folder
Write-Host "Copying SSMS installer to temporary location..."
try {
    Copy-Item -Path $mediaPath -Destination $installerPath -Force -ErrorAction Stop
    Write-Host "Installer copied successfully to: $installerPath"
} catch {
    Write-Error "Failed to copy SSMS installer: $_"
    exit 1
}

# Verify the installer file exists in the temp location
if (-not (Test-Path -Path $installerPath -PathType Leaf)) {
    Write-Error "Installer not found at: $installerPath"
    exit 1
}

# Installation arguments
$silentArgs = @(
    "/quiet",
    "/install",
    "/norestart",
    "/log `"$logFile`""
)

# Valid exit codes
$validExitCodes = @(0, 3010, 1641)

# Install SSMS
Write-Host "Installing SQL Server Management Studio..."
try {
    $installProcess = Start-Process -FilePath $installerPath -ArgumentList $silentArgs -PassThru -Wait -NoNewWindow -ErrorAction Stop
    $exitCode = $installProcess.ExitCode
    
    # Check exit code
    if ($exitCode -notin $validExitCodes) {
        Write-Error "Installation failed with exit code: $exitCode"
        Write-Host "Installation log available at: $logFile"
        exit 1
    } else {
        if ($exitCode -eq 3010 -or $exitCode -eq 1641) {
            Write-Host "Installation successful. A system restart is required."
        } else {
            Write-Host "Installation successful."
        }
    }
} catch {
    Write-Error "Failed to start installation process: $_"
    exit 1
}

# Remove installer but keep log
Write-Host "Removing installer file..."
try {
    if (Test-Path -Path $installerPath -PathType Leaf) {
        Remove-Item -Path $installerPath -Force -ErrorAction Stop
        Write-Host "Installer removed successfully"
    }

    # Confirm log file exists
    if (Test-Path -Path $logFile -PathType Leaf) {
        Write-Host "Installation log saved at: $logFile"
    } else {
        Write-Warning "Log file not found at expected location: $logFile"
    }
} catch {
    Write-Warning "Failed to remove installer file: $_"
}

Write-Host "SQL Server Management Studio installation complete"