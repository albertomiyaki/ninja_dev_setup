# TailWm setup

# This script creates tailwm.bat and tailwm.config files for tailing log files with a configurable tool
# and adds the batch file to the system PATH

# Configuration section - Edit these variables as needed
$tailToolName = "baretail"                          # Default tool - will be configurable in the bat file
$batchFilePath = "$env:USERPROFILE\Tools\tailwm\"   # Where to save the files
$configFileName = "tailwm.config" # Name of the companion configuration file

# Sample log mappings (will be written to the config file if it doesn't exist)
$defaultLogMappings = @{
    "myapp1" = "c:\program files\myapp1\log\server.log"
    "myapp2" = "c:\program files\myapp2\log\server.log"
    "myapp3" = "c:\program files\myapp3\log\server.log"
    # Add more mappings as needed
}

# Create directory if it doesn't exist
if (-not (Test-Path -Path $batchFilePath)) {
    New-Item -ItemType Directory -Path $batchFilePath -Force | Out-Null
    Write-Host "Created directory: $batchFilePath" -ForegroundColor Green
}

# Define file paths
$batchFile = Join-Path -Path $batchFilePath -ChildPath "tailwm.bat"
$configFile = Join-Path -Path $batchFilePath -ChildPath $configFileName

# Check if config file already exists and prompt for confirmation if it does
$createConfigFile = $true
if (Test-Path -Path $configFile) {
    $response = Read-Host "Configuration file already exists. Do you want to overwrite it? (Y/N)"
    $createConfigFile = $response -in 'Y', 'y', 'Yes', 'yes'
}

# Create the config file if needed
if ($createConfigFile) {
    # Build config file content
    $configContent = "# TailWm configuration file`r`n"
    $configContent += "# Format: ""appname"" = ""path\to\logfile.log""`r`n"
    $configContent += "# Lines starting with # are ignored`r`n`r`n"
    
    foreach ($key in $defaultLogMappings.Keys) {
        $value = $defaultLogMappings[$key]
        $configContent += """$key"" = ""$value""`r`n"
    }
    
    # Write config file
    $configContent | Out-File -FilePath $configFile -Encoding UTF8 -Force
    Write-Host "Created/Updated configuration file: $configFile" -ForegroundColor Green
} else {
    Write-Host "Keeping existing configuration file: $configFile" -ForegroundColor Cyan
}

# Build batch file content
$batchContent = @"
@echo off
:: Configuration - Change this to your preferred log viewer
set LOGTOOL=$tailToolName
:: Log mapping logic using companion config file: $configFileName

:: Check if config file exists
if not exist "%~dp0$configFileName" (
    echo Error: Configuration file not found: %~dp0$configFileName
    echo Please run the TailWm setup script again to create it.
    exit /b 1
)

:: Process the parameter
if "%~1"=="" goto :help

:: Read config file and find the requested log file
setlocal EnableDelayedExpansion
set FOUND=0
for /f "usebackq tokens=1,2 delims==" %%a in ("%~dp0$configFileName") do (
    set line=%%a
    :: Skip comment lines
    if not "!line:~0,1!"=="#" (
        :: Remove quotes and trim spaces
        set key=%%a
        set key=!key:"=!
        set key=!key: =!
        
        if "!key!"=="%~1" (
            set FOUND=1
            :: Remove quotes from the value
            set logpath=%%b
            set logpath=!logpath:"=!
            set logpath=!logpath: =!
            
            :: Launch the log viewer
            start "" %LOGTOOL% "!logpath!"
            echo Tailing !logpath!
            exit /b 0
        )
    )
)

if %FOUND%==0 (
    echo Unknown logfile: %1
    goto :help
)
exit /b 0

:help
echo Available options:
for /f "usebackq tokens=1,2 delims==" %%a in ("%~dp0$configFileName") do (
    set line=%%a
    :: Skip comment lines
    if not "!line:~0,1!"=="#" (
        :: Remove quotes and trim spaces
        set key=%%a
        set key=!key:"=!
        set key=!key: =!
        
        :: Remove quotes from the value
        set logpath=%%b
        set logpath=!logpath:"=!
        set logpath=!logpath: =!
        
        echo !key!  --^> !logpath!
    )
)
echo.
echo Current log tool: %LOGTOOL%
echo To change the log tool, edit the LOGTOOL variable at the top of this batch file
echo To add more log files, edit the configuration file: %~dp0$configFileName
exit /b 1
"@

# Check if batch file already exists and prompt for confirmation if it does
$createBatchFile = $true
if (Test-Path -Path $batchFile) {
    $response = Read-Host "Batch file already exists. Do you want to overwrite it? (Y/N)"
    $createBatchFile = $response -in 'Y', 'y', 'Yes', 'yes'
}

# Write batch file if confirmed
if ($createBatchFile) {
    $batchContent | Out-File -FilePath $batchFile -Encoding ASCII -Force
    Write-Host "Created/Updated batch file: $batchFile" -ForegroundColor Green
    Write-Host "Default log tool set to: $tailToolName" -ForegroundColor Green
} else {
    Write-Host "Keeping existing batch file: $batchFile" -ForegroundColor Cyan
}

# Add to PATH if not already there
$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($currentPath -notlike "*$batchFilePath*") {
    [Environment]::SetEnvironmentVariable("PATH", "$currentPath;$batchFilePath", "User")
    Write-Host "Added $batchFilePath to user PATH" -ForegroundColor Green
    Write-Host "Please restart your command prompt or PowerShell for PATH changes to take effect" -ForegroundColor Yellow
} else {
    Write-Host "$batchFilePath is already in PATH" -ForegroundColor Cyan
}

# Show usage information
Write-Host "`nSetup complete! You can now use:" -ForegroundColor Green
Write-Host "tailwm <logfile>" -ForegroundColor White
Write-Host "Where <logfile> is one of the entries in the configuration file:" -ForegroundColor White

# Read and display the actual configuration file entries
if (Test-Path -Path $configFile) {
    $configEntries = Get-Content -Path $configFile | Where-Object { $_ -match '^\s*".*"\s*=\s*".*"' }
    foreach ($entry in $configEntries) {
        if ($entry -match '^\s*"(.*)"\s*=\s*"(.*)"') {
            $app = $matches[1]
            $logPath = $matches[2]
            Write-Host "  $app  -->  $logPath" -ForegroundColor White
        }
    }
}

Write-Host "`nTo change the log tool:" -ForegroundColor Cyan
Write-Host "Edit the LOGTOOL variable at the top of $batchFile" -ForegroundColor White
Write-Host "`nTo modify log mappings:" -ForegroundColor Cyan
Write-Host "Edit the configuration file: $configFile" -ForegroundColor White
Write-Host "Format: ""appname"" = ""path\to\logfile.log""" -ForegroundColor White