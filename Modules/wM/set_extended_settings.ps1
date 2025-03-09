# Set Extended Settings

# Reads configuration files (*.conf) in server_settings folder and applies settings to target files
# First line must be a comment with target path

# Check if the script is running with administrator privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Self-elevate the script if required
if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    # Build arguments to pass to the elevated process
    $argString = ""
    if ($MyInvocation.BoundParameters.Count -gt 0) {
        $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $value = $MyInvocation.BoundParameters[$key]
            if ($value -is [System.String]) {
                $argString += " -$key `"$value`""
            } else {
                $argString += " -$key $value"
            }
        }
    } else {
        $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    }
    
    # Start PowerShell as administrator with this script
    Start-Process -FilePath PowerShell.exe -ArgumentList $argString -Verb RunAs
    
    # Exit the non-elevated script
    exit
}

# Get the script's root directory
$scriptRoot = $PSScriptRoot
$serverSettingsPath = Join-Path $scriptRoot "server_settings"

# Function for logging to console
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Information", "Warning", "Error", "Success")]
        [string]$Type = "Information"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # Write to console with color coding
    switch ($Type) {
        "Information" { Write-Host $logMessage -ForegroundColor Gray }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }
}

# Function to read a setting file and extract destination path and settings
function Get-SettingsFromFile {
    param (
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Log "Settings file not found: $FilePath" -Type Error
            return $null
        }
        
        $lines = Get-Content -Path $FilePath
        
        if ($lines.Count -eq 0) {
            Write-Log "Settings file is empty: $FilePath" -Type Warning
            return $null
        }
        
        # First line should be a comment with the target path
        $targetPath = $null
        if ($lines[0] -match '^#\s*(.+)$') {
            $targetPath = $matches[1].Trim()
        } else {
            Write-Log "Invalid settings file format. First line must be a comment with target path: $FilePath" -Type Error
            return $null
        }
        
        # Parse settings (all non-empty, non-comment lines after the first line)
        $settings = @()
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            
            # Skip empty lines and comments
            if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith("#")) {
                continue
            }
            
            # Parse setting line (key=value)
            if ($line -match '^([^=]+)=(.*)$') {
                $key = $matches[1].Trim()
                $value = $matches[2]
                
                $settings += @{
                    Key = $key
                    Value = $value
                    LineNumber = $i + 1  # +1 for human-readable line number
                }
            } else {
                Write-Log "Invalid setting format at line $($i+1): $line" -Type Warning
            }
        }
        
        return @{
            SourceFile = $FilePath
            TargetPath = $targetPath
            Settings = $settings
        }
    }
    catch {
        Write-Log "Error processing settings file $FilePath: $_" -Type Error
        return $null
    }
}

# Function to get current value of a setting from a file
function Get-CurrentSettingValue {
    param (
        [string]$FilePath,
        [string]$SettingKey
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            return "<file not found>"
        }
        
        $content = Get-Content -Path $FilePath -ErrorAction SilentlyContinue
        
        if ($null -eq $content) {
            return "<unable to read file>"
        }
        
        # Look for the setting matching exactly at the beginning of a line
        foreach ($line in $content) {
            if ($line -match "^$([regex]::Escape($SettingKey))=(.*)$") {
                return $matches[1]
            }
        }
        
        return "<not set>"
    }
    catch {
        return "<error: $_>"
    }
}

# Function to update settings in target file
function Update-TargetFile {
    param (
        [string]$TargetPath,
        [array]$SelectedSettings
    )
    
    try {
        # Check if target file exists
        $targetExists = Test-Path $TargetPath
        
        if ($targetExists) {
            # Read the entire file content
            $content = Get-Content -Path $TargetPath -Raw
            
            # If file doesn't end with newline, add one
            if (-not $content.EndsWith("`n")) {
                $content += "`r`n"
            }
        } else {
            # Create a new file with empty content
            $content = ""
            
            # Ensure target directory exists
            $targetDir = Split-Path -Parent $TargetPath
            if (-not (Test-Path $targetDir)) {
                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                Write-Log "Created directory: $targetDir" -Type Information
            }
        }
        
        $updatedCount = 0
        $addedCount = 0
        
        # Process each selected setting
        foreach ($setting in $SelectedSettings) {
            $settingLine = "$($setting.Key)=$($setting.Value)"
            $pattern = "^$([regex]::Escape($setting.Key))=.*$"
            
            # Check if the setting already exists in the file
            if ($content -match "(?m)$pattern") {
                # Update existing setting
                $content = $content -replace "(?m)$pattern", $settingLine
                $updatedCount++
            } else {
                # Add new setting
                $content += "$settingLine`r`n"
                $addedCount++
            }
        }
        
        # Write the updated content back to the file
        Set-Content -Path $TargetPath -Value $content
        
        return @{
            Success = $true
            Updated = $updatedCount
            Added = $addedCount
            Message = "Updated $updatedCount setting(s), added $addedCount setting(s)"
        }
    }
    catch {
        Write-Log "Error updating target file $TargetPath: $_" -Type Error
        return @{
            Success = $false
            Updated = 0
            Added = 0
            Message = "Error: $_"
        }
    }
}

# Function to find all server configuration files
function Get-ServerConfigFiles {
    try {
        if (-not (Test-Path $serverSettingsPath)) {
            Write-Log "Server settings folder not found: $serverSettingsPath" -Type Error
            return @()
        }
        
        # Find all .conf files in the server_settings folder
        $configFiles = Get-ChildItem -Path $serverSettingsPath -Filter "*.conf" -File
        
        if ($configFiles.Count -eq 0) {
            Write-Log "No configuration files found in $serverSettingsPath" -Type Warning
            return @()
        }
        
        Write-Log "Found $($configFiles.Count) configuration files" -Type Success
        return $configFiles
    }
    catch {
        Write-Log "Error searching for configuration files: $_" -Type Error
        return @()
    }
}

# Function to prompt user for selection from a list
function Get-UserSelection {
    param (
        [array]$Items,
        [string]$Prompt,
        [switch]$MultiSelect
    )
    
    if ($Items.Count -eq 0) {
        return @()
    }
    
    Write-Host "`n$Prompt" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host "[$($i+1)] $($Items[$i])"
    }
    
    if ($MultiSelect) {
        Write-Host "`nEnter numbers separated by commas, 'all' to select all, or 'none' to cancel: " -ForegroundColor Yellow -NoNewline
        $selection = Read-Host
        
        if ($selection -eq "none") {
            return @()
        }
        
        if ($selection -eq "all") {
            return $Items
        }
        
        $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 }
        return $selectedIndices | Where-Object { $_ -ge 0 -and $_ -lt $Items.Count } | ForEach-Object { $Items[$_] }
    }
    else {
        Write-Host "`nEnter a number, or 'none' to cancel: " -ForegroundColor Yellow -NoNewline
        $selection = Read-Host
        
        if ($selection -eq "none") {
            return $null
        }
        
        if ($selection -match '^\d+$') {
            $index = [int]$selection - 1
            if ($index -ge 0 -and $index -lt $Items.Count) {
                return $Items[$index]
            }
        }
        
        Write-Host "Invalid selection" -ForegroundColor Red
        return $null
    }
}

# Main function to run the interactive settings updater
function Start-InteractiveSettingsUpdater {
    Clear-Host
    Write-Host "=== Interactive Settings Updater ===" -ForegroundColor Cyan
    Write-Host "Script location: $scriptRoot" -ForegroundColor Gray
    Write-Host "Server settings folder: $serverSettingsPath`n" -ForegroundColor Gray
    
    # Get all configuration files
    $configFiles = Get-ServerConfigFiles
    
    if ($configFiles.Count -eq 0) {
        Write-Host "No configuration files found. Exiting." -ForegroundColor Red
        return
    }
    
    # Display list of configuration files and let user select one
    $fileOptions = $configFiles | ForEach-Object { $_.Name }
    $selectedFileName = Get-UserSelection -Items $fileOptions -Prompt "Select a configuration file to process:"
    
    if ($null -eq $selectedFileName) {
        Write-Host "No file selected. Exiting." -ForegroundColor Yellow
        return
    }
    
    $selectedFilePath = Join-Path $serverSettingsPath $selectedFileName
    Write-Host "Processing $selectedFileName..." -ForegroundColor Cyan
    
    # Parse the selected configuration file
    $fileSettings = Get-SettingsFromFile -FilePath $selectedFilePath
    
    if ($null -eq $fileSettings) {
        Write-Host "Failed to process configuration file. Exiting." -ForegroundColor Red
        return
    }
    
    Write-Host "`nTarget file: $($fileSettings.TargetPath)" -ForegroundColor Green
    Write-Host "Found $($fileSettings.Settings.Count) settings in configuration file`n" -ForegroundColor Green
    
    # Get current values from target file for comparison
    $targetExists = Test-Path $fileSettings.TargetPath
    if ($targetExists) {
        Write-Host "Target file exists. Getting current values for comparison..." -ForegroundColor Cyan
    } else {
        Write-Host "Target file doesn't exist. It will be created." -ForegroundColor Yellow
    }
    
    # Display settings with current values and let user select which to apply
    $settingOptions = @()
    foreach ($setting in $fileSettings.Settings) {
        $currentValue = $targetExists ? (Get-CurrentSettingValue -FilePath $fileSettings.TargetPath -SettingKey $setting.Key) : "<file will be created>"
        $settingOptions += "[Line $($setting.LineNumber)] $($setting.Key) = $($setting.Value) (Current: $currentValue)"
        $setting.CurrentValue = $currentValue
    }
    
    $selectedSettingOptions = Get-UserSelection -Items $settingOptions -Prompt "Select settings to apply:" -MultiSelect
    
    if ($selectedSettingOptions.Count -eq 0) {
        Write-Host "No settings selected. Exiting." -ForegroundColor Yellow
        return
    }
    
    # Map selected options back to actual settings
    $selectedSettings = @()
    foreach ($option in $selectedSettingOptions) {
        $lineNumberMatch = $option -match '\[Line (\d+)\]'
        if ($lineNumberMatch) {
            $lineNumber = [int]$matches[1]
            $matchingSetting = $fileSettings.Settings | Where-Object { $_.LineNumber -eq $lineNumber } | Select-Object -First 1
            if ($matchingSetting) {
                $selectedSettings += $matchingSetting
            }
        }
    }
    
    # Ask for final confirmation
    Write-Host "`nYou are about to update $($selectedSettings.Count) settings in $($fileSettings.TargetPath)" -ForegroundColor Yellow
    $confirmation = Read-Host "Are you sure you want to proceed? (y/n)"
    
    if ($confirmation -ne "y") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Apply the selected settings
    $result = Update-TargetFile -TargetPath $fileSettings.TargetPath -SelectedSettings $selectedSettings
    
    if ($result.Success) {
        Write-Host "`nUpdated target file successfully!" -ForegroundColor Green
        Write-Host "Updated $($result.Updated) existing setting(s)" -ForegroundColor Green
        Write-Host "Added $($result.Added) new setting(s)" -ForegroundColor Green
        Write-Log "Successfully updated $($fileSettings.TargetPath) with $($selectedSettings.Count) settings from $selectedFileName" -Type Success
    } else {
        Write-Host "`nFailed to update target file: $($result.Message)" -ForegroundColor Red
        Write-Log "Failed to update $($fileSettings.TargetPath): $($result.Message)" -Type Error
    }
}

# Main execution
try {
    Start-InteractiveSettingsUpdater
    
    # Prompt user to press any key before exiting
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
    Write-Log "Unhandled error: $_" -Type Error
    
    # Prompt user to press any key before exiting
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}