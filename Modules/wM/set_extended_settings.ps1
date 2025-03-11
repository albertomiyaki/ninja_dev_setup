# Set Extended Settings

# Reads configuration files (*.conf) in extended_settings folder and applies settings to target files
# First line must be a comment with target path

# Define user profile paths
$userProfileFolder = $env:USERPROFILE
$userSettingsBasePath = Join-Path -Path $userProfileFolder -ChildPath "Tools"
$userSettingsPath = Join-Path -Path $userSettingsBasePath -ChildPath "extended_settings"

# Check if the script is running with administrator privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to wait for keypress before exiting
function Wait-ForKeyPress {
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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

# Get the script's root directory (repo path)
$scriptRoot = $PSScriptRoot
$repoSettingsPath = Join-Path $scriptRoot "extended_settings"

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

# Function to display a config file in the menu
function Display-ConfigItem {
    param (
        [string]$ConfigFile,
        [bool]$IsHighlighted,
        [bool]$IsSelected
    )
    
    if ($IsHighlighted) {
        Write-Host "$ConfigFile" -ForegroundColor White
    } else {
        Write-Host "$ConfigFile" -ForegroundColor Gray
    }
}

# Function to display a setting in the menu
function Display-SettingItem {
    param (
        [string]$Setting,
        [bool]$IsHighlighted,
        [bool]$IsSelected
    )
    
    if ($IsHighlighted) {
        Write-Host "$Setting" -ForegroundColor White
    } else {
        Write-Host "$Setting" -ForegroundColor Gray
    }
}

# Improved function to display a single-select menu
function Show-SelectMenu {
    param(
        [string]$Title,
        [array]$Options,
        [scriptblock]$DisplayFunction
    )
    
    $currentIndex = 0
    $exitMenu = $false
    
    while (-not $exitMenu) {
        Clear-Host
        Write-Host "===================================" -ForegroundColor Cyan
        Write-Host " $Title" -ForegroundColor Cyan
        Write-Host "===================================" -ForegroundColor Cyan
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $currentIndex) {
                Write-Host " >> " -NoNewline -ForegroundColor Green
                & $DisplayFunction $Options[$i] $true $false
            } else {
                Write-Host "    " -NoNewline
                & $DisplayFunction $Options[$i] $false $false
            }
        }
        
        Write-Host "`nNavigation:" -ForegroundColor Cyan
        Write-Host " ↑/↓      - Navigate up/down" -ForegroundColor Gray
        Write-Host " Enter    - Select item" -ForegroundColor Gray
        Write-Host " Esc      - Cancel" -ForegroundColor Gray
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentIndex -gt 0) {
                    $currentIndex--
                } else {
                    $currentIndex = $Options.Count - 1
                }
            }
            40 { # Down arrow
                if ($currentIndex -lt ($Options.Count - 1)) {
                    $currentIndex++
                } else {
                    $currentIndex = 0
                }
            }
            13 { # Enter
                return $Options[$currentIndex]
            }
            27 { # Escape
                return $null
            }
        }
    }
    
    return $null
}

# Improved function to display a multi-select menu
function Show-MultiSelectMenu {
    param(
        [string]$Title,
        [array]$Options,
        [scriptblock]$DisplayFunction
    )
    
    $selectedIndices = @()
    $currentIndex = 0
    $exitMenu = $false
    
    while (-not $exitMenu) {
        Clear-Host
        Write-Host "===================================" -ForegroundColor Cyan
        Write-Host " $Title" -ForegroundColor Cyan
        Write-Host "===================================" -ForegroundColor Cyan
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            $selected = $selectedIndices -contains $i
            $selectionIndicator = if ($selected) { "[X]" } else { "[ ]" }
            
            if ($i -eq $currentIndex) {
                Write-Host " >> " -NoNewline -ForegroundColor Green
                Write-Host "$selectionIndicator " -NoNewline -ForegroundColor Yellow
                & $DisplayFunction $Options[$i] $true $selected
            } else {
                Write-Host "    " -NoNewline
                Write-Host "$selectionIndicator " -NoNewline -ForegroundColor Gray
                & $DisplayFunction $Options[$i] $false $selected
            }
        }
        
        Write-Host "`nSelected: $($selectedIndices.Count) of $($Options.Count)" -ForegroundColor Magenta
        
        Write-Host "`nNavigation:" -ForegroundColor Cyan
        Write-Host " ↑/↓      - Navigate up/down" -ForegroundColor Gray
        Write-Host " Spacebar - Select/Deselect item" -ForegroundColor Gray
        Write-Host " A        - Select All" -ForegroundColor Gray
        Write-Host " N        - Select None" -ForegroundColor Gray
        Write-Host " Enter    - Confirm and proceed" -ForegroundColor Gray
        Write-Host " Esc      - Cancel" -ForegroundColor Gray
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentIndex -gt 0) {
                    $currentIndex--
                } else {
                    $currentIndex = $Options.Count - 1
                }
            }
            40 { # Down arrow
                if ($currentIndex -lt ($Options.Count - 1)) {
                    $currentIndex++
                } else {
                    $currentIndex = 0
                }
            }
            32 { # Spacebar
                if ($selectedIndices -contains $currentIndex) {
                    $selectedIndices = $selectedIndices | Where-Object { $_ -ne $currentIndex }
                } else {
                    $selectedIndices += $currentIndex
                }
            }
            65 { # 'A' key
                $selectedIndices = @(0..($Options.Count - 1))
            }
            78 { # 'N' key
                $selectedIndices = @()
            }
            13 { # Enter
                if ($selectedIndices.Count -gt 0) {
                    return $selectedIndices | ForEach-Object { $Options[$_] }
                } else {
                    Write-Host "`nNo items selected. Please select at least one item." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                }
            }
            27 { # Escape
                return @()
            }
        }
    }
    
    return @()
}

# Function to confirm the selection
function Confirm-Selection {
    param(
        [string]$TargetPath,
        [array]$SelectedSettings
    )
    
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Confirm Settings Update" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "You are about to update the following settings in:`n" -ForegroundColor Yellow
    Write-Host "$TargetPath" -ForegroundColor White
    
    Write-Host "`nSelected settings:" -ForegroundColor Yellow
    foreach ($setting in $SelectedSettings) {
        $key = $setting.Key
        $value = $setting.Value
        $currentValue = $setting.CurrentValue
        
        Write-Host " • $key = " -NoNewline -ForegroundColor White
        Write-Host "$value" -NoNewline -ForegroundColor Green
        Write-Host " (Current: $currentValue)" -ForegroundColor Gray
    }
    
    Write-Host "`nDo you want to proceed? (Y/N) " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    return $confirmation -eq "Y" -or $confirmation -eq "y"
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
            Write-Log "Settings file is empty: $($FilePath)" -Type Warning
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
        Write-Log "Error processing settings file $($FilePath): $_" -Type Error
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
        Write-Log "Error updating target file $($TargetPath): $_" -Type Error
        return @{
            Success = $false
            Updated = 0
            Added = 0
            Message = "Error: $_"
        }
    }
}

# Function to find all server configuration files
function Get-ConfigFiles {
    param (
        [string]$SettingsPath
    )
    
    try {
        if (-not (Test-Path $SettingsPath)) {
            Write-Log "Settings folder not found: $SettingsPath" -Type Error
            return @()
        }
        
        # Find all .conf files in the settings folder
        $configFiles = Get-ChildItem -Path $SettingsPath -Filter "*.conf" -File
        
        if ($configFiles.Count -eq 0) {
            Write-Log "No configuration files found in $SettingsPath" -Type Warning
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

# Function to ensure user settings directory exists and copy template config files if needed
function Initialize-UserSettingsDirectory {
    # Create user settings directory if it doesn't exist
    if (-not (Test-Path $userSettingsPath)) {
        try {
            New-Item -Path $userSettingsBasePath -ItemType Directory -Force | Out-Null
            New-Item -Path $userSettingsPath -ItemType Directory -Force | Out-Null
            Write-Log "Created user settings directory: $userSettingsPath" -Type Success
        }
        catch {
            Write-Log "Failed to create user settings directory: $_" -Type Error
            return $false
        }
    }
    
    # Check for configuration files in repo
    $repoConfigFiles = Get-ConfigFiles -SettingsPath $repoSettingsPath
    
    if ($repoConfigFiles.Count -eq 0) {
        Write-Log "No template configuration files found in repository" -Type Warning
        return $true
    }
    
    # Copy repo config files to user directory if they don't exist
    $copiedCount = 0
    foreach ($file in $repoConfigFiles) {
        $destPath = Join-Path $userSettingsPath $file.Name
        if (-not (Test-Path $destPath)) {
            try {
                Copy-Item -Path $file.FullName -Destination $destPath -Force
                $copiedCount++
                Write-Log "Copied template configuration file: $($file.Name)" -Type Success
            }
            catch {
                Write-Log "Failed to copy template configuration file $($file.Name): $_" -Type Error
            }
        }
    }
    
    if ($copiedCount -gt 0) {
        Write-Host "Copied $copiedCount template configuration files to user directory" -ForegroundColor Green
    }
    
    return $true
}

# Main function to run the interactive settings updater
function Start-InteractiveSettingsUpdater {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Interactive Settings Updater" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "Script location (repository): $scriptRoot" -ForegroundColor Gray
    Write-Host "User settings folder: $userSettingsPath`n" -ForegroundColor Gray
    
    # Initialize user settings directory and copy template files if needed
    $initResult = Initialize-UserSettingsDirectory
    if (-not $initResult) {
        Write-Host "Failed to initialize user settings directory. Exiting." -ForegroundColor Red
        Wait-ForKeyPress
        return
    }
    
    # Get all configuration files from user directory
    $configFiles = Get-ConfigFiles -SettingsPath $userSettingsPath
    
    if ($configFiles.Count -eq 0) {
        Write-Host "No configuration files found in user directory. Exiting." -ForegroundColor Red
        Wait-ForKeyPress
        return
    }
    
    # Display list of configuration files and let user select one
    $fileOptions = $configFiles | ForEach-Object { $_.Name }
    $selectedFileName = Show-SelectMenu -Title "Select a Configuration File" -Options $fileOptions -DisplayFunction ${function:Display-ConfigItem}
    
    if ($null -eq $selectedFileName) {
        Write-Host "No file selected. Exiting." -ForegroundColor Yellow
        Wait-ForKeyPress
        return
    }
    
    $selectedFilePath = Join-Path $userSettingsPath $selectedFileName
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host " Processing Configuration File" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "Selected file: $selectedFileName" -ForegroundColor Yellow
    
    # Parse the selected configuration file
    $fileSettings = Get-SettingsFromFile -FilePath $selectedFilePath
    
    if ($null -eq $fileSettings) {
        Write-Host "Failed to process configuration file. Exiting." -ForegroundColor Red
        Wait-ForKeyPress
        return
    }
    
    Write-Host "`nTarget file: $($fileSettings.TargetPath)" -ForegroundColor Green
    Write-Host "Found $($fileSettings.Settings.Count) settings in configuration file" -ForegroundColor Green
    
    # Get current values from target file for comparison
    $targetExists = Test-Path $($fileSettings.TargetPath)
    if ($targetExists) {
        Write-Host "Target file exists. Getting current values for comparison..." -ForegroundColor Cyan
    } else {
        Write-Host "Target file doesn't exist. It will be created." -ForegroundColor Yellow
    }
    
    # Display settings with current values and let user select which to apply
    $settingOptions = @()
    $settingsWithInfo = @()
    
    foreach ($setting in $fileSettings.Settings) {
        $currentValue = if ($targetExists) { 
            Get-CurrentSettingValue -FilePath $($fileSettings.TargetPath) -SettingKey $($setting.Key) 
        } else { 
            "<file will be created>" 
        }
        
        $displayText = "[Line $($setting.LineNumber)] $($setting.Key) = $($setting.Value) (Current: $currentValue)"
        $settingOptions += $displayText
        
        # Store setting with current value and display text for later reference
        $settingWithInfo = $setting.Clone()
        $settingWithInfo.CurrentValue = $currentValue
        $settingWithInfo.DisplayText = $displayText
        $settingsWithInfo += $settingWithInfo
    }
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host " Select Settings to Apply" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    $selectedSettingOptions = Show-MultiSelectMenu -Title "Available Settings" -Options $settingOptions -DisplayFunction ${function:Display-SettingItem}
    
    if ($selectedSettingOptions.Count -eq 0) {
        Write-Host "No settings selected. Exiting." -ForegroundColor Yellow
        Wait-ForKeyPress
        return
    }
    
    # Map selected options back to actual settings
    $selectedSettings = @()
    foreach ($option in $selectedSettingOptions) {
        $lineNumberMatch = $option -match '\[Line (\d+)\]'
        if ($lineNumberMatch) {
            $lineNumber = [int]$matches[1]
            $matchingSetting = $settingsWithInfo | Where-Object { $_.LineNumber -eq $lineNumber } | Select-Object -First 1
            if ($matchingSetting) {
                $selectedSettings += $matchingSetting
            }
        }
    }
    
    # Ask for final confirmation
    $confirmed = Confirm-Selection -TargetPath $fileSettings.TargetPath -SelectedSettings $selectedSettings
    
    if (-not $confirmed) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Wait-ForKeyPress
        return
    }
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host " Applying Settings" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    # Apply the selected settings
    $result = Update-TargetFile -TargetPath $($fileSettings.TargetPath) -SelectedSettings $selectedSettings
    
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host " Update Results" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    if ($result.Success) {
        Write-Host "Target file updated successfully!" -ForegroundColor Green
        Write-Host " • Updated $($result.Updated) existing setting(s)" -ForegroundColor Green
        Write-Host " • Added $($result.Added) new setting(s)" -ForegroundColor Green
        Write-Log "Successfully updated $($fileSettings.TargetPath) with $($selectedSettings.Count) settings from $selectedFileName" -Type Success
    } else {
        Write-Host "Failed to update target file: $($result.Message)" -ForegroundColor Red
        Write-Log "Failed to update $($fileSettings.TargetPath): $($result.Message)" -Type Error
    }
}

# Main execution
try {
    Start-InteractiveSettingsUpdater
    Wait-ForKeyPress
}
catch {
    Write-Host "`n===================================" -ForegroundColor Red
    Write-Host " Error" -ForegroundColor Red
    Write-Host "===================================" -ForegroundColor Red
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
    Write-Log "Unhandled error: $_" -Type Error
    Wait-ForKeyPress
}