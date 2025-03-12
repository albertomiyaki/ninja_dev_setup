# Add to Environment PATH
# 
# Reads paths from add_to_env_path_config.json and adds selected ones to the system PATH
# First copies itself to the Tools directory to avoid committing changes back to the repo

# Function to check if running as administrator
function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to elevate script if needed
function Invoke-ElevatedScript {
    if (-not (Test-IsAdmin)) {
        Write-Host "Elevating privileges to run as administrator..." -ForegroundColor Yellow
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

# Function to wait for keypress before exiting
function Wait-ForKeyPress {
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to get the script directory
function Get-ScriptDirectory {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) {
        Write-Host "Warning: Unable to determine script path using MyInvocation.MyCommand.Path" -ForegroundColor Yellow
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            Write-Host "Warning: Unable to determine script path. Using current directory." -ForegroundColor Yellow
            return (Get-Location).Path
        }
    }
    return Split-Path $scriptPath -Parent
}

# Function to display a path item in the menu
function Display-PathItem {
    param (
        [PSCustomObject]$PathItem,
        [bool]$IsHighlighted,
        [bool]$IsSelected
    )
    
    $pathName = $PathItem.Name
    $pathValue = $PathItem.Path
    
    if ($IsHighlighted) {
        Write-Host "$pathName " -NoNewline -ForegroundColor White
        Write-Host "($pathValue)" -ForegroundColor Yellow
    } else {
        Write-Host "$pathName " -NoNewline -ForegroundColor Gray
        Write-Host "($pathValue)" -ForegroundColor DarkGray
    }
}

# Function to display multi-select menu
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
        Write-Host " Up/Down  - Navigate up/down" -ForegroundColor Gray
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
                $exitMenu = $true
            }
            27 { # Escape
                return @()
            }
        }
    }
    
    return $selectedIndices
}

# Function to confirm selection before proceeding
function Confirm-Selection {
    param(
        [array]$SelectedPaths
    )
    
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Confirm Selection" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "The following paths will be added to the PATH environment variable:" -ForegroundColor Yellow
    
    foreach ($path in $SelectedPaths) {
        Write-Host " - $($path.Name) " -NoNewline -ForegroundColor White
        Write-Host "($($path.Path))" -ForegroundColor Gray
    }
    
    Write-Host "`nDo you want to proceed? (Y/N) " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    return $confirmation -eq "Y" -or $confirmation -eq "y"
}

# Function to validate if a path exists
function Test-PathExists {
    param (
        [string]$Path
    )
    
    # Handle network paths with more robustness
    if ($Path -match "^\\\\") {
        # This is a network path, try to test it but don't let it hang
        try {
            # Use a timeout approach for network paths
            $job = Start-Job -ScriptBlock { 
                param($p) 
                Test-Path -Path $p -ErrorAction SilentlyContinue 
            } -ArgumentList $Path
            
            # Wait for the job with timeout
            if (Wait-Job $job -Timeout 3) {
                $result = Receive-Job $job
                Remove-Job $job
                return $result
            } else {
                # If it times out, consider the path unavailable
                Remove-Job $job -Force
                return $false
            }
        } catch {
            return $false
        }
    } else {
        # Local path, just test it normally
        return Test-Path -Path $Path -ErrorAction SilentlyContinue
    }
}

# Function to add a path to the environment PATH
function Add-ToEnvPath {
    param (
        [string]$Path
    )
    
    try {
        # Validate the path
        if (-not (Test-PathExists -Path $Path)) {
            return @{
                Success = $false
                Message = "Path not found or not accessible"
            }
        }
        
        # Get current user PATH
        $currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Check if path is already in PATH
        if ($currentPath -like "*$Path*" -or $currentPath -like "*$Path;*") {
            return @{
                Success = $true
                Message = "Path is already in PATH"
                AlreadyAdded = $true
            }
        }
        
        # Add the path to PATH
        [Environment]::SetEnvironmentVariable("PATH", "$currentPath;$Path", "User")
        
        # Verify it was added successfully
        $newPath = [Environment]::GetEnvironmentVariable("PATH", "User")
        if ($newPath -like "*$Path*" -or $newPath -like "*$Path;*") {
            return @{
                Success = $true
                Message = "Successfully added to PATH"
                AlreadyAdded = $false
            }
        } else {
            return @{
                Success = $false
                Message = "Failed to add to PATH"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error: $_"
        }
    }
}

# Main script execution
try {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Add to Environment PATH" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    # Check if running as admin and elevate if needed
    Invoke-ElevatedScript
    
    # Get script directory and determine if this is running from the Tools directory
    $scriptDir = Get-ScriptDirectory
    $toolsDir = "$env:USERPROFILE\Tools"
    $toolsScriptDir = Join-Path -Path $toolsDir -ChildPath "EnvPathManager"
    $isRunningFromTools = $scriptDir -eq $toolsScriptDir
    
    # Create the Tools directory if it doesn't exist
    if (-not (Test-Path -Path $toolsScriptDir)) {
        New-Item -ItemType Directory -Path $toolsScriptDir -Force | Out-Null
        Write-Host "Created Tools directory: $toolsScriptDir" -ForegroundColor Green
    }
    
    # If not running from the Tools directory, copy the script and config
    if (-not $isRunningFromTools) {
        Write-Host "Running from repository. Preparing to copy to Tools directory..." -ForegroundColor Yellow
        
        # Script path variables
        $scriptPath = $MyInvocation.MyCommand.Path
        $destScriptPath = Join-Path -Path $toolsScriptDir -ChildPath (Split-Path $scriptPath -Leaf)
        
        # Config path variables
        $configPath = Join-Path -Path $scriptDir -ChildPath "add_to_env_path_config.json"
        $destConfigPath = Join-Path -Path $toolsScriptDir -ChildPath "add_to_env_path_config.json"
        
        # Check if script already exists in destination and prompt for confirmation
        $copyScript = $true
        if (Test-Path $destScriptPath) {
            $response = Read-Host "Script already exists in Tools directory. Overwrite? (Y/N)"
            $copyScript = $response -in 'Y', 'y', 'Yes', 'yes'
        }
        
        if ($copyScript) {
            Copy-Item -Path $scriptPath -Destination $destScriptPath -Force
            Write-Host "Copied script to $destScriptPath" -ForegroundColor Green
        } else {
            Write-Host "Using existing script in Tools directory" -ForegroundColor Cyan
        }
        
        # Handle config file
        if (Test-Path $configPath) {
            $copyConfig = $true
            if (Test-Path $destConfigPath) {
                $response = Read-Host "Configuration file already exists in Tools directory. Overwrite? (Y/N)"
                $copyConfig = $response -in 'Y', 'y', 'Yes', 'yes'
            }
            
            if ($copyConfig) {
                Copy-Item -Path $configPath -Destination $destConfigPath -Force
                Write-Host "Copied configuration file to $destConfigPath" -ForegroundColor Green
            } else {
                Write-Host "Using existing configuration file in Tools directory" -ForegroundColor Cyan
            }
        } else {
            # Create a sample config file if it doesn't exist at the destination
            if (-not (Test-Path $destConfigPath)) {
                $sampleConfig = @"
[
    {
        "Name": "jq",
        "Path": "C:\\Program Files\\jq"
    },
    {
        "Name": "Sublime Text",
        "Path": "C:\\Program Files\\Sublime Text"
    },
    {
        "Name": "Baretail",
        "Path": "C:\\Program Files\\Baretail"
    },
    {
        "Name": "glogg",
        "Path": "C:\\Program Files\\glogg"
    },
    {
        "Name": "Go Binaries",
        "Path": "C:\\Go\\bin"
    },
    {
        "Name": "Custom Tools",
        "Path": "C:\\Users\\Username\\Tools\\bin"
    }
]
"@
                $sampleConfig | Out-File -FilePath $destConfigPath -Encoding UTF8
                Write-Host "Created sample configuration file at $destConfigPath" -ForegroundColor Green
            }
        }
        
        # Launch the copied script
        Write-Host "Launching script from Tools directory..." -ForegroundColor Yellow
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$destScriptPath`""
        exit
    }
    
    # If we're here, we're running from the Tools directory
    $jsonPath = Join-Path -Path $scriptDir -ChildPath "add_to_env_path_config.json"
    
    # Check if JSON file exists
    if (-not (Test-Path $jsonPath)) {
        Write-Host "Error: add_to_env_path_config.json not found at '$jsonPath'" -ForegroundColor Red
        Write-Host "Please create a add_to_env_path_config.json file with this format:" -ForegroundColor Yellow
        Write-Host @"
[
    {
        "Name": "Python Scripts",
        "Path": "C:\\Users\\Username\\AppData\\Local\\Programs\\Python\\Python39\\Scripts"
    },
    {
        "Name": "Node.js",
        "Path": "C:\\Program Files\\nodejs"
    }
]
"@ -ForegroundColor Gray
        Wait-ForKeyPress
        exit
    }
    
    # Read paths from JSON file
    try {
        $pathsData = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
        Write-Host "Successfully loaded paths from '$jsonPath'" -ForegroundColor Green
    }
    catch {
        Write-Host "Error reading JSON file: $_" -ForegroundColor Red
        Wait-ForKeyPress
        exit
    }
    
    # Validate JSON structure
    if ($null -eq $pathsData -or $pathsData.Count -eq 0) {
        Write-Host "Error: No paths found in the JSON file" -ForegroundColor Red
        Wait-ForKeyPress
        exit
    }
    
    # Check if paths have the required properties
    $validPaths = @()
    foreach ($item in $pathsData) {
        if ($item.PSObject.Properties.Name -contains "Name" -and $item.PSObject.Properties.Name -contains "Path") {
            $validPaths += $item
        }
    }
    
    if ($validPaths.Count -eq 0) {
        Write-Host "Error: No valid paths found in the JSON file" -ForegroundColor Red
        Write-Host "Each path entry must have 'Name' and 'Path' properties" -ForegroundColor Yellow
        Wait-ForKeyPress
        exit
    }
    
    # Show multi-select menu to select paths
    $selectedIndices = Show-MultiSelectMenu -Title "Select Paths to Add to PATH" -Options $validPaths -DisplayFunction ${function:Display-PathItem}
    $selectedPaths = $selectedIndices | ForEach-Object { $validPaths[$_] }
    
    if ($selectedPaths.Count -eq 0) {
        Write-Host "No paths selected. Exiting." -ForegroundColor Yellow
        Wait-ForKeyPress
        exit
    }
    
    # Confirm selection
    $confirmed = Confirm-Selection -SelectedPaths $selectedPaths
    
    if (-not $confirmed) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Wait-ForKeyPress
        exit
    }
    
    # Process selected paths
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host " Adding Paths to Environment PATH" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    $successCount = 0
    $skippedCount = 0
    $failedCount = 0
    $results = @()
    
    foreach ($path in $selectedPaths) {
        Write-Host "`nProcessing: $($path.Name) ($($path.Path))" -ForegroundColor Yellow
        $result = Add-ToEnvPath -Path $path.Path
        
        if ($result.Success) {
            if ($result.AlreadyAdded) {
                Write-Host "Already in PATH" -ForegroundColor Cyan
                $skippedCount++
            } else {
                Write-Host "Successfully added to PATH" -ForegroundColor Green
                $successCount++
            }
        } else {
            Write-Host "Failed: $($result.Message)" -ForegroundColor Red
            $failedCount++
        }
        
        $pathResult = @{
            Name = $path.Name
            Path = $path.Path
            Success = $result.Success
            Message = $result.Message
            AlreadyAdded = $result.AlreadyAdded
        }
        
        $results += $pathResult
    }
    
    # Display summary
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Environment PATH Update Summary" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    Write-Host "Total paths selected: $($selectedPaths.Count)" -ForegroundColor White
    Write-Host "Successfully added: $successCount" -ForegroundColor Green
    Write-Host "Already in PATH: $skippedCount" -ForegroundColor Cyan
    Write-Host "Failed: $failedCount" -ForegroundColor Red
    
    if ($failedCount -gt 0) {
        Write-Host "`nFailed paths:" -ForegroundColor Red
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host " - $($_.Name) " -NoNewline -ForegroundColor Red
            Write-Host "($($_.Path)) - $($_.Message)" -ForegroundColor Gray
        }
    }
    
    if ($successCount -gt 0) {
        Write-Host "`nSuccessfully added paths:" -ForegroundColor Green
        $results | Where-Object { $_.Success -and -not $_.AlreadyAdded } | ForEach-Object {
            Write-Host " - $($_.Name) " -NoNewline -ForegroundColor Green
            Write-Host "($($_.Path))" -ForegroundColor Gray
        }
    }
    
    if ($skippedCount -gt 0) {
        Write-Host "`nAlready in PATH:" -ForegroundColor Cyan
        $results | Where-Object { $_.AlreadyAdded } | ForEach-Object {
            Write-Host " - $($_.Name) " -NoNewline -ForegroundColor Cyan
            Write-Host "($($_.Path))" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nNote: You may need to restart your command prompt or PowerShell for" -ForegroundColor Yellow
    Write-Host "PATH changes to take effect in the current session." -ForegroundColor Yellow
}
catch {
    Write-Host "`n===================================" -ForegroundColor Red
    Write-Host " Error" -ForegroundColor Red
    Write-Host "===================================" -ForegroundColor Red
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
}

# Wait for user to press a key before exiting
Wait-ForKeyPress