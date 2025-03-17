# GitHub Repo Cloner

# Check if running as administrator and self-elevate if needed
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

# Set strict mode and error preferences for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper Functions
function Write-ColorMessage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    # Set console colors based on message level
    switch ($Level) {
        "Success" { $color = "Green"; $prefix = "SUCCESS: " }
        "Warning" { $color = "Yellow"; $prefix = "WARNING: " }
        "Error" { $color = "Red"; $prefix = "ERROR: " }
        default { $color = "Cyan"; $prefix = "INFO: " }
    }
    
    # Write to console
    Write-Host "$prefix$Message" -ForegroundColor $color
}

function Get-UserConfirmation {
    param (
        [string]$Message = "Do you want to continue?"
    )
    
    $confirmation = Read-Host "$Message [Y/N]"
    return $confirmation.ToLower() -eq 'y'
}

function Test-CommandExists {
    param (
        [string]$Command
    )
    
    $exists = $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
    return $exists
}

function Install-GithubCLI {
    Write-ColorMessage "GitHub CLI not found. Need to install it first." -Level Warning
    
    $installChoice = Read-Host @"
How would you like to install GitHub CLI?
1. Use local/network MSI installer
2. Use Chocolatey (if installed)
Enter choice (1 or 2)
"@

    switch ($installChoice) {
        "2" {
            # Check if chocolatey is installed
            if (Test-CommandExists "choco") {
                Write-ColorMessage "Installing GitHub CLI via Chocolatey..." -Level Info
                
                try {
                    Start-Process -FilePath "choco" -ArgumentList "install gh -y" -Verb RunAs -Wait
                    # Refresh environment variables
                    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                    
                    if (Test-CommandExists "gh") {
                        Write-ColorMessage "GitHub CLI installed successfully via Chocolatey." -Level Success
                        return $true
                    } else {
                        Write-ColorMessage "GitHub CLI installation via Chocolatey completed, but command not found." -Level Warning
                        Write-ColorMessage "You may need to restart your terminal or refresh your environment." -Level Info
                        return $false
                    }
                }
                catch {
                    Write-ColorMessage "Failed to install GitHub CLI via Chocolatey: $_" -Level Error
                    return $false
                }
            } else {
                Write-ColorMessage "Chocolatey is not installed. Please choose another option." -Level Error
                return Install-GithubCLI
            }
        }
        
        "1" {
            # Default MSI path - adjust this to your organization's standard location
            $defaultMsiPath = "C:\Installers\gh_2.67.0_windows_amd64.msi"
            
            # Ask for local/network path to MSI file with default value
            $msiPathPrompt = Read-Host "Enter the path to the GitHub CLI MSI file (default: $defaultMsiPath)"
            
            # Use default if user just pressed Enter
            $msiPath = if ([string]::IsNullOrWhiteSpace($msiPathPrompt)) { $defaultMsiPath } else { $msiPathPrompt }
            
            # Validate path exists
            if (-not (Test-Path $msiPath)) {
                Write-ColorMessage "The specified MSI file does not exist: $msiPath" -Level Error
                if (Get-UserConfirmation "Would you like to specify a different path?") {
                    return Install-GithubCLI
                }
                return $false
            }
            
            try {
                Write-ColorMessage "Installing GitHub CLI from: $msiPath" -Level Info
                Start-Process -FilePath "msiexec" -ArgumentList "/i `"$msiPath`" /quiet" -Verb RunAs -Wait
                
                # Refresh environment variables
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                
                # Verify installation
                if (Test-CommandExists "gh") {
                    Write-ColorMessage "GitHub CLI installed successfully." -Level Success
                    return $true
                }
                else {
                    Write-ColorMessage "GitHub CLI installation completed, but command not found." -Level Warning
                    Write-ColorMessage "You may need to restart your terminal or refresh your environment." -Level Info
                    
                    if (Get-UserConfirmation "Would you like to continue anyway?") {
                        return $true
                    }
                    return $false
                }
            }
            catch {
                Write-ColorMessage "Failed to install GitHub CLI: $_" -Level Error
                return $false
            }
        }
        
        default {
            Write-ColorMessage "Invalid choice. Please select 1 or 2." -Level Error
            return Install-GithubCLI
        }
    }
}
#endregion

#region Main Script
Clear-Host
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "     GitHub Repository Cloning Tool" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host

# Check if Git is installed
if (-not (Test-CommandExists "git")) {
    Write-ColorMessage "Git is not installed. Please install Git for Windows first." -Level Error
    Write-ColorMessage "You can download Git from: https://git-scm.com/download/win" -Level Info
    exit 1
}

# Check if GitHub CLI is installed, if not, install it
if (-not (Test-CommandExists "gh")) {
    $installed = Install-GithubCLI
    if (-not $installed) {
        Write-ColorMessage "GitHub CLI is required but could not be installed. Exiting." -Level Error
        exit 1
    }
}

# Verify GitHub CLI is authenticated
Write-ColorMessage "Checking GitHub authentication status..." -Level Info

# First, verify the GitHub CLI is working
$ghVersion = & gh --version
Write-ColorMessage "GitHub CLI is installed: $($ghVersion[0])" -Level Success

# Now handle authentication
# We need to capture STDERR without triggering an error - PowerShell makes this tricky
$authStatus = Start-Process -FilePath "gh" -ArgumentList "auth status" -Wait -NoNewWindow -PassThru
$authExitCode = $authStatus.ExitCode

if ($authExitCode -ne 0) {
    Write-ColorMessage "You need to authenticate with GitHub CLI." -Level Info
    Write-ColorMessage "Running 'gh auth login' command. Please follow the prompts." -Level Info
    
    # Run the login command interactively
    Start-Process -FilePath "gh" -ArgumentList "auth login" -Wait -NoNewWindow
    
    # Verify authentication again
    $authStatus = Start-Process -FilePath "gh" -ArgumentList "auth status" -Wait -NoNewWindow -PassThru
    $authExitCode = $authStatus.ExitCode
    
    if ($authExitCode -ne 0) {
        Write-ColorMessage "Authentication failed after login attempt. Please run the script again." -Level Error
        exit 1
    }
}

Write-ColorMessage "Successfully authenticated with GitHub." -Level Success

# Get user inputs
$orgName = Read-Host "Enter your GitHub organization name"
$repoLimit = Read-Host "Enter maximum number of repositories to list (default: 100)"

if ([string]::IsNullOrWhiteSpace($repoLimit)) {
    $repoLimit = 100
}
else {
    try {
        $repoLimit = [int]$repoLimit
    }
    catch {
        Write-ColorMessage "Invalid number. Using default limit of 100." -Level Warning
        $repoLimit = 100
    }
}

# Default clone directory
$defaultCloneDir = Join-Path $env:USERPROFILE "git"
$cloneDir = Read-Host "Enter the directory to clone repositories into (default: $defaultCloneDir)"

if ([string]::IsNullOrWhiteSpace($cloneDir)) {
    $cloneDir = $defaultCloneDir
}

# Ensure the directory exists
if (-not (Test-Path $cloneDir)) {
    New-Item -ItemType Directory -Path $cloneDir -Force | Out-Null
    Write-ColorMessage "Created directory: $cloneDir" -Level Info
}

# List repositories
Write-ColorMessage "Listing repositories from $orgName organization (limit: $repoLimit)..." -Level Info
try {
    # Change to the clone directory
    $originalLocation = Get-Location
    Set-Location -Path $cloneDir
    
    # Get the list of repositories
    $repos = @()
    $repoList = gh repo list $orgName --limit $repoLimit
    
    if (-not $repoList) {
        Write-ColorMessage "No repositories found in $orgName organization." -Level Warning
        Set-Location -Path $originalLocation
        exit 0
    }
    
    foreach ($line in $repoList) {
        # Parse repository information (name, description, etc.)
        $repoInfo = $line -split '\t'
        $repoName = ($repoInfo[0] -split '\s+')[0]  # Get just the repo name
        
        # Create custom object with properties
        $repos += [PSCustomObject]@{
            Name = $repoName
            Description = if ($repoInfo.Length -gt 1) { $repoInfo[1] } else { "" }
            FullLine = $line
        }
    }
    
    # Allow user to select repositories
    Write-ColorMessage "Found $($repos.Count) repositories. Please select which ones to clone:" -Level Info
    $selectedRepos = $repos | Out-GridView -Title "Select repositories to clone from $orgName" -PassThru
    
    if (-not $selectedRepos) {
        Write-ColorMessage "No repositories selected. Exiting." -Level Warning
        Set-Location -Path $originalLocation
        exit 0
    }
    
    # Fix: Ensure $selectedRepos is always treated as an array
    if ($selectedRepos -isnot [System.Array]) {
        $selectedRepos = @($selectedRepos)
    }
    
    # Clone selected repositories
    Write-ColorMessage "Cloning $($selectedRepos.Count) selected repositories..." -Level Info
    
    $successCount = 0
    $failCount = 0
    
    foreach ($repo in $selectedRepos) {
        $repoName = $repo.Name
        # Extract just the repository name without the organization prefix
        $repoNameOnly = $repoName.Split('/')[-1]
        $localRepoPath = Join-Path $cloneDir $repoNameOnly
        
        # Check if repo already exists locally
        if (Test-Path $localRepoPath) {
            Write-ColorMessage "Repository '$repoNameOnly' already exists in $localRepoPath, skipping..." -Level Warning
            continue
        }
        
        Write-ColorMessage "Cloning $repoName to $localRepoPath..." -Level Info
        
        try {
            # Clone directly to the specified location
            gh repo clone "$repoName" "$localRepoPath"
            
            if ($LASTEXITCODE -eq 0) {
                Write-ColorMessage "Successfully cloned $repoName to $localRepoPath" -Level Success
                $successCount++
            } else {
                Write-ColorMessage "Failed to clone $repoName" -Level Error
                $failCount++
            }
        } catch {
            Write-ColorMessage "Error cloning $repoName : $($_.Exception.Message)" -Level Error
            $failCount++
        }
    }
    
    # Return to original directory
    Set-Location -Path $originalLocation
    
    # Summary
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "             Cloning Summary" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-ColorMessage "Successfully cloned: $successCount repositories" -Level Success
    Write-ColorMessage "Failed to clone: $failCount repositories" -Level $(if ($failCount -gt 0) { "Error" } else { "Info" })
    Write-ColorMessage "All repositories were cloned directly to: $cloneDir" -Level Info
    
}
catch {
    # Return to original directory in case of error
    Set-Location -Path $originalLocation
    Write-ColorMessage "Error: $_" -Level Error
    exit 1
}
#endregion

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host "      GitHub Cloning Process Complete" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan