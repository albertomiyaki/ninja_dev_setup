# Create multiple user databases

$ErrorActionPreference = 'Stop'

# Function to check if running as Administrator
Function Test-Administrator {
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
        Wait-ForKeyPress
        exit 1
    }
    
    # Exit the non-elevated script
    exit
}

# Configuration
$sqlInstance = "$env:COMPUTERNAME\SQLEXPRESS"
$dbOwner = "wmuser" 

# Import SqlServer module
Write-Host "===================================" -ForegroundColor Cyan
Write-Host " SQL Server Module Check" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Importing SqlServer module..." -ForegroundColor Cyan
try {
    Import-Module -Name "SqlServer" -ErrorAction Stop
    Write-Host "SqlServer module imported successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to import SqlServer module: $_" -ForegroundColor Red
    Write-Host "Please ensure the SqlServer module is installed. You can install it with:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name SqlServer" -ForegroundColor Yellow
    Wait-ForKeyPress
    exit 1
}

# Check if database owner exists
try {
    $server = New-Object Microsoft.SqlServer.Management.Smo.Server($sqlInstance)
    $loginExists = $server.Logins | Where-Object { $_.Name -eq $dbOwner }
    
    if (-not $loginExists) {
        Write-Host "Warning: Database owner '$dbOwner' does not exist on the server." -ForegroundColor Yellow
        Write-Host "Databases will be created but ownership change may fail." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Failed to check if database owner exists: $_" -ForegroundColor Yellow
}

# Define available databases
$availableDbs = @(
    "user1_1015",
    "user2_1015", 
    "user3_1015", 
    "user4_1015",
    "user5_1015",
    "user6_1015"
)

# Function to display a database in the menu
function Display-DatabaseItem {
    param (
        [string]$Database,
        [bool]$IsHighlighted,
        [bool]$IsSelected
    )
    
    if ($IsHighlighted) {
        Write-Host "$Database" -ForegroundColor White
    } else {
        Write-Host "$Database" -ForegroundColor Gray
    }
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
        }
    }
    
    # Return the selected items by their indices
    return $selectedIndices
}

# Function to confirm selection
function Confirm-Selection {
    param(
        [array]$SelectedDatabases
    )
    
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Confirm Selection" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "The following databases will be created:" -ForegroundColor Yellow
    
    foreach ($db in $SelectedDatabases) {
        Write-Host " • $db" -ForegroundColor White
    }
    
    Write-Host "`nDo you want to proceed? (Y/N) " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    return $confirmation -eq "Y" -or $confirmation -eq "y"
}

# Show multi-select menu and get chosen databases
Write-Host "`nLoading database selection menu..." -ForegroundColor Cyan
$selectedIndices = Show-MultiSelectMenu -Title "Select Databases to Create" -Options $availableDbs -DisplayFunction ${function:Display-DatabaseItem}
$selectedDbs = $selectedIndices | ForEach-Object { $availableDbs[$_] }

if ($selectedDbs.Count -eq 0) {
    Write-Host "No databases selected. Exiting." -ForegroundColor Yellow
    Wait-ForKeyPress
    exit 0
}

# Display confirmation
$confirmed = Confirm-Selection -SelectedDatabases $selectedDbs

if (-not $confirmed) {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Wait-ForKeyPress
    exit 0
}

# Create the selected databases
$successCount = 0
$failCount = 0
$failedDbs = @()

Write-Host "`n===================================" -ForegroundColor Cyan
Write-Host " Database Creation Progress" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan

foreach ($dbName in $selectedDbs) {
    Write-Host "`nCreating database '$dbName'..." -ForegroundColor Cyan
    
    try {
        # Check if database already exists
        $server = New-Object Microsoft.SqlServer.Management.Smo.Server($sqlInstance)
        if ($server.Databases[$dbName]) {
            Write-Host "Database '$dbName' already exists. Skipping." -ForegroundColor Yellow
            $failCount++
            $failedDbs += $dbName
            continue
        }
        
        # Create database
        $database = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $server, $dbName
        $database.Create()
        Write-Host "Database created successfully." -ForegroundColor Green
        
        # Set recovery model
        $database.RecoveryModel = [Microsoft.SqlServer.Management.Smo.RecoveryModel]::Simple
        $database.Alter()
        Write-Host "Recovery model set to Simple." -ForegroundColor Green
        
        # Set owner
        try {
            $database.SetOwner($dbOwner)
            Write-Host "Database owner set to '$dbOwner'." -ForegroundColor Green
        }
        catch {
            Write-Host "Warning: Could not set database owner: $_" -ForegroundColor Yellow
            Write-Host "Database was created but ownership was not changed." -ForegroundColor Yellow
        }
        
        $successCount++
    }
    catch {
        Write-Host "Error creating database '$dbName': $_" -ForegroundColor Red
        $failCount++
        $failedDbs += $dbName
    }
}

# Summary
Clear-Host
Write-Host "===================================" -ForegroundColor Cyan
Write-Host " Database Creation Summary" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Total databases selected: $($selectedDbs.Count)" -ForegroundColor White
Write-Host "Successfully created: $successCount" -ForegroundColor Green

if ($failCount -gt 0) {
    Write-Host "Failed to create: $failCount" -ForegroundColor Red
    Write-Host "`nFailed databases:" -ForegroundColor Red
    foreach ($db in $failedDbs) {
        Write-Host " • $db" -ForegroundColor Red
    }
}

if ($successCount -gt 0) {
    Write-Host "`nSuccessfully created databases:" -ForegroundColor Green
    foreach ($db in $selectedDbs) {
        if (-not ($failedDbs -contains $db)) {
            Write-Host " • $db" -ForegroundColor Green
        }
    }
}

Write-Host "`nDatabase creation process completed." -ForegroundColor Cyan

# Wait for user to press a key before exiting
Wait-ForKeyPress