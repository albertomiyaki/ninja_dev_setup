# Create multiple user databases

$ErrorActionPreference = 'Stop'

# Function to check if running as Administrator
Function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

# Configuration
$sqlInstance = "$env:COMPUTERNAME\SQLEXPRESS"
$dbOwner = "wmuser" 

# Import SqlServer module
Write-Host "Importing SqlServer module..." -ForegroundColor Cyan
try {
    Import-Module -Name "SqlServer" -ErrorAction Stop
    Write-Host "SqlServer module imported successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to import SqlServer module: $_" -ForegroundColor Red
    Write-Host "Please ensure the SqlServer module is installed. You can install it with: Install-Module -Name SqlServer" -ForegroundColor Yellow
    exit 1
}

# Check if database owner exists
try {
    $server = New-Object Microsoft.SqlServer.Management.Smo.Server($sqlInstance)
    $loginExists = $server.Logins | Where-Object { $_.Name -eq $dbOwner }
    
    if (-not $loginExists) {
        Write-Host "Warning: Database owner '$dbOwner' does not exist on the server. Databases will be created but ownership change may fail." -ForegroundColor Yellow
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

# Function to display a multi-select menu
function Show-MultiSelectMenu {
    param (
        [array]$Items
    )
    
    $selected = @()
    $currentPosition = 0
    $selectionDone = $false
    
    while (-not $selectionDone) {
        Clear-Host
        Write-Host "===== Select Databases to Create =====" -ForegroundColor Cyan
        Write-Host "Use arrow keys to navigate, space to select/deselect, enter to confirm" -ForegroundColor Yellow
        Write-Host "Press 'A' to select all, 'N' to select none" -ForegroundColor Yellow
        Write-Host ""
        
        for ($i = 0; $i -lt $Items.Count; $i++) {
            $isSelected = $selected -contains $i
            $isCurrent = $i -eq $currentPosition
            
            $prefix = if ($isCurrent) { ">" } else { " " }
            $mark = if ($isSelected) { "[X]" } else { "[ ]" }
            $color = if ($isCurrent) { "Cyan" } else { "White" }
            
            Write-Host "$prefix $mark $($Items[$i])" -ForegroundColor $color
        }
        
        Write-Host "`nSelected: $($selected.Count) of $($Items.Count)" -ForegroundColor Magenta
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentPosition -gt 0) { $currentPosition-- }
            }
            40 { # Down arrow
                if ($currentPosition -lt $Items.Count - 1) { $currentPosition++ }
            }
            32 { # Space
                if ($selected -contains $currentPosition) {
                    $selected = $selected | Where-Object { $_ -ne $currentPosition }
                }
                else {
                    $selected += $currentPosition
                }
            }
            65 { # A key - select all
                $selected = 0..($Items.Count - 1)
            }
            78 { # N key - select none
                $selected = @()
            }
            13 { # Enter
                $selectionDone = $true
            }
        }
    }
    
    # Return the selected items by their values
    return $selected | ForEach-Object { $Items[$_] }
}

# Show multi-select menu and get chosen databases
Write-Host "Loading database selection menu..." -ForegroundColor Cyan
$selectedDbs = Show-MultiSelectMenu -Items $availableDbs

if ($selectedDbs.Count -eq 0) {
    Write-Host "No databases selected. Exiting." -ForegroundColor Yellow
    exit 0
}

# Display confirmation
Write-Host "`nYou selected the following databases to create:" -ForegroundColor Cyan
foreach ($db in $selectedDbs) {
    Write-Host " - $db" -ForegroundColor White
}

$confirmation = Read-Host "`nProceed with creation? (Y/N)"
if ($confirmation -ne "Y" -and $confirmation -ne "y") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    exit 0
}

# Create the selected databases
$successCount = 0
$failCount = 0

foreach ($dbName in $selectedDbs) {
    Write-Host "`nCreating database '$dbName'..." -ForegroundColor Cyan
    
    try {
        # Check if database already exists
        $server = New-Object Microsoft.SqlServer.Management.Smo.Server($sqlInstance)
        if ($server.Databases[$dbName]) {
            Write-Host "Database '$dbName' already exists. Skipping." -ForegroundColor Yellow
            $failCount++
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
    }
}

# Summary
Write-Host "`n===== Creation Summary =====" -ForegroundColor Cyan
Write-Host "Total databases selected: $($selectedDbs.Count)" -ForegroundColor White
Write-Host "Successfully created: $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "Failed to create: $failCount" -ForegroundColor Red
}

Write-Host "`nDatabase creation process completed." -ForegroundColor Cyan