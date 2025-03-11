# Run DBConfigurator

# Parameters
param(
    [string]$ConfigFilePath = "run_dbConfigurator_config.json",
    [string]$ExecutablePath = "C:\Program Files\dbTools\dbConfigurator.bat"
)

# Helper Functions
function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-ElevatedScript {
    if (-not (Test-IsAdmin)) {
        Write-Host "Elevating privileges to run as administrator..." -ForegroundColor Yellow
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

function Load-Configuration {
    param (
        [string]$ConfigPath
    )
    
    if (-not (Test-Path $ConfigPath)) {
        Write-Host "Error: Configuration file not found at: $ConfigPath" -ForegroundColor Red
        Wait-ForKeyPress
        exit 1
    }
    
    try {
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        return $config
    }
    catch {
        Write-Host "Error: Failed to parse configuration file: $_" -ForegroundColor Red
        Wait-ForKeyPress
        exit 1
    }
}

function Wait-ForKeyPress {
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

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
    
    return $selectedIndices
}

function Display-InstanceItem {
    param(
        [PSCustomObject]$Instance,
        [bool]$IsHighlighted,
        [bool]$IsSelected
    )
    
    if ($IsHighlighted) {
        Write-Host "$($Instance.instanceName) " -NoNewline -ForegroundColor White
        Write-Host "(DB: $($Instance.databaseName))" -ForegroundColor Yellow
    } else {
        Write-Host "$($Instance.instanceName) " -NoNewline -ForegroundColor Gray
        Write-Host "(DB: $($Instance.databaseName))" -ForegroundColor DarkGray
    }
}

function Confirm-Selection {
    param(
        [array]$SelectedInstances
    )
    
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Confirm Selection" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "The following database instances will be configured:" -ForegroundColor Yellow
    
    foreach ($instance in $SelectedInstances) {
        Write-Host " • $($instance.instanceName) " -NoNewline -ForegroundColor White
        Write-Host "(DB: $($instance.databaseName))" -ForegroundColor Gray
    }
    
    Write-Host "`nDo you want to proceed? (Y/N) " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    return $confirmation -eq "Y" -or $confirmation -eq "y"
}

function Configure-DatabaseInstance {
    param(
        [PSCustomObject]$Instance,
        [string]$Username,
        [string]$Password,
        [string]$ExePath
    )
    
    $cmdArgs = "-a create -d sqlserver -l `"jdbc:wm:sqlserver://localhost:1433;databaseName=$($Instance.databaseName)`" -u `"$Username`" -p `"$Password`" -v latest -c `"$($Instance.componentsList)`""
    
    Write-Host "Configuring $($Instance.instanceName)..." -ForegroundColor Cyan
    
    try {
        $process = Start-Process -FilePath $ExePath -ArgumentList $cmdArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Configuration successful for $($Instance.instanceName)" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Configuration failed for $($Instance.instanceName) with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Error configuring $($Instance.instanceName): $_" -ForegroundColor Red
        return $false
    }
}

# Main Script Execution
try {
    # Ensure we're running as administrator
    Invoke-ElevatedScript
    
    # Validate executable path
    if (-not (Test-Path $ExecutablePath)) {
        Write-Host "Error: dbConfigurator.bat not found at: $ExecutablePath" -ForegroundColor Red
        Wait-ForKeyPress
        exit 1
    }
    
    # Load configuration
    $config = Load-Configuration -ConfigPath $ConfigFilePath
    
    if (-not $config -or -not $config.instances -or $config.instances.Count -eq 0) {
        Write-Host "Error: No database instances found in configuration file" -ForegroundColor Red
        Wait-ForKeyPress
        exit 1
    }
    
    # Show menu and get user selection
    $selectedIndices = Show-MultiSelectMenu -Title "Database Instance Selection" -Options $config.instances -DisplayFunction ${function:Display-InstanceItem}
    $selectedInstances = $selectedIndices | ForEach-Object { $config.instances[$_] }
    
    # If no instances selected, exit
    if ($selectedInstances.Count -eq 0) {
        Write-Host "No instances selected. Exiting." -ForegroundColor Yellow
        Wait-ForKeyPress
        exit
    }
    
    # Confirm selection
    $confirmed = Confirm-Selection -SelectedInstances $selectedInstances
    
    if (-not $confirmed) {
        Write-Host "Configuration cancelled by user." -ForegroundColor Yellow
        Wait-ForKeyPress
        exit
    }

    # Prompt for username (with default)
    $defaultUsername = "sa"
    Write-Host "Enter database username [$defaultUsername]: " -NoNewline -ForegroundColor Cyan
    $username = Read-Host
    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = $defaultUsername
    }
    
    # Prompt for password with confirmation
    $passwordsMatch = $false
    $password = $null
    
    while (-not $passwordsMatch) {
        Write-Host "Enter database password: " -NoNewline -ForegroundColor Cyan
        $securePassword = Read-Host -AsSecureString
        
        Write-Host "Confirm database password: " -NoNewline -ForegroundColor Cyan
        $secureConfirmPassword = Read-Host -AsSecureString
        
        # Convert secure strings for comparison
        $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR1)
        
        $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfirmPassword)
        $confirmPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR2)
        
        # Compare passwords
        if ($password -eq $confirmPassword) {
            $passwordsMatch = $true
            Write-Host "Passwords match" -ForegroundColor Green
        } else {
            Write-Host "Passwords do not match. Please try again." -ForegroundColor Red
        }
        
        # Clean up
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR1)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR2)
        $confirmPassword = $null
    }
    
    # Configure each selected instance
    $results = @()
    
    foreach ($instance in $selectedInstances) {
        $success = Configure-DatabaseInstance -Instance $instance -Username $username -Password $password -ExePath $ExecutablePath
        $results += @{
            Instance = $instance.instanceName
            Success = $success
        }
    }
    
    # Clean up sensitive data
    $password = $null
    
    # Display summary
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host " Configuration Summary" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    
    $successCount = ($results | Where-Object { $_.Success }).Count
    $failCount = ($results | Where-Object { -not $_.Success }).Count
    
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failCount" -ForegroundColor Red
    Write-Host ""
    
    if ($failCount -gt 0) {
        Write-Host "Failed Instances:" -ForegroundColor Red
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host " • $($_.Instance)" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    if ($successCount -gt 0) {
        Write-Host "Successful Instances:" -ForegroundColor Green
        $results | Where-Object { $_.Success } | ForEach-Object {
            Write-Host " • $($_.Instance)" -ForegroundColor Green
        }
    }
    
    # Wait for user to press a key before exiting
    Wait-ForKeyPress
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Wait-ForKeyPress
    exit 1
}