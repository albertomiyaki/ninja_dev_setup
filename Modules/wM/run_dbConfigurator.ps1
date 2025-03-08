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
        Write-Host "Elevating privileges to run as administrator..."
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
        exit 1
    }
    
    try {
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        return $config
    }
    catch {
        Write-Host "Error: Failed to parse configuration file: $_" -ForegroundColor Red
        exit 1
    }
}

function Show-Menu {
    param(
        [array]$Instances
    )
    
    $selectedIndices = @()
    $currentIndex = 0
    $exitMenu = $false
    
    while (-not $exitMenu) {
        Clear-Host
        Write-Host "=== Database Instance Selection Menu ===" -ForegroundColor Cyan
        Write-Host "Use arrow keys to navigate, Spacebar to select/deselect, A to select all, N to deselect all, Enter to confirm"
        Write-Host ""
        
        for ($i = 0; $i -lt $Instances.Count; $i++) {
            $instance = $Instances[$i]
            $selected = $selectedIndices -contains $i
            $indicator = if ($selected) { "[X]" } else { "[ ]" }
            $highlight = if ($i -eq $currentIndex) { "->" } else { "  " }
            
            Write-Host "$highlight $indicator $($instance.instanceName) (DB: $($instance.databaseName))"
        }
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentIndex -gt 0) {
                    $currentIndex--
                }
            }
            40 { # Down arrow
                if ($currentIndex -lt ($Instances.Count - 1)) {
                    $currentIndex++
                }
            }
            32 { # Spacebar
                if ($selectedIndices -contains $currentIndex) {
                    $selectedIndices = $selectedIndices | Where-Object { $_ -ne $currentIndex }
                }
                else {
                    $selectedIndices += $currentIndex
                }
            }
            65 { # 'A' key
                $selectedIndices = @(0..($Instances.Count - 1))
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

function Confirm-Selection {
    param(
        [array]$SelectedInstances
    )
    
    Clear-Host
    Write-Host "=== Confirm Selection ===" -ForegroundColor Cyan
    Write-Host "The following database instances will be configured:" -ForegroundColor Yellow
    
    foreach ($instance in $SelectedInstances) {
        Write-Host " - $($instance.instanceName) (DB: $($instance.databaseName))"
    }
    
    Write-Host ""
    $confirmation = Read-Host "Do you want to proceed? (Y/N)"
    
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
        }
        else {
            Write-Host "Configuration failed for $($Instance.instanceName) with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
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
        exit 1
    }
    
    # Load configuration
    $config = Load-Configuration -ConfigPath $ConfigFilePath
    
    if (-not $config -or -not $config.instances -or $config.instances.Count -eq 0) {
        Write-Host "Error: No database instances found in configuration file" -ForegroundColor Red
        exit 1
    }
    
    # Show menu and get user selection
    $selectedIndices = Show-Menu -Instances $config.instances
    $selectedInstances = $selectedIndices | ForEach-Object { $config.instances[$_] }
    
    # If no instances selected, exit
    if ($selectedInstances.Count -eq 0) {
        Write-Host "No instances selected. Exiting." -ForegroundColor Yellow
        exit
    }
    
    # Confirm selection
    $confirmed = Confirm-Selection -SelectedInstances $selectedInstances
    
    if (-not $confirmed) {
        Write-Host "Configuration cancelled by user." -ForegroundColor Yellow
        exit
    }
    
    # Final confirmation
    Write-Host ""
    $finalConfirmation = Read-Host "Are you absolutely sure you want to proceed with the configuration? (Y/N)"
    
    if ($finalConfirmation -ne "Y" -and $finalConfirmation -ne "y") {
        Write-Host "Configuration cancelled by user." -ForegroundColor Yellow
        exit
    }
    
    # Prompt for username (with default)
    $defaultUsername = "sa"
    $username = Read-Host "Enter database username [$defaultUsername]"
    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = $defaultUsername
    }
    
    # Prompt for password with confirmation
    $passwordsMatch = $false
    $password = $null
    
    while (-not $passwordsMatch) {
        $securePassword = Read-Host "Enter database password" -AsSecureString
        $secureConfirmPassword = Read-Host "Confirm database password" -AsSecureString
        
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
    Write-Host "=== Configuration Summary ===" -ForegroundColor Cyan
    
    $successCount = ($results | Where-Object { $_.Success }).Count
    $failCount = ($results | Where-Object { -not $_.Success }).Count
    
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failCount" -ForegroundColor Red
    Write-Host ""
    
    if ($failCount -gt 0) {
        Write-Host "Failed Instances:" -ForegroundColor Red
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host " - $($_.Instance)" -ForegroundColor Red
        }
    }
    
    if ($successCount -gt 0) {
        Write-Host "Successful Instances:" -ForegroundColor Green
        $results | Where-Object { $_.Success } | ForEach-Object {
            Write-Host " - $($_.Instance)" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}