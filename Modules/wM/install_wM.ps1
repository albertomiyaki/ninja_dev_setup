# Install wM products

# Parameters for file paths
param(
    [string]$BaseLocalPath = "C:\Temp\myProject",
    [string]$ConfigFile = "files-config.json"
)

# Helper Functions

function Get-ScriptDirectory {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) {
        Write-Host "Warning: Unable to determine script path using MyInvocation.MyCommand.Path" -ForegroundColor Yellow
        # Try alternate method
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            Write-Host "Warning: Unable to determine script path using PSCommandPath" -ForegroundColor Yellow
            # Fallback to current directory
            return (Get-Location).Path
        }
    }
    return Split-Path $scriptPath -Parent
}

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


# File Handling Functions

function Get-FileHash {
    param(
        [string]$FilePath,
        [string]$Algorithm = "SHA256"
    )
    
    try {
        $hash = Get-FileHash -Path $FilePath -Algorithm $Algorithm
        return $hash.Hash
    }
    catch {
        Write-Host "Failed to calculate hash for $FilePath : $_" -ForegroundColor Yellow
        return $null
    }
}

function Ensure-FileExists {
    param(
        [string]$BaseLocalPath,
        [PSCustomObject]$FileInfo
    )
    
    if (-not $BaseLocalPath) {
        Write-Host "Error: Base local path is null or empty" -ForegroundColor Red
        return $null
    }
    
    if (-not $FileInfo -or -not $FileInfo.relativePath) {
        Write-Host "Error: File information is missing or invalid" -ForegroundColor Red
        return $null
    }
    
    # Construct full local path
    $fullLocalPath = Join-Path -Path $BaseLocalPath -ChildPath $FileInfo.relativePath
    
    if (-not (Test-Path $fullLocalPath)) {
        Write-Host "$($FileInfo.description) not found at $fullLocalPath" -ForegroundColor Yellow
        
        if (-not $FileInfo.fallbackPath) {
            Write-Host "Error: Fallback path is missing for $($FileInfo.relativePath)" -ForegroundColor Red
            return $null
        }
        
        Write-Host "Attempting to copy from $($FileInfo.fallbackPath)..." -ForegroundColor Yellow
        
        # Create directory if it doesn't exist
        $targetDir = Split-Path $fullLocalPath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            Write-Host "Created directory: $targetDir" -ForegroundColor Green
        }
        
        # Copy the file
        try {
            if (-not (Test-Path $FileInfo.fallbackPath)) {
                Write-Host "Error: Fallback path does not exist: $($FileInfo.fallbackPath)" -ForegroundColor Red
                return $null
            }
            
            Copy-Item -Path $FileInfo.fallbackPath -Destination $fullLocalPath -Force
            Write-Host "$($FileInfo.description) copied successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to copy $($FileInfo.description): $_" -ForegroundColor Red
            return $null
        }
    }
    
    # Verify hash if provided
    if ($FileInfo.verifyHash -and $FileInfo.verifyHash -ne "") {
        $fileHash = Get-FileHash -FilePath $fullLocalPath
        if ($fileHash -and $fileHash -ne $FileInfo.verifyHash) {
            Write-Host "Hash verification failed for $($FileInfo.description) !" -ForegroundColor Red
            Write-Host "Expected: $($FileInfo.verifyHash)" -ForegroundColor Yellow
            Write-Host "Actual  : $fileHash" -ForegroundColor Yellow
            
            # Try to recopy the file
            try {
                if (-not (Test-Path $FileInfo.fallbackPath)) {
                    Write-Host "Error: Fallback path does not exist: $($FileInfo.fallbackPath)" -ForegroundColor Red
                    return $null
                }
                
                Write-Host "Attempting to recopy file..." -ForegroundColor Yellow
                Copy-Item -Path $FileInfo.fallbackPath -Destination $fullLocalPath -Force
                
                # Verify hash again
                $fileHash = Get-FileHash -FilePath $fullLocalPath
                if ($fileHash -and $fileHash -ne $FileInfo.verifyHash) {
                    Write-Host "Hash verification failed again for $($FileInfo.description) !" -ForegroundColor Red
                    return $null
                }
            }
            catch {
                Write-Host "Failed to recopy $($FileInfo.description): $_" -ForegroundColor Red
                return $null
            }
        }
    }
    
    return $fullLocalPath
}


# Product Management Functions

function Get-ProductDescription {
    param(
        [string]$FilePath
    )
    
    if (-not $FilePath -or -not (Test-Path $FilePath)) {
        return "Unknown product"
    }
    
    try {
        $firstLine = Get-Content -Path $FilePath -TotalCount 1 -ErrorAction Stop
        if ($firstLine -match "^#(.*)") {
            return $Matches[1].Trim()
        }
        else {
            return (Split-Path $FilePath -Leaf)
        }
    }
    catch {
        Write-Host "Warning: Could not read product description: $_" -ForegroundColor Yellow
        return (Split-Path $FilePath -Leaf)
    }
}

function Test-NeedsPassword {
    param(
        [string]$FilePath
    )
    
    if (-not $FilePath -or -not (Test-Path $FilePath)) {
        return $false
    }
    
    try {
        # Check if file contains a line starting with "adminPassword="
        $content = Get-Content -Path $FilePath -ErrorAction Stop
        return ($content -match "^adminPassword=")
    }
    catch {
        Write-Host "Warning: Could not check if product needs password: $_" -ForegroundColor Yellow
        return $false
    }
}

function Create-TempProductCopy {
    param(
        [string]$FilePath,
        [string]$Password
    )
    
    # Create a temp file path
    $tempDir = [System.IO.Path]::GetTempPath()
    $tempFileName = [System.IO.Path]::GetRandomFileName() + ".tmp"
    $tempFilePath = Join-Path -Path $tempDir -ChildPath $tempFileName
    
    # Read the content and replace adminPassword line
    $content = Get-Content -Path $FilePath
    $newContent = $content -replace "^adminPassword=.*", "adminPassword=$Password"
    
    # Write to temp file
    $newContent | Set-Content -Path $tempFilePath
    
    Write-Host "Created temporary copy of product file with password" -ForegroundColor Green
    return $tempFilePath
}

function Install-Product {
    param(
        [string]$InstallerPath,
        [string]$ScriptPath,
        [string]$InstallerArgument
    )
    
    try {
        Write-Host "Installing using script: $ScriptPath..." -ForegroundColor Cyan
        Write-Host "Using installer: $InstallerPath" -ForegroundColor Cyan
        $process = Start-Process -FilePath $InstallerPath -ArgumentList "$InstallerArgument `"$ScriptPath`"" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            return $true
        }
        else {
            Write-Host "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error running installer: $_" -ForegroundColor Red
        return $false
    }
}


# User Interface Functions

function Show-Menu {
    param(
        [array]$Products
    )
    
    $selectedIndices = @()
    $currentIndex = 0
    $exitMenu = $false
    
    while (-not $exitMenu) {
        Clear-Host
        Write-Host "=== Product Selection Menu ===" -ForegroundColor Cyan
        Write-Host "Use arrow keys to navigate, Spacebar to select/deselect, A to select all, N to deselect all, Enter to confirm"
        Write-Host ""
        
        for ($i = 0; $i -lt $Products.Count; $i++) {
            $product = $Products[$i]
            $selected = $selectedIndices -contains $i
            $indicator = if ($selected) { "[X]" } else { "[ ]" }
            $highlight = if ($i -eq $currentIndex) { "->" } else { "  " }
            
            Write-Host "$highlight $indicator $($product.Description)"
        }
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($currentIndex -gt 0) {
                    $currentIndex--
                }
            }
            40 { # Down arrow
                if ($currentIndex -lt ($Products.Count - 1)) {
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
                $selectedIndices = @(0..($Products.Count - 1))
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
        [array]$SelectedProducts
    )
    
    Clear-Host
    Write-Host "=== Confirm Selection ===" -ForegroundColor Cyan
    Write-Host "The following products will be installed:" -ForegroundColor Yellow
    
    foreach ($product in $SelectedProducts) {
        Write-Host " - $($product.Description)"
    }
    
    Write-Host ""
    $confirmation = Read-Host "Do you want to proceed? (Y/N)"
    
    return $confirmation -eq "Y" -or $confirmation -eq "y"
}


# Main Script Execution

try {
    # Ensure we're running as administrator
    Invoke-ElevatedScript
    
    # Get script directory and set paths
    $scriptDir = Get-ScriptDirectory
    
    if (-not $scriptDir) {
        throw "Unable to determine script directory."
    }
    
    $configFilePath = Join-Path -Path $scriptDir -ChildPath $ConfigFile
    
    if (-not $ConfigFile) {
        throw "Config file name is null or empty."
    }
    
    $productsDir = Join-Path -Path $scriptDir -ChildPath "products"
    
    # Load configuration file
    if (-not (Test-Path $configFilePath)) {
        throw "Configuration file not found at: $configFilePath"
    }
    
    try {
        $configContent = Get-Content -Path $configFilePath -Raw -ErrorAction Stop
        $config = $configContent | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Failed to read or parse configuration file: $_"
    }
    
    if (-not $config.requiredFiles) {
        throw "Invalid configuration file. Missing 'requiredFiles' section."
    }
    
    # Set installer arguments from config
    $installerArgs = if ($config.installerArgs) { $config.installerArgs } else { "-readScript" }
    
    # Ensure products directory exists
    if (-not (Test-Path $productsDir)) {
        throw "Products directory not found at: $productsDir"
    }
    
    # Ensure base local path exists
    if (-not (Test-Path $BaseLocalPath)) {
        New-Item -ItemType Directory -Path $BaseLocalPath -Force | Out-Null
        Write-Host "Created base directory: $BaseLocalPath" -ForegroundColor Green
    }
    
    # Ensure all required files exist
    $requiredPaths = @{}
    foreach ($fileInfo in $config.requiredFiles) {
        $filePath = Ensure-FileExists -BaseLocalPath $BaseLocalPath -FileInfo $fileInfo
        
        if (-not $filePath) {
            throw "Required file '$($fileInfo.relativePath)' could not be copied or verified."
        }
        
        # Store the path for later use
        $requiredPaths[$fileInfo.relativePath] = $filePath
    }
    
    # Find installer path in required paths
    $installerPath = $null
    $installerKey = $requiredPaths.Keys | Where-Object { $_ -like "*Installer*.exe" } | Select-Object -First 1
    
    if ($installerKey) {
        $installerPath = $requiredPaths[$installerKey]
    }
    
    if (-not $installerPath) {
        throw "Could not find installer executable in required files."
    }
    
    # Get all product files
    $productFiles = Get-ChildItem -Path $productsDir -File
    
    if ($productFiles.Count -eq 0) {
        throw "No product files found in: $productsDir"
    }
    
    # Gather product information
    $products = @()
    foreach ($file in $productFiles) {
        $description = Get-ProductDescription -FilePath $file.FullName
        $products += @{
            Path = $file.FullName
            Description = $description
        }
    }
    
    # Show menu and get user selection
    $selectedIndices = Show-Menu -Products $products
    $selectedProducts = $selectedIndices | ForEach-Object { $products[$_] }
    
    # If no products selected, exit
    if ($selectedProducts.Count -eq 0) {
        Write-Host "No products selected. Exiting." -ForegroundColor Yellow
        exit
    }
    
    # Confirm selection
    $confirmed = Confirm-Selection -SelectedProducts $selectedProducts
    
    if (-not $confirmed) {
        Write-Host "Installation cancelled by user." -ForegroundColor Yellow
        exit
    }
    
    # Check if any selected products need a password
    $passwordNeeded = $false
    foreach ($product in $selectedProducts) {
        if (Test-NeedsPassword -FilePath $product.Path) {
            $passwordNeeded = $true
            break
        }
    }
    
    # If password needed, prompt once with confirmation
    $adminPassword = $null
    if ($passwordNeeded) {
        $passwordsMatch = $false
        while (-not $passwordsMatch) {
            Write-Host "Admin password is required for one or more products" -ForegroundColor Yellow
            $securePassword = Read-Host "Enter admin password" -AsSecureString
            $secureConfirmPassword = Read-Host "Confirm admin password" -AsSecureString
            
            # Convert secure strings to plain text for comparison
            $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR1)
            
            $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfirmPassword)
            $plainConfirmPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR2)
            
            # Compare passwords
            if ($plainPassword -eq $plainConfirmPassword) {
                $passwordsMatch = $true
                $adminPassword = $plainPassword
                Write-Host "Passwords match" -ForegroundColor Green
            } else {
                Write-Host "Passwords do not match. Please try again." -ForegroundColor Red
            }
            
            # Clean up
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR1)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR2)
            $plainPassword = $null
            $plainConfirmPassword = $null
        }
    }
    
    # Install each selected product
    $results = @()
    $tempFiles = @()
    
    foreach ($product in $selectedProducts) {
        $scriptPath = $product.Path
        $needsPassword = Test-NeedsPassword -FilePath $scriptPath
        
        # If password is needed, create temp copy with the admin password
        if ($needsPassword -and $adminPassword) {
            Write-Host "Using admin password for $($product.Description)" -ForegroundColor Yellow
            
            # Create temp copy with password
            $tempPath = Create-TempProductCopy -FilePath $scriptPath -Password $adminPassword
            $tempFiles += $tempPath
            $scriptPath = $tempPath
        }
        
        # Run the installer with the appropriate script path
        $success = Install-Product -InstallerPath $installerPath -ScriptPath $scriptPath -InstallerArgument $installerArgs
        $results += @{
            Product = $product.Description
            Success = $success
        }
    }
    
    # Clean up any temp files
    foreach ($tempFile in $tempFiles) {
        if (Test-Path $tempFile) {
            Remove-Item -Path $tempFile -Force
            Write-Host "Removed temporary file: $tempFile" -ForegroundColor Green
        }
    }
    
    # Clear the password from memory
    $adminPassword = $null
    
    # Display summary
    Clear-Host
    Write-Host "=== Installation Summary ===" -ForegroundColor Cyan
    
    $successCount = ($results | Where-Object { $_.Success }).Count
    $failCount = ($results | Where-Object { -not $_.Success }).Count
    
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failCount" -ForegroundColor Red
    Write-Host ""
    
    if ($failCount -gt 0) {
        Write-Host "Failed Products:" -ForegroundColor Red
        $results | Where-Object { -not $_.Success } | ForEach-Object {
            Write-Host " - $($_.Product)" -ForegroundColor Red
        }
    }
    
    if ($successCount -gt 0) {
        Write-Host "Successful Products:" -ForegroundColor Green
        $results | Where-Object { $_.Success } | ForEach-Object {
            Write-Host " - $($_.Product)" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
