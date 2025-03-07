# Install wM products

# Parameters for installer and image paths
param(
    [string]$InstallerPath = "C:\Temp\installer\Installer20240626-w64.exe",
    [string]$FallbackInstallerPath = "\\network\share\Installer20240626-w64.exe",
    [string]$ImagePath = "C:\Temp\installer\image.zip",
    [string]$FallbackImagePath = "\\network\share\image.zip"
)

function Get-ScriptDirectory {
    $scriptPath = $MyInvocation.MyCommand.Path
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

function Get-ProductDescription {
    param(
        [string]$FilePath
    )
    
    $firstLine = Get-Content -Path $FilePath -TotalCount 1
    if ($firstLine -match "^#(.*)") {
        return $Matches[1].Trim()
    }
    else {
        return (Split-Path $FilePath -Leaf)
    }
}

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

function Ensure-FileExists {
    param(
        [string]$FilePath,
        [string]$FallbackPath,
        [string]$FileDescription = "File"
    )
    
    if (-not (Test-Path $FilePath)) {
        Write-Host "$FileDescription not found at $FilePath" -ForegroundColor Yellow
        Write-Host "Attempting to copy from $FallbackPath..." -ForegroundColor Yellow
        
        # Create directory if it doesn't exist
        $targetDir = Split-Path $FilePath -Parent
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        
        # Copy the file
        try {
            Copy-Item -Path $FallbackPath -Destination $FilePath -Force
            Write-Host "$FileDescription copied successfully." -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Failed to copy $FileDescription: $_" -ForegroundColor Red
            return $false
        }
    }
    
    return $true
}

function Install-Product {
    param(
        [string]$InstallerPath,
        [string]$ScriptPath
    )
    
    try {
        Write-Host "Installing $ScriptPath..." -ForegroundColor Cyan
        $process = Start-Process -FilePath $InstallerPath -ArgumentList "-readScript `"$ScriptPath`"" -Wait -PassThru -NoNewWindow
        
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

# Main script execution
try {
    # Ensure we're running as administrator
    Invoke-ElevatedScript
    
    # Get script directory and products folder
    $scriptDir = Get-ScriptDirectory
    $productsDir = Join-Path -Path $scriptDir -ChildPath "products"
    
    if (-not (Test-Path $productsDir)) {
        throw "Products directory not found at: $productsDir"
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
    
    # Ensure installer and image.zip exist
    $installerExists = Ensure-FileExists -FilePath $InstallerPath -FallbackPath $FallbackInstallerPath -FileDescription "Installer"
    
    if (-not $installerExists) {
        throw "Installer not found and could not be copied."
    }
    
    $imageExists = Ensure-FileExists -FilePath $ImagePath -FallbackPath $FallbackImagePath -FileDescription "Image file"
    
    if (-not $imageExists) {
        throw "Image file not found and could not be copied."
    }
    
    # Install each selected product
    $results = @()
    foreach ($product in $selectedProducts) {
        $success = Install-Product -InstallerPath $InstallerPath -ScriptPath $product.Path
        $results += @{
            Product = $product.Description
            Success = $success
        }
    }
    
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