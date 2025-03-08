# Export Personal Certificate

# Script to select and export personal certificates with private keys to PKCS12 format

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

function Show-Menu {
    param (
        [string]$Title,
        [array]$Options,
        [scriptblock]$DisplayFunction
    )
    
    $selection = 0
    $maxIndex = $Options.Count - 1
    $enterPressed = $false
    
    while (-not $enterPressed) {
        Clear-Host
        Write-Host "`n $Title`n" -ForegroundColor Cyan
        Write-Host " Use up/down arrow keys to navigate, Enter to select`n" -ForegroundColor Gray
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $selection) {
                Write-Host " > " -NoNewline -ForegroundColor Green
                & $DisplayFunction $Options[$i] $true
            } else {
                Write-Host "   " -NoNewline
                & $DisplayFunction $Options[$i] $false
            }
        }
        
        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($selection -gt 0) { $selection-- }
                else { $selection = $maxIndex }
            }
            40 { # Down arrow
                if ($selection -lt $maxIndex) { $selection++ }
                else { $selection = 0 }
            }
            13 { # Enter
                $enterPressed = $true
            }
        }
    }
    
    return $Options[$selection]
}

function Display-Certificate {
    param (
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [bool]$Selected
    )
    
    $color = if ($Selected) { "Green" } else { "White" }
    $subject = $Certificate.Subject
    $issuer = $Certificate.Issuer
    $thumbprint = $Certificate.Thumbprint
    $expiration = $Certificate.NotAfter.ToString("yyyy-MM-dd")
    
    # Format the subject for better display
    $subjectCN = if ($subject -match "CN=([^,]+)") { $matches[1] } else { $subject }
    
    Write-Host "$subjectCN" -ForegroundColor $color
    Write-Host "      Thumbprint: $thumbprint" -ForegroundColor Gray
    Write-Host "      Expiration: $expiration" -ForegroundColor Gray
    Write-Host "      Issuer: $issuer" -ForegroundColor Gray
    Write-Host ""
}

function Get-ExportableCertificates {
    # Get certificates from the personal store that have exportable private keys
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadOnly")
    
    $exportableCerts = @()
    
    foreach ($cert in $store.Certificates) {
        if ($cert.HasPrivateKey) {
            try {
                $key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
                if ($key -and $key.Key.ExportPolicy -ne "NonExportable") {
                    $exportableCerts += $cert
                }
            } catch {
                # Skip certificates where we can't check exportability
            }
        }
    }
    
    $store.Close()
    return $exportableCerts
}

function Get-Password {
    $passwordMatch = $false
    $password = $null
    
    while (-not $passwordMatch) {
        $securePassword = Read-Host "Enter password for certificate protection" -AsSecureString
        $securePasswordConfirm = Read-Host "Confirm password" -AsSecureString
        
        $passwordText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
        $passwordConfirmText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePasswordConfirm))
        
        if ($passwordText -eq $passwordConfirmText) {
            $passwordMatch = $true
            $password = $securePassword
        } else {
            Write-Host "Passwords do not match. Please try again." -ForegroundColor Red
        }
    }
    
    return $password
}

function Export-CertificateToP12 {
    param (
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [System.Security.SecureString]$Password,
        [string]$FilePath,
        [string]$Algorithm
    )
    
    try {
        # Create export options based on the selected algorithm
        $exportParams = @{
            Cert = $Certificate
            FilePath = $FilePath
            Password = $Password
            Force = $true
        }
        
        # Set encryption algorithm if specified
        if ($Algorithm -ne "Default") {
            $exportParams.ChainOption = "EndEntityCertOnly"
            
            switch ($Algorithm) {
                "AES256" {
                    $exportParams.CryptoAlgorithmOption = "AES256_SHA256"
                }
                "3DES" {
                    $exportParams.CryptoAlgorithmOption = "TripleDES_SHA1"
                }
            }
        }
        
        # Export the certificate
        Export-PfxCertificate @exportParams | Out-Null
        
        Write-Host "`nCertificate successfully exported to: $FilePath" -ForegroundColor Green
        Write-Host "Remember to store the password securely!" -ForegroundColor Yellow
    } catch {
        Write-Host "`nError exporting certificate: $_" -ForegroundColor Red
    }
}

# Main Script

Write-Host "Export Personal Certificate with Private Key" -ForegroundColor Cyan
Write-Host "----------------------------------------------`n" 

# Get exportable certificates
$exportableCerts = Get-ExportableCertificates

if ($exportableCerts.Count -eq 0) {
    Write-Host "No exportable certificates with private keys were found in your personal store." -ForegroundColor Red
    Write-Host "Please ensure you have certificates with exportable private keys installed." -ForegroundColor Yellow
    exit
}

# Display the menu and let user select a certificate
Write-Host "Retrieving certificates with exportable private keys..." -ForegroundColor Yellow
Start-Sleep -Seconds 1

$selectedCert = Show-Menu -Title "Select a certificate to export:" -Options $exportableCerts -DisplayFunction ${function:Display-Certificate}

# Ask for password
$password = Get-Password

# Select encryption algorithm
$algorithms = @("Default", "AES256", "3DES")
$selectedAlgorithm = Show-Menu -Title "Select encryption algorithm:" -Options $algorithms -DisplayFunction {
    param($Algorithm, $Selected)
    $color = if ($Selected) { "Green" } else { "White" }
    Write-Host "$Algorithm" -ForegroundColor $color
    Write-Host ""
}

# Determine default file path
$defaultFileName = "$($selectedCert.Subject -replace 'CN=|[,=].*', '').p12"
$defaultFileName = $defaultFileName -replace '[\\/:*?"<>|]', '_' # Remove invalid filename characters
$defaultPath = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $defaultFileName

# Ask for file path
Write-Host "`nExport Location" -ForegroundColor Cyan
Write-Host "Default path: $defaultPath" -ForegroundColor Yellow
$userPath = Read-Host "Enter export path (or press Enter for default)"

$filePath = if ([string]::IsNullOrWhiteSpace($userPath)) { $defaultPath } else { $userPath }

# Export the certificate
Write-Host "`nExporting certificate..." -ForegroundColor Yellow
Export-CertificateToP12 -Certificate $selectedCert -Password $password -FilePath $filePath -Algorithm $selectedAlgorithm

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null