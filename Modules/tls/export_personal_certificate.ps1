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

# Verify keytool is available
function Test-KeytoolAvailable {
    try {
        $null = & keytool -help 2>&1
        return $true
    } catch {
        return $false
    }
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
        Write-Host "===================================" -ForegroundColor Cyan
        Write-Host "        $Title" -ForegroundColor Cyan
        Write-Host "===================================" -ForegroundColor Cyan
        
        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $selection) {
                Write-Host "  [>>] " -NoNewline -ForegroundColor Green
                & $DisplayFunction $Options[$i] $true
            } else {
                Write-Host "  [  ] " -NoNewline -ForegroundColor Gray
                & $DisplayFunction $Options[$i] $false
            }
        }
        
        Write-Host "(Use ↑ and ↓ arrow keys to navigate, press Enter to select)" -ForegroundColor Gray
        
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
    $keyType = "Unknown"
    
    # Try to determine key type
    try {
        if ($Certificate.HasPrivateKey) {
            if ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)) {
                $keyType = "RSA"
            } elseif ([System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($Certificate)) {
                $keyType = "ECDSA"
            } elseif ([System.Security.Cryptography.X509Certificates.DSACertificateExtensions]::GetDSAPrivateKey($Certificate)) {
                $keyType = "DSA"
            }
        } else {
            $keyType = "No Private Key"
        }
    } catch {
        # Ignore errors when determining key type
    }
    
    # Format the subject for better display
    $subjectCN = if ($subject -match "CN=([^,]+)") { $matches[1] } else { $subject }
    
    Write-Host "$subjectCN" -ForegroundColor $color
    Write-Host "      Thumbprint: $thumbprint" -ForegroundColor $color
    Write-Host "      Key Type: $keyType" -ForegroundColor $color
    Write-Host "      Expiration: $expiration" -ForegroundColor $color
    Write-Host "      Issuer: $issuer" -ForegroundColor $color
}

function Get-ExportableCertificates {
    # Get certificates from the personal store that have exportable private keys
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadOnly")
    
    $exportableCerts = @()
    
    foreach ($cert in $store.Certificates) {
        if ($cert.HasPrivateKey) {
            try {
                # Create a temporary in-memory PFX to test if we can export the key
                $tempPassword = ConvertTo-SecureString -String "TempPassword123!" -Force -AsPlainText
                $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $tempPassword)
                
                # If we get here without exception, it's exportable
                $exportableCerts += $cert
            } catch {
                # Certificate's private key is not exportable
                Write-Debug "Certificate $($cert.Subject) has a non-exportable private key: $_"
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
    
    return $password, $passwordText
}

function Export-CertificateToPfxWithKeytool {
    param(
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Password,
        [string]$FilePath,
        [string]$KeyAlias,
        [string]$Algorithm = "AES256"
    )
    
    try {
        # First export the certificate with private key to a temporary PFX file
        $tempDir = [System.IO.Path]::GetTempPath()
        $tempPfxFile = Join-Path -Path $tempDir -ChildPath "temp_$([System.Guid]::NewGuid().ToString()).pfx"
        $securePassword = ConvertTo-SecureString -String $Password -Force -AsPlainText
        
        Write-Host "Exporting certificate to temporary file..." -ForegroundColor Yellow
        Export-PfxCertificate -Cert $Certificate -FilePath $tempPfxFile -Password $securePassword -Force | Out-Null
        
        # Create the destination directory if it doesn't exist
        $destinationFolder = Split-Path -Path $FilePath -Parent
        if (-not (Test-Path -Path $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $destinationFolder" -ForegroundColor Yellow
        }
        
        # Prepare keytool import command with the proper algorithm
        $keystoreType = "PKCS12"
        
        # Set the appropriate keystore algorithm based on user selection
        $keystoreAlgorithm = switch ($Algorithm) {
            "AES256" { "PBEWithHmacSHA256AndAES_256" }
            "3DES" { "PBEWithSHA1AndDESede" }
            default { "PBEWithHmacSHA256AndAES_256" }
        }
        
        # Get the current alias from the PFX file using the verbose flag for reliable parsing
        Write-Host "Identifying original alias in certificate store..." -ForegroundColor Yellow
        $listAliasCmd = "keytool -list -v -keystore `"$tempPfxFile`" -storetype PKCS12 -storepass `"$Password`""
        $listOutput = Invoke-Expression $listAliasCmd 2>&1
        
        $originalAlias = $null
        
        # Parse the verbose output to find the PrivateKeyEntry alias
        if ($listOutput -match "Alias name:\s*([^\r\n]+)[\r\n]+[^\r\n]*[\r\n]+Entry type: PrivateKeyEntry") {
            $originalAlias = $matches[1].Trim()
            Write-Host "Found original alias: $originalAlias" -ForegroundColor Green
        } else {
            # Fallback to non-verbose parsing if verbose parsing fails
            $listAliasCmd = "keytool -list -keystore `"$tempPfxFile`" -storetype PKCS12 -storepass `"$Password`""
            $listOutput = Invoke-Expression $listAliasCmd 2>&1
            
            # Look for pattern like "alias, date, PrivateKeyEntry,"
            if ($listOutput -match "([^,]+),\s*[^,]+,\s*PrivateKeyEntry,") {
                $originalAlias = $matches[1].Trim()
                Write-Host "Found original alias (non-verbose): $originalAlias" -ForegroundColor Green
            } else {
                throw "Could not determine original alias in PFX file"
            }
        }
        
        # Import the PFX to the destination keystore
        Write-Host "Importing certificate to keystore..." -ForegroundColor Yellow
        
        $keytoolImportArgs = @(
            "-importkeystore",
            "-srckeystore", "`"$tempPfxFile`"",
            "-srcstoretype", "PKCS12",
            "-srcstorepass", "`"$Password`"",
            "-destkeystore", "`"$FilePath`"",
            "-deststoretype", "$keystoreType",
            "-deststorepass", "`"$Password`"",
            "-destkeypass", "`"$Password`"",
            "-srcalias", "`"$originalAlias`"",
            "-destalias", "`"$KeyAlias`"",
            "-noprompt"
        )
        
        # If algorithm is not default, add the appropriate flags
        if ($Algorithm -ne "Default") {
            $keytoolImportArgs += @(
                "-J-Dkeystore.pkcs12.keyProtectionAlgorithm=$keystoreAlgorithm"
            )
        }
        
        $keytoolImportCommand = "keytool $($keytoolImportArgs -join ' ')"
        
        # Execute the command
        $keytoolOutput = Invoke-Expression $keytoolImportCommand 2>&1
        
        # Check if the command was successful
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Keytool import failed with exit code $LASTEXITCODE" -ForegroundColor Red
            Write-Host "Output: $keytoolOutput" -ForegroundColor Red
            throw "Failed to import certificate with keytool"
        }
        
        # Cleanup temporary file
        if (Test-Path -Path $tempPfxFile) {
            Remove-Item -Path $tempPfxFile -Force
        }
        
        Write-Host "`nCertificate successfully exported to: $FilePath" -ForegroundColor Green
        Write-Host "Private key alias: $KeyAlias" -ForegroundColor Green
        Write-Host "Remember to store the password securely!" -ForegroundColor Yellow
    }
    catch {
        Write-Host "`nError exporting certificate with keytool: $_" -ForegroundColor Red
        
        # Fallback to standard Export-PfxCertificate if keytool method fails
        try {
            Write-Host "Falling back to standard Windows export method (without custom alias)..." -ForegroundColor Yellow
            $securePassword = ConvertTo-SecureString -String $Password -Force -AsPlainText
            Export-PfxCertificate -Cert $Certificate -FilePath $FilePath -Password $securePassword -Force | Out-Null
            Write-Host "Certificate exported successfully with default alias." -ForegroundColor Green
        } catch {
            Write-Host "Fallback export failed: $_" -ForegroundColor Red
        }
    }
}

# Main Script

Write-Host "Export Personal Certificate " -ForegroundColor Cyan
Write-Host "-------------------------------------`n" 

# Check if keytool is available
if (-not (Test-KeytoolAvailable)) {
    Write-Host "Error: 'keytool' command not found in PATH." -ForegroundColor Red
    Write-Host "Please make sure Java is installed and keytool is available in your PATH." -ForegroundColor Yellow
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    exit
}

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

$selectedCert = Show-Menu -Title "Select Certificate to Export" -Options $exportableCerts -DisplayFunction ${function:Display-Certificate}

# Ask for password
$securePassword, $passwordText = Get-Password

# Select encryption algorithm
$algorithms = @("Default", "AES256", "3DES")
$selectedAlgorithm = Show-Menu -Title "Select Encryption Algorithm" -Options $algorithms -DisplayFunction {
    param($Algorithm, $Selected)
    $color = if ($Selected) { "Green" } else { "White" }
    Write-Host "Encryption: $Algorithm" -ForegroundColor $color
}

# Get hostname for default path and alias
$hostname = [System.Net.Dns]::GetHostName().ToLower()

# Create default directory and file path
$defaultDirectory = Join-Path -Path $env:USERPROFILE -ChildPath "tls"
$defaultFileName = "${hostname}_keystore.p12"
$defaultPath = Join-Path -Path $defaultDirectory -ChildPath $defaultFileName

# Ask for file path
Write-Host "`nExport Location" -ForegroundColor Cyan
Write-Host "Default path: $defaultPath" -ForegroundColor Yellow
$userPath = Read-Host "Enter export path (or press Enter for default)"
$filePath = if ([string]::IsNullOrWhiteSpace($userPath)) { $defaultPath } else { $userPath }

# Use hostname as the key alias
$keyAlias = $hostname

# Export the certificate
Write-Host "`nExporting certificate with alias '$keyAlias'..." -ForegroundColor Yellow
Export-CertificateToPfxWithKeytool -Certificate $selectedCert -Password $passwordText -FilePath $filePath -KeyAlias $keyAlias -Algorithm $selectedAlgorithm

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null