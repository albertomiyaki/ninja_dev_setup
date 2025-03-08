# CSR Request - CMD Tool

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

# Certificate template to use (change as needed)
$certificateTemplate = "WebServer" # Replace with your enterprise template name

# Get hostname information for default values
$hostname = [System.Net.Dns]::GetHostByName(($env:computerName)).HostName
$shortHostname = $env:COMPUTERNAME

function Show-Menu {
    Clear-Host
    Write-Host "===== Certificate Signing Request Generator ====="
    Write-Host "This script will create a CSR and submit it for signing."
    Write-Host
}

function Get-UserInput {
    # Prompt for Common Name with hostname as default
    $commonName = Read-Host "Enter Common Name (default: $hostname)"
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = $hostname }
    
    # Prompt for Friendly Name with default value
    $friendlyName = Read-Host "Enter Friendly Name (default: $shortHostname Certificate)"
    if ([string]::IsNullOrWhiteSpace($friendlyName)) { $friendlyName = "$shortHostname Certificate" }
    
    # Prompt for certificate usage
    Write-Host "`nSelect certificate usage:"
    Write-Host "1. Client Authentication (clientAuth)"
    Write-Host "2. Server Authentication (serverAuth)"
    Write-Host "3. Both Client and Server Authentication"
    
    $usageChoice = Read-Host "Enter your choice (1-3, default: 3)"
    if ([string]::IsNullOrWhiteSpace($usageChoice) -or $usageChoice -notin @("1", "2", "3")) { 
        $usageChoice = "3" 
        Write-Host "Using default: Both Client and Server Authentication" -ForegroundColor Yellow
    }
    
    # Map choice to actual usage
    switch($usageChoice) {
        "1" { $certUsage = "clientAuth" }
        "2" { $certUsage = "serverAuth" }
        "3" { $certUsage = "both" }
    }
    
    # Prompt for additional Subject Alternative Names (SANs)
    Write-Host "`nThe following SANs will be added automatically:"
    Write-Host " - $hostname"
    Write-Host " - $shortHostname"
    $additionalSANs = Read-Host "Enter additional SANs (comma-separated, leave blank for none)"
    
    # Prompt for email address
    $email = Read-Host "Enter your email address (leave blank to omit)"
    
    # Prompt for organization details - all optional
    Write-Host "`nThe following fields are optional. Press Enter to leave any field empty."
    
    $organization = Read-Host "Enter Organization Name (optional)"
    $organizationalUnit = Read-Host "Enter Organizational Unit (optional)"
    $city = Read-Host "Enter City/Locality (optional)"
    $state = Read-Host "Enter State/Province (optional)"
    
    $country = Read-Host "Enter Country (two-letter code, default: US)"
    if ([string]::IsNullOrWhiteSpace($country)) { $country = "US" }
    
    # Return all collected information as a hashtable
    return @{
        CommonName = $commonName
        FriendlyName = $friendlyName
        AdditionalSANs = $additionalSANs
        Email = $email
        Organization = $organization
        OrganizationalUnit = $organizationalUnit
        City = $city
        State = $state
        Country = $country
        CertUsage = $certUsage
    }
}

function Create-CSR {
    param (
        [hashtable]$UserInfo
    )
    
    Write-Host "Creating Certificate Signing Request..." -ForegroundColor Yellow
    
    # Prepare the SAN extension
    $sanList = @($UserInfo.CommonName, $hostname, $shortHostname)
    
    # Add additional SANs if provided
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.AdditionalSANs)) {
        $additionalSANArray = $UserInfo.AdditionalSANs -split ',' | ForEach-Object { $_.Trim() }
        $sanList += $additionalSANArray
    }
    
    # Remove duplicates from SAN list
    $sanList = $sanList | Select-Object -Unique
    
    # Create a new SAN extension
    $sanExtension = New-Object -ComObject X509Enrollment.CX509ExtensionAlternativeNames
    $sanCollection = New-Object -ComObject X509Enrollment.CAlternativeNames
    
    foreach ($san in $sanList) {
        if (-not [string]::IsNullOrWhiteSpace($san)) {
            $altName = New-Object -ComObject X509Enrollment.CAlternativeName
            $altName.InitializeFromString(0x3, $san) # 0x3 represents DNS name
            $sanCollection.Add($altName)
        }
    }
    
    $sanExtension.InitializeEncode($sanCollection)
    
    # Create Enhanced Key Usage extension based on user selection
    $ekuOids = New-Object -ComObject X509Enrollment.CObjectIds
    
    # Add selected usages based on user choice
    if ($UserInfo.CertUsage -eq "clientAuth" -or $UserInfo.CertUsage -eq "both") {
        # Client Authentication OID: 1.3.6.1.5.5.7.3.2
        $clientAuthOid = New-Object -ComObject X509Enrollment.CObjectId
        $clientAuthOid.InitializeFromValue("1.3.6.1.5.5.7.3.2")
        $ekuOids.Add($clientAuthOid)
    }
    
    if ($UserInfo.CertUsage -eq "serverAuth" -or $UserInfo.CertUsage -eq "both") {
        # Server Authentication OID: 1.3.6.1.5.5.7.3.1
        $serverAuthOid = New-Object -ComObject X509Enrollment.CObjectId
        $serverAuthOid.InitializeFromValue("1.3.6.1.5.5.7.3.1")
        $ekuOids.Add($serverAuthOid)
    }
    
    # Create the extension
    $ekuExtension = New-Object -ComObject X509Enrollment.CX509ExtensionEnhancedKeyUsage
    $ekuExtension.InitializeEncode($ekuOids)
    
    # Create the Distinguished Name with only non-empty fields
    $dn = New-Object -ComObject X509Enrollment.CX500DistinguishedName
    
    $dnParts = @()
    $dnParts += "CN=$($UserInfo.CommonName)"
    
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.Organization)) { 
        $dnParts += "O=$($UserInfo.Organization)" 
    }
    
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.OrganizationalUnit)) { 
        $dnParts += "OU=$($UserInfo.OrganizationalUnit)" 
    }
    
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.City)) { 
        $dnParts += "L=$($UserInfo.City)" 
    }
    
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.State)) { 
        $dnParts += "S=$($UserInfo.State)" 
    }
    
    $dnParts += "C=$($UserInfo.Country)"
    
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.Email)) { 
        $dnParts += "E=$($UserInfo.Email)" 
    }
    
    $dnString = $dnParts -join ", "
    $dn.Encode($dnString, 0x0) # 0x0 represents XCN_CERT_NAME_STR_NONE
    
    # Create a private key
    $privateKey = New-Object -ComObject X509Enrollment.CX509PrivateKey
    $privateKey.ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
    $privateKey.KeySpec = 1 # XCN_AT_KEYEXCHANGE
    $privateKey.Length = 4096
    $privateKey.MachineContext = $true
    $privateKey.ExportPolicy = 1 # XCN_NCRYPT_ALLOW_EXPORT_FLAG - Private key is exportable
    $privateKey.Create()
    
    # Create certificate request object
    $cert = New-Object -ComObject X509Enrollment.CX509CertificateRequestPkcs10
    $cert.InitializeFromPrivateKey(1, $privateKey, "") # 1 represents context = machine
    $cert.Subject = $dn
    $cert.FriendlyName = $UserInfo.FriendlyName
    
    # Add the SAN extension to the certificate
    $cert.X509Extensions.Add($sanExtension)
    
    # Add the Enhanced Key Usage extension
    $cert.X509Extensions.Add($ekuExtension)
    
    # Create the enrollment request
    $enrollment = New-Object -ComObject X509Enrollment.CX509Enrollment
    $enrollment.InitializeFromRequest($cert)
    
    try {
        # Create the request
        $csr = $enrollment.CreateRequest(0x1) # 0x1 represents base64 encoding
        
        # Submit the request directly to the enterprise CA
        Write-Host "Submitting CSR to the Enterprise CA..." -ForegroundColor Yellow
        
        try {
            # Try to submit the request and install the response immediately
            $enrollment.InstallResponse(2, $csr, 1, $certificateTemplate) # 2 represents submit to CA directly
            
            Write-Host "Certificate successfully requested and installed in the Personal certificate store." -ForegroundColor Green
            Write-Host "You can find it in the Certificates MMC snap-in (certmgr.msc) under 'Personal > Certificates'."
        }
        catch [System.Runtime.InteropServices.COMException] {
            # Check if this is a pending request error
            if ($_.Exception.HResult -eq 0x80094012 -or $_.Exception.Message -match "pending") {
                Write-Host "Certificate request has been submitted but is pending approval." -ForegroundColor Yellow
                Write-Host "The request has been queued for processing by the CA administrator."
                
                # Save the request as a file for later reference
                $requestId = [guid]::NewGuid().ToString()
                $requestFilePath = "$env:TEMP\CSR_$requestId.req"
                $csr | Out-File -FilePath $requestFilePath -Encoding ascii
                
                Write-Host "The CSR has been saved to: $requestFilePath" -ForegroundColor Cyan
                Write-Host "Once approved, the certificate will be installed via auto-enrollment or can be manually installed."
            }
            else {
                # Re-throw other errors
                throw
            }
        }
        
        # Optionally export the certificate info for reference
        $certSubject = "CN=$($UserInfo.CommonName)"
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $certSubject } | Select-Object -First 1
        
        if ($cert) {
            Write-Host "`nCertificate Details:" -ForegroundColor Yellow
            Write-Host "  Subject: $($cert.Subject)"
            Write-Host "  Issuer:  $($cert.Issuer)"
            Write-Host "  Thumbprint: $($cert.Thumbprint)"
            Write-Host "  Valid from: $($cert.NotBefore) to $($cert.NotAfter)"
        }
    }
    catch {
        Write-Host "Error submitting CSR to the CA: $_" -ForegroundColor Red
        Write-Host "The CSR might have been created but not submitted successfully." -ForegroundColor Yellow
    }
}

# Main script execution
Show-Menu
$userInputData = Get-UserInput
Create-CSR -UserInfo $userInputData

Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")