# CSR Request - GUI Tool
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})

# Available certificate templates (customize this list as needed)
$availableTemplates = @(
    @{Name = "WebServer"; Description = "Web Server Certificate"},
    @{Name = "Workstation"; Description = "Workstation Authentication Certificate"},
    @{Name = "User"; Description = "User Authentication Certificate"},
    @{Name = "IPSECIntermediateOffline"; Description = "IPSec Certificate"},
    @{Name = "SmartcardLogon"; Description = "Smartcard Logon Certificate"},
    @{Name = "ClientAuth"; Description = "Client Authentication Certificate"}
)

# Default template will be set later based on UI selection

# Get hostname information for default values
$hostname = [System.Net.Dns]::GetHostByName(($env:computerName)).HostName
$shortHostname = $env:COMPUTERNAME
$defaultEmail = "user@domain.com" # Replace with your default domain

# Create log function for centralized logging
function Write-ActivityLog {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Information", "Warning", "Error", "Success")]
        [string]$Type = "Information"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # Add to the log text box if UI is loaded
    if ($sync.LogTextBox -ne $null) {
        $sync.LogTextBox.Dispatcher.Invoke([action]{
            $sync.LogTextBox.AppendText("$logMessage`r`n")
            $sync.LogTextBox.ScrollToEnd()
        })
    }
    
    # Also write to console for debugging purposes
    switch ($Type) {
        "Information" { Write-Host $logMessage -ForegroundColor Gray }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }
}

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CSR Request Tool" 
    Height="680" 
    Width="700"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="DefaultButton">
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#003C6A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,3,0,10"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
        </Style>
        
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,3,0,10"/>
        </Style>
        
        <Style TargetType="Label">
            <Setter Property="Margin" Value="0,5,0,2"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="CSR Request Tool" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Generate and submit Certificate Signing Requests" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="15">
            <StackPanel>
                <!-- Certificate Information -->
                <GroupBox Header="Certificate Information" Padding="10" Margin="0,0,0,15">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Label Grid.Row="0" Grid.Column="0" Content="Common Name:"/>
                        <TextBox Grid.Row="0" Grid.Column="1" x:Name="CommonNameTextBox" ToolTip="The fully qualified domain name for the certificate"/>
                        
                        <Label Grid.Row="1" Grid.Column="0" Content="Friendly Name:"/>
                        <TextBox Grid.Row="1" Grid.Column="1" x:Name="FriendlyNameTextBox" ToolTip="A descriptive name for this certificate"/>
                        
                        <Label Grid.Row="2" Grid.Column="0" Content="Certificate Template:"/>
                        <ComboBox Grid.Row="2" Grid.Column="1" x:Name="TemplateComboBox" ToolTip="Select the certificate template to use">
                            <!-- Templates will be added programmatically -->
                        </ComboBox>
                        
                        <Label Grid.Row="3" Grid.Column="0" Content="Certificate Usage:"/>
                        <ComboBox Grid.Row="3" Grid.Column="1" x:Name="CertUsageComboBox">
                            <ComboBoxItem Content="Client Authentication" Tag="clientAuth"/>
                            <ComboBoxItem Content="Server Authentication" Tag="serverAuth"/>
                            <ComboBoxItem Content="Both Client and Server" Tag="both" IsSelected="True"/>
                        </ComboBox>
                        
                        <Label Grid.Row="4" Grid.Column="0" Content="Additional SANs:"/>
                        <TextBox Grid.Row="4" Grid.Column="1" x:Name="AdditionalSANsTextBox" 
                                 ToolTip="Comma-separated list of additional Subject Alternative Names"/>
                    </Grid>
                </GroupBox>
                
                <!-- Organization Information -->
                <GroupBox Header="Organization Information" Padding="10" Margin="0,0,0,15">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Label Grid.Row="0" Grid.Column="0" Content="Email Address:"/>
                        <TextBox Grid.Row="0" Grid.Column="1" x:Name="EmailTextBox" ToolTip="Your email address"/>
                        
                        <Label Grid.Row="1" Grid.Column="0" Content="Organization:"/>
                        <TextBox Grid.Row="1" Grid.Column="1" x:Name="OrganizationTextBox" ToolTip="Your organization name"/>
                        
                        <Label Grid.Row="2" Grid.Column="0" Content="Org. Unit:"/>
                        <TextBox Grid.Row="2" Grid.Column="1" x:Name="OrgUnitTextBox" ToolTip="Your organizational unit or department"/>
                        
                        <Label Grid.Row="3" Grid.Column="0" Content="City/Locality:"/>
                        <TextBox Grid.Row="3" Grid.Column="1" x:Name="CityTextBox" ToolTip="Your city or locality"/>
                        
                        <Label Grid.Row="4" Grid.Column="0" Content="State/Province:"/>
                        <TextBox Grid.Row="4" Grid.Column="1" x:Name="StateTextBox" ToolTip="Your state or province"/>
                        
                        <Label Grid.Row="5" Grid.Column="0" Content="Country:"/>
                        <TextBox Grid.Row="5" Grid.Column="1" x:Name="CountryTextBox" ToolTip="Two-letter country code (e.g. US, UK, CA)"/>
                    </Grid>
                </GroupBox>
                
                <!-- Preview -->
                <GroupBox Header="Preview" Padding="10" Margin="0,0,0,15">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Grid.Row="0" Text="The following information will be used to create your certificate:" Margin="0,0,0,5"/>
                        <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3" Background="#F8F8F8">
                            <TextBlock x:Name="PreviewTextBlock" FontFamily="Consolas" Padding="10" TextWrapping="Wrap"/>
                        </Border>
                    </Grid>
                </GroupBox>
                
                <!-- Actions -->
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,10">
                    <Button x:Name="RefreshPreviewButton" Content="Refresh Preview" 
                            Style="{StaticResource DefaultButton}" Width="150"/>
                    <Button x:Name="CreateCSRButton" Content="Create and Submit CSR" 
                            Style="{StaticResource DefaultButton}" Width="200" Margin="15,5,5,5"/>
                </StackPanel>
            </StackPanel>
        </ScrollViewer>
        
        <!-- Log Panel -->
        <Grid Grid.Row="2" Margin="15,0,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="100"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Activity Log" FontWeight="Bold" FontSize="14" Margin="0,0,0,5"/>
            <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBox x:Name="LogTextBox" IsReadOnly="True" TextWrapping="Wrap" 
                             FontFamily="Consolas" Background="#F8F8F8" BorderThickness="0"
                             Padding="10"/>
                </ScrollViewer>
            </Border>
        </Grid>
        
        <!-- Status Bar -->
        <Border Grid.Row="3" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
            <TextBlock x:Name="StatusBar" Text="Ready" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@

# Create a form object from the XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

# Store references in the sync hashtable
$sync.Window = $window
$sync.CommonNameTextBox = $window.FindName("CommonNameTextBox")
$sync.FriendlyNameTextBox = $window.FindName("FriendlyNameTextBox")
$sync.TemplateComboBox = $window.FindName("TemplateComboBox")
$sync.CertUsageComboBox = $window.FindName("CertUsageComboBox")
$sync.AdditionalSANsTextBox = $window.FindName("AdditionalSANsTextBox")
$sync.EmailTextBox = $window.FindName("EmailTextBox")
$sync.OrganizationTextBox = $window.FindName("OrganizationTextBox")
$sync.OrgUnitTextBox = $window.FindName("OrgUnitTextBox")
$sync.CityTextBox = $window.FindName("CityTextBox")
$sync.StateTextBox = $window.FindName("StateTextBox")
$sync.CountryTextBox = $window.FindName("CountryTextBox")
$sync.PreviewTextBlock = $window.FindName("PreviewTextBlock")
$sync.RefreshPreviewButton = $window.FindName("RefreshPreviewButton")
$sync.CreateCSRButton = $window.FindName("CreateCSRButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")

# Set default values
$sync.CommonNameTextBox.Text = $hostname
$sync.FriendlyNameTextBox.Text = "$shortHostname Certificate"
$sync.EmailTextBox.Text = $defaultEmail
$sync.CountryTextBox.Text = "US"

# Populate certificate templates dropdown
foreach ($template in $availableTemplates) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $template.Description
    $item.Tag = $template.Name
    $sync.TemplateComboBox.Items.Add($item)
}

# Select the first template by default
if ($sync.TemplateComboBox.Items.Count -gt 0) {
    $sync.TemplateComboBox.SelectedIndex = 0
}

# Function to generate a preview of the certificate details
function Update-Preview {
    $commonName = $sync.CommonNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = $hostname }
    
    $selectedItem = $sync.CertUsageComboBox.SelectedItem
    $certUsage = $selectedItem.Tag
    
    $selectedTemplateItem = $sync.TemplateComboBox.SelectedItem
    $templateName = $selectedTemplateItem.Tag
    $templateDescription = $selectedTemplateItem.Content
    
    $sanList = @($commonName, $hostname, $shortHostname)
    
    # Add additional SANs if provided
    if (-not [string]::IsNullOrWhiteSpace($sync.AdditionalSANsTextBox.Text)) {
        $additionalSANArray = $sync.AdditionalSANsTextBox.Text -split ',' | ForEach-Object { $_.Trim() }
        $sanList += $additionalSANArray
    }
    
    # Remove duplicates from SAN list
    $sanList = $sanList | Select-Object -Unique
    
    $preview = "Subject: CN=$commonName"
    
    if (-not [string]::IsNullOrWhiteSpace($sync.OrganizationTextBox.Text)) {
        $preview += ", O=$($sync.OrganizationTextBox.Text)"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($sync.OrgUnitTextBox.Text)) {
        $preview += ", OU=$($sync.OrgUnitTextBox.Text)"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($sync.CityTextBox.Text)) {
        $preview += ", L=$($sync.CityTextBox.Text)"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($sync.StateTextBox.Text)) {
        $preview += ", S=$($sync.StateTextBox.Text)"
    }
    
    $preview += ", C=$($sync.CountryTextBox.Text)"
    
    if (-not [string]::IsNullOrWhiteSpace($sync.EmailTextBox.Text)) {
        $preview += ", E=$($sync.EmailTextBox.Text)"
    }
    
    $preview += "`n`nFriendly Name: $($sync.FriendlyNameTextBox.Text)"
    $preview += "`n`nCertificate Usage: "
    
    switch($certUsage) {
        "clientAuth" { $preview += "Client Authentication" }
        "serverAuth" { $preview += "Server Authentication" }
        "both" { $preview += "Both Client and Server Authentication" }
    }
    
    $preview += "`n`nTemplate: $templateDescription ($templateName)"
    
    $preview += "`n`nSubject Alternative Names (SANs):"
    foreach ($san in $sanList) {
        $preview += "`n- $san"
    }
    
    $preview += "`n`nKey Size: 4096 bits"
    $preview += "`n`nKey Storage Provider: Microsoft Software Key Storage Provider"
    $preview += "`n`nPrivate Key: Exportable"
    
    $sync.PreviewTextBlock.Text = $preview
}

# Function to collect user input from the form
function Get-FormInput {
    $commonName = $sync.CommonNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = $hostname }
    
    $friendlyName = $sync.FriendlyNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($friendlyName)) { $friendlyName = "$shortHostname Certificate" }
    
    $selectedItem = $sync.CertUsageComboBox.SelectedItem
    $certUsage = $selectedItem.Tag
    
    $selectedTemplateItem = $sync.TemplateComboBox.SelectedItem
    $templateName = $selectedTemplateItem.Tag
    
    $country = $sync.CountryTextBox.Text
    if ([string]::IsNullOrWhiteSpace($country)) { $country = "US" }
    
    # Return all collected information as a hashtable
    return @{
        CommonName = $commonName
        FriendlyName = $friendlyName
        AdditionalSANs = $sync.AdditionalSANsTextBox.Text
        Email = $sync.EmailTextBox.Text
        Organization = $sync.OrganizationTextBox.Text
        OrganizationalUnit = $sync.OrgUnitTextBox.Text
        City = $sync.CityTextBox.Text
        State = $sync.StateTextBox.Text
        Country = $country
        CertUsage = $certUsage
        CertificateTemplate = $templateName
    }
}

# Function to create and submit the CSR
function Create-CSR {
    param (
        [hashtable]$UserInfo
    )
    
    Write-ActivityLog "Creating Certificate Signing Request..." -Type Information
    
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
    
    # Create Key Usage extension for digital signature and key encipherment
    $keyUsageExtension = New-Object -ComObject X509Enrollment.CX509ExtensionKeyUsage
    # 0x20 = keyEncipherment, 0x80 = digitalSignature, combined = 0xA0
    $keyUsageExtension.InitializeEncode(0xA0)
    
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
    
    # Create a private key using Microsoft Software Key Storage Provider
    $privateKey = New-Object -ComObject X509Enrollment.CX509PrivateKey
    $privateKey.ProviderName = "Microsoft Software Key Storage Provider"
    $privateKey.KeySpec = 1 # XCN_AT_KEYEXCHANGE
    $privateKey.Length = 4096
    $privateKey.MachineContext = $true
    $privateKey.ExportPolicy = 3 # 3 = Exportable and allows plaintext export
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
    
    # Add the Key Usage extension
    $cert.X509Extensions.Add($keyUsageExtension)
    
    # Create the enrollment request
    $enrollment = New-Object -ComObject X509Enrollment.CX509Enrollment
    $enrollment.InitializeFromRequest($cert)
    
    try {
        # Create the request
        $csr = $enrollment.CreateRequest(0x1) # 0x1 represents base64 encoding
        
        # Generate a unique ID for the request
        $requestId = [guid]::NewGuid().ToString()
        $requestFilePath = "$env:TEMP\CSR_$requestId.req"
        
        # Save the CSR to a file
        $csr | Out-File -FilePath $requestFilePath -Encoding ascii
        Write-ActivityLog "CSR saved to: $requestFilePath" -Type Information
        
        # Submit the request to the enterprise CA without waiting for immediate installation
        Write-ActivityLog "Submitting CSR to the Enterprise CA..." -Type Information
        
        try {
            Write-ActivityLog "Submitting CSR to the Enterprise CA..." -Type Information
            $enrollment.Enroll()  # This submits the request and registers it as pending
            Write-ActivityLog "Certificate request submitted. It is now pending approval." -Type Success
            
            [System.Windows.MessageBox]::Show(
                "Certificate request has been submitted to the CA and is awaiting approval.`n`nThe CSR has been saved to: $requestFilePath`n`nThe request is registered in the Certificate Enrollment system and will be automatically installed when approved if auto-enrollment is enabled.",
                "Request Submitted",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        catch {
            Write-ActivityLog "Error submitting CSR to the CA: $_" -Type Error
            [System.Windows.MessageBox]::Show(
                "Error submitting CSR to the CA: $_`n`nThe CSR has been saved to: $requestFilePath and can be submitted manually.",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
        
    }
    catch {
        Write-ActivityLog "Error creating CSR: $_" -Type Error
        $sync.StatusBar.Text = "Error creating CSR"
        
        [System.Windows.MessageBox]::Show(
            "Error creating CSR: $_",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Set up event handlers
$sync.RefreshPreviewButton.Add_Click({
    Update-Preview
    Write-ActivityLog "Preview refreshed" -Type Information
})

$sync.CreateCSRButton.Add_Click({
    try {
        $userInfo = Get-FormInput
        Create-CSR -UserInfo $userInfo
    }
    catch {
        Write-ActivityLog "Error: $_" -Type Error
        $sync.StatusBar.Text = "Error occurred"
        
        [System.Windows.MessageBox]::Show(
            "An error occurred: $_",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# Add change handlers to update the preview automatically
$updatePreviewControls = @(
    $sync.CommonNameTextBox,
    $sync.FriendlyNameTextBox,
    $sync.TemplateComboBox,
    $sync.CertUsageComboBox,
    $sync.AdditionalSANsTextBox,
    $sync.EmailTextBox,
    $sync.OrganizationTextBox,
    $sync.OrgUnitTextBox,
    $sync.CityTextBox,
    $sync.StateTextBox,
    $sync.CountryTextBox
)

foreach ($control in $updatePreviewControls) {
    if ($control -is [System.Windows.Controls.TextBox]) {
        $control.Add_TextChanged({
            Update-Preview
        })
    }
    elseif ($control -is [System.Windows.Controls.ComboBox]) {
        $control.Add_SelectionChanged({
            Update-Preview
        })
    }
}

# Initialize the preview
Update-Preview

# Log startup
Write-ActivityLog "CSR Request Tool started" -Type Information
$sync.StatusBar.Text = "Ready"

# Show the window
$sync.Window.ShowDialog() | Out-Null