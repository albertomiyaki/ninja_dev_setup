# CSR Request - GUI Tool with certreq
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Define script variables
$CA_Server = "-"  # Use "-" for default behavior or set to "YourCAServer\CAName"
$CertificateStoragePath = "C:\temp\tls\csr_gui\"

# System information variables
$hostname = ([System.Net.Dns]::GetHostEntry($env:computerName)).HostName
$shortHostname = $env:COMPUTERNAME
$defaultEmail = "user@domain.com"

function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Definition
    $argString = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    if ($MyInvocation.BoundParameters.Count -gt 0) {
        foreach ($key in $MyInvocation.BoundParameters.Keys) {
            $value = $MyInvocation.BoundParameters[$key]
            if ($value -is [System.String]) {
                $argString += " -$key `"$value`""
            }
            else {
                $argString += " -$key $value"
            }
        }
    }
    Start-Process -FilePath PowerShell.exe -ArgumentList "-NoProfile $argString" -Verb RunAs
    exit
}

$sync = [hashtable]::Synchronized(@{})

$availableTemplates = @(
    @{Name = "WebServer"; Description = "Web Server Certificate"},
    @{Name = "Workstation"; Description = "Workstation Authentication Certificate"},
    @{Name = "User"; Description = "User Authentication Certificate"},
    @{Name = "IPSECIntermediateOffline"; Description = "IPSec Certificate"},
    @{Name = "SmartcardLogon"; Description = "Smartcard Logon Certificate"},
    @{Name = "ClientAuth"; Description = "Client Authentication Certificate"}
)

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
    if ($sync.LogTextBox -ne $null) {
        $sync.LogTextBox.Dispatcher.Invoke([action]{
            $sync.LogTextBox.AppendText("$logMessage`r`n")
            $sync.LogTextBox.ScrollToEnd()
        })
    }
    switch ($Type) {
        "Information" { Write-Host $logMessage -ForegroundColor Gray }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }
}

[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CSR Request Tool" 
    Height="720" 
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
        
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Background" Value="#F0F0F0"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="CSR Request Tool" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Generate and submit Certificate Signing Requests using certreq" 
                               Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <TabControl Grid.Row="1" x:Name="MainTabControl" Margin="15">
            <TabItem Header="Create CSR">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel>
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
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="CommonNameTextBox" 
                                         ToolTip="The fully qualified domain name for the certificate"/>
                                
                                <Label Grid.Row="1" Grid.Column="0" Content="Friendly Name:"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="FriendlyNameTextBox" 
                                         ToolTip="A descriptive name for this certificate"/>
                                
                                <Label Grid.Row="2" Grid.Column="0" Content="Certificate Template:"/>
                                <ComboBox Grid.Row="2" Grid.Column="1" x:Name="TemplateComboBox" 
                                          ToolTip="Select the certificate template to use"/>
                                
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
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="EmailTextBox" 
                                         ToolTip="Your email address"/>
                                
                                <Label Grid.Row="1" Grid.Column="0" Content="Organization:"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="OrganizationTextBox" 
                                         ToolTip="Your organization name"/>
                                
                                <Label Grid.Row="2" Grid.Column="0" Content="Org. Unit:"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="OrgUnitTextBox" 
                                         ToolTip="Your organizational unit or department"/>
                                
                                <Label Grid.Row="3" Grid.Column="0" Content="City/Locality:"/>
                                <TextBox Grid.Row="3" Grid.Column="1" x:Name="CityTextBox" 
                                         ToolTip="Your city or locality"/>
                                
                                <Label Grid.Row="4" Grid.Column="0" Content="State/Province:"/>
                                <TextBox Grid.Row="4" Grid.Column="1" x:Name="StateTextBox" 
                                         ToolTip="Your state or province"/>
                                
                                <Label Grid.Row="5" Grid.Column="0" Content="Country:"/>
                                <TextBox Grid.Row="5" Grid.Column="1" x:Name="CountryTextBox" 
                                         ToolTip="Two-letter country code (e.g. US, UK, CA)"/>
                            </Grid>
                        </GroupBox>
                        
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
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,10">
                            <Button x:Name="RefreshPreviewButton" Content="Refresh Preview" 
                                    Style="{StaticResource DefaultButton}" Width="150"/>
                            <Button x:Name="CreateCSRButton" Content="Create CSR" 
                                    Style="{StaticResource DefaultButton}" Width="150" Margin="15,5,5,5"/>
                            <Button x:Name="SubmitCSRButton" Content="Submit CSR" 
                                    Style="{StaticResource DefaultButton}" Width="150" Margin="5,5,5,5"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Retrieve Certificate">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" 
                               Text="Retrieve and install your signed certificate" 
                               FontSize="16" 
                               FontWeight="Bold"
                               Margin="0,0,0,15"/>
                    
                    <GroupBox Grid.Row="1" Header="Certificate Request Details" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <Label Grid.Row="0" Grid.Column="0" Content="Request ID:"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="RequestIdTextBox" 
                                     ToolTip="Enter the Request ID from your certificate request"/>
                            <Button Grid.Row="0" Grid.Column="2" x:Name="CheckStatusButton" 
                                    Content="Check Status" Style="{StaticResource DefaultButton}" 
                                    Width="110" Margin="10,3,0,10"/>
                            
                            <Label Grid.Row="1" Grid.Column="0" Content="Certificate File:"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="CertFileTextBox" 
                                     ToolTip="Path to a certificate file (.cer, .crt, .p7b)"/>
                            <Button Grid.Row="1" Grid.Column="2" x:Name="BrowseCertButton" 
                                    Content="Browse" Style="{StaticResource DefaultButton}" 
                                    Width="110" Margin="10,3,0,10"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="2" Header="Certificate Status" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <Label Grid.Row="0" Grid.Column="0" Content="Status:"/>
                            <TextBlock Grid.Row="0" Grid.Column="1" x:Name="CertStatusTextBlock" 
                                       Text="Not checked yet" Margin="5"/>
                            
                            <Label Grid.Row="1" Grid.Column="0" Content="Disposition:"/>
                            <TextBlock Grid.Row="1" Grid.Column="1" x:Name="DispositionTextBlock" 
                                       Text="-" Margin="5"/>
                            
                            <Label Grid.Row="2" Grid.Column="0" Content="Last Checked:"/>
                            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="LastCheckedTextBlock" 
                                       Text="-" Margin="5"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="3" Header="Certificate Actions" Padding="10" Margin="0,0,0,15">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="RetrieveButton" Content="Retrieve Certificate" 
                                    Style="{StaticResource DefaultButton}" Width="160" Margin="5,5,15,5"/>
                            <Button x:Name="InstallButton" Content="Install Certificate" 
                                    Style="{StaticResource DefaultButton}" Width="160" Margin="5"/>
                            <Button x:Name="ExportButton" Content="Export with Private Key" 
                                    Style="{StaticResource DefaultButton}" Width="160" Margin="15,5,5,5"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="4" Header="Certificate Details" Padding="10" Margin="0,0,0,15">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBlock x:Name="CertDetailsTextBlock" FontFamily="Consolas" 
                                       TextWrapping="Wrap" Padding="5"/>
                        </ScrollViewer>
                    </GroupBox>
                </Grid>
            </TabItem>
        </TabControl>
        
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
        
        <Border Grid.Row="3" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
            <TextBlock x:Name="StatusBar" Text="Ready" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $_", "XAML Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

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
$sync.SubmitCSRButton = $window.FindName("SubmitCSRButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.MainTabControl = $window.FindName("MainTabControl")

$sync.RequestIdTextBox = $window.FindName("RequestIdTextBox")
$sync.CertFileTextBox = $window.FindName("CertFileTextBox")
$sync.CheckStatusButton = $window.FindName("CheckStatusButton")
$sync.BrowseCertButton = $window.FindName("BrowseCertButton")
$sync.CertStatusTextBlock = $window.FindName("CertStatusTextBlock")
$sync.DispositionTextBlock = $window.FindName("DispositionTextBlock")
$sync.LastCheckedTextBlock = $window.FindName("LastCheckedTextBlock")
$sync.RetrieveButton = $window.FindName("RetrieveButton")
$sync.InstallButton = $window.FindName("InstallButton")
$sync.ExportButton = $window.FindName("ExportButton")
$sync.CertDetailsTextBlock = $window.FindName("CertDetailsTextBlock")

$sync.CommonNameTextBox.Text = $hostname
$sync.FriendlyNameTextBox.Text = "$shortHostname Certificate"
$sync.EmailTextBox.Text = $defaultEmail
$sync.CountryTextBox.Text = "US"

$sync.SubmitCSRButton.IsEnabled = $false

foreach ($template in $availableTemplates) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $template.Description
    $item.Tag = $template.Name
    $sync.TemplateComboBox.Items.Add($item)
}
if ($sync.TemplateComboBox.Items.Count -gt 0) {
    $sync.TemplateComboBox.SelectedIndex = 0
}

function Update-Preview {
    $commonName = $sync.CommonNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = $hostname }
    
    $selectedUsageItem = $sync.CertUsageComboBox.SelectedItem
    $certUsage = $selectedUsageItem.Tag
    
    $selectedTemplateItem = $sync.TemplateComboBox.SelectedItem
    $templateName = $selectedTemplateItem.Tag
    $templateDescription = $selectedTemplateItem.Content
    
    $sanList = @($commonName, $hostname, $shortHostname)
    if (-not [string]::IsNullOrWhiteSpace($sync.AdditionalSANsTextBox.Text)) {
        $additionalSANArray = $sync.AdditionalSANsTextBox.Text -split ',' | ForEach-Object { $_.Trim() }
        $sanList += $additionalSANArray
    }
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

function Get-FormInput {
    $commonName = $sync.CommonNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($commonName)) { $commonName = $hostname }
    
    $friendlyName = $sync.FriendlyNameTextBox.Text
    if ([string]::IsNullOrWhiteSpace($friendlyName)) { $friendlyName = "$shortHostname Certificate" }
    
    $selectedUsageItem = $sync.CertUsageComboBox.SelectedItem
    $certUsage = $selectedUsageItem.Tag
    
    $selectedTemplateItem = $sync.TemplateComboBox.SelectedItem
    $templateName = $selectedTemplateItem.Tag
    
    $country = $sync.CountryTextBox.Text
    if ([string]::IsNullOrWhiteSpace($country)) { $country = "US" }
    
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

function Create-CSRConfigFile {
    param (
        [hashtable]$UserInfo
    )
    Write-ActivityLog "Creating CSR configuration file..." -Type Information
    $requestId = [Guid]::NewGuid().ToString()
    $sync.CurrentRequestId = $requestId
    
    # Ensure certificate storage directory exists
    if (-not (Test-Path $CertificateStoragePath)) {
        New-Item -Path $CertificateStoragePath -ItemType Directory | Out-Null
    }
    
    $configFile = "$CertificateStoragePath\CSR_Config_$requestId.inf"
    $csrFile = "$CertificateStoragePath\CSR_$requestId.req"
    $sync.CurrentConfigFile = $configFile
    $sync.CurrentCSRFile = $csrFile
    
    $sanList = @($UserInfo.CommonName, $hostname, $shortHostname)
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.AdditionalSANs)) {
        $additionalSANArray = $UserInfo.AdditionalSANs -split ',' | ForEach-Object { $_.Trim() }
        $sanList += $additionalSANArray
    }
    $sanList = $sanList | Select-Object -Unique | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    $subjectString = "CN=`"$($UserInfo.CommonName)`""
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.Organization)) {
        $subjectString += ", O=`"$($UserInfo.Organization)`""
    }
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.OrganizationalUnit)) {
        $subjectString += ", OU=`"$($UserInfo.OrganizationalUnit)`""
    }
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.City)) {
        $subjectString += ", L=`"$($UserInfo.City)`""
    }
    if (-not [string]::IsNullOrWhiteSpace($UserInfo.State)) {
        $subjectString += ", S=`"$($UserInfo.State)`""
    }
    $subjectString += ", C=$($UserInfo.Country)"
    
    $infContent = @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject="$subjectString"
KeySpec=1
KeyLength=4096
Exportable=TRUE
MachineKeySet=TRUE
SMIME=False
PrivateKeyArchive=FALSE
UserProtected=FALSE
UseExistingKeySet=FALSE
ProviderName="Microsoft Software Key Storage Provider"
ProviderType=12
RequestType=PKCS10
KeyUsage=0xA0
FriendlyName="$($UserInfo.FriendlyName)"

[RequestAttributes]
CertificateTemplate="$($UserInfo.CertificateTemplate)"
"@
    $infContent += "`r`n[EnhancedKeyUsageExtension]`r`n"
    switch ($UserInfo.CertUsage) {
        "serverAuth" {
            $infContent += "OID=1.3.6.1.5.5.7.3.1 ; Server Authentication`r`n"
        }
        "clientAuth" {
            $infContent += "OID=1.3.6.1.5.5.7.3.2 ; Client Authentication`r`n"
        }
        "both" {
            $infContent += "OID=1.3.6.1.5.5.7.3.1 ; Server Authentication`r`n"
            $infContent += "OID=1.3.6.1.5.5.7.3.2 ; Client Authentication`r`n"
        }
    }
    if ($sanList.Count -gt 0) {
        $infContent += "`r`n[Extensions]`r`n"
        $infContent += '2.5.29.17 = "{text}"' + "`r`n"
        foreach ($san in $sanList) {
            $infContent += '_continue_ = "dns=' + $san + '&"' + "`r`n"
        }
    }
    $infContent | Out-File -FilePath $configFile -Encoding ascii
    Write-ActivityLog "Configuration file created at: $configFile" -Type Success
    return $configFile
}

function Create-CSR {
    param (
        [hashtable]$UserInfo
    )
    try {
        $configFile = Create-CSRConfigFile -UserInfo $UserInfo
        Write-ActivityLog "Generating CSR using certreq..." -Type Information
        $csrFile = $sync.CurrentCSRFile
        $certreqResult = & certreq -new "$configFile" "$csrFile" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ActivityLog "CSR successfully created at: $csrFile" -Type Success
            $sync.StatusBar.Text = "CSR created successfully"
            $sync.SubmitCSRButton.IsEnabled = $true
            [System.Windows.MessageBox]::Show("CSR has been created successfully at: $csrFile`n`nYou can now submit the CSR to your Certificate Authority.", "CSR Created", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return $true
        }
        else {
            Write-ActivityLog "Error creating CSR: $certreqResult" -Type Error
            $sync.StatusBar.Text = "Error creating CSR"
            [System.Windows.MessageBox]::Show("Error creating CSR: $certreqResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
    }
    catch {
        Write-ActivityLog "Error creating CSR: $_" -Type Error
        $sync.StatusBar.Text = "Error creating CSR"
        [System.Windows.MessageBox]::Show("Error creating CSR: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

function Submit-CSR {
    try {
        if (-not (Test-Path $sync.CurrentCSRFile)) {
            Write-ActivityLog "No CSR file found. Please create a CSR first." -Type Error
            $sync.StatusBar.Text = "CSR file not found"
            [System.Windows.MessageBox]::Show("No CSR file found. Please create a CSR first.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
        Write-ActivityLog "Submitting CSR to the Certificate Authority..." -Type Information
        $sync.StatusBar.Text = "Submitting CSR..."
        
        $responseFile = $sync.CurrentCSRFile -replace "\.req$", ".rsp"
        $certFile = $sync.CurrentCSRFile -replace "\.req$", ".cer"
        $sync.CurrentResponseFile = $responseFile
        $sync.CurrentCertFile = $certFile
        
        $submitResult = & certreq -submit -config "$CA_Server" "$($sync.CurrentCSRFile)" "$responseFile" "$certFile" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $requestIdMatch = [regex]::Match($submitResult, "RequestId:\s*(\d+)")
            if ($requestIdMatch.Success) {
                $caRequestId = $requestIdMatch.Groups[1].Value
                $sync.CARequestId = $caRequestId
                Write-ActivityLog "CSR submitted successfully. CA Request ID: $caRequestId" -Type Success
                
                $requestInfoFile = $sync.CurrentCSRFile -replace "\.req$", ".info"
                @{
                    RequestId    = $caRequestId
                    ConfigFile   = $sync.CurrentConfigFile
                    CSRFile      = $sync.CurrentCSRFile
                    ResponseFile = $responseFile
                    CertFile     = $certFile
                    SubmittedAt  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                } | ConvertTo-Json | Out-File -FilePath $requestInfoFile -Encoding utf8
                
                $sync.StatusBar.Text = "CSR submitted (Request ID: $caRequestId)"
                [System.Windows.MessageBox]::Show("CSR has been submitted successfully to the CA.`n`nRequest ID: $caRequestId`n`nThis Request ID has been saved and can be used to retrieve the certificate later.", "CSR Submitted", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $sync.MainTabControl.SelectedIndex = 1
                return $true
            }
            else {
                Write-ActivityLog "CSR submitted but could not determine Request ID" -Type Warning
                $sync.StatusBar.Text = "CSR submitted (unknown Request ID)"
                [System.Windows.MessageBox]::Show("CSR has been submitted, but the Request ID could not be determined.`n`nOutput: $submitResult", "CSR Submitted", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return $true
            }
        }
        else {
            Write-ActivityLog "Error submitting CSR: $submitResult" -Type Error
            $sync.StatusBar.Text = "Error submitting CSR"
            [System.Windows.MessageBox]::Show("Error submitting CSR: $submitResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
    }
    catch {
        Write-ActivityLog "Error submitting CSR: $_" -Type Error
        $sync.StatusBar.Text = "Error submitting CSR"
        [System.Windows.MessageBox]::Show("Error submitting CSR: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

$sync.RefreshPreviewButton.Add_Click({ Update-Preview })
$sync.CreateCSRButton.Add_Click({
    $userInfo = Get-FormInput
    if (Create-CSR -UserInfo $userInfo) {
        Update-Preview
    }
})
$sync.SubmitCSRButton.Add_Click({ Submit-CSR })

$window.ShowDialog() | Out-Null