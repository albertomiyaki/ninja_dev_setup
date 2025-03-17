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

function Check-CertificateStatus {
    try {
        $requestId = $sync.RequestIdTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($requestId)) {
            Write-ActivityLog "Please enter a valid Request ID" -Type Warning
            return
        }

        Write-ActivityLog "Checking certificate status for Request ID: $requestId" -Type Information
        $sync.StatusBar.Text = "Checking certificate status..."
        
        $checkResult = & certutil -config "$CA_Server" -getrequeststate $requestId 2>&1
        if ($LASTEXITCODE -eq 0) {
            $dispositionMatch = [regex]::Match($checkResult, "RequestStatus:\s*(\d+).*RequestStatusString:\s*(.+?)\r?\n")
            if ($dispositionMatch.Success) {
                $dispositionCode = $dispositionMatch.Groups[1].Value
                $dispositionText = $dispositionMatch.Groups[2].Value.Trim()
                
                $sync.CertStatusTextBlock.Text = $dispositionText
                $sync.DispositionTextBlock.Text = "Disposition: $dispositionCode"
                $sync.LastCheckedTextBlock.Text = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                
                Write-ActivityLog "Certificate status: $dispositionText (Code: $dispositionCode)" -Type Information
                
                switch ($dispositionCode) {
                    "0" { # Pending
                        $sync.RetrieveButton.IsEnabled = $false
                    }
                    "2" { # Denied
                        $sync.RetrieveButton.IsEnabled = $false
                    }
                    "3" { # Issued
                        $sync.RetrieveButton.IsEnabled = $true
                        Write-ActivityLog "Certificate is ready to be retrieved" -Type Success
                    }
                    "20" { # Certificate issued
                        $sync.RetrieveButton.IsEnabled = $true
                        Write-ActivityLog "Certificate is ready to be retrieved" -Type Success
                    }
                    default {
                        $sync.RetrieveButton.IsEnabled = $false
                    }
                }
            }
            else {
                Write-ActivityLog "Could not determine certificate status from response" -Type Warning
                $sync.CertStatusTextBlock.Text = "Unknown"
                $sync.DispositionTextBlock.Text = "Could not determine"
                $sync.LastCheckedTextBlock.Text = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        else {
            Write-ActivityLog "Error checking certificate status: $checkResult" -Type Error
            $sync.CertStatusTextBlock.Text = "Error"
            $sync.DispositionTextBlock.Text = "Error checking status"
            $sync.LastCheckedTextBlock.Text = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        $sync.StatusBar.Text = "Ready"
    }
    catch {
        Write-ActivityLog "Error checking certificate status: $_" -Type Error
        $sync.StatusBar.Text = "Error checking certificate status"
    }
}

function Browse-CertificateFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Certificate File"
    $openFileDialog.Filter = "Certificate Files (*.cer;*.crt;*.p7b)|*.cer;*.crt;*.p7b|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = $CertificateStoragePath
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $sync.CertFileTextBox.Text = $openFileDialog.FileName
        Get-CertificateDetails -CertPath $openFileDialog.FileName
        $sync.InstallButton.IsEnabled = $true
    }
}

function Retrieve-Certificate {
    try {
        $requestId = $sync.RequestIdTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($requestId)) {
            Write-ActivityLog "Please enter a valid Request ID" -Type Warning
            return
        }
        
        Write-ActivityLog "Retrieving certificate for Request ID: $requestId" -Type Information
        $sync.StatusBar.Text = "Retrieving certificate..."
        
        # Create certificate file path
        $certFileName = "Certificate_$requestId.cer"
        $certFilePath = Join-Path -Path $CertificateStoragePath -ChildPath $certFileName
        
        # Retrieve the certificate
        $retrieveResult = & certreq -retrieve -config "$CA_Server" $requestId "$certFilePath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ActivityLog "Certificate retrieved successfully: $certFilePath" -Type Success
            $sync.CertFileTextBox.Text = $certFilePath
            $sync.StatusBar.Text = "Certificate retrieved successfully"
            
            # Display certificate details
            Get-CertificateDetails -CertPath $certFilePath
            
            # Enable install button
            $sync.InstallButton.IsEnabled = $true
            
            [System.Windows.MessageBox]::Show("Certificate has been successfully retrieved and saved to:`n$certFilePath", "Certificate Retrieved", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        else {
            Write-ActivityLog "Error retrieving certificate: $retrieveResult" -Type Error
            $sync.StatusBar.Text = "Error retrieving certificate"
            [System.Windows.MessageBox]::Show("Error retrieving certificate: $retrieveResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    catch {
        Write-ActivityLog "Error retrieving certificate: $_" -Type Error
        $sync.StatusBar.Text = "Error retrieving certificate"
        [System.Windows.MessageBox]::Show("Error retrieving certificate: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Install-Certificate {
    try {
        $certPath = $sync.CertFileTextBox.Text.Trim()
        if (-not (Test-Path $certPath)) {
            Write-ActivityLog "Certificate file not found: $certPath" -Type Error
            [System.Windows.MessageBox]::Show("Certificate file not found: $certPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        Write-ActivityLog "Installing certificate from: $certPath" -Type Information
        $sync.StatusBar.Text = "Installing certificate..."
        
        # Install the certificate
        $installResult = & certreq -accept "$certPath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ActivityLog "Certificate installed successfully" -Type Success
            $sync.StatusBar.Text = "Certificate installed successfully"
            $sync.ExportButton.IsEnabled = $true
            [System.Windows.MessageBox]::Show("Certificate has been successfully installed in the certificate store.", "Certificate Installed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        else {
            Write-ActivityLog "Error installing certificate: $installResult" -Type Error
            $sync.StatusBar.Text = "Error installing certificate"
            [System.Windows.MessageBox]::Show("Error installing certificate: $installResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    catch {
        Write-ActivityLog "Error installing certificate: $_" -Type Error
        $sync.StatusBar.Text = "Error installing certificate"
        [System.Windows.MessageBox]::Show("Error installing certificate: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Get-CertificateDetails {
    param (
        [string]$CertPath
    )
    
    try {
        if (-not (Test-Path $CertPath)) {
            $sync.CertDetailsTextBlock.Text = "Certificate file not found"
            return
        }
        
        Write-ActivityLog "Getting certificate details from: $CertPath" -Type Information
        
        $certDetails = & certutil -verify -silent "$CertPath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            # Filter and format the certificate details
            $detailsText = $certDetails -join "`n"
            $sync.CertDetailsTextBlock.Text = $detailsText
        }
        else {
            $sync.CertDetailsTextBlock.Text = "Error retrieving certificate details: $certDetails"
            Write-ActivityLog "Error retrieving certificate details: $certDetails" -Type Error
        }
    }
    catch {
        $sync.CertDetailsTextBlock.Text = "Error retrieving certificate details: $_"
        Write-ActivityLog "Error retrieving certificate details: $_" -Type Error
    }
}

function Export-CertificateWithPrivateKey {
    try {
        # Get current certificate information from the file
        $certPath = $sync.CertFileTextBox.Text.Trim()
        if (-not (Test-Path $certPath)) {
            Write-ActivityLog "Certificate file not found: $certPath" -Type Error
            [System.Windows.MessageBox]::Show("Certificate file not found: $certPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        # Get certificate details to find the subject
        $certSubject = & certutil -dump "$certPath" | Select-String -Pattern "Subject:" | ForEach-Object { $_ -replace ".*Subject: ", "" }
        if (-not $certSubject) {
            Write-ActivityLog "Could not determine certificate subject" -Type Error
            [System.Windows.MessageBox]::Show("Could not determine certificate subject", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        # Create password dialog
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = "Enter PFX Password"
        $passwordForm.Size = New-Object System.Drawing.Size(350, 200)
        $passwordForm.StartPosition = "CenterScreen"
        $passwordForm.FormBorderStyle = "FixedDialog"
        $passwordForm.MaximizeBox = $false
        $passwordForm.MinimizeBox = $false
        
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "Enter password to protect the PFX file:"
        $passwordLabel.Size = New-Object System.Drawing.Size(300, 20)
        $passwordLabel.Location = New-Object System.Drawing.Point(10, 20)
        
        $passwordTextBox = New-Object System.Windows.Forms.MaskedTextBox
        $passwordTextBox.PasswordChar = '*'
        $passwordTextBox.Size = New-Object System.Drawing.Size(300, 20)
        $passwordTextBox.Location = New-Object System.Drawing.Point(10, 50)
        
        $confirmLabel = New-Object System.Windows.Forms.Label
        $confirmLabel.Text = "Confirm password:"
        $confirmLabel.Size = New-Object System.Drawing.Size(300, 20)
        $confirmLabel.Location = New-Object System.Drawing.Point(10, 80)
        
        $confirmTextBox = New-Object System.Windows.Forms.MaskedTextBox
        $confirmTextBox.PasswordChar = '*'
        $confirmTextBox.Size = New-Object System.Drawing.Size(300, 20)
        $confirmTextBox.Location = New-Object System.Drawing.Point(10, 110)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Location = New-Object System.Drawing.Point(75, 140)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.Location = New-Object System.Drawing.Point(165, 140)
        
        $passwordForm.Controls.Add($passwordLabel)
        $passwordForm.Controls.Add($passwordTextBox)
        $passwordForm.Controls.Add($confirmLabel)
        $passwordForm.Controls.Add($confirmTextBox)
        $passwordForm.Controls.Add($okButton)
        $passwordForm.Controls.Add($cancelButton)
        $passwordForm.AcceptButton = $okButton
        $passwordForm.CancelButton = $cancelButton
        
        # Show password dialog
        $result = $passwordForm.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
        
        # Validate passwords match
        if ($passwordTextBox.Text -ne $confirmTextBox.Text) {
            [System.Windows.MessageBox]::Show("Passwords do not match", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        # Create save file dialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Title = "Save PFX File"
        $saveFileDialog.Filter = "PFX Files (*.pfx)|*.pfx|All Files (*.*)|*.*"
        $saveFileDialog.InitialDirectory = $CertificateStoragePath
        $saveFileDialog.FileName = [System.IO.Path]::GetFileNameWithoutExtension($certPath) + ".pfx"
        
        if ($saveFileDialog.ShowDialog() -ne "OK") {
            return
        }
        
        $pfxPath = $saveFileDialog.FileName
        Write-ActivityLog "Exporting certificate with private key to: $pfxPath" -Type Information
        $sync.StatusBar.Text = "Exporting certificate with private key..."
        
        # Export the certificate with private key
        $exportCmd = @"
        # Find certificate by subject
        `$cert = Get-ChildItem -Path Cert:\LocalMachine\My -Recurse | Where-Object { `$_.Subject -like "*$certSubject*" } | Select-Object -First 1
        if (`$cert) {
            # Export to PFX
            Export-PfxCertificate -Cert `$cert -FilePath '$pfxPath' -Password (ConvertTo-SecureString -String '$($passwordTextBox.Text)' -Force -AsPlainText) -ChainOption BuildChain
            if (Test-Path '$pfxPath') {
                Write-Output "SUCCESS: Certificate exported to $pfxPath"
            } else {
                Write-Output "ERROR: Failed to create PFX file"
            }
        } else {
            Write-Output "ERROR: Certificate not found in store with subject: $certSubject"
        }
"@
        
        $exportResult = PowerShell -Command $exportCmd 2>&1
        
        if ($exportResult -like "*SUCCESS:*") {
            Write-ActivityLog "Certificate exported successfully: $pfxPath" -Type Success
            $sync.StatusBar.Text = "Certificate exported successfully"
            [System.Windows.MessageBox]::Show("Certificate with private key has been successfully exported to:`n$pfxPath", "Export Successful", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        else {
            Write-ActivityLog "Error exporting certificate: $exportResult" -Type Error
            $sync.StatusBar.Text = "Error exporting certificate"
            [System.Windows.MessageBox]::Show("Error exporting certificate: $exportResult", "Export Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    catch {
        Write-ActivityLog "Error exporting certificate: $_" -Type Error
        $sync.StatusBar.Text = "Error exporting certificate"
        [System.Windows.MessageBox]::Show("Error exporting certificate: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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

# Certificate retrieval tab button handlers
$sync.CheckStatusButton.Add_Click({ Check-CertificateStatus })
$sync.BrowseCertButton.Add_Click({ Browse-CertificateFile })
$sync.RetrieveButton.Add_Click({ Retrieve-Certificate })
$sync.InstallButton.Add_Click({ Install-Certificate })
$sync.ExportButton.Add_Click({ Export-CertificateWithPrivateKey })

# Initialize button states for the certificate tab
$sync.RetrieveButton.IsEnabled = $false
$sync.InstallButton.IsEnabled = $false
$sync.ExportButton.IsEnabled = $false

# Set up initial log message
Write-ActivityLog "CSR Request Tool started" -Type Information

$window.ShowDialog() | Out-Null