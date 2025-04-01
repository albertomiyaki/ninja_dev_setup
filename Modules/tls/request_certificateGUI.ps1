# CSR GUI Tool
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
# $defaultEmail = "firstname.lastname@example.com"
# $defaultEmail = Get-ADUser -Identity $env:USERNAME | Select-Object -ExpandProperty UserPrincipalName
$objUser = New-Object DirectoryServices.DirectorySearcher
$objUser.Filter = "(&(objectCategory=User)(sAMAccountName=$env:USERNAME))"
$defaultEmail = $objUser.FindOne().Properties.userprincipalname


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
    Title="Certificate Request Tool" 
    Height="920" 
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
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="Certificate Request Tool" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Generate and manage certificate requests" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Folder Selection Bar -->
        <Grid Grid.Row="1" Background="#f0f0f0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Label Grid.Column="0" Content="Working Directory:" VerticalAlignment="Center" Margin="10,5"/>
            <TextBox Grid.Column="1" x:Name="WorkingDirTextBox" Margin="5,5"/>
            <Button Grid.Column="2" x:Name="BrowseFolderButton" 
                    Content="Browse" Style="{StaticResource DefaultButton}" 
                    Width="80" Margin="5,5"/>
        </Grid>
        
        <!-- Main Content -->
        <TabControl Grid.Row="2" x:Name="MainTabControl" Margin="10">
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
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <Label Grid.Row="0" Grid.Column="0" Content="Email Address:"/>
                                <TextBox Grid.Row="0" Grid.Column="1" x:Name="EmailTextBox" 
                                         ToolTip="Your email address"/>
                                
                                <Label Grid.Row="1" Grid.Column="0" Content="Organization:"/>
                                <TextBox Grid.Row="1" Grid.Column="1" x:Name="OrganizationTextBox" 
                                         ToolTip="Your organization name"/>
                                
                                <Label Grid.Row="2" Grid.Column="0" Content="Organizational Unit:"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="OrgUnitTextBox" 
                                         ToolTip="Your organizational unit or department"/>
                                
                                <Label Grid.Row="3" Grid.Column="0" Content="City:"/>
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
                        
                        <!-- Add buttons at the bottom of the Create CSR tab -->
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,10,0,0">
                            <Button x:Name="GenerateCSRButton" Content="Generate CSR" 
                                    Style="{StaticResource DefaultButton}" Margin="0,0,10,0"/>
                            <Button x:Name="SendToCAButton" Content="Submit to CA" 
                                    Style="{StaticResource DefaultButton}"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Retrieve Certificate">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Request ID Input -->
                    <GroupBox Grid.Row="0" Header="Certificate Request Details" Padding="10" 
                             Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <Label Grid.Column="0" Content="Request ID:"/>
                            <TextBox Grid.Column="1" x:Name="RequestIdTextBox"/>
                            <Button Grid.Column="2" x:Name="RetrieveCertButton" 
                                    Content="Retrieve Certificate" Style="{StaticResource DefaultButton}" 
                                    Width="150" Margin="10,0,0,0"/>
                        </Grid>
                    </GroupBox>
                    
                    <!-- Status Display -->
                    <GroupBox Grid.Row="1" Header="Status" Padding="10" Margin="0,0,0,10">
                        <TextBlock x:Name="StatusTextBlock" TextWrapping="Wrap"/>
                    </GroupBox>
                </Grid>
            </TabItem>
            
            <TabItem Header="Export PKCS#12">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="120"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Grid.Row="0" Grid.ColumnSpan="2" TextWrapping="Wrap" Margin="0,0,0,20"
                               Text="Export certificate with private key from your personal store to PKCS#12 format."/>

                    <Label Grid.Row="1" Content="Export Path:"/>
                    <TextBox Grid.Row="1" Grid.Column="1" Name="ExportPathTextBox" Margin="0,0,0,10"/>

                    <Label Grid.Row="2" Content="Password:"/>
                    <PasswordBox Grid.Row="2" Grid.Column="1" Name="ExportPasswordTextBox" Margin="0,0,0,10"/>

                    <Label Grid.Row="3" Content="Confirm Password:"/>
                    <PasswordBox Grid.Row="3" Grid.Column="1" Name="ConfirmPasswordTextBox" Margin="0,0,0,10"/>

                    <Button Grid.Row="4" Grid.Column="1" Content="Export PKCS#12" 
                            Name="ExportButton" HorizontalAlignment="Left" Width="150" Margin="0,10,0,0"/>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Log Area -->
        <GroupBox Grid.Row="3" Header="Activity Log" Margin="10">
            <TextBox x:Name="LogTextBox" Height="100" IsReadOnly="True" 
                     VerticalScrollBarVisibility="Auto" FontFamily="Consolas"/>
        </GroupBox>
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
$sync.GenerateCSRButton = $window.FindName("GenerateCSRButton")
$sync.SendToCAButton = $window.FindName("SendToCAButton")
$sync.RetrieveCertButton = $window.FindName("RetrieveCertButton")
$sync.InstallButton = $window.FindName("InstallButton")
$sync.ExportButton = $window.FindName("ExportButton")
$sync.RefreshPreviewButton = $window.FindName("RefreshPreviewButton")
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
$sync.CertDetailsTextBlock = $window.FindName("CertDetailsTextBlock")

$sync.ExportPathTextBox = $window.FindName("ExportPathTextBox")
$sync.ExportPasswordTextBox = $window.FindName("ExportPasswordTextBox")
$sync.ConfirmPasswordTextBox = $window.FindName("ConfirmPasswordTextBox")

# Initialize form with default values
$sync.CommonNameTextBox.Text = $hostname
$sync.FriendlyNameTextBox.Text = "$shortHostname"
$sync.EmailTextBox.Text = $defaultEmail
$sync.OrganizationTextBox.Text = "YourCompany"
$sync.OrgUnitTextBox.Text = "YourOrg"
$sync.CityTextBox.Text = "YourCity"
$sync.StateTextBox.Text = "YourState"
$sync.CountryTextBox.Text = "YourCountry"
$sync.AdditionalSANsTextBox.Text = "$hostname,$shortHostname"

foreach ($template in $availableTemplates) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $template.Description
    $item.Tag = $template.Name
    $sync.TemplateComboBox.Items.Add($item)
}
if ($sync.TemplateComboBox.Items.Count -gt 0) {
    $sync.TemplateComboBox.SelectedIndex = 0
}

# Initialize working directory with CertificateStoragePath
$sync.WorkingDirTextBox = $window.FindName("WorkingDirTextBox")
$sync.BrowseFolderButton = $window.FindName("BrowseFolderButton")
$sync.WorkingDirTextBox.Text = $CertificateStoragePath

# Add function to handle working directory selection
function Select-WorkingDirectory {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Working Directory"
    $folderBrowser.SelectedPath = $sync.WorkingDirTextBox.Text
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $sync.WorkingDirTextBox.Text = $folderBrowser.SelectedPath
        Write-ActivityLog "Working directory changed to: $($folderBrowser.SelectedPath)" -Type Information
        
        # Create directory if it doesn't exist
        if (-not (Test-Path $folderBrowser.SelectedPath)) {
            New-Item -Path $folderBrowser.SelectedPath -ItemType Directory | Out-Null
            Write-ActivityLog "Created working directory" -Type Information
        }
    }
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
    
    # Use working directory instead of CertificateStoragePath
    $workingDir = $sync.WorkingDirTextBox.Text
    if (-not (Test-Path $workingDir)) {
        New-Item -Path $workingDir -ItemType Directory | Out-Null
    }
    
    $configFile = Join-Path $workingDir "CSR_Config_$requestId.inf"
    $csrFile = Join-Path $workingDir "CSR_$requestId.req"
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
            [System.Windows.MessageBox]::Show("CSR has been created successfully at: $csrFile`n`nYou can now submit the CSR to your Certificate Authority.", "CSR Created", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return $true
        }
        else {
            Write-ActivityLog "Error creating CSR: $certreqResult" -Type Error
            [System.Windows.MessageBox]::Show("Error creating CSR: $certreqResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
    }
    catch {
        Write-ActivityLog "Error creating CSR: $_" -Type Error
        [System.Windows.MessageBox]::Show("Error creating CSR: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

function Submit-CSR {
    if (-not $sync.CurrentCSRFile -or -not (Test-Path $sync.CurrentCSRFile)) {
        Write-ActivityLog "No CSR file selected. Please generate or select a CSR first." -Type Error
        return $false
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to submit this CSR to the Certificate Authority?",
        "Confirm Submission",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::No) {
        return $false
    }
    
    try {
        Write-ActivityLog "Submitting CSR to the Certificate Authority..." -Type Information
        
        $workingDir = $sync.WorkingDirTextBox.Text
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($sync.CurrentCSRFile)
        $responseFile = Join-Path $workingDir "$baseFileName.rsp"
        $certFile = Join-Path $workingDir "$baseFileName.cer"
        
        $submitResult = & certreq -submit -config "$CA_Server" "$($sync.CurrentCSRFile)" "$responseFile" "$certFile" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $requestIdMatch = [regex]::Match($submitResult, "RequestId:\s*(\d+)")
            if ($requestIdMatch.Success) {
                $caRequestId = $requestIdMatch.Groups[1].Value
                Write-ActivityLog "CSR submitted successfully. Request ID: $caRequestId" -Type Success
                
                # Update status and GUI
                $sync.RequestIdTextBox.Text = $caRequestId
                
                # Save request information
                $requestInfoFile = Join-Path $workingDir "$baseFileName.info"
                @{
                    RequestId = $caRequestId
                    ConfigFile = $configFile
                    CSRFile = $sync.CurrentCSRFile
                    ResponseFile = $responseFile
                    CertFile = $certFile
                    SubmittedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                } | ConvertTo-Json | Out-File -FilePath $requestInfoFile -Encoding utf8
                
                # Switch to retrieve tab and populate request ID
                $sync.MainTabControl.SelectedIndex = 1
                $sync.RequestIdTextBox.Text = $caRequestId
                
                [System.Windows.MessageBox]::Show(
                    "CSR has been submitted successfully.`n`nRequest ID: $caRequestId`n`nPlease note this ID for retrieving your certificate later.",
                    "CSR Submitted",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                return $true
            }
        }
        
        Write-ActivityLog "Error submitting CSR: $submitResult" -Type Error
        [System.Windows.MessageBox]::Show("Error submitting CSR: $submitResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
    catch {
        Write-ActivityLog "Error submitting CSR: $_" -Type Error
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
        
    }
    catch {
        Write-ActivityLog "Error checking certificate status: $_" -Type Error
    }
}

function Browse-CertificateFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Certificate File"
    $openFileDialog.Filter = "Certificate Files (*.cer;*.crt;*.p7b)|*.cer;*.crt;*.p7b|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = $sync.WorkingDirTextBox.Text
    
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
            return $null
        }
        
        Write-ActivityLog "Retrieving certificate for Request ID: $requestId" -Type Information
        
        # Use working directory for temporary certificate file
        $workingDir = $sync.WorkingDirTextBox.Text
        $tempCertFile = Join-Path -Path $workingDir -ChildPath "temp_$requestId.cer"
        
        # First check if we have info file for this request
        $infoFiles = Get-ChildItem -Path $workingDir -Filter "*.info"
        $matchingInfo = $null
        foreach ($file in $infoFiles) {
            $info = Get-Content $file.FullName | ConvertFrom-Json
            if ($info.RequestId -eq $requestId) {
                $tempCertFile = $info.CertFile
                Write-ActivityLog "Found existing request info, using saved certificate path" -Type Information
                break
            }
        }
        
        # Retrieve the certificate to temporary file
        $retrieveResult = & certreq -retrieve -config "$CA_Server" $requestId "$tempCertFile" 2>&1
        if ($LASTEXITCODE -eq 0) {
            # Read the certificate to get the subject
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempCertFile)
            $subject = $cert.Subject
            
            # Extract CN from subject
            $cnMatch = [regex]::Match($subject, "CN=([^,]+)")
            $commonName = if ($cnMatch.Success) { $cnMatch.Groups[1].Value.Trim('"') } else { "cert_$requestId" }
            
            # Create final certificate path with common name
            $finalCertFile = Join-Path -Path $workingDir -ChildPath "$commonName.cer"
            
            # If file already exists, add number suffix
            $counter = 1
            while (Test-Path $finalCertFile) {
                $finalCertFile = Join-Path -Path $workingDir -ChildPath "$commonName($counter).cer"
                $counter++
            }
            
            # Move the temporary file to final location
            Move-Item -Path $tempCertFile -Destination $finalCertFile -Force
            
            Write-ActivityLog "Certificate retrieved successfully: $finalCertFile" -Type Success
            $sync.CertFileTextBox.Text = $finalCertFile
            
            return $finalCertFile
        }
        else {
            Write-ActivityLog "Error retrieving certificate: $retrieveResult" -Type Error
            [System.Windows.MessageBox]::Show("Error retrieving certificate: $retrieveResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $null
        }
    }
    catch {
        Write-ActivityLog "Error retrieving certificate: $_" -Type Error
        return $null
    }
}

function Install-Certificate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CertPath
    )
    
    try {
        if (-not (Test-Path $CertPath)) {
            Write-ActivityLog "Certificate file not found: $CertPath" -Type Error
            [System.Windows.MessageBox]::Show("Certificate file not found: $CertPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
        
        Write-ActivityLog "Installing certificate from: $CertPath" -Type Information
                
        # Install the certificate
        $installResult = & certreq -accept -machine "$CertPath" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ActivityLog "Certificate installed successfully" -Type Success
            [System.Windows.MessageBox]::Show("Certificate has been successfully installed in the certificate store.", "Certificate Installed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            return $true
        }
        else {
            Write-ActivityLog "Error installing certificate: $installResult" -Type Error
            [System.Windows.MessageBox]::Show("Error installing certificate: $installResult", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
    }
    catch {
        Write-ActivityLog "Error installing certificate: $_" -Type Error
        [System.Windows.MessageBox]::Show("Error installing certificate: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
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

function Export-Pkcs12 {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CertificatePath,
        
        [Parameter(Mandatory=$true)]
        [string]$ExportPath,
        
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$Password,
        
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$ConfirmPassword
    )

    try {
        # Convert passwords to strings for comparison (will be cleared immediately)
        $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $plainPassword1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR1)
        $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ConfirmPassword)
        $plainPassword2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR2)

        # Clear BSTRs immediately
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR1)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR2)

        if ($plainPassword1 -ne $plainPassword2) {
            throw "Passwords do not match"
        }

        # Clear plaintext passwords
        $plainPassword1 = $null
        $plainPassword2 = $null
        [System.GC]::Collect()

        # Get certificate from file to match
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath)
        $certThumbprint = $cert.Thumbprint

        # Find matching certificate with private key
        $matchingCert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {
            $_.Thumbprint -eq $certThumbprint -and $_.HasPrivateKey
        } | Select-Object -First 1

        if (-not $matchingCert) {
            throw "No certificate with matching thumbprint and private key found in your personal store."
        }

        # Export the certificate
        Export-PfxCertificate -Cert $matchingCert -FilePath $ExportPath -Password $Password -CryptoAlgorithmOption AES256_SHA256
        return $true
    }
    catch {
        Write-ActivityLog "Error exporting certificate: $_" -Type Error
        [System.Windows.MessageBox]::Show("Error exporting certificate: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Find-ExistingCSR {
    $workingDir = $sync.WorkingDirTextBox.Text
    $csrFiles = Get-ChildItem -Path $workingDir -Filter "*.req" -ErrorAction SilentlyContinue
    
    if ($csrFiles.Count -gt 0) {
        $result = [System.Windows.MessageBox]::Show(
            "Found existing CSR file(s) in the working directory. Would you like to use an existing CSR?",
            "Existing CSR Found",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Title = "Select CSR File"
            $openFileDialog.Filter = "CSR Files (*.req)|*.req|All Files (*.*)|*.*"
            $openFileDialog.InitialDirectory = $workingDir
            
            if ($openFileDialog.ShowDialog() -eq "OK") {
                $sync.CurrentCSRFile = $openFileDialog.FileName
                return $true
            }
        }
    }
    return $false
}

function Generate-CSR {
    if (Find-ExistingCSR) {
        Write-ActivityLog "Using existing CSR file: $($sync.CurrentCSRFile)" -Type Information
        $sync.SendToCAButton.IsEnabled = $true
        return $true
    }
    
    $userInfo = Get-FormInput
    
    if (Create-CSR -UserInfo $userInfo) {
        $sync.SendToCAButton.IsEnabled = $true
        return $true
    }
    
    return $false
}

function Retrieve-CertificateManual {
    try {
        $requestId = $sync.RequestIdTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($requestId)) {
            Write-ActivityLog "Please enter a valid Request ID" -Type Warning
            return
        }

        # First check if we have info file for this request
        $infoFiles = Get-ChildItem -Path $sync.WorkingDirTextBox.Text -Filter "*.info"
        $matchingInfo = $null
        foreach ($file in $infoFiles) {
            $info = Get-Content $file.FullName | ConvertFrom-Json
            if ($info.RequestId -eq $requestId) {
                $matchingInfo = $info
                break
            }
        }

        # If we found matching info, use those paths
        if ($matchingInfo) {
            $certFile = $matchingInfo.CertFile
            Write-ActivityLog "Found existing request info, using saved certificate path" -Type Information
        }
        else {
            # Let user choose where to save the certificate
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Title = "Save Certificate File"
            $saveFileDialog.Filter = "Certificate Files (*.cer)|*.cer|All Files (*.*)|*.*"
            $saveFileDialog.InitialDirectory = $sync.WorkingDirTextBox.Text
            $saveFileDialog.FileName = "Certificate_$requestId.cer"
            
            if ($saveFileDialog.ShowDialog() -ne "OK") {
                return
            }
            $certFile = $saveFileDialog.FileName
        }

        # Retrieve the certificate
        Write-ActivityLog "Retrieving certificate for Request ID: $requestId" -Type Information
        $retrieveResult = & certreq -retrieve -config "$CA_Server" $requestId "$certFile" 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-ActivityLog "Certificate retrieved successfully: $certFile" -Type Success
            
            # Ask if user wants to install the certificate
            $installResult = [System.Windows.MessageBox]::Show(
                "Certificate has been retrieved successfully. Would you like to install it now?",
                "Install Certificate",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($installResult -eq [System.Windows.MessageBoxResult]::Yes) {
                Install-Certificate -CertPath $certFile
            }
        }
        else {
            Write-ActivityLog "Error retrieving certificate: $retrieveResult" -Type Error
            [System.Windows.MessageBox]::Show("Error retrieving certificate: $retrieveResult", "Error")
        }
    }
    catch {
        Write-ActivityLog "Error retrieving certificate: $_" -Type Error
    }
}

# Button click handlers
$sync.GenerateCSRButton.Add_Click({
    $userInfo = Get-FormInput
    if (Create-CSR -UserInfo $userInfo) {
        Update-Preview
        $sync.SendToCAButton.IsEnabled = $true
    }
})

$sync.SendToCAButton.Add_Click({ Submit-CSR })

$sync.RetrieveCertButton.Add_Click({
    $certPath = Retrieve-Certificate
    if ($certPath) {
        $result = [System.Windows.MessageBox]::Show(
            "Certificate retrieved successfully. Would you like to install it now?",
            "Install Certificate",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Install-Certificate -CertPath $certPath
        }
    }
})

$sync.InstallButton.Add_Click({
    if ($sync.CertFileTextBox.Text) {
        Install-Certificate -CertPath $sync.CertFileTextBox.Text
    }
})

# Add handler for preview refresh
$sync.RefreshPreviewButton.Add_Click({ Update-Preview })

# Add event handler for the Export button
$sync.ExportButton.Add_Click({
    $exportPath = $sync.ExportPathTextBox.Text
    
    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        [System.Windows.MessageBox]::Show("Please specify an export path.", "Error", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error)
        return
    }

    $password = $sync.ExportPasswordTextBox.SecurePassword
    $confirmPassword = $sync.ConfirmPasswordTextBox.SecurePassword
    $cerPath = "$workingDir\$($sync.CommonNameTextBox.Text).cer"

    if (-not (Test-Path $cerPath)) {
        [System.Windows.MessageBox]::Show("Certificate file not found. Please request a certificate first.", "Error", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Error)
        return
    }

    $result = Export-Pkcs12 -CertificatePath $cerPath -ExportPath $exportPath -Password $password -ConfirmPassword $confirmPassword

    if ($result) {
        [System.Windows.MessageBox]::Show("Certificate successfully exported to $exportPath", "Success", 
            [System.Windows.MessageBoxButton]::OK, 
            [System.Windows.MessageBoxImage]::Information)
    }
})

# Update export path when CN changes
if ($sync.CommonNameTextBox) {
    $sync.CommonNameTextBox.Add_TextChanged({
        if ($sync.ExportPathTextBox) {
            $sync.ExportPathTextBox.Text = "$workingDir\$($sync.CommonNameTextBox.Text).p12"
        }
    })
}

# Initialize button states
if ($sync.SendToCAButton) {
    $sync.SendToCAButton.IsEnabled = $false
}
if ($sync.RetrieveCertButton) {
    $sync.RetrieveCertButton.IsEnabled = $true
}

# Update the working directory initialization
if (-not (Test-Path $CertificateStoragePath)) {
    New-Item -Path $CertificateStoragePath -ItemType Directory -Force | Out-Null
}
$sync.WorkingDirTextBox.Text = $CertificateStoragePath

# Set up initial log message
Write-ActivityLog "CSR Request Tool started" -Type Information

# Add handler for Browse button
$sync.BrowseFolderButton.Add_Click({ Select-WorkingDirectory })

$window.ShowDialog() | Out-Null