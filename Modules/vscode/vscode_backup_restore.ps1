# VSCode Backup & Restore Tool

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.VSCodeUserPath = "$env:APPDATA\Code\User"
$sync.VSCodeExtensionsPath = "$env:USERPROFILE\.vscode\extensions"
$sync.BackupItems = @{
    "Settings" = @{
        Path = "settings.json"
        Description = "User preferences and settings"
    }
    "Extensions" = @{
        Path = "extensions-list.json"
        Description = "Installed VS Code extensions"
        IsExtension = $true
    }
}

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
    Title="VSCode Backup &amp; Restore" 
    Height="900" 
    Width="700"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="DefaultButton">
            <Setter Property="Background" Value="#007ACC"/>
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
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="15,5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,8,0,8"/>
            <Setter Property="FontSize" Value="13"/>
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
        <Border Grid.Row="0" Background="#252526" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="VSCode Backup &amp; Restore" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Backup and restore your Visual Studio Code settings and extensions" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="15" x:Name="MainTabControl">
            <!-- Backup Tab -->
            <TabItem Header="Backup">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Header row with text and refresh button -->
                    <Grid Grid.Row="0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Select what to backup:" 
                                FontSize="16" 
                                FontWeight="SemiBold" 
                                Margin="0,0,0,10" 
                                Grid.Column="0"/>
                        <Button x:Name="RefreshBackupButton" 
                                Content="Refresh" 
                                Style="{StaticResource DefaultButton}" 
                                Width="150" 
                                Margin="0,0,0,10" 
                                Grid.Column="1"/>
                    </Grid>
                    
                    <!-- Rest of the backup UI -->
                    <Border Grid.Row="1" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3" Padding="15" Margin="0,0,0,15">
                        <!-- (Existing grid and controls for backup options) -->
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <CheckBox Grid.Row="0" x:Name="BackupSettingsCheckBox" Content="Settings - User preferences and settings" IsChecked="True" Margin="0,0,0,10" />
                            <CheckBox Grid.Row="1" x:Name="BackupExtensionsCheckBox" Content="Extensions - Installed VS Code extensions" IsChecked="True" Margin="0,0,0,10" />
                            
                            <Border Grid.Row="2" BorderBrush="#E5E5E5" BorderThickness="1" CornerRadius="3" Padding="10" Margin="20,0,0,0">
                                <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                                    <StackPanel x:Name="ExtensionsListPanel">
                                        <TextBlock Text="Loading extensions..." Foreground="#666666" Margin="0,5,0,5"/>
                                    </StackPanel>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </Border>
                    
                    <Button Grid.Row="2" x:Name="BackupButton" Content="Create Backup" 
                            Style="{StaticResource DefaultButton}" 
                            Width="150" 
                            HorizontalAlignment="Right"/>
                </Grid>
            </TabItem>

            
            <!-- Restore Tab -->
            <TabItem Header="Restore">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Select backup file to restore from:" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    
                    <Grid Grid.Row="1" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="RestoreFilePathTextBox" IsReadOnly="True" Padding="8" BorderThickness="1" BorderBrush="#CCCCCC"/>
                        <Button Grid.Column="1" x:Name="BrowseRestoreButton" Content="Browse..." 
                                Style="{StaticResource DefaultButton}" 
                                Width="100"/>
                    </Grid>
                    
                    <Border Grid.Row="2" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="3" Padding="15" Margin="0,0,0,15">
                        <StackPanel x:Name="RestoreOptionsPanel">
                            <TextBlock Text="Select backup file to view available restore options" 
                                       Foreground="#666666" 
                                       HorizontalAlignment="Center" 
                                       VerticalAlignment="Center" 
                                       FontStyle="Italic"
                                       Margin="0,20,0,0"/>
                        </StackPanel>
                    </Border>
                    
                    <Button Grid.Row="3" x:Name="RestoreButton" Content="Restore Selected" 
                            Style="{StaticResource DefaultButton}" 
                            Width="150" HorizontalAlignment="Right" IsEnabled="False"/>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Log Panel -->
        <Grid Grid.Row="2" Margin="15,0,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="120"/>
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
$sync.MainTabControl = $window.FindName("MainTabControl")
$sync.BackupSettingsCheckBox = $window.FindName("BackupSettingsCheckBox")
$sync.BackupExtensionsCheckBox = $window.FindName("BackupExtensionsCheckBox")
$sync.ExtensionsListPanel = $window.FindName("ExtensionsListPanel")
$sync.BackupButton = $window.FindName("BackupButton")
$sync.RestoreFilePathTextBox = $window.FindName("RestoreFilePathTextBox")
$sync.BrowseRestoreButton = $window.FindName("BrowseRestoreButton")
$sync.RestoreOptionsPanel = $window.FindName("RestoreOptionsPanel")
$sync.RestoreButton = $window.FindName("RestoreButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.InstalledExtensions = @()
$sync.SelectedExtensions = @()

$sync.RefreshBackupButton = $window.FindName("RefreshBackupButton")


# Function to check if VSCode paths exist
function Test-VSCodePaths {
    if (!(Test-Path -Path $sync.VSCodeUserPath)) {
        Write-ActivityLog "VSCode user settings path not found: $($sync.VSCodeUserPath)" -Type Warning
        return $false
    }
    
    return $true
}

# Function to get installed extensions
function Get-VSCodeExtensions {
    try {
        $extensions = @()
        
        # Try command line method first
        try {
            $codeCommand = Get-Command code -ErrorAction SilentlyContinue
            if ($codeCommand) {
                # Get full extensions info using --show-versions
                $cliExtensions = Invoke-Expression "code --list-extensions --show-versions" | Where-Object { $_ -ne "" }
                foreach ($ext in $cliExtensions) {
                    $extParts = $ext -split '@'
                    $extId = $extParts[0]
                    $extVersion = $extParts[1]
                    
                    # Create display name from extension ID
                    $extNameParts = $extId -split '\.'
                    $publisherName = $extNameParts[0]
                    $extensionName = $extNameParts[-1]  # Get last part
                    $displayName = "$extensionName ($publisherName)"
                    
                    $extensions += [PSCustomObject]@{
                        Id = $extId
                        Name = $displayName
                        Version = $extVersion
                    }
                }
            }
        }
        catch {
            Write-ActivityLog "Error getting extensions from CLI: $_" -Type Warning
        }
        
        # If we couldn't get extensions from CLI, try the filesystem method
        if ($extensions.Count -eq 0 -and (Test-Path -Path $sync.VSCodeExtensionsPath)) {
            $extensionDirs = Get-ChildItem -Path $sync.VSCodeExtensionsPath -Directory
            
            foreach ($dir in $extensionDirs) {
                $packageJsonPath = Join-Path $dir.FullName "package.json"
                
                if (Test-Path $packageJsonPath) {
                    try {
                        $packageJson = Get-Content $packageJsonPath -Raw | ConvertFrom-Json
                        if ($packageJson.name) {
                            $displayName = if ($packageJson.displayName) { 
                                "$($packageJson.displayName) ($($packageJson.publisher))" 
                            } else { 
                                "$($packageJson.name -split '\.' | Select-Object -Last 1) ($($packageJson.publisher))" 
                            }
                            
                            $extensions += [PSCustomObject]@{
                                Id = $packageJson.name
                                Name = $displayName
                                Version = $packageJson.version
                            }
                        }
                    }
                    catch {
                        Write-ActivityLog "Error parsing extension metadata: $($dir.Name)" -Type Warning
                    }
                }
            }
        }
        
        return $extensions
    }
    catch {
        Write-ActivityLog "Error getting installed extensions: $_" -Type Error
        return @()
    }
}

# Function to populate extensions list
function Initialize-ExtensionsList {
    $sync.ExtensionsListPanel.Children.Clear()
    $sync.SelectedExtensions = @()
    
    if ($sync.InstalledExtensions.Count -eq 0) {
        $noExtText = New-Object System.Windows.Controls.TextBlock
        $noExtText.Text = "No extensions found"
        $noExtText.Foreground = "#666666"
        $noExtText.Margin = "0,5,0,5"
        $sync.ExtensionsListPanel.Children.Add($noExtText)
        return
    }
    
    # Add select all checkbox
    $selectAllCheckBox = New-Object System.Windows.Controls.CheckBox
    $selectAllCheckBox.Content = "Select All Extensions"
    $selectAllCheckBox.Margin = "0,5,0,10"
    $selectAllCheckBox.FontWeight = "Bold"
    $selectAllCheckBox.IsChecked = $true
    
    $selectAllCheckBox.Add_Checked({
        foreach ($child in $sync.ExtensionsListPanel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                $child.IsChecked = $true
            }
        }
    })
    
    $selectAllCheckBox.Add_Unchecked({
        foreach ($child in $sync.ExtensionsListPanel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                $child.IsChecked = $false
            }
        }
    })
    
    $sync.ExtensionsListPanel.Children.Add($selectAllCheckBox)
    
    # Add individual extension checkboxes
    foreach ($ext in $sync.InstalledExtensions) {
        $extCheckBox = New-Object System.Windows.Controls.CheckBox
        $extCheckBox.Content = $ext.Name
        $extCheckBox.Tag = $ext.Id
        $extCheckBox.Margin = "20,3,0,3"
        $extCheckBox.IsChecked = $true
        $extCheckBox.ToolTip = "$($ext.Id) v$($ext.Version)"
        
        # Add event handler
        $extCheckBox.Add_Checked({
            $ext = $this.Tag
            if ($sync.SelectedExtensions -notcontains $ext) {
                $sync.SelectedExtensions += $ext
            }
        })
        
        $extCheckBox.Add_Unchecked({
            $ext = $this.Tag
            $sync.SelectedExtensions = $sync.SelectedExtensions | Where-Object { $_ -ne $ext }
        })
        
        # Add to selected extensions list initially
        $sync.SelectedExtensions += $ext.Id
        
        $sync.ExtensionsListPanel.Children.Add($extCheckBox)
    }
}

# Function to create backup
function New-VSCodeBackup {
    param (
        [string]$BackupPath
    )
    
    # Check which items to backup
    $backupSettings = $sync.BackupSettingsCheckBox.IsChecked
    $backupExtensions = $sync.BackupExtensionsCheckBox.IsChecked
    
    if (-not $backupSettings -and -not $backupExtensions) {
        Write-ActivityLog "No items selected for backup" -Type Warning
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to backup",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    try {
        # Create a temp directory for backup files
        $tempDir = Join-Path $env:TEMP "VSCodeBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        Write-ActivityLog "Created temporary directory for backup: $tempDir" -Type Information
        
        # Backup settings if selected
        if ($backupSettings) {
            $settingsPath = Join-Path $sync.VSCodeUserPath "settings.json"
            $settingsDest = Join-Path $tempDir "Settings"
            
            if (Test-Path -Path $settingsPath) {
                Copy-Item -Path $settingsPath -Destination $settingsDest -Force
                Write-ActivityLog "Backed up settings.json" -Type Success
            } else {
                Write-ActivityLog "Settings file not found: $settingsPath" -Type Warning
            }
        }
        
        # Backup extensions if selected
        if ($backupExtensions) {
            if ($sync.SelectedExtensions.Count -gt 0) {
                $extensionsDir = Join-Path $tempDir "Extensions"
                New-Item -ItemType Directory -Path $extensionsDir -Force | Out-Null
                
                # Save extensions list as JSON with names
                $extensionsList = @{
                    Extensions = $sync.SelectedExtensions
                    Count = $sync.SelectedExtensions.Count
                    ExtensionNames = @{}
                }
                
                # Create a map of extension IDs to display names
                foreach ($ext in $sync.InstalledExtensions) {
                    if ($sync.SelectedExtensions -contains $ext.Id) {
                        $extensionsList.ExtensionNames[$ext.Id] = $ext.Name
                    }
                }
                
                $extensionsList | ConvertTo-Json | Out-File (Join-Path $extensionsDir "extensions-list.json") -Force
                
                # Also save as text file for easy install
                $sync.SelectedExtensions | Out-File (Join-Path $extensionsDir "extensions-list.txt") -Force
                
                Write-ActivityLog "Backed up list of $($sync.SelectedExtensions.Count) extensions" -Type Success
            } else {
                Write-ActivityLog "No extensions selected to backup" -Type Warning
            }
        }
        
        # Create a metadata file
        $metadata = @{
            CreatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            BackupSettings = $backupSettings
            BackupExtensions = ($backupExtensions -and ($sync.SelectedExtensions.Count -gt 0))
            VSCodeVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Code.exe" -ErrorAction SilentlyContinue).Version
        }
        $metadata | ConvertTo-Json | Out-File (Join-Path $tempDir "metadata.json") -Force
        
        # Create the ZIP file
        if (Test-Path $BackupPath) {
            Remove-Item -Path $BackupPath -Force
        }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $BackupPath)
        Write-ActivityLog "Backup saved to: $BackupPath" -Type Success
        
        # Clean up temp directory
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ActivityLog "Temporary files cleaned up" -Type Information
        
        $sync.StatusBar.Text = "Backup completed successfully"
        
        [System.Windows.MessageBox]::Show(
            "Backup completed successfully!`nFile saved to: $BackupPath",
            "Backup Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    } catch {
        Write-ActivityLog "Error creating backup: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while creating the backup:`n$_",
            "Backup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function to initialize restore options from backup file
function Initialize-RestoreOptions {
    param (
        [string]$BackupPath
    )
    
    try {
        $sync.RestoreOptionsPanel.Children.Clear()
        
        # Extract metadata from the zip file
        $tempDir = Join-Path $env:TEMP "VSCodeRestore_Temp"
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Extract files from the backup
        $zip = [System.IO.Compression.ZipFile]::OpenRead($BackupPath)
        $metadataEntry = $zip.Entries | Where-Object { $_.Name -eq "metadata.json" }
        
        if ($null -eq $metadataEntry) {
            Write-ActivityLog "Invalid backup file: Metadata not found" -Type Error
            $zip.Dispose()
            return
        }
        
        $metadataPath = Join-Path $tempDir "metadata.json"
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($metadataEntry, $metadataPath, $true)
        
        # Extract extensions list if it exists
        $extensionsEntry = $zip.Entries | Where-Object { $_.FullName -match "Extensions[/\\]extensions-list.json" }
        $extensionsListPath = $null
        if ($null -ne $extensionsEntry) {
            $extensionsListPath = Join-Path $tempDir "extensions-list.json"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($extensionsEntry, $extensionsListPath, $true)
        }
        
        $zip.Dispose()
        
        # Read metadata
        $metadata = Get-Content $metadataPath -Raw | ConvertFrom-Json
        
        # Create text block with backup info
        $infoBlock = New-Object System.Windows.Controls.TextBlock
        $infoBlock.Margin = "0,0,0,15"
        $infoBlock.TextWrapping = "Wrap"
        $infoBlock.Text = "Backup created on: $($metadata.CreatedOn)`nVSCode version: $($metadata.VSCodeVersion)"
        $sync.RestoreOptionsPanel.Children.Add($infoBlock)
        
        # Add separator
        $separator = New-Object System.Windows.Controls.Separator
        $separator.Margin = "0,0,0,15"
        $sync.RestoreOptionsPanel.Children.Add($separator)
        
        # Add options label
        $optionsLabel = New-Object System.Windows.Controls.TextBlock
        $optionsLabel.Text = "Select items to restore:"
        $optionsLabel.FontWeight = "SemiBold"
        $optionsLabel.FontSize = 14
        $optionsLabel.Margin = "0,0,0,10"
        $sync.RestoreOptionsPanel.Children.Add($optionsLabel)
        
        # Add checkboxes for available items
        if ($metadata.BackupSettings) {
            $settingsCheck = New-Object System.Windows.Controls.CheckBox
            $settingsCheck.Content = "Settings - User preferences and settings"
            $settingsCheck.Tag = "Settings"
            $settingsCheck.IsChecked = $true
            $settingsCheck.Margin = "0,5,0,15"
            $sync.RestoreOptionsPanel.Children.Add($settingsCheck)
        }
        
        if ($metadata.BackupExtensions -and $null -ne $extensionsListPath) {
            $extensionsMainCheck = New-Object System.Windows.Controls.CheckBox
            $extensionsMainCheck.Content = "Extensions - Installed VS Code extensions"
            $extensionsMainCheck.Tag = "Extensions"
            $extensionsMainCheck.IsChecked = $true
            $extensionsMainCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($extensionsMainCheck)
            
            # Create extensions list panel
            $extensionsListBorder = New-Object System.Windows.Controls.Border
            $extensionsListBorder.BorderBrush = "#E5E5E5"
            $extensionsListBorder.BorderThickness = "1"
            $extensionsListBorder.CornerRadius = "3"
            $extensionsListBorder.Padding = "10"
            $extensionsListBorder.Margin = "20,0,0,10"
            
            $extensionsScroll = New-Object System.Windows.Controls.ScrollViewer
            $extensionsScroll.MaxHeight = "150"
            $extensionsScroll.VerticalScrollBarVisibility = "Auto"
            
            $extensionsStackPanel = New-Object System.Windows.Controls.StackPanel
            
            # Read extensions list
            $extensionsList = Get-Content $extensionsListPath -Raw | ConvertFrom-Json
            $sync.BackupExtensions = $extensionsList.Extensions
            
            # Get extension names if available
            $extNameMap = @{}
            if ($extensionsList.ExtensionNames) {
                $extNameMap = $extensionsList.ExtensionNames
            } else {
                # Generate names from IDs
                foreach ($extId in $extensionsList.Extensions) {
                    $extNameParts = $extId -split '\.'
                    $publisherName = $extNameParts[0]
                    $extensionName = $extNameParts[-1]
                    $extNameMap[$extId] = "$extensionName ($publisherName)"
                }
            }
            
            # Add select all checkbox
            $selectAllCheckBox = New-Object System.Windows.Controls.CheckBox
            $selectAllCheckBox.Content = "Select All Extensions"
            $selectAllCheckBox.Margin = "0,5,0,10"
            $selectAllCheckBox.FontWeight = "Bold"
            $selectAllCheckBox.IsChecked = $true
            
            $selectAllCheckBox.Add_Checked({
                foreach ($child in $extensionsStackPanel.Children) {
                    if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                        $child.IsChecked = $true
                    }
                }
            })
            
            $selectAllCheckBox.Add_Unchecked({
                foreach ($child in $extensionsStackPanel.Children) {
                    if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                        $child.IsChecked = $false
                    }
                }
            })
            
            $extensionsStackPanel.Children.Add($selectAllCheckBox)
            
            # Add individual extension checkboxes
            foreach ($ext in $extensionsList.Extensions) {
                $extCheckBox = New-Object System.Windows.Controls.CheckBox
                $displayName = if ($extNameMap.PSObject.Properties.Name -contains $ext) { $extNameMap.$ext } else { $ext }
                $extCheckBox.Content = $displayName
                $extCheckBox.Tag = $ext
                $extCheckBox.Margin = "20,3,0,3"
                $extCheckBox.IsChecked = $true
                $extCheckBox.ToolTip = $ext
                
                $extensionsStackPanel.Children.Add($extCheckBox)
            }
            
            $extensionsScroll.Content = $extensionsStackPanel
            $extensionsListBorder.Child = $extensionsScroll
            $sync.RestoreOptionsPanel.Children.Add($extensionsListBorder)
            
            # Link main extension checkbox to enable/disable extension selection
            $extensionsMainCheck.Add_Checked({
                $extensionsListBorder.IsEnabled = $true
            })
            
            $extensionsMainCheck.Add_Unchecked({
                $extensionsListBorder.IsEnabled = $false
            })
        }
        
        # Enable restore button
        $sync.RestoreButton.IsEnabled = $true
        
        # Clean up
        Remove-Item -Path $tempDir -Recurse -Force
        
        Write-ActivityLog "Loaded restore options from backup file" -Type Success
    } catch {
        Write-ActivityLog "Error loading restore options: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while loading the backup file:`n$_",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function to restore from backup
function Restore-VSCodeSettings {
    param (
        [string]$BackupPath
    )
    
    # Get selected restore options
    $selectedOptions = @()
    foreach ($child in $sync.RestoreOptionsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked -eq $true) {
            $selectedOptions += $child.Tag
        }
    }
    
    if ($selectedOptions.Count -eq 0) {
        Write-ActivityLog "No restore options selected" -Type Warning
        [System.Windows.MessageBox]::Show(
            "Please select at least one item to restore",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }
    
    try {
        # Check if VSCode is running
        $vscodeProcess = Get-Process -Name "Code" -ErrorAction SilentlyContinue
        if ($vscodeProcess) {
            $result = [System.Windows.MessageBox]::Show(
                "Visual Studio Code is currently running. It must be closed before restoring settings.`n`nWould you like to close VS Code now?",
                "Close VS Code",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-ActivityLog "Closing Visual Studio Code" -Type Information
                $vscodeProcess | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 2
                
                $vscodeProcess = Get-Process -Name "Code" -ErrorAction SilentlyContinue
                if ($vscodeProcess) {
                    Write-ActivityLog "Forcing VS Code to close" -Type Warning
                    $vscodeProcess | Stop-Process -Force
                    Start-Sleep -Seconds 2
                }
            } else {
                Write-ActivityLog "Restore cancelled - VS Code must be closed first" -Type Warning
                return
            }
        }
        
        # Extract files from the backup
        $tempDir = Join-Path $env:TEMP "VSCodeRestore_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($BackupPath, $tempDir)
        Write-ActivityLog "Extracted backup to temporary location" -Type Information
        
        # Restore settings if selected
        if ($selectedOptions -contains "Settings") {
            $settingsSource = Join-Path $tempDir "Settings"
            $settingsDest = Join-Path $sync.VSCodeUserPath "settings.json"
            
            if (Test-Path $settingsSource) {
                # Create backup of current settings
                if (Test-Path $settingsDest) {
                    Copy-Item -Path $settingsDest -Destination "$settingsDest.bak" -Force
                    Write-ActivityLog "Created backup of current settings" -Type Information
                }
                
                # Copy new settings
                Copy-Item -Path $settingsSource -Destination $settingsDest -Force
                Write-ActivityLog "Restored settings.json" -Type Success
            } else {
                Write-ActivityLog "Settings not found in backup" -Type Warning
            }
        }
        
        # Restore extensions if selected
        if ($selectedOptions -contains "Extensions") {
            # Get selected extensions from checkboxes
            $selectedExtensions = @()
            
            foreach ($child in $sync.RestoreOptionsPanel.Children) {
                if ($child -is [System.Windows.Controls.Border]) {
                    $scrollViewer = $child.Child
                    if ($scrollViewer -is [System.Windows.Controls.ScrollViewer]) {
                        $stackPanel = $scrollViewer.Content
                        if ($stackPanel -is [System.Windows.Controls.StackPanel]) {
                            foreach ($item in $stackPanel.Children) {
                                if ($item -is [System.Windows.Controls.CheckBox] -and $item.IsChecked -eq $true -and $item.Content -ne "Select All Extensions") {
                                    $selectedExtensions += $item.Tag
                                }
                            }
                        }
                    }
                }
            }
            
            if ($selectedExtensions.Count -gt 0) {
                # Check if code command is available
                $codeCommand = Get-Command code -ErrorAction SilentlyContinue
                if ($codeCommand) {
                    $installedCount = 0
                    $failedCount = 0
                    
                    foreach ($ext in $selectedExtensions) {
                        Write-ActivityLog "Installing extension: $ext" -Type Information
                        try {
                            $process = Start-Process -FilePath "code" -ArgumentList "--install-extension $ext" -Wait -PassThru -NoNewWindow
                            if ($process.ExitCode -eq 0) {
                                Write-ActivityLog "Successfully installed extension: $ext" -Type Success
                                $installedCount++
                            } else {
                                Write-ActivityLog "Failed to install extension: $ext" -Type Error
                                $failedCount++
                            }
                        } catch {
                            Write-ActivityLog "Error installing extension $ext`: $_" -Type Error
                            $failedCount++
                        }
                    }
                    
                    Write-ActivityLog "Extensions installation complete: $installedCount installed, $failedCount failed" -Type Information
                } else {
                    # Save selected extensions to a file
                    $selectedExtensionsPath = "$env:USERPROFILE\Desktop\vscode-extensions-to-install.txt"
                    $selectedExtensions | Out-File -FilePath $selectedExtensionsPath -Force
                    
                    Write-ActivityLog "VS Code command not found. Extensions list saved to Desktop for manual installation." -Type Warning
                    [System.Windows.MessageBox]::Show(
                        "Could not find 'code' command to install extensions automatically.`n`nAn extensions list has been saved to your Desktop. You can install them manually using:`n`ncode --install-extension EXTENSION_ID",
                        "Extensions Installation",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                }
            } else {
                Write-ActivityLog "No extensions selected for installation" -Type Warning
            }
        }
        
        # Clean up temp directory
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ActivityLog "Temporary files cleaned up" -Type Information
        
        $sync.StatusBar.Text = "Restore completed successfully"
        
        [System.Windows.MessageBox]::Show(
            "Restore completed successfully!`n`nYou may need to restart Visual Studio Code for all changes to take effect.",
            "Restore Complete",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    } catch {
        Write-ActivityLog "Error restoring settings: $_" -Type Error
        [System.Windows.MessageBox]::Show(
            "An error occurred while restoring settings:`n$_",
            "Restore Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Set up event handlers
$sync.BackupButton.Add_Click({
    # Show save file dialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Zip Files (*.zip)|*.zip"
    $saveFileDialog.Title = "Save VSCode Settings Backup"
    $saveFileDialog.FileName = "VSCodeBackup_$(Get-Date -Format 'yyyyMMdd').zip"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.StatusBar.Text = "Creating backup..."
        Write-ActivityLog "Creating backup to $($saveFileDialog.FileName)" -Type Information
        New-VSCodeBackup -BackupPath $saveFileDialog.FileName
    }
})

$sync.BrowseRestoreButton.Add_Click({
    # Show open file dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Zip Files (*.zip)|*.zip"
    $openFileDialog.Title = "Select VSCode Settings Backup"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.RestoreFilePathTextBox.Text = $openFileDialog.FileName
        $sync.StatusBar.Text = "Loading backup file..."
        Write-ActivityLog "Loading backup file: $($openFileDialog.FileName)" -Type Information
        Initialize-RestoreOptions -BackupPath $openFileDialog.FileName
    }
})

$sync.RestoreButton.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "This will replace your current VSCode settings with the ones from the backup.`n`nDo you want to continue?",
        "Confirm Restore",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $sync.StatusBar.Text = "Restoring settings..."
        Write-ActivityLog "Restoring settings from $($sync.RestoreFilePathTextBox.Text)" -Type Information
        Restore-VSCodeSettings -BackupPath $sync.RestoreFilePathTextBox.Text
    }
})


# Event handler for BackupExtensionsCheckBox
$sync.BackupExtensionsCheckBox.Add_Checked({
    $sync.ExtensionsListPanel.IsEnabled = $true
})

$sync.BackupExtensionsCheckBox.Add_Unchecked({
    $sync.ExtensionsListPanel.IsEnabled = $false
})

if ($sync.RefreshBackupButton -ne $null) {
    $sync.RefreshBackupButton.Add_Click({
        Write-ActivityLog "Refreshing installed extensions and settings..." -Type Information
        # Re-read the installed extensions from VSCode
        $sync.InstalledExtensions = Get-VSCodeExtensions
        # Rebuild the extensions list checkboxes
        Initialize-ExtensionsList
    })
} else {
    Write-ActivityLog "RefreshBackupButton not found in XAML." -Type Error
}


# Initialize the application
if (Test-VSCodePaths) {
    Write-ActivityLog "VSCode paths found. Initializing..." -Type Information
    
    # Get and display installed extensions
    $sync.InstalledExtensions = Get-VSCodeExtensions
    Initialize-ExtensionsList
    Write-ActivityLog "Found $($sync.InstalledExtensions.Count) installed extensions" -Type Information
} else {
    Write-ActivityLog "VSCode installation not found or paths incorrect" -Type Warning
    [System.Windows.MessageBox]::Show(
        "VSCode installation not found or paths are incorrect. Some features may not work properly.",
        "Initialization Warning",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    )
}

# Start with welcome message
Write-ActivityLog "VSCode Backup & Restore Tool started." -Type Information
$sync.StatusBar.Text = "Ready"

# Show the window
$sync.Window.ShowDialog() | Out-Null