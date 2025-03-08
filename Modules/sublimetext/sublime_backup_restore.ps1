# Sublime Text Backup & Restore Tool

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.SublimeDataPath = "$env:APPDATA\Sublime Text"
$sync.SublimeInstalledPackagesPath = "$env:APPDATA\Sublime Text\Installed Packages"
$sync.SublimePackagesPath = "$env:APPDATA\Sublime Text\Packages"
$sync.SublimeUserPath = "$env:APPDATA\Sublime Text\Packages\User"
$sync.BackupItems = @{
    "Settings" = @{
        Path = "Preferences.sublime-settings"
        Description = "User preferences and settings"
    }
    "KeyBindings" = @{
        Path = "Default (Windows).sublime-keymap"
        Description = "Custom key bindings"
    }
    "Packages" = @{
        Path = "Package Control.sublime-settings"
        Description = "Installed Sublime Text packages"
        IsPackage = $true
    }
    "Workspace" = @{
        Path = "*.sublime-workspace"
        Description = "Current workspace and open files"
    }
    "ProjectData" = @{
        Path = "*.sublime-project"
        Description = "Project configuration"
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
    Title="Sublime Text Backup &amp; Restore" 
    Height="900" 
    Width="700"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="DefaultButton">
            <Setter Property="Background" Value="#FF9800"/>
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
                                <Setter Property="Background" Value="#F57C00"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#E65100"/>
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
        <Border Grid.Row="0" Background="#272822" Padding="15">
            <Grid>
                <StackPanel>
                    <TextBlock Text="Sublime Text Backup &amp; Restore" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Backup and restore your Sublime Text settings, packages, and workspaces" Foreground="#CCCCCC" Margin="0,5,0,0"/>
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
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <CheckBox Grid.Row="0" x:Name="BackupSettingsCheckBox" Content="Settings - User preferences and settings" IsChecked="True" Margin="0,0,0,10" />
                            <CheckBox Grid.Row="1" x:Name="BackupKeymapCheckBox" Content="Key Bindings - Custom keyboard shortcuts" IsChecked="True" Margin="0,0,0,10" />
                            <CheckBox Grid.Row="2" x:Name="BackupPackagesCheckBox" Content="Packages - Installed Sublime Text packages" IsChecked="True" Margin="0,0,0,10" />
                            <CheckBox Grid.Row="3" x:Name="BackupWorkspaceCheckBox" Content="Workspace - Current workspace and open files" IsChecked="True" Margin="0,0,0,10" />
                            <CheckBox Grid.Row="4" x:Name="BackupProjectCheckBox" Content="Project Data - Project configuration" IsChecked="True" Margin="0,0,0,10" />
                            
                            <Border Grid.Row="5" BorderBrush="#E5E5E5" BorderThickness="1" CornerRadius="3" Padding="10" Margin="20,0,0,0">
                                <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                                    <StackPanel x:Name="PackagesListPanel">
                                        <TextBlock Text="Loading packages..." Foreground="#666666" Margin="0,5,0,5"/>
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
$sync.BackupKeymapCheckBox = $window.FindName("BackupKeymapCheckBox")
$sync.BackupPackagesCheckBox = $window.FindName("BackupPackagesCheckBox") 
$sync.BackupWorkspaceCheckBox = $window.FindName("BackupWorkspaceCheckBox")
$sync.BackupProjectCheckBox = $window.FindName("BackupProjectCheckBox")
$sync.PackagesListPanel = $window.FindName("PackagesListPanel")
$sync.BackupButton = $window.FindName("BackupButton")
$sync.RestoreFilePathTextBox = $window.FindName("RestoreFilePathTextBox")
$sync.BrowseRestoreButton = $window.FindName("BrowseRestoreButton")
$sync.RestoreOptionsPanel = $window.FindName("RestoreOptionsPanel")
$sync.RestoreButton = $window.FindName("RestoreButton")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.InstalledPackages = @()
$sync.SelectedPackages = @()

$sync.RefreshBackupButton = $window.FindName("RefreshBackupButton")

# Function to check if Sublime Text paths exist
function Test-SublimePaths {
    if (!(Test-Path -Path $sync.SublimeDataPath)) {
        Write-ActivityLog "Sublime Text data path not found: $($sync.SublimeDataPath)" -Type Warning
        return $false
    }
    
    return $true
}

# Function to get installed packages
function Get-SublimePackages {
    try {
        $packages = @()
        
        # Get packages from Package Control settings
        $packageControlPath = Join-Path $sync.SublimeUserPath "Package Control.sublime-settings"
        
        if (Test-Path $packageControlPath) {
            $packageControlSettings = Get-Content $packageControlPath -Raw | ConvertFrom-Json
            
            if ($packageControlSettings.installed_packages) {
                foreach ($package in $packageControlSettings.installed_packages) {
                    $packages += [PSCustomObject]@{
                        Id = $package
                        Name = $package -replace "([a-z])([A-Z])", '$1 $2' # Add spaces between words
                    }
                }
            }
        }
        
        # If no packages found in package control, try fallback to directory listing
        if ($packages.Count -eq 0 -and (Test-Path -Path $sync.SublimeInstalledPackagesPath)) {
            $packageDirs = Get-ChildItem -Path $sync.SublimeInstalledPackagesPath -Filter "*.sublime-package"
            
            foreach ($pkg in $packageDirs) {
                $packageName = $pkg.BaseName
                $packages += [PSCustomObject]@{
                    Id = $packageName
                    Name = $packageName -replace "([a-z])([A-Z])", '$1 $2' # Add spaces between words
                }
            }
            
            # Also check the Packages directory for unpacked packages
            if (Test-Path -Path $sync.SublimePackagesPath) {
                $unpacked = Get-ChildItem -Path $sync.SublimePackagesPath -Directory | 
                            Where-Object { $_.Name -ne "User" -and $_.Name -ne "Default" }
                            
                foreach ($pkg in $unpacked) {
                    if ($packages.Id -notcontains $pkg.Name) {
                        $packages += [PSCustomObject]@{
                            Id = $pkg.Name
                            Name = $pkg.Name -replace "([a-z])([A-Z])", '$1 $2'
                        }
                    }
                }
            }
        }
        
        return $packages
    }
    catch {
        Write-ActivityLog "Error getting installed packages: $_" -Type Error
        return @()
    }
}

# Function to populate packages list
function Initialize-PackagesList {
    $sync.PackagesListPanel.Children.Clear()
    $sync.SelectedPackages = @()
    
    if ($sync.InstalledPackages.Count -eq 0) {
        $noPackagesText = New-Object System.Windows.Controls.TextBlock
        $noPackagesText.Text = "No packages found"
        $noPackagesText.Foreground = "#666666"
        $noPackagesText.Margin = "0,5,0,5"
        $sync.PackagesListPanel.Children.Add($noPackagesText)
        return
    }
    
    # Add select all checkbox
    $selectAllCheckBox = New-Object System.Windows.Controls.CheckBox
    $selectAllCheckBox.Content = "Select All Packages"
    $selectAllCheckBox.Margin = "0,5,0,10"
    $selectAllCheckBox.FontWeight = "Bold"
    $selectAllCheckBox.IsChecked = $true
    
    $selectAllCheckBox.Add_Checked({
        foreach ($child in $sync.PackagesListPanel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                $child.IsChecked = $true
            }
        }
    })
    
    $selectAllCheckBox.Add_Unchecked({
        foreach ($child in $sync.PackagesListPanel.Children) {
            if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                $child.IsChecked = $false
            }
        }
    })
    
    $sync.PackagesListPanel.Children.Add($selectAllCheckBox)
    
    # Add individual package checkboxes
    foreach ($pkg in $sync.InstalledPackages) {
        $pkgCheckBox = New-Object System.Windows.Controls.CheckBox
        $pkgCheckBox.Content = $pkg.Name
        $pkgCheckBox.Tag = $pkg.Id
        $pkgCheckBox.Margin = "20,3,0,3"
        $pkgCheckBox.IsChecked = $true
        $pkgCheckBox.ToolTip = $pkg.Id
        
        # Add event handler
        $pkgCheckBox.Add_Checked({
            $pkg = $this.Tag
            if ($sync.SelectedPackages -notcontains $pkg) {
                $sync.SelectedPackages += $pkg
            }
        })
        
        $pkgCheckBox.Add_Unchecked({
            $pkg = $this.Tag
            $sync.SelectedPackages = $sync.SelectedPackages | Where-Object { $_ -ne $pkg }
        })
        
        # Add to selected packages list initially
        $sync.SelectedPackages += $pkg.Id
        
        $sync.PackagesListPanel.Children.Add($pkgCheckBox)
    }
}

# Function to create backup
function New-SublimeBackup {
    param (
        [string]$BackupPath
    )
    
    # Check which items to backup
    $backupSettings = $sync.BackupSettingsCheckBox.IsChecked
    $backupKeymap = $sync.BackupKeymapCheckBox.IsChecked
    $backupPackages = $sync.BackupPackagesCheckBox.IsChecked
    $backupWorkspace = $sync.BackupWorkspaceCheckBox.IsChecked
    $backupProject = $sync.BackupProjectCheckBox.IsChecked
    
    if (-not $backupSettings -and -not $backupKeymap -and -not $backupPackages -and -not $backupWorkspace -and -not $backupProject) {
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
        $tempDir = Join-Path $env:TEMP "SublimeBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        Write-ActivityLog "Created temporary directory for backup: $tempDir" -Type Information
        
        # Create User directory for settings
        $userDir = Join-Path $tempDir "User"
        New-Item -ItemType Directory -Path $userDir -Force | Out-Null
        
        # Backup settings if selected
        if ($backupSettings) {
            $settingsPath = Join-Path $sync.SublimeUserPath "Preferences.sublime-settings"
            
            if (Test-Path -Path $settingsPath) {
                Copy-Item -Path $settingsPath -Destination $userDir -Force
                Write-ActivityLog "Backed up Preferences.sublime-settings" -Type Success
            } else {
                Write-ActivityLog "Settings file not found: $settingsPath" -Type Warning
            }
        }
        
        # Backup keymap if selected
        if ($backupKeymap) {
            $keymapPath = Join-Path $sync.SublimeUserPath "Default (Windows).sublime-keymap"
            
            if (Test-Path -Path $keymapPath) {
                Copy-Item -Path $keymapPath -Destination $userDir -Force
                Write-ActivityLog "Backed up Default (Windows).sublime-keymap" -Type Success
            } else {
                Write-ActivityLog "Keymap file not found: $keymapPath" -Type Warning
            }
            
            # Also backup other platform keymaps if they exist
            $otherKeyMaps = @(
                "Default (OSX).sublime-keymap",
                "Default (Linux).sublime-keymap"
            )
            
            foreach ($keymap in $otherKeyMaps) {
                $keymapPath = Join-Path $sync.SublimeUserPath $keymap
                if (Test-Path -Path $keymapPath) {
                    Copy-Item -Path $keymapPath -Destination $userDir -Force
                    Write-ActivityLog "Backed up $keymap" -Type Success
                }
            }
        }
        
        # Backup packages if selected
        if ($backupPackages) {
            $packageControlPath = Join-Path $sync.SublimeUserPath "Package Control.sublime-settings"
            
            if (Test-Path -Path $packageControlPath) {
                # Read the original file
                $packageControlContent = Get-Content -Path $packageControlPath -Raw | ConvertFrom-Json
                
                # Filter packages based on selection
                if ($sync.SelectedPackages.Count -gt 0) {
                    $packageControlContent.installed_packages = @($sync.SelectedPackages)
                }
                
                # Save the modified file
                $packageControlContent | ConvertTo-Json -Depth 10 | Out-File (Join-Path $userDir "Package Control.sublime-settings") -Force
                
                Write-ActivityLog "Backed up list of $($sync.SelectedPackages.Count) packages" -Type Success
                
                # Also create a plain text list for reference
                $sync.SelectedPackages | Out-File (Join-Path $userDir "packages-list.txt") -Force
            } else {
                Write-ActivityLog "Package Control settings not found: $packageControlPath" -Type Warning
            }
        }
        
        # Backup workspace if selected
        if ($backupWorkspace) {
            $workspaceFiles = Get-ChildItem -Path $sync.SublimeUserPath -Filter "*.sublime-workspace"
            
            if ($workspaceFiles.Count -gt 0) {
                foreach ($workspace in $workspaceFiles) {
                    Copy-Item -Path $workspace.FullName -Destination $userDir -Force
                }
                Write-ActivityLog "Backed up $($workspaceFiles.Count) workspace files" -Type Success
            } else {
                Write-ActivityLog "No workspace files found" -Type Warning
            }
        }
        
        # Backup project data if selected
        if ($backupProject) {
            $projectFiles = Get-ChildItem -Path $sync.SublimeUserPath -Filter "*.sublime-project"
            
            if ($projectFiles.Count -gt 0) {
                foreach ($project in $projectFiles) {
                    Copy-Item -Path $project.FullName -Destination $userDir -Force
                }
                Write-ActivityLog "Backed up $($projectFiles.Count) project files" -Type Success
            } else {
                Write-ActivityLog "No project files found" -Type Warning
            }
        }
        
        # Backup additional important settings files that are commonly customized
        $additionalSettingsFiles = @(
            "Distraction Free.sublime-settings",
            "Package Control.last-run",
            "Package Control.merged-ca-bundle",
            "Package Control.user-ca-bundle",
            "Syntax Highlighting for Sass.sublime-settings",
            "SublimeLinter.sublime-settings",
            "ColorHighlighter.sublime-settings",
            "Side Bar.sublime-settings"
        )
        
        foreach ($file in $additionalSettingsFiles) {
            $filePath = Join-Path $sync.SublimeUserPath $file
            if (Test-Path -Path $filePath) {
                Copy-Item -Path $filePath -Destination $userDir -Force
                Write-ActivityLog "Backed up additional settings: $file" -Type Success
            }
        }
        
        # Create a metadata file
        $metadata = @{
            CreatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            BackupSettings = $backupSettings
            BackupKeymap = $backupKeymap
            BackupPackages = $backupPackages
            BackupWorkspace = $backupWorkspace
            BackupProject = $backupProject
            SublimeVersion = (Get-ItemProperty -Path "HKCU:\Software\Sublime Text" -ErrorAction SilentlyContinue).Version
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
        $tempDir = Join-Path $env:TEMP "SublimeRestore_Temp"
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
        
        # Extract package list if it exists
        $packageControlEntry = $zip.Entries | Where-Object { $_.FullName -match "User[/\\]Package Control.sublime-settings" }
        $packageListPath = $null
        if ($null -ne $packageControlEntry) {
            $packageListPath = Join-Path $tempDir "Package Control.sublime-settings"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($packageControlEntry, $packageListPath, $true)
        }
        
        $zip.Dispose()
        
        # Read metadata
        $metadata = Get-Content $metadataPath -Raw | ConvertFrom-Json
        
        # Create text block with backup info
        $infoBlock = New-Object System.Windows.Controls.TextBlock
        $infoBlock.Margin = "0,0,0,15"
        $infoBlock.TextWrapping = "Wrap"
        $infoBlock.Text = "Backup created on: $($metadata.CreatedOn)`nSublime Text version: $($metadata.SublimeVersion)"
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
            $settingsCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($settingsCheck)
        }
        
        if ($metadata.BackupKeymap) {
            $keymapCheck = New-Object System.Windows.Controls.CheckBox
            $keymapCheck.Content = "Key Bindings - Custom key mappings"
            $keymapCheck.Tag = "Keymap"
            $keymapCheck.IsChecked = $true
            $keymapCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($keymapCheck)
        }
        
        if ($metadata.BackupPackages -and $null -ne $packageListPath) {
            $packagesMainCheck = New-Object System.Windows.Controls.CheckBox
            $packagesMainCheck.Content = "Packages - Installed Sublime Text packages"
            $packagesMainCheck.Tag = "Packages"
            $packagesMainCheck.IsChecked = $true
            $packagesMainCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($packagesMainCheck)
            
            # Create packages list panel
            $packagesListBorder = New-Object System.Windows.Controls.Border
            $packagesListBorder.BorderBrush = "#E5E5E5"
            $packagesListBorder.BorderThickness = "1"
            $packagesListBorder.CornerRadius = "3"
            $packagesListBorder.Padding = "10"
            $packagesListBorder.Margin = "20,0,0,10"
            
            $packagesScroll = New-Object System.Windows.Controls.ScrollViewer
            $packagesScroll.MaxHeight = "150"
            $packagesScroll.VerticalScrollBarVisibility = "Auto"
            
            $packagesStackPanel = New-Object System.Windows.Controls.StackPanel
            
            # Read packages list
            $packagesList = Get-Content $packageListPath -Raw | ConvertFrom-Json
            $sync.BackupPackages = $packagesList.installed_packages
            
            # Add select all checkbox
            $selectAllCheckBox = New-Object System.Windows.Controls.CheckBox
            $selectAllCheckBox.Content = "Select All Packages"
            $selectAllCheckBox.Margin = "0,5,0,10"
            $selectAllCheckBox.FontWeight = "Bold"
            $selectAllCheckBox.IsChecked = $true
            
            $selectAllCheckBox.Add_Checked({
                foreach ($child in $packagesStackPanel.Children) {
                    if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                        $child.IsChecked = $true
                    }
                }
            })
            
            $selectAllCheckBox.Add_Unchecked({
                foreach ($child in $packagesStackPanel.Children) {
                    if ($child -is [System.Windows.Controls.CheckBox] -and $child -ne $this) {
                        $child.IsChecked = $false
                    }
                }
            })
            
            $packagesStackPanel.Children.Add($selectAllCheckBox)
            
            # Add individual package checkboxes
            foreach ($pkg in $packagesList.installed_packages) {
                $pkgCheckBox = New-Object System.Windows.Controls.CheckBox
                $displayName = $pkg -replace "([a-z])([A-Z])", '$1 $2'
                $pkgCheckBox.Content = $displayName
                $pkgCheckBox.Tag = $pkg
                $pkgCheckBox.Margin = "20,3,0,3"
                $pkgCheckBox.IsChecked = $true
                $pkgCheckBox.ToolTip = $pkg
                
                $packagesStackPanel.Children.Add($pkgCheckBox)
            }
            
            $packagesScroll.Content = $packagesStackPanel
            $packagesListBorder.Child = $packagesScroll
            $sync.RestoreOptionsPanel.Children.Add($packagesListBorder)
            
            # Link main package checkbox to enable/disable package selection
            $packagesMainCheck.Add_Checked({
                $packagesListBorder.IsEnabled = $true
            })
            
            $packagesMainCheck.Add_Unchecked({
                $packagesListBorder.IsEnabled = $false
            })
        }
        
        if ($metadata.BackupWorkspace) {
            $workspaceCheck = New-Object System.Windows.Controls.CheckBox
            $workspaceCheck.Content = "Workspace - Current workspace and open files"
            $workspaceCheck.Tag = "Workspace"
            $workspaceCheck.IsChecked = $true
            $workspaceCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($workspaceCheck)
        }
        
        if ($metadata.BackupProject) {
            $projectCheck = New-Object System.Windows.Controls.CheckBox
            $projectCheck.Content = "Project Data - Project configuration"
            $projectCheck.Tag = "Project"
            $projectCheck.IsChecked = $true
            $projectCheck.Margin = "0,5,0,5"
            $sync.RestoreOptionsPanel.Children.Add($projectCheck)
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
function Restore-SublimeSettings {
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
        # Check if Sublime Text is running
        $sublimeProcess = Get-Process -Name "sublime_text" -ErrorAction SilentlyContinue
        if ($sublimeProcess) {
            $result = [System.Windows.MessageBox]::Show(
                "Sublime Text is currently running. It must be closed before restoring settings.`n`nWould you like to close Sublime Text now?",
                "Close Sublime Text",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-ActivityLog "Closing Sublime Text" -Type Information
                $sublimeProcess | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep -Seconds 2
                
                $sublimeProcess = Get-Process -Name "sublime_text" -ErrorAction SilentlyContinue
                if ($sublimeProcess) {
                    Write-ActivityLog "Forcing Sublime Text to close" -Type Warning
                    $sublimeProcess | Stop-Process -Force
                    Start-Sleep -Seconds 2
                }
            } else {
                Write-ActivityLog "Restore cancelled - Sublime Text must be closed first" -Type Warning
                return
            }
        }
        
        # Extract files from the backup
        $tempDir = Join-Path $env:TEMP "SublimeRestore_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($BackupPath, $tempDir)
        Write-ActivityLog "Extracted backup to temporary location" -Type Information
        
        # Check if User directory exists in the backup
        $userSourceDir = Join-Path $tempDir "User"
        if (!(Test-Path $userSourceDir)) {
            Write-ActivityLog "User directory not found in backup" -Type Error
            [System.Windows.MessageBox]::Show(
                "Invalid backup file: User directory not found",
                "Restore Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            Remove-Item -Path $tempDir -Recurse -Force
            return
        }
        
        # Ensure user directory exists
        if (!(Test-Path $sync.SublimeUserPath)) {
            New-Item -ItemType Directory -Path $sync.SublimeUserPath -Force | Out-Null
            Write-ActivityLog "Created Sublime Text User directory" -Type Information
        }
        
        # Create backup of current settings
        $backupDir = Join-Path $env:TEMP "SublimeOriginal_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        
        if (Test-Path $sync.SublimeUserPath) {
            Copy-Item -Path "$($sync.SublimeUserPath)\*" -Destination $backupDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-ActivityLog "Created backup of current settings" -Type Information
        }
        
        # Restore settings if selected
        if ($selectedOptions -contains "Settings") {
            $settingsSource = Join-Path $userSourceDir "Preferences.sublime-settings"
            $settingsDest = Join-Path $sync.SublimeUserPath "Preferences.sublime-settings"
            
            if (Test-Path $settingsSource) {
                Copy-Item -Path $settingsSource -Destination $settingsDest -Force
                Write-ActivityLog "Restored Preferences.sublime-settings" -Type Success
            } else {
                Write-ActivityLog "Settings not found in backup" -Type Warning
            }
        }
        
        # Restore keymap if selected
        if ($selectedOptions -contains "Keymap") {
            $keymapFiles = @(
                "Default (Windows).sublime-keymap",
                "Default (OSX).sublime-keymap",
                "Default (Linux).sublime-keymap"
            )
            
            foreach ($keymap in $keymapFiles) {
                $keymapSource = Join-Path $userSourceDir $keymap
                if (Test-Path $keymapSource) {
                    Copy-Item -Path $keymapSource -Destination $sync.SublimeUserPath -Force
                    Write-ActivityLog "Restored $keymap" -Type Success
                }
            }
        }
        
        # Restore packages if selected
        if ($selectedOptions -contains "Packages") {
            # Get selected packages from checkboxes
            $selectedPackages = @()
            
            foreach ($child in $sync.RestoreOptionsPanel.Children) {
                if ($child -is [System.Windows.Controls.Border]) {
                    $scrollViewer = $child.Child
                    if ($scrollViewer -is [System.Windows.Controls.ScrollViewer]) {
                        $stackPanel = $scrollViewer.Content
                        if ($stackPanel -is [System.Windows.Controls.StackPanel]) {
                            foreach ($item in $stackPanel.Children) {
                                if ($item -is [System.Windows.Controls.CheckBox] -and $item.IsChecked -eq $true -and $item.Content -ne "Select All Packages") {
                                    $selectedPackages += $item.Tag
                                }
                            }
                        }
                    }
                }
            }
            
            if ($selectedPackages.Count -gt 0) {
                # Restore Package Control settings with selected packages
                $packageControlSource = Join-Path $userSourceDir "Package Control.sublime-settings"
                $packageControlDest = Join-Path $sync.SublimeUserPath "Package Control.sublime-settings"
                
                if (Test-Path $packageControlSource) {
                    # Read package control settings
                    $packageControlSettings = Get-Content -Path $packageControlSource -Raw | ConvertFrom-Json
                    
                    # Update with selected packages
                    $packageControlSettings.installed_packages = @($selectedPackages)
                    
                    # Save updated settings
                    $packageControlSettings | ConvertTo-Json -Depth 10 | Out-File $packageControlDest -Force
                    
                    Write-ActivityLog "Restored Package Control settings with $($selectedPackages.Count) packages" -Type Success
                    
                    # Also restore Package Control metadata files
                    $packageControlFiles = @(
                        "Package Control.last-run",
                        "Package Control.ca-list",
                        "Package Control.ca-bundle",
                        "Package Control.system-ca-bundle",
                        "Package Control.cache",
                        "Package Control.merged-ca-bundle",
                        "Package Control.user-ca-bundle"
                    )
                    
                    foreach ($file in $packageControlFiles) {
                        $filePath = Join-Path $userSourceDir $file
                        if (Test-Path $filePath) {
                            Copy-Item -Path $filePath -Destination $sync.SublimeUserPath -Force
                        }
                    }
                    
                    # Create a notification that packages need to be installed
                    $installationNote = @"
You'll need to manually install the packages after restarting Sublime Text.

When you first start Sublime Text, Package Control will attempt to install the packages you've selected.
If it doesn't start automatically, open the Command Palette (Ctrl+Shift+P) and run:
"Package Control: Install Package"

Then, install the following packages:
$($selectedPackages -join "`n")
"@
                    
                    [System.Windows.MessageBox]::Show(
                        $installationNote,
                        "Package Installation Instructions",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                } else {
                    Write-ActivityLog "Package Control settings not found in backup" -Type Warning
                }
            } else {
                Write-ActivityLog "No packages selected for restoration" -Type Warning
            }
        }
        
        # Restore workspace if selected
        if ($selectedOptions -contains "Workspace") {
            $workspaceFiles = Get-ChildItem -Path $userSourceDir -Filter "*.sublime-workspace"
            
            if ($workspaceFiles.Count -gt 0) {
                foreach ($workspace in $workspaceFiles) {
                    Copy-Item -Path $workspace.FullName -Destination $sync.SublimeUserPath -Force
                }
                Write-ActivityLog "Restored $($workspaceFiles.Count) workspace files" -Type Success
            } else {
                Write-ActivityLog "No workspace files found in backup" -Type Warning
            }
        }
        
        # Restore project if selected
        if ($selectedOptions -contains "Project") {
            $projectFiles = Get-ChildItem -Path $userSourceDir -Filter "*.sublime-project"
            
            if ($projectFiles.Count -gt 0) {
                foreach ($project in $projectFiles) {
                    Copy-Item -Path $project.FullName -Destination $sync.SublimeUserPath -Force
                }
                Write-ActivityLog "Restored $($projectFiles.Count) project files" -Type Success
            } else {
                Write-ActivityLog "No project files found in backup" -Type Warning
            }
        }
        
        # Restore additional settings files that might be useful
        $additionalSettingsFiles = @(
            "Distraction Free.sublime-settings",
            "SublimeLinter.sublime-settings",
            "ColorHighlighter.sublime-settings",
            "Side Bar.sublime-settings",
            "Syntax Highlighting for Sass.sublime-settings"
        )
        
        foreach ($file in $additionalSettingsFiles) {
            $filePath = Join-Path $userSourceDir $file
            if (Test-Path $filePath) {
                Copy-Item -Path $filePath -Destination $sync.SublimeUserPath -Force
                Write-ActivityLog "Restored additional settings: $file" -Type Success
            }
        }
        
        # Clean up temp directory
        Remove-Item -Path $tempDir -Recurse -Force
        Write-ActivityLog "Temporary files cleaned up" -Type Information
        
        $sync.StatusBar.Text = "Restore completed successfully"
        
        [System.Windows.MessageBox]::Show(
            "Restore completed successfully!`n`nYou may need to restart Sublime Text for all changes to take effect.",
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
    $saveFileDialog.Title = "Save Sublime Text Settings Backup"
    $saveFileDialog.FileName = "SublimeBackup_$(Get-Date -Format 'yyyyMMdd').zip"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.StatusBar.Text = "Creating backup..."
        Write-ActivityLog "Creating backup to $($saveFileDialog.FileName)" -Type Information
        New-SublimeBackup -BackupPath $saveFileDialog.FileName
    }
})

$sync.BrowseRestoreButton.Add_Click({
    # Show open file dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Zip Files (*.zip)|*.zip"
    $openFileDialog.Title = "Select Sublime Text Settings Backup"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.RestoreFilePathTextBox.Text = $openFileDialog.FileName
        $sync.StatusBar.Text = "Loading backup file..."
        Write-ActivityLog "Loading backup file: $($openFileDialog.FileName)" -Type Information
        Initialize-RestoreOptions -BackupPath $openFileDialog.FileName
    }
})

$sync.RestoreButton.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "This will replace your current Sublime Text settings with the ones from the backup.`n`nDo you want to continue?",
        "Confirm Restore",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $sync.StatusBar.Text = "Restoring settings..."
        Write-ActivityLog "Restoring settings from $($sync.RestoreFilePathTextBox.Text)" -Type Information
        Restore-SublimeSettings -BackupPath $sync.RestoreFilePathTextBox.Text
    }
})

# Event handlers for checkboxes
$sync.BackupPackagesCheckBox.Add_Checked({
    $sync.PackagesListPanel.IsEnabled = $true
})

$sync.BackupPackagesCheckBox.Add_Unchecked({
    $sync.PackagesListPanel.IsEnabled = $false
})

$sync.RefreshBackupButton.Add_Click({
    Write-ActivityLog "Refreshing installed packages and settings..." -Type Information
    # Re-read the installed packages from Sublime Text
    $sync.InstalledPackages = Get-SublimePackages
    # Rebuild the packages list checkboxes
    Initialize-PackagesList
})

# Initialize the application
if (Test-SublimePaths) {
    Write-ActivityLog "Sublime Text paths found. Initializing..." -Type Information
    
    # Get and display installed packages
    $sync.InstalledPackages = Get-SublimePackages
    Initialize-PackagesList
    Write-ActivityLog "Found $($sync.InstalledPackages.Count) installed packages" -Type Information
} else {
    Write-ActivityLog "Sublime Text installation not found or paths incorrect" -Type Warning
    [System.Windows.MessageBox]::Show(
        "Sublime Text installation not found or paths are incorrect. Some features may not work properly.",
        "Initialization Warning",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    )
}

# Start with welcome message
Write-ActivityLog "Sublime Text Backup & Restore Tool started." -Type Information
$sync.StatusBar.Text = "Ready"

# Show the window
$sync.Window.ShowDialog() | Out-Null