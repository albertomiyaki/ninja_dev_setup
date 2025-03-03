# File Manager - Copy files based on JSON configuration
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Create a synchronized hashtable for sharing data
$sync = [hashtable]::Synchronized(@{})
$sync.CopyInProgress = $false
$sync.CurrentFile = ""

# Create log function for centralized logging
function Write-FileLog {
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
    
    # Write to log file if configured
    if ($sync.LogFile) {
        Add-Content -Path $sync.LogFile -Value $logMessage
    }
}

# Function to resolve file paths
function Resolve-FilePath {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [string]$BasePath = $moduleRoot
    )
    
    # Check if path is already absolute or UNC
    if ([System.IO.Path]::IsPathRooted($Path) -or $Path -match "^\\\\") {
        return $Path
    }
    
    # Otherwise, combine with base path
    return [System.IO.Path]::Combine($BasePath, $Path)
}

# Load file configurations from JSON
function Get-FileConfigurations {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ConfigFile
    )
    
    try {
        if (-not (Test-Path $ConfigFile)) {
            Write-FileLog "Error: Configuration file not found at $ConfigFile" -Type Error
            return @()
        }
        
        $configs = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        
        if (-not $configs) {
            Write-FileLog "No file configurations found in the JSON file" -Type Warning
            return @()
        }
        
        # Process each file config to ensure paths are resolved
        foreach ($config in $configs) {
            # Add status property to track copy operations
            $config | Add-Member -NotePropertyName Status -NotePropertyValue "Ready" -Force
            
            # Add property to store resolved source path
            $config | Add-Member -NotePropertyName ResolvedSource -NotePropertyValue (Resolve-FilePath -Path $config.Source) -Force
        }
        
        Write-FileLog "Successfully loaded $($configs.Count) file configurations" -Type Success
        return $configs
    }
    catch {
        Write-FileLog "Error loading configuration file: $_" -Type Error
        return @()
    }
}

# Copy a single file based on configuration
function Copy-ConfigFile {
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$FileConfig
    )
    
    try {
        $sync.CopyInProgress = $true
        $sync.CurrentFile = $FileConfig.Name
        $sync.StatusBar.Text = "Copying file: $($FileConfig.Name)..."
        
        # Update file status in the list
        $fileItem = $sync.FileListView.Items | Where-Object { $_.Name -eq $FileConfig.Name }
        if ($fileItem) {
            $fileItem.Status = "Copying"
            $sync.FileListView.Items.Refresh()
        }
        
        $sourcePath = $FileConfig.ResolvedSource
        $destinationPath = $FileConfig.Destination
        
        # Validate source file
        if (-not (Test-Path $sourcePath)) {
            Write-FileLog "Error: Source file '$sourcePath' not found" -Type Error
            
            if ($fileItem) {
                $fileItem.Status = "Failed - Source not found"
                $sync.FileListView.Items.Refresh()
            }
            
            return $false
        }
        
        # Create destination directory if it doesn't exist
        $destinationDir = Split-Path -Parent $destinationPath
        if (-not (Test-Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
            Write-FileLog "Created destination directory: $destinationDir" -Type Information
        }
        
        # Check if destination file exists and overwrite is disabled
        if ((Test-Path $destinationPath) -and -not $FileConfig.Overwrite) {
            Write-FileLog "File already exists at '$destinationPath' and overwrite is disabled" -Type Warning
            
            if ($fileItem) {
                $fileItem.Status = "Skipped - File exists"
                $sync.FileListView.Items.Refresh()
            }
            
            return $false
        }
        
        # Perform the file copy
        Write-FileLog "Copying '$($FileConfig.Name)' from '$sourcePath' to '$destinationPath'" -Type Information
        Copy-Item -Path $sourcePath -Destination $destinationPath -Force
        
        # Verify the file was copied successfully
        if (Test-Path $destinationPath) {
            Write-FileLog "File '$($FileConfig.Name)' copied successfully" -Type Success
            
            if ($fileItem) {
                $fileItem.Status = "Copied"
                $sync.FileListView.Items.Refresh()
            }
            
            return $true
        } else {
            Write-FileLog "Failed to copy file '$($FileConfig.Name)' to destination" -Type Error
            
            if ($fileItem) {
                $fileItem.Status = "Failed - Copy error"
                $sync.FileListView.Items.Refresh()
            }
            
            return $false
        }
    }
    catch {
        Write-FileLog "Error copying file '$($FileConfig.Name)': $_" -Type Error
        
        if ($fileItem) {
            $fileItem.Status = "Failed - Error"
            $sync.FileListView.Items.Refresh()
        }
        
        return $false
    }
    finally {
        $sync.CopyInProgress = $false
        $sync.StatusBar.Text = "Copy operation completed for: $($FileConfig.Name)"
    }
}

# Copy multiple files
function Copy-SelectedFiles {
    param (
        [array]$SelectedFiles
    )
    
    if ($SelectedFiles.Count -eq 0) {
        Write-FileLog "No files selected for copying" -Type Warning
        $sync.StatusBar.Text = "No files selected"
        return
    }
    
    $sync.ProgressBar.Maximum = $SelectedFiles.Count
    $sync.ProgressBar.Value = 0
    
    $copyCount = 0
    $skipCount = 0
    $failCount = 0
    
    foreach ($file in $SelectedFiles) {
        $result = Copy-ConfigFile -FileConfig $file
        
        if ($result) {
            $copyCount++
        }
        elseif ($file.Status -match "Skipped") {
            $skipCount++
        }
        else {
            $failCount++
        }
        
        $sync.ProgressBar.Value++
    }
    
    # Update status
    $sync.StatusBar.Text = "Copy completed: $copyCount copied, $skipCount skipped, $failCount failed"
    Write-FileLog "Copy operation completed for $($SelectedFiles.Count) files: $copyCount copied, $skipCount skipped, $failCount failed" -Type Information
}

# Get the module directory dynamically
$moduleRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultConfigFile = Join-Path $moduleRoot "file_manager.json"

# Create a log file
$sync.LogFile = Join-Path $moduleRoot "file_manager.log"
New-Item -ItemType File -Path $sync.LogFile -Force | Out-Null

# Define the XAML UI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="File Manager" 
    Height="900" 
    Width="1200"
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
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#2d2d30" Padding="15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock Text="File Manager" FontSize="22" Foreground="White" FontWeight="Bold"/>
                    <TextBlock Text="Copy files based on JSON configuration" Foreground="#999999" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="LoadConfigButton" Grid.Column="1" 
                        Style="{StaticResource DefaultButton}"
                        Content="Load Config" 
                        ToolTip="Load configuration file"
                        VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Configuration Path -->
        <Border Grid.Row="1" Background="#F5F5F5" Padding="15" BorderThickness="0,0,0,1" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="Configuration File: " VerticalAlignment="Center" FontWeight="Medium"/>
                <TextBox x:Name="ConfigPathTextBox" Grid.Column="1" Padding="8,5" Margin="5,0" Text=""/>
                <Button x:Name="BrowseConfigButton" Grid.Column="2" Content="Browse" Padding="8,5"/>
            </Grid>
        </Border>
        
        <!-- Search and Filter -->
        <Grid Grid.Row="2" Margin="15,15,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Search Box -->
            <Grid Grid.Row="0" Margin="0,0,0,15">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="🔍 Search: " VerticalAlignment="Center" Margin="0,0,10,0" FontSize="14"/>
                <TextBox x:Name="SearchBox" Grid.Column="1" Padding="8" FontSize="14" BorderThickness="1" BorderBrush="#CCCCCC">
                    <TextBox.Resources>
                        <Style TargetType="{x:Type Border}">
                            <Setter Property="CornerRadius" Value="3"/>
                        </Style>
                    </TextBox.Resources>
                </TextBox>
                <Button x:Name="SelectAllButton" Grid.Column="2" Content="Select All" Margin="10,0,0,0" Padding="8,5"/>
            </Grid>
            
            <!-- File List -->
            <ListView x:Name="FileListView" Grid.Row="1" 
                      BorderThickness="1" 
                      BorderBrush="#CCCCCC"
                      Padding="0"
                      FontSize="14"
                      SelectionMode="Multiple">
                <ListView.Resources>
                    <Style TargetType="{x:Type ListViewItem}">
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="BorderThickness" Value="0,0,0,1"/>
                        <Setter Property="BorderBrush" Value="#EEEEEE"/>
                        <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#E3F2FD"/>
                                <Setter Property="BorderBrush" Value="#BBDEFB"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#F5F5F5"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </ListView.Resources>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Name" Width="200">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Name}" TextWrapping="NoWrap"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Source" Width="250">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ResolvedSource}" TextWrapping="Wrap" MaxWidth="240"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Destination" Width="250">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Destination}" TextWrapping="Wrap" MaxWidth="240"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Overwrite" Width="80">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox IsChecked="{Binding Overwrite}" IsHitTestVisible="False"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Status" Width="120">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding Status}" TextWrapping="NoWrap">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Status}" Value="Copied">
                                                        <Setter Property="Foreground" Value="Green"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Copying">
                                                        <Setter Property="Foreground" Value="Blue"/>
                                                        <Setter Property="FontWeight" Value="Bold"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed - Source not found">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed - Copy error">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed - Error">
                                                        <Setter Property="Foreground" Value="Red"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Skipped - File exists">
                                                        <Setter Property="Foreground" Value="Orange"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                    </GridView>
                </ListView.View>
            </ListView>
        </Grid>
        
        <!-- Progress Bar -->
        <Grid Grid.Row="3" Margin="15,15,15,0">
            <ProgressBar x:Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Value="0" />
        </Grid>
        
        <!-- Clone Status -->
        <Border Grid.Row="4" Background="#F5F5F5" Padding="15" BorderThickness="0,1,0,1" BorderBrush="#DDDDDD" Margin="0,15,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="SelectedFileCount" Text="No files selected" FontWeight="Bold" FontSize="14"/>
                    <TextBlock x:Name="FileStatusTextBlock" Text="Ready to copy files" Foreground="#666666" FontSize="12" Margin="0,5,0,0"/>
                </StackPanel>
                
                <Button x:Name="CopyButton" 
                        Grid.Column="1" 
                        Content="Copy Selected Files" 
                        Style="{StaticResource DefaultButton}"
                        FontSize="14"/>
            </Grid>
        </Border>
        
        <!-- Log Panel -->
        <Grid Grid.Row="5" Margin="15,15,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="150"/>
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
        <Border Grid.Row="6" Background="#E0E0E0" Padding="10,8" Margin="0,15,0,0">
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
$sync.LoadConfigButton = $window.FindName("LoadConfigButton")
$sync.BrowseConfigButton = $window.FindName("BrowseConfigButton")
$sync.ConfigPathTextBox = $window.FindName("ConfigPathTextBox")
$sync.FileListView = $window.FindName("FileListView")
$sync.SearchBox = $window.FindName("SearchBox")
$sync.SelectAllButton = $window.FindName("SelectAllButton")
$sync.CopyButton = $window.FindName("CopyButton")
$sync.ProgressBar = $window.FindName("ProgressBar")
$sync.StatusBar = $window.FindName("StatusBar")
$sync.LogTextBox = $window.FindName("LogTextBox")
$sync.SelectedFileCount = $window.FindName("SelectedFileCount")
$sync.FileStatusTextBlock = $window.FindName("FileStatusTextBlock")

# Set default config path
$sync.ConfigPathTextBox.Text = $defaultConfigFile

# Function to update the selected file count
function Update-SelectedFileCount {
    $selectedCount = $sync.FileListView.SelectedItems.Count
    
    if ($selectedCount -eq 0) {
        $sync.SelectedFileCount.Text = "No files selected"
        $sync.CopyButton.IsEnabled = $false
    } else {
        $sync.SelectedFileCount.Text = "$selectedCount files selected"
        $sync.CopyButton.IsEnabled = $true
    }
}

# Function to filter files based on search text
function Update-FileSearch {
    $searchText = $sync.SearchBox.Text.ToLower()
    $allFiles = $sync.AllFiles
    
    if (-not $allFiles) {
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        $filteredFiles = $allFiles
    } else {
        $filteredFiles = $allFiles | Where-Object { 
            $_.Name.ToLower().Contains($searchText) -or 
            $_.Source.ToLower().Contains($searchText) -or
            $_.Destination.ToLower().Contains($searchText)
        }
    }
    
    # Update the ListView
    $sync.FileListView.Items.Clear()
    
    foreach ($file in $filteredFiles) {
        $sync.FileListView.Items.Add($file)
    }
    
    $sync.StatusBar.Text = "Showing $($filteredFiles.Count) of $($allFiles.Count) files"
}

# Function to load configuration file
function Load-ConfigurationFile {
    param (
        [string]$ConfigPath
    )
    
    $sync.FileListView.Items.Clear()
    $files = Get-FileConfigurations -ConfigFile $ConfigPath
    
    if ($files.Count -eq 0) {
        $sync.StatusBar.Text = "No files found in configuration file"
        return
    }
    
    # Store all files
    $sync.AllFiles = $files
    
    # Update file search to populate list
    Update-FileSearch
    
    $sync.StatusBar.Text = "Loaded $($files.Count) files from configuration file"
}

# Set up event handlers
$sync.LoadConfigButton.Add_Click({
    $configPath = $sync.ConfigPathTextBox.Text
    
    if ([string]::IsNullOrWhiteSpace($configPath)) {
        [System.Windows.MessageBox]::Show("Please enter a configuration file path.", "Missing Path", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if (-not (Test-Path $configPath)) {
        [System.Windows.MessageBox]::Show("Configuration file not found: $configPath", "File Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Load-ConfigurationFile -ConfigPath $configPath
})

$sync.BrowseConfigButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $fileDialog.Title = "Select File Manager Configuration File"
    $fileDialog.InitialDirectory = $moduleRoot
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sync.ConfigPathTextBox.Text = $fileDialog.FileName
        Load-ConfigurationFile -ConfigPath $fileDialog.FileName
    }
})

$sync.SearchBox.Add_TextChanged({
    Update-FileSearch
})

$sync.SelectAllButton.Add_Click({
    if ($sync.FileListView.Items.Count -gt 0) {
        $sync.FileListView.SelectAll()
        Update-SelectedFileCount
    }
})

$sync.CopyButton.Add_Click({
    $selectedFiles = $sync.FileListView.SelectedItems
    
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one file to copy.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Confirm
    $confirmResult = [System.Windows.MessageBox]::Show(
        "Are you sure you want to copy $($selectedFiles.Count) files to their destinations?",
        "Confirm Copy",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($confirmResult -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }
    
    # Start copying
    Copy-SelectedFiles -SelectedFiles $selectedFiles
})

$sync.FileListView.Add_SelectionChanged({
    Update-SelectedFileCount
})

# Initial UI setup
$sync.CopyButton.IsEnabled = $false
$sync.ProgressBar.Value = 0
Write-FileLog "Application started" -Type Information
$sync.StatusBar.Text = "Ready"

# Try to load default config file if it exists
if (Test-Path $defaultConfigFile) {
    Write-FileLog "Loading default configuration file: $defaultConfigFile" -Type Information
    Load-ConfigurationFile -ConfigPath $defaultConfigFile
} else {
    Write-FileLog "Default configuration file not found: $defaultConfigFile" -Type Warning
}

# Show the window
$sync.Window.ShowDialog() | Out-Null